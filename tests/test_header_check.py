"""Header check for CSV and TSV import: the first row is compared with the table's columns.

100% match → header, 0% → no header, anything in between → warning and the user decides.
"""

import io
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch

import app as importer
from test_csv_sql_export import FakeConnection

COLUMNS = ["CODE", "AMOUNT", "DESCRIPTION", "ACTIVE"]
METADATA = [("CODE", "varchar", "YES", None), ("AMOUNT", "varchar", "YES", None),
            ("DESCRIPTION", "varchar", "YES", None), ("ACTIVE", "varchar", "YES", None)]
CONFIG = dict(host="example", port=1433, database="test", username="user", password="secret")


class CheckHeaderTests(unittest.TestCase):
    def test_every_value_is_a_column_name_is_a_header(self):
        check = importer.check_header(["CODE", "AMOUNT", "DESCRIPTION", "ACTIVE"], COLUMNS)
        self.assertEqual((check["status"], check["header"], check["percent"]), ("header", True, 100))
        self.assertEqual(check["unknown"], [])

    def test_header_in_any_order_full_width_lower_case_and_padding(self):
        check = importer.check_header(["ａｍｏｕｎｔ", " code ", "", ""], COLUMNS)
        self.assertEqual((check["status"], check["matched"], check["cells"]), ("header", 2, 2))
        self.assertEqual(check["missing"], ["DESCRIPTION", "ACTIVE"])

    def test_no_value_is_a_column_name_is_no_header(self):
        check = importer.check_header(["0001", "12.50", "text", "1"], COLUMNS)
        self.assertEqual((check["status"], check["header"], check["percent"]), ("none", False, 0))

    def test_header_with_a_typo_is_a_warning_and_defaults_to_header(self):
        check = importer.check_header(["CODE", "AMOUNT", "DESCRIPTON", "ACTIVE"], COLUMNS)
        self.assertEqual((check["status"], check["header"], check["percent"]), ("partial", True, 75))
        self.assertEqual(check["unknown"], ["DESCRIPTON"])

    def test_one_matching_value_is_a_warning_and_defaults_to_data(self):
        check = importer.check_header(["0001", "CODE", "text", "1"], COLUMNS)
        self.assertEqual((check["status"], check["header"], check["percent"]), ("partial", False, 25))

    def test_a_tiny_partial_match_still_shows_at_least_one_percent(self):
        check = importer.check_header(["CODE"] + [str(n) for n in range(199)], COLUMNS)
        self.assertEqual((check["status"], check["percent"]), ("partial", 1))

    def test_duplicates_are_reported(self):
        check = importer.check_header(["CODE", "code", "AMOUNT"], COLUMNS)
        self.assertEqual(check["duplicates"], ["code"])

    def test_user_choice_overrides_the_automatic_one(self):
        check = importer.check_header(["0001", "CODE", "text", "1"], COLUMNS)
        self.assertFalse(importer.resolve_header(check, "auto"))
        self.assertTrue(importer.resolve_header(check, "yes"))
        self.assertFalse(importer.resolve_header(check, "no"))
        with self.assertRaises(importer.RouteError):
            importer.header_mode_arg("maybe")


class HeaderModeImportTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        with importer.IMPORT_JOBS_LOCK:
            importer.IMPORT_JOBS.clear()
        self.addCleanup(self.clear_jobs)

    def clear_jobs(self):
        with importer.IMPORT_JOBS_LOCK:
            for job in importer.IMPORT_JOBS.values():
                if job.get("sql_path"):
                    Path(job["sql_path"]).unlink(missing_ok=True)
            importer.IMPORT_JOBS.clear()

    def run_csv(self, raw, header_mode):
        path = Path(self.directory.name) / "MDNNOP.csv"
        path.write_bytes(raw)
        connection = FakeConnection(metadata=METADATA)
        job = importer._new_import_job("header-mode", ["MDNNOP"])
        importer.IMPORT_JOBS["header-mode"] = job
        importer._run_csv_import("header-mode", connection, str(path), "MDNNOP", ["MDNNOP"], "lvapdbf",
                                 ",", header_mode)
        return connection, job["results"][0]

    def test_typo_header_used_as_header_by_default_with_a_warning(self):
        connection, result = self.run_csv(b"CODE,AMOUNT,DESCRIPTON,ACTIVE\n0001,1,x,1\n", "auto")
        self.assertEqual(connection.committed, [("0001", "1", "1")])   # DESCRIPTON skipped
        self.assertIn("3 of 4 values are column names (75%)", result["message"])
        self.assertIn("used as header", result["message"])

    def test_user_can_import_a_partial_match_row_as_data(self):
        connection, result = self.run_csv(b"CODE,AMOUNT,DESCRIPTON,ACTIVE\n0001,1,x,1\n", "no")
        self.assertEqual(connection.committed,
                         [("CODE", "AMOUNT", "DESCRIPTON", "ACTIVE"), ("0001", "1", "x", "1")])
        self.assertIn("imported as data", result["message"])

    def test_user_can_use_a_low_match_row_as_header(self):
        connection, result = self.run_csv(b"CODE,X1,X2,X3\n0001,a,b,c\n", "yes")
        self.assertEqual(connection.committed, [("0001",)])
        self.assertIn("used as header", result["message"])

    def test_full_match_and_no_match_show_no_warning(self):
        _, header = self.run_csv(b"CODE,AMOUNT\n0001,1\n", "auto")
        self.assertNotIn("⚠", header["message"])
        _, data = self.run_csv(b"0001,1\n", "auto")
        self.assertNotIn("⚠", data["message"])


class HeaderCheckRouteTests(unittest.TestCase):
    def patched(self, connection):
        return (patch.object(importer, "load_config", return_value=CONFIG),
                patch.object(importer, "_build_conn_str", return_value=("mock", None)),
                patch.object(importer.pyodbc, "connect", return_value=connection))

    def test_csv_header_check_uses_the_chosen_schema_table(self):
        client = importer.app.test_client()
        connection = FakeConnection(metadata=METADATA)
        p1, p2, p3 = self.patched(connection)
        with p1, p2, p3:
            upload = client.post("/upload", data={"file": (io.BytesIO(b"CODE,AMOUNT,OOPS\n1,2,3\n"), "MDNNOP.csv")},
                                 content_type="multipart/form-data")
            self.assertEqual(upload.status_code, 200)
            with client.session_transaction() as browser_session:
                self.addCleanup(Path(browser_session["tmp_path"]).unlink, missing_ok=True)
            data = client.post("/csv/header-check", json={"schema": "lvapdbf"}).get_json()
            missing_schema = client.post("/csv/header-check", json={"schema": ""})
        self.assertEqual((data["schema"], data["table"], data["table_columns"], data["file_values"]),
                         ("lvapdbf", "MDNNOP", 4, 3))
        self.assertEqual((data["check"]["status"], data["check"]["percent"], data["check"]["unknown"]),
                         ("partial", 66, ["OOPS"]))
        self.assertTrue(any("TABLE_SCHEMA = ?" in sql for sql in connection.statements))
        self.assertEqual(missing_schema.status_code, 400)

    def test_csv_import_rejects_an_unknown_header_mode(self):
        client = importer.app.test_client()
        p1, p2, p3 = self.patched(FakeConnection(metadata=METADATA))
        with p1, p2, p3:
            client.post("/upload", data={"file": (io.BytesIO(b"1,2\n"), "MDNNOP.csv")},
                        content_type="multipart/form-data")
            with client.session_transaction() as browser_session:
                self.addCleanup(Path(browser_session["tmp_path"]).unlink, missing_ok=True)
            response = client.post("/import", json={"tables": ["MDNNOP"], "header": "maybe"})
        self.assertEqual(response.status_code, 400)

    def test_tsv_header_check_and_import_follow_the_users_choice(self):
        client = importer.app.test_client()
        raw = b"CODE\tAMOUNT\tDESCRIPTON\tACTIVE\n0001\t1\tx\t1\n"
        connection = FakeConnection(metadata=METADATA)
        p1, p2, p3 = self.patched(connection)
        with p1, p2, p3:
            client.post("/upload", data={"file": (io.BytesIO(raw), "MDNNOP.tsv")},
                        content_type="multipart/form-data")
            with client.session_transaction() as browser_session:
                self.addCleanup(Path(browser_session["tmp_path"]).unlink, missing_ok=True)
            auto = client.post("/csv/header-check", json={"schema": "lvapdbf"}).get_json()
            as_data = client.post("/csv/header-check", json={"schema": "lvapdbf", "header": "no"}).get_json()
            job = client.post("/import", json={"tables": ["MDNNOP"], "schema": "lvapdbf", "header": "no"}).get_json()
            deadline = time.monotonic() + 3
            while time.monotonic() < deadline:
                status = client.get(f"/import/status/{job['job_id']}").get_json()
                if status["state"] in {"done", "failed"}:
                    break
                time.sleep(0.01)
        self.addCleanup(lambda: Path(importer.IMPORT_JOBS.get(job["job_id"], {}).get("sql_path", "/nonexistent"))
                        .unlink(missing_ok=True))
        self.assertEqual((auto["check"]["status"], auto["header"], auto["records"]), ("partial", True, 1))
        self.assertEqual((as_data["header"], as_data["records"], as_data["header_mode"]), (False, 2, "no"))
        self.assertEqual(status["state"], "done")
        self.assertEqual(len(connection.committed), 2)   # the first row was imported as data


if __name__ == "__main__":
    unittest.main()
