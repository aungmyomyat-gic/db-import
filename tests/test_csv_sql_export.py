"""CSV/SQL parity and download lifecycle tests; no live database is modified."""

import csv
import io
import tempfile
import time
import unittest
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch

import pyodbc

import app as importer


METADATA = [
    ("CODE", "varchar", "NO", None),
    ("AMOUNT", "decimal", "YES", None),
    ("DESCRIPTION", "nvarchar", "YES", None),
    ("ACTIVE", "bit", "NO", None),
    ("CREATED", "datetime", "NO", None),
    ("REQUIRED", "varchar", "NO", None),
    ("SERVER_DEFAULT", "int", "NO", "((7))"),
]


class FakeCursor:
    def __init__(self, connection):
        self.connection = connection
        self.fast_executemany = False
        self.result = []
        self.savepoint = 0

    def execute(self, sql, *parameters):
        connection = self.connection
        connection.statements.append(sql)
        if sql.startswith("SELECT TABLE_SCHEMA"):
            self.result = [(connection.schema,)] if connection.table_exists else []
        elif sql.startswith("SELECT COLUMN_NAME"):
            self.result = connection.metadata
        elif sql.startswith("TRUNCATE"):
            if connection.truncate_blocked:
                raise pyodbc.Error("TRUNCATE blocked by foreign key")
            connection.pending = []
        elif sql.startswith("DELETE FROM"):
            connection.pending = []
        elif sql.startswith("SAVE TRANSACTION"):
            self.savepoint = len(connection.pending)
        elif sql.startswith("ROLLBACK TRANSACTION"):
            connection.pending = connection.pending[:self.savepoint]
        elif sql.startswith("INSERT"):
            values = tuple(parameters[0])
            if values[0] == connection.rejected_code:
                raise pyodbc.Error("Rejected row")
            connection.pending.append(values)
        return self

    def executemany(self, sql, rows):
        connection = self.connection
        connection.statements.append(sql)
        if connection.fallback or any(row[0] == connection.rejected_code for row in rows):
            # A real driver can insert some rows before raising. The savepoint
            # must undo these before the importer's row-by-row fallback.
            if rows:
                connection.pending.append(tuple(rows[0]))
            raise pyodbc.Error("Batch failed")
        connection.pending.extend(tuple(row) for row in rows)

    def fetchall(self):
        return self.result


class FakeConnection:
    def __init__(self, schema="lvapdbf", metadata=None, truncate_blocked=False,
                 rejected_code=None, fallback=False, commit_error=False, table_exists=True):
        self.schema = schema
        self.metadata = metadata or METADATA
        self.truncate_blocked = truncate_blocked
        self.rejected_code = rejected_code
        self.fallback = fallback
        self.commit_error = commit_error
        self.table_exists = table_exists
        self.pending = []
        self.committed = None
        self.rollbacks = 0
        self.statements = []
        self.closed = False
        self._cursor = FakeCursor(self)

    def cursor(self):
        return self._cursor

    def commit(self):
        if self.commit_error:
            raise pyodbc.Error("Commit failed")
        self.committed = list(self.pending)

    def rollback(self):
        self.rollbacks += 1
        self.pending = []

    def close(self):
        self.closed = True


class CsvSqlExportTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.csv_path = Path(self.directory.name) / "MDNNOP.csv"
        self.job_id = "test-csv-sql-export"
        with importer.IMPORT_JOBS_LOCK:
            importer.IMPORT_JOBS.clear()
        self.addCleanup(self.clear_jobs)

    def clear_jobs(self):
        with importer.IMPORT_JOBS_LOCK:
            for job in importer.IMPORT_JOBS.values():
                if job.get("sql_path"):
                    Path(job["sql_path"]).unlink(missing_ok=True)
            importer.IMPORT_JOBS.clear()
        (Path(importer.SQL_EXPORT_DIR.name) / f"{self.job_id}.sql").unlink(missing_ok=True)

    def write_csv(self, rows, encoding="utf-8-sig"):
        with self.csv_path.open("w", encoding=encoding, newline="") as stream:
            csv.writer(stream).writerows(rows)
        return self.csv_path.read_bytes()

    def run_csv(self, connection, rows=None, schema="lvapdbf", encoding="utf-8-sig"):
        self.write_csv(rows or [["0001", "12.50", "O'Brien 日本語", "", ""]], encoding)
        job = importer._new_import_job(self.job_id, ["MDNNOP"])
        importer.IMPORT_JOBS[self.job_id] = job
        importer._set_job_state(self.job_id, "running")
        importer._run_csv_import(self.job_id, connection, str(self.csv_path), "MDNNOP", ["MDNNOP"], schema)
        # Publishing the path alone must not make a running job downloadable.
        self.assertNotIn("sql_download_url", importer._job_snapshot(job))
        importer._finish_import_job(self.job_id)
        return job

    def test_result_reports_record_count_and_elapsed_ms(self):
        rows = [["0001", "12.50", "a", "", ""], ["0002", "", "b", "1", "2026-11-22"]]
        job = self.run_csv(FakeConnection(), rows)
        result = importer._job_snapshot(job)["results"][0]
        self.assertEqual(result["status"], "ok")
        self.assertEqual(result["records"], 2)
        self.assertIsInstance(result["elapsed_ms"], int)
        self.assertGreaterEqual(result["elapsed_ms"], 0)

    def test_success_script_matches_odbc_values_and_defaults_across_batches(self):
        connection = FakeConnection()
        rows = [["0001", "12.50", "O'Brien 日本語", "", ""],
                ["0002", "", "line\r\nGO\r\ntext", "1", "2026-11-22"]]
        with patch.object(importer, "IMPORT_BATCH_SIZE", 1):
            job = self.run_csv(connection, rows)
        self.assertEqual(connection.committed, [
            ("0001", "12.50", "O'Brien 日本語", 0, "1900-01-01", ""),
            ("0002", None, "line\r\nGO\r\ntext", "1", "2026-11-22", ""),
        ])
        path = Path(job["sql_path"])
        self.assertTrue(path.read_bytes().startswith(b"\xef\xbb\xbf"))
        sql = path.read_text(encoding="utf-8-sig")
        self.assertTrue(sql.startswith("DELETE FROM lvapdbf.MDNNOP;\nINSERT INTO lvapdbf.MDNNOP "))
        self.assertIn("(CODE, AMOUNT, DESCRIPTION, ACTIVE, CREATED, REQUIRED)\nVALUES\n", sql)
        self.assertIn("(N'0001', 12.50, N'O''Brien 日本語', 0, N'1900-01-01', N''),\n", sql)
        self.assertIn("(N'0002', NULL,", sql)
        self.assertIn("NCHAR(13)", sql)
        self.assertIn("NCHAR(10)", sql)
        self.assertNotIn("\nGO\n", sql)
        self.assertNotIn("SERVER_DEFAULT", sql)
        self.assertEqual(sql.count("INSERT INTO"), 1)
        self.assertTrue(sql.endswith(");\n"))
        self.assertNotIn("BEGIN TRY", sql)
        self.assertNotIn("TRUNCATE", sql)
        self.assertIn("sql_download_url", importer._job_snapshot(job))
        self.assertEqual(job["results"][0]["status"], "ok")

    def test_shift_jis_and_successful_row_fallback(self):
        connection = FakeConnection(fallback=True)
        job = self.run_csv(connection, [["0001", "0", "日本語", "1", "2026-11-22"]], encoding="cp932")
        self.assertEqual(len(connection.committed), 1)
        self.assertIn("N'日本語'", Path(job["sql_path"]).read_text(encoding="utf-8-sig"))

    def test_delete_fallback_and_bracketed_identifiers_match_export(self):
        metadata = [("CO]DE", "varchar", "NO", None), ("OMITTED", "int", "YES", None)]
        connection = FakeConnection(schema="schema]name", metadata=metadata, truncate_blocked=True)
        job = self.run_csv(connection, [["0001"]], schema="schema]name")
        sql = Path(job["sql_path"]).read_text(encoding="utf-8-sig")
        self.assertIn("DELETE FROM [schema]]name].MDNNOP;", sql)
        self.assertNotIn("TRUNCATE TABLE", sql)
        self.assertIn("([CO]]DE])\nVALUES\n(N'0001');", sql)
        self.assertEqual(connection.rollbacks, 1)

    def test_full_width_csv_matches_requested_script_format(self):
        metadata = [("CODE", "varchar", "NO", None), ("QTY", "int", "NO", None),
                    ("PRICE", "decimal", "NO", None), ("DESCRIPTION", "nvarchar", "NO", None)]
        connection = FakeConnection(metadata=metadata)
        job = self.run_csv(connection, [["00001", "0", "1.25", ""],
                                        ["00002", "-2", "3.50", "日本語 O'Brien"]])
        sql = Path(job["sql_path"]).read_text(encoding="utf-8-sig")
        self.assertEqual(sql, (
            "DELETE FROM lvapdbf.MDNNOP;\n"
            "INSERT INTO lvapdbf.MDNNOP\n"
            "VALUES\n"
            "(N'00001', 0, 1.25, N''),\n"
            "(N'00002', -2, 3.50, N'日本語 O''Brien');\n"
        ))

    def test_large_exports_split_values_into_valid_batches(self):
        columns = {"CODE": {"name": "CODE", "type": "varchar", "nullable": False}}
        plan = [("CODE", ("sheet", 0))]
        path = Path(self.directory.name) / "large.sql"
        for count, expected_sizes in ((1, [1]), (1000, [1000]), (1001, [1000, 1]),
                                      (2000, [1000, 1000]), (2001, [1000, 1000, 1])):
            with self.subTest(rows=count):
                rows = [[str(index).zfill(5)] for index in range(count)]
                importer._write_import_sql(path, "lvapdbf", "MSHBRP", plan, rows, columns)
                sql = path.read_text(encoding="utf-8-sig")
                self.assertEqual(sql.count("DELETE FROM lvapdbf.MSHBRP;"), 1)
                blocks = sql.split("INSERT INTO lvapdbf.MSHBRP\nVALUES\n")[1:]
                self.assertEqual([len(block.strip().splitlines()) for block in blocks], expected_sizes)
                self.assertTrue(all(block.rstrip().endswith(");") for block in blocks))
                self.assertEqual(sum(block.count("(N'") for block in blocks), count)

    def test_reserved_schema_and_table_names_are_delimited(self):
        path = Path(self.directory.name) / "reserved.sql"
        columns = {"CODE": {"name": "CODE", "type": "varchar", "nullable": False}}
        importer._write_import_sql(path, "order", "SELECT", [("CODE", ("sheet", 0))], [["0001"]], columns)
        self.assertEqual(path.read_text(encoding="utf-8-sig"), (
            "DELETE FROM [order].[SELECT];\nINSERT INTO [order].[SELECT]\nVALUES\n(N'0001');\n"
        ))

    def test_partial_import_keeps_existing_warning_behavior_without_download(self):
        connection = FakeConnection(rejected_code="BAD")
        job = self.run_csv(connection, [["GOOD"], ["BAD"]])
        self.assertEqual(len(connection.committed), 1)
        self.assertEqual(job["results"][0]["status"], "warn")
        self.assertNotIn("sql_download_url", importer._job_snapshot(job))
        self.assertFalse((Path(importer.SQL_EXPORT_DIR.name) / f"{self.job_id}.sql").exists())

    def test_commit_failure_rolls_back_and_discards_script(self):
        connection = FakeConnection(commit_error=True)
        job = self.run_csv(connection)
        self.assertIsNone(connection.committed)
        self.assertEqual(connection.rollbacks, 1)
        self.assertEqual(job["results"][0]["status"], "failed")
        self.assertNotIn("sql_download_url", importer._job_snapshot(job))
        self.assertFalse((Path(importer.SQL_EXPORT_DIR.name) / f"{self.job_id}.sql").exists())

    def test_export_write_failure_prevents_commit(self):
        connection = FakeConnection()
        with patch.object(importer, "_write_import_sql", side_effect=OSError("Disk full")):
            job = self.run_csv(connection)
        self.assertIsNone(connection.committed)
        self.assertEqual(connection.rollbacks, 1)
        self.assertEqual(job["results"][0]["status"], "failed")
        self.assertNotIn("sql_download_url", importer._job_snapshot(job))

    def test_missing_table_is_skipped_without_sql(self):
        connection = FakeConnection(table_exists=False)
        job = self.run_csv(connection)
        self.assertEqual(job["results"][0]["status"], "skip")
        self.assertIsNone(connection.committed)
        self.assertNotIn("sql_download_url", importer._job_snapshot(job))

    def test_excel_import_still_runs_without_an_export(self):
        connection = FakeConnection()
        result = importer._import_one(connection, "MDNNOP", ["CODE", "DESCRIPTION"],
                                      [["0001", "Excel 日本語"]])
        self.assertEqual(result["status"], "ok")
        self.assertEqual(connection.committed[0][0], "0001")
        self.assertEqual(connection.committed[0][1], "Excel 日本語")

    def test_header_row_maps_columns_by_name_in_any_order(self):
        connection = FakeConnection()
        job = self.run_csv(connection, [["description", "ＣＯＤＥ", "EXTRA", "AMOUNT"],
                                        ["日本語", "0001", "ignored", "12.50"],
                                        ["", "0002", "", ""]])
        self.assertEqual(connection.committed, [
            ("0001", "12.50", "日本語", 0, "1900-01-01", ""),
            ("0002", None, None, 0, "1900-01-01", ""),
        ])
        message = job["results"][0]["message"]
        self.assertIn("2 rows inserted", message)
        self.assertIn("header row found — columns matched by name", message)
        self.assertIn("1 CSV cols skipped — not in table", message)
        sql = Path(job["sql_path"]).read_text(encoding="utf-8-sig")
        self.assertIn("(CODE, AMOUNT, DESCRIPTION, ACTIVE, CREATED, REQUIRED)\nVALUES\n", sql)
        self.assertNotIn("description", sql)

    def test_header_row_with_every_column_matches_headerless_import(self):
        metadata = [("CODE", "varchar", "NO", None), ("QTY", "int", "NO", None)]
        with_header = FakeConnection(metadata=metadata)
        self.run_csv(with_header, [["CODE", "QTY"], ["00001", "3"]])
        self.clear_jobs()
        without_header = FakeConnection(metadata=metadata)
        self.run_csv(without_header, [["00001", "3"]])
        self.assertEqual(with_header.committed, without_header.committed)
        self.assertEqual(with_header.committed, [("00001", "3")])

    def test_data_row_containing_a_column_name_is_not_a_header(self):
        connection = FakeConnection(metadata=[("CODE", "varchar", "NO", None),
                                              ("NOTE", "varchar", "YES", None),
                                              ("QTY", "varchar", "YES", None)])
        job = self.run_csv(connection, [["CODE", "abc", "1"], ["0002", "def", "2"]])
        self.assertEqual(connection.committed, [("CODE", "abc", "1"), ("0002", "def", "2")])
        self.assertNotIn("header row", job["results"][0]["message"])

    def test_csv_header_detection_and_preview_counts(self):
        names = ["CODE", "AMOUNT", "DESCRIPTION"]
        self.assertTrue(importer.csv_has_header(["code", " AMOUNT ", ""], names))
        self.assertTrue(importer.csv_has_header(["CODE", "AMOUNT", "OTHER"], names))
        self.assertFalse(importer.csv_has_header(["CODE", "1", "x"], names))
        self.assertFalse(importer.csv_has_header(["", ""], names))
        rows = [["CODE", "AMOUNT"], ["1", "2"], ["3", "4"]]
        preview = importer._csv_preview("MDNNOP", rows, 2, True)
        self.assertEqual((preview["records"], preview["row_range"], preview["col_range"]),
                         (2, "2:3", "Header row"))
        preview = importer._csv_preview("MDNNOP", rows, 2, False)
        self.assertEqual((preview["records"], preview["row_range"]), (3, "1:3"))

    def test_sql_literals_keep_quotes_unicode_numeric_precision_and_binary(self):
        self.assertEqual(importer._sql_literal("O'Brien 日本語"), "N'O''Brien 日本語'")
        self.assertEqual(importer._sql_literal(Decimal("12345678901234567890.1234")), "12345678901234567890.1234")
        self.assertEqual(importer._sql_literal(Decimal("1E+30")), "1000000000000000000000000000000")
        self.assertEqual(importer._sql_literal(None), "NULL")
        self.assertEqual(importer._sql_literal(b"\x00\xff"), "0x00FF")
        self.assertEqual(importer._sql_literal(""), "N''")
        self.assertEqual(importer._sql_literal(False), "0")
        self.assertIn("nvarchar(max)", importer._sql_literal(("a" * 5000) + "\nGO\n"))

    def test_numeric_strings_use_destination_types_without_changing_text_codes(self):
        self.assertEqual(importer._sql_literal("00001", "varchar"), "N'00001'")
        self.assertEqual(importer._sql_literal("00001", "int"), "1")
        self.assertEqual(importer._sql_literal("0012.50", "decimal"), "12.50")
        self.assertEqual(importer._sql_literal("+1.25e2", "float"), "1.25E+2")
        self.assertEqual(importer._sql_literal("1e200", "float"), "1E+200")
        self.assertEqual(importer._sql_literal("TRUE", "bit"), "1")
        self.assertEqual(importer._sql_literal("false", "bit"), "0")
        self.assertEqual(importer._sql_literal("20171021", "int"), "20171021")
        self.assertEqual(importer._sql_literal("1e9999", "decimal"), "N'1e9999'")
        self.assertEqual(importer._sql_literal("1e999999999999999999999999", "float"),
                         "N'1e999999999999999999999999'")
        self.assertEqual(importer._sql_literal("0); DROP TABLE DATA;--", "int"),
                         "N'0); DROP TABLE DATA;--'")

    def test_http_upload_import_download_and_session_isolation(self):
        raw = self.write_csv([["0001", "12.50", "日本語 O'Brien", "", ""]])
        connection = FakeConnection()
        client = importer.app.test_client()
        config = dict(host="example", port=1433, database="test", username="user", password="secret")
        with patch.object(importer, "load_config", return_value=config), \
             patch.object(importer, "_build_conn_str", return_value=("mock connection", None)), \
             patch.object(importer.pyodbc, "connect", return_value=connection):
            upload = client.post("/upload", data={"file": (io.BytesIO(raw), "MDNNOP.csv")})
            self.assertEqual(upload.status_code, 200)
            with client.session_transaction() as browser_session:
                uploaded_path = Path(browser_session["tmp_path"])
            self.addCleanup(uploaded_path.unlink, missing_ok=True)
            self.assertEqual(client.post("/scan", json={"sheet": "MDNNOP"}).status_code, 200)
            response = client.post("/import", json={"tables": ["MDNNOP"], "schema": "lvapdbf"})
            self.assertEqual(response.status_code, 202)
            job_id = response.json["job_id"]
            deadline = time.monotonic() + 3
            while time.monotonic() < deadline:
                snapshot = client.get(f"/import/status/{job_id}").json
                if snapshot["state"] in {"done", "failed"}:
                    break
                time.sleep(0.01)
            self.assertEqual(snapshot["state"], "done")
            self.assertIn("sql_download_url", snapshot)
            self.assertNotIn("sql_path", snapshot)
            self.assertNotIn("owner", snapshot)
            download_url = snapshot["sql_download_url"]
            download = client.get(download_url)
            self.assertEqual(download.status_code, 200)
            self.assertIn("attachment;", download.headers["Content-Disposition"])
            self.assertIn(".sql", download.headers["Content-Disposition"])
            self.assertIn("no-store", download.headers["Cache-Control"])
            sql = download.data.decode("utf-8-sig")
            download.close()
            self.assertIn("N'日本語 O''Brien'", sql)
            self.assertNotIn("secret", sql)
            stranger = importer.app.test_client()
            self.assertEqual(stranger.get(download_url).status_code, 404)
            self.assertEqual(stranger.get(f"/import/status/{job_id}").status_code, 404)
            job = importer.IMPORT_JOBS[job_id]
            job["state"] = "running"
            self.assertEqual(client.get(download_url).status_code, 409)
            self.assertNotIn("sql_download_url", client.get(f"/import/status/{job_id}").json)
            job["state"] = "done"
            export_path = Path(job["sql_path"])
            job["updated_at"] = 0
            self.assertEqual(client.get(download_url).status_code, 404)
            self.assertFalse(export_path.exists())


class SqlGenerateTests(unittest.TestCase):
    """SQL Generate writes the import's SQL file but never changes the database."""

    CONFIG = dict(host="example", port=1433, database="test", username="user", password="secret")

    def setUp(self):
        with importer.IMPORT_JOBS_LOCK:
            importer.IMPORT_JOBS.clear()
        self.addCleanup(self.clear_jobs)

    def clear_jobs(self):
        with importer.IMPORT_JOBS_LOCK:
            for job in importer.IMPORT_JOBS.values():
                if job.get("sql_path"):
                    Path(job["sql_path"]).unlink(missing_ok=True)
            importer.IMPORT_JOBS.clear()

    def generate(self, connection, raw, filename="MDNNOP.csv", schema="lvapdbf"):
        client = importer.app.test_client()
        with patch.object(importer, "load_config", return_value=self.CONFIG), \
             patch.object(importer, "_build_conn_str", return_value=("mock connection", None)), \
             patch.object(importer.pyodbc, "connect", return_value=connection):
            response = client.post("/sql/generate", data={"file": (io.BytesIO(raw), filename), "schema": schema})
            if response.status_code != 202:
                return client, response, None
            job_id = response.json["job_id"]
            deadline = time.monotonic() + 3
            while time.monotonic() < deadline:
                snapshot = client.get(f"/import/status/{job_id}").json
                if snapshot["state"] in {"done", "failed"}:
                    break
                time.sleep(0.01)
        return client, response, snapshot

    @staticmethod
    def csv_bytes(rows, delimiter=","):
        stream = io.StringIO(newline="")
        csv.writer(stream, delimiter=delimiter).writerows(rows)
        return stream.getvalue().encode("utf-8")

    def test_generated_sql_matches_import_export_and_db_is_untouched(self):
        rows = [["0001", "12.50", "O'Brien 日本語", "", ""], ["0002", "", "text", "1", "2026-11-22"]]
        connection = FakeConnection()
        client, response, snapshot = self.generate(connection, self.csv_bytes(rows))
        self.assertEqual(response.status_code, 202)
        self.assertEqual(snapshot["state"], "done")
        self.assertEqual(snapshot["results"][0]["status"], "ok")
        self.assertIn("SQL generated for 2 rows", snapshot["results"][0]["message"])
        download = client.get(snapshot["sql_download_url"])
        self.assertEqual(download.status_code, 200)
        generated = download.data.decode("utf-8-sig")
        download.close()
        # Only metadata was read: nothing cleared, inserted or committed.
        self.assertTrue(all(sql.startswith("SELECT") for sql in connection.statements))
        self.assertIsNone(connection.committed)
        self.assertTrue(connection.closed)

        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "MDNNOP.csv"
            path.write_bytes(self.csv_bytes(rows))
            job = importer._new_import_job("parity", ["MDNNOP"])
            importer.IMPORT_JOBS["parity"] = job
            importer._run_csv_import("parity", FakeConnection(), str(path), "MDNNOP", ["MDNNOP"], "lvapdbf")
            imported = Path(job["sql_path"]).read_text(encoding="utf-8-sig")
        self.assertEqual(generated, imported)

    def test_tsv_and_header_row_are_supported(self):
        rows = [["CREATED", "CODE", "DESCRIPTION"], ["2026-11-22", "0001", "a\tb"]]
        _, _, snapshot = self.generate(FakeConnection(), self.csv_bytes(rows, "\t"), "MDNNOP.tsv")
        self.assertEqual(snapshot["state"], "done")
        self.assertIn("header row found", snapshot["results"][0]["message"])
        sql = Path(importer.IMPORT_JOBS[snapshot["job_id"]]["sql_path"]).read_text(encoding="utf-8-sig")
        self.assertIn("NCHAR(9)", sql)   # the quoted tab is kept inside the value
        self.assertIn("N'0001'", sql)

    def test_missing_table_is_skipped_without_download(self):
        _, _, snapshot = self.generate(FakeConnection(table_exists=False), self.csv_bytes([["0001"]]))
        self.assertEqual(snapshot["state"], "done")
        self.assertEqual(snapshot["results"][0]["status"], "skip")
        self.assertNotIn("sql_download_url", snapshot)

    def test_invalid_requests_are_rejected(self):
        raw = self.csv_bytes([["0001"]])
        self.assertEqual(self.generate(FakeConnection(), raw, "MDNNOP.xlsx")[1].status_code, 400)
        self.assertEqual(self.generate(FakeConnection(), raw, "123.csv")[1].status_code, 400)
        self.assertEqual(self.generate(FakeConnection(), raw, schema="")[1].status_code, 400)
        self.assertEqual(self.generate(FakeConnection(), b"0001,a\x81", "MDNNOP.csv")[1].status_code, 400)


if __name__ == "__main__":
    unittest.main()
