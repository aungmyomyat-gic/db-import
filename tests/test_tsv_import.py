"""TSV import: same engine as CSV import, values separated by tabs."""

import io
import time
import unittest
from pathlib import Path
from unittest.mock import patch

import app as importer
from test_csv_sql_export import FakeConnection

METADATA = [("CODE", "varchar", "YES", None), ("FILE_NAME", "varchar", "YES", None),
            ("QTY", "varchar", "YES", None)]
CONFIG = dict(host="example", port=1433, database="test", username="user", password="secret")


class TsvImportTests(unittest.TestCase):
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

    def patched(self, connection):
        return (patch.object(importer, "load_config", return_value=CONFIG),
                patch.object(importer, "_build_conn_str", return_value=("mock", None)),
                patch.object(importer.pyodbc, "connect", return_value=connection))

    def post(self, client, url, raw, filename="MSHHKP.tsv", schema="lvapdbf"):
        return client.post(url, data={"schema": schema, "file": (io.BytesIO(raw), filename)},
                           content_type="multipart/form-data")

    def test_tabs_split_values_and_quoted_values_keep_their_tabs(self):
        raw = '1\ta,b,c.pdf\t3\n\n2\t"memo\twith tab"\t"say ""hi"""\n'.encode("cp932")
        self.assertEqual(importer.parse_delimited(raw, "\t"),
                         [(1, ["1", "a,b,c.pdf", "3"]), (3, ["2", "memo\twith tab", 'say "hi"'])])

    def test_what_csv_to_tsv_writes_is_read_back_exactly(self):
        values = ["06800", "→\t8885008400920。\t\tマ", 'a "b".pdf', "a,b,c.pdf", " ", ""]
        line = importer._tsv_line(values)
        self.assertEqual(importer.parse_delimited(line.encode("utf-8"), "\t"), [(1, values)])
        # The comma CSV reader is unchanged.
        self.assertEqual(importer.parse_delimited(b'1,"a,b.pdf",3\n'), [(1, ["1", "a,b.pdf", "3"])])

    def test_preview_reports_header_and_rows_with_the_wrong_number_of_values(self):
        connection = FakeConnection(metadata=METADATA)
        raw = "1\ta,b.pdf\t3\n2\tx.pdf\n3\ty.pdf\t4\textra\n".encode()
        client = importer.app.test_client()
        p1, p2, p3 = self.patched(connection)
        with p1, p2, p3:
            data = self.post(client, "/tsv/preview", raw, "MSHHKP (1).tsv").get_json()
            header = self.post(client, "/tsv/preview", b"CODE\tFILE_NAME\tOTHER\n1\ta.pdf\tz\n").get_json()
            wrong = self.post(client, "/tsv/preview", raw, "MSHHKP.csv")
        self.assertEqual((data["table"], data["records"], data["columns"], data["header"]),
                         ("MSHHKP", 3, 3, False))
        self.assertEqual([(m["line"], m["cells"]) for m in data["mismatched"]], [(2, 2), (3, 4)])
        self.assertEqual((header["header"], header["records"], header["unknown_header_columns"]),
                         (True, 1, ["OTHER"]))
        self.assertEqual(wrong.status_code, 400)
        self.assertIn(".tsv", wrong.get_json()["error"])

    def test_import_replaces_rows_and_offers_sql_download(self):
        connection = FakeConnection(metadata=METADATA)
        raw = ("CODE\tFILE_NAME\tQTY\r\n1\t熱衝撃試験（1-08,2-08と併用).pdf\t3\r\n2\t \t\r\n"
               '3\t"→\t8885"\t4\r\n').encode("cp932")
        client = importer.app.test_client()
        p1, p2, p3 = self.patched(connection)
        with p1, p2, p3:
            response = self.post(client, "/tsv/import", raw)
            self.assertEqual(response.status_code, 202)
            job_id = response.get_json()["job_id"]
            deadline = time.monotonic() + 3
            while time.monotonic() < deadline:
                snapshot = client.get(f"/import/status/{job_id}").get_json()
                if snapshot["state"] in {"done", "failed"}:
                    break
                time.sleep(0.01)
            self.assertEqual(snapshot["state"], "done")
            result = snapshot["results"][0]
            self.assertEqual((result["name"], result["status"], result["schema"]),
                             ("MSHHKP", "ok", "lvapdbf"))
            self.assertIn("3 rows inserted", result["message"])
            self.assertIn("header row found", result["message"])
            download = client.get(snapshot["sql_download_url"])
            sql = download.data.decode("utf-8-sig")
            download.close()
        self.assertEqual(connection.committed, [("1", "熱衝撃試験（1-08,2-08と併用).pdf", "3"),
                                                ("2", " ", None), ("3", "→\t8885", "4")])
        self.assertIn("N'熱衝撃試験（1-08,2-08と併用).pdf'", sql)
        self.assertTrue(any(s.startswith("TRUNCATE") for s in connection.statements))

    def test_import_needs_schema_and_a_tsv_named_after_the_table(self):
        client = importer.app.test_client()
        p1, p2, p3 = self.patched(FakeConnection())
        with p1, p2, p3:
            self.assertIn("Schema is required",
                          self.post(client, "/tsv/import", b"1\t2\n", schema="").get_json()["error"])
            self.assertIn("must start with the table name",
                          self.post(client, "/tsv/import", b"1\t2\n", "123.tsv").get_json()["error"])
        self.assertEqual(importer.IMPORT_JOBS, {})


if __name__ == "__main__":
    unittest.main()
