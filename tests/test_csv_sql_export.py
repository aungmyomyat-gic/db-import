"""CSV/SQL parity and download lifecycle tests; no live database is modified."""

import csv
import io
import tempfile
import time
import threading
import unittest
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch

import pyodbc

import app as importer


CONFIG = dict(host="example", port=1433, database="test", username="user", password="secret")


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
                 rejected_code=None, fallback=False, commit_error=False, table_exists=True,
                 existing_rows=None):
        self.schema = schema
        self.metadata = metadata or METADATA
        self.truncate_blocked = truncate_blocked
        self.rejected_code = rejected_code
        self.fallback = fallback
        self.commit_error = commit_error
        self.table_exists = table_exists
        self.persisted = list(existing_rows or [])
        self.pending = list(self.persisted)
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
        self.persisted = list(self.pending)

    def rollback(self):
        self.rollbacks += 1
        self.pending = list(self.persisted)

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
        plan = [("CODE", ("csv", 0))]
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
        importer._write_import_sql(path, "order", "SELECT", [("CODE", ("csv", 0))], [["0001"]], columns)
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

    def test_non_csv_and_empty_uploads_are_rejected(self):
        client = importer.app.test_client()
        for filename, content in [("data.xlsx", b"workbook"), ("MDNNOP.csv", b"\n,\n")]:
            with self.subTest(filename=filename):
                response = client.post("/upload", data={"file": (io.BytesIO(content), filename)})
                self.assertEqual(response.status_code, 400)
                self.assertIn("error", response.json)
        with client.session_transaction() as browser_session:
            self.assertNotIn("uploads", browser_session)

    def test_csv_read_progress_preserves_encodings_multiline_fields_and_padding(self):
        rows = [["0001", "日本語\nsecond line", "3"], ["0002", "短い行"]] * 1100
        for encoding in ("utf-8-sig", "cp932"):
            with self.subTest(encoding=encoding):
                self.write_csv(rows, encoding)
                updates = []
                actual, width = importer.read_csv_rows(
                    self.csv_path, progress_cb=lambda *update: updates.append(update))
                self.assertEqual(width, 3)
                self.assertEqual(actual, [row + [""] * (3 - len(row)) for row in rows])
                self.assertEqual(updates[0][0], 0)
                self.assertEqual(updates[-1][0], updates[-1][1])
                self.assertEqual(updates[-1][2], len(rows))
                self.assertEqual([u[0] for u in updates], sorted(u[0] for u in updates))
                self.assertTrue(any(0 < done < total for done, total, _ in updates))

    def test_background_csv_preview_reports_progress_and_blocks_use_until_ready(self):
        client = importer.app.test_client()
        stranger = importer.app.test_client()
        started = threading.Event()
        release = threading.Event()
        finished = threading.Event()
        original_read = importer.read_csv_rows

        def slow_read(path, cancel_cb=None, progress_cb=None):
            progress_cb(4, 10, 1)
            started.set()
            if not release.wait(3):
                raise RuntimeError("Test did not release preview worker")
            return original_read(path, cancel_cb=cancel_cb, progress_cb=progress_cb)

        original_worker = importer._run_csv_preview

        def worker(*args):
            try:
                original_worker(*args)
            finally:
                finished.set()

        raw = b"FIRST,1\nSECOND,2\n"
        with patch.object(importer, "read_csv_rows", side_effect=slow_read), \
             patch.object(importer, "_run_csv_preview", side_effect=worker), \
             patch.object(importer, "load_config", return_value=CONFIG):
            try:
                response = client.post("/upload", data={
                    "file": (io.BytesIO(raw), "MDNNOP.csv"), "purpose": "generate", "background": "1"})
                self.assertEqual(response.status_code, 202)
                self.assertTrue(started.wait(3))
                status_url = response.json["status_url"]
                status = client.get(status_url).json
                self.assertEqual(status["state"], "running")
                self.assertEqual(status["file_size"], len(raw))
                self.assertEqual(status["results"][0]["progress"], {"done": 4, "total": 10, "records": 1})
                self.assertNotIn("preview", status)
                self.assertNotIn("owner", status)
                self.assertNotIn("sql_download_url", status)
                self.assertEqual(stranger.get(status_url).status_code, 404)
                self.assertEqual(client.post("/scan", json={"purpose": "generate"}).status_code, 400)
                self.assertEqual(client.post("/generate-sql", json={}).status_code, 400)
                self.assertEqual(client.post("/import/cancel/" + status["job_id"]).status_code, 409)
            finally:
                release.set()
                self.assertTrue(finished.wait(3))
        status = client.get(status_url).json
        self.assertEqual(status["state"], "done")
        self.assertEqual(status["preview"]["records"], 2)
        self.assertEqual(status["preview"]["columns"], 2)
        with client.session_transaction() as browser_session:
            path = Path(browser_session["uploads"]["generate"]["path"])
        self.addCleanup(path.unlink, missing_ok=True)
        self.assertEqual(path.read_bytes(), raw)

    def test_background_csv_preview_failure_discards_input_and_allows_retry(self):
        client = importer.app.test_client()
        for content in (b"\n,\n", b"\x81"):
            with self.subTest(content=content):
                response = client.post("/upload", data={
                    "file": (io.BytesIO(content), "MDNNOP.csv"), "background": "1"})
                snapshot = self.await_job(client, response)
                self.assertEqual(snapshot["state"], "failed")
                self.assertIn("Cannot open CSV", snapshot["error"])
                self.assertNotIn("preview", snapshot)
                with client.session_transaction() as browser_session:
                    path = Path(browser_session["uploads"]["import"]["path"])
                self.assertFalse(path.exists())
        response = client.post("/upload", data={
            "file": (io.BytesIO(b"VALID,1\n"), "MDNNOP.csv"), "background": "1"})
        self.assertEqual(self.await_job(client, response)["state"], "done")
        with client.session_transaction() as browser_session:
            path = Path(browser_session["uploads"]["import"]["path"])
        self.addCleanup(path.unlink, missing_ok=True)

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
                uploaded_path = Path(browser_session["uploads"]["import"]["path"])
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


    def upload_csv(self, client, raw, purpose="import", filename="MDNNOP.csv"):
        response = client.post("/upload", data={"file": (io.BytesIO(raw), filename), "purpose": purpose})
        self.assertEqual(response.status_code, 200, response.json)
        with client.session_transaction() as browser_session:
            path = Path(browser_session["uploads"][purpose]["path"])
        self.addCleanup(path.unlink, missing_ok=True)
        return response.json

    def await_job(self, client, response):
        self.assertEqual(response.status_code, 202, response.json)
        deadline = time.monotonic() + 3
        while time.monotonic() < deadline:
            status = client.get(response.json["status_url"])
            self.assertEqual(status.status_code, 200)
            if status.json["state"] in {"done", "failed", "cancelled"}:
                return status.json
            time.sleep(0.01)
        self.fail("CSV job did not finish")

    def test_standalone_generation_matches_import_sql_without_database_writes(self):
        raw = self.write_csv([["0001", "12.50", "日本語 O'Brien", "", ""],
                              ["0002", "", "next row", "1", "2026-11-22"]])
        imported = self.run_csv(FakeConnection(), [["0001", "12.50", "日本語 O'Brien", "", ""],
                                                  ["0002", "", "next row", "1", "2026-11-22"]])
        expected = Path(imported["sql_path"]).read_bytes()
        client = importer.app.test_client()
        self.upload_csv(client, raw, "generate")
        connection = FakeConnection()
        with patch.object(importer, "load_config", return_value=CONFIG), \
             patch.object(importer, "_build_conn_str", return_value=("mock", None)), \
             patch.object(importer.pyodbc, "connect", return_value=connection):
            snapshot = self.await_job(client, client.post("/generate-sql", json={"schema": "lvapdbf"}))
        self.assertEqual(snapshot["state"], "done")
        self.assertEqual(snapshot["operation"], "generate")
        self.assertIsNone(connection.committed)
        self.assertEqual(connection.pending, [])
        self.assertEqual(connection.rollbacks, 0)
        self.assertTrue(connection.closed)
        self.assertTrue(all(sql.startswith("SELECT ") for sql in connection.statements))
        download = client.get(snapshot["sql_download_url"])
        self.assertEqual(download.status_code, 200)
        self.assertEqual(download.data, expected)
        download.close()
        stranger = importer.app.test_client()
        self.assertEqual(stranger.get(snapshot["sql_download_url"]).status_code, 404)
        self.assertEqual(stranger.get(snapshot["status_url"]).status_code, 404)
        job = importer.IMPORT_JOBS[snapshot["job_id"]]
        job["updated_at"] = 0
        export = Path(job["sql_path"])
        self.assertEqual(client.get(snapshot["sql_download_url"]).status_code, 404)
        self.assertFalse(export.exists())

    def test_both_tabs_keep_their_own_csv(self):
        client = importer.app.test_client()
        self.upload_csv(client, b"IMPORT_ONLY\n", "import")
        self.upload_csv(client, b"GENERATE_ONLY\n", "generate", "OTHER.csv")
        imported, generated = FakeConnection(), FakeConnection()
        with patch.object(importer, "load_config", return_value=CONFIG), \
             patch.object(importer, "_build_conn_str", return_value=("mock", None)), \
             patch.object(importer.pyodbc, "connect", side_effect=[generated, imported]):
            sql_job = self.await_job(client, client.post("/generate-sql", json={"schema": "lvapdbf"}))
            import_job = self.await_job(client, client.post("/import", json={"schema": "lvapdbf"}))
        self.assertEqual(imported.committed[0][0], "IMPORT_ONLY")
        download = client.get(sql_job["sql_download_url"])
        self.assertIn("INSERT INTO lvapdbf.OTHER", download.data.decode("utf-8-sig"))
        self.assertIn("N'GENERATE_ONLY'", download.data.decode("utf-8-sig"))
        self.assertNotIn("IMPORT_ONLY", download.data.decode("utf-8-sig"))
        download.close()
        self.assertEqual(import_job["results"][0]["name"], "MDNNOP")
        self.assertIsNone(generated.committed)

    def test_reupload_cannot_replace_a_queued_jobs_input(self):
        client = importer.app.test_client()
        self.upload_csv(client, b"ORIGINAL\n")
        original_worker = importer._run_import_job
        release_worker = threading.Event()
        worker_finished = threading.Event()
        captured = []

        def delayed_worker(*args):
            captured.append(args[2])
            try:
                if release_worker.wait(timeout=3):
                    original_worker(*args)
            finally:
                worker_finished.set()

        connection = FakeConnection()
        with patch.object(importer, "load_config", return_value=CONFIG), \
             patch.object(importer, "_build_conn_str", return_value=("mock", None)), \
             patch.object(importer.pyodbc, "connect", return_value=connection), \
             patch.object(importer, "_run_import_job", side_effect=delayed_worker):
            response = client.post("/import", json={"schema": "lvapdbf"})
            self.assertEqual(response.status_code, 202)
            self.upload_csv(client, b"REPLACEMENT\n")
            release_worker.set()
            self.assertTrue(worker_finished.wait(timeout=3))
            snapshot = self.await_job(client, response)
        self.assertEqual(snapshot["state"], "done")
        self.assertEqual(connection.committed[0][0], "ORIGINAL")
        self.assertTrue(captured)
        self.assertFalse(Path(captured[0]).exists())

    def test_elapsed_time_increases_then_freezes_on_completion(self):
        with patch.object(importer.time, "perf_counter", return_value=100):
            job = importer._new_import_job(self.job_id, ["MDNNOP"])
        importer.IMPORT_JOBS[self.job_id] = job
        with patch.object(importer.time, "perf_counter", return_value=101.125):
            self.assertEqual(importer._job_snapshot(job)["elapsed_seconds"], 1.125)
            importer._set_job_state(self.job_id, "running")
        with patch.object(importer.time, "perf_counter", return_value=105.5):
            importer._finish_import_job(self.job_id)
        with patch.object(importer.time, "perf_counter", return_value=999):
            snapshot = importer._job_snapshot(job)
            self.assertEqual(snapshot["elapsed_seconds"], 5.5)
        self.assertNotIn("started_at", snapshot)
        self.assertNotIn("finished_at", snapshot)

    def test_elapsed_time_freezes_on_failure(self):
        with patch.object(importer.time, "perf_counter", return_value=100):
            job = importer._new_import_job(self.job_id, ["MDNNOP"])
        importer.IMPORT_JOBS[self.job_id] = job
        with patch.object(importer.time, "perf_counter", return_value=102.25):
            importer._fail_import_job(self.job_id, "Connection lost")
        with patch.object(importer.time, "perf_counter", return_value=999):
            snapshot = importer._job_snapshot(job)
        self.assertEqual(snapshot["elapsed_seconds"], 2.25)
        self.assertEqual(snapshot["state"], "failed")
        self.assertNotIn("sql_download_url", snapshot)

    def test_generator_failure_keeps_database_unchanged_and_hides_download(self):
        for failure in ("missing_table", "write_error"):
            with self.subTest(failure=failure):
                client = importer.app.test_client()
                self.upload_csv(client, b"0001\n", "generate")
                connection = FakeConnection(table_exists=failure != "missing_table")
                writer = importer._write_import_sql
                if failure == "write_error":
                    def writer(path, *args):
                        Path(path).write_text("partial SQL")
                        raise OSError("Disk full")
                with patch.object(importer, "load_config", return_value=CONFIG), \
                     patch.object(importer, "_build_conn_str", return_value=("mock", None)), \
                     patch.object(importer.pyodbc, "connect", return_value=connection), \
                     patch.object(importer, "_write_import_sql", side_effect=writer):
                    snapshot = self.await_job(client, client.post("/generate-sql", json={"schema": "lvapdbf"}))
                self.assertEqual(snapshot["state"], "failed")
                self.assertEqual(snapshot["results"][0]["status"], "failed")
                self.assertNotIn("sql_download_url", snapshot)
                self.assertFalse((Path(importer.SQL_EXPORT_DIR.name) / f"{snapshot['job_id']}.sql").exists())
                self.assertEqual(connection.rollbacks, 0)
                self.assertIsNone(connection.committed)
                self.assertTrue(connection.closed)
                self.assertTrue(all(sql.startswith("SELECT ") for sql in connection.statements))

    def test_connection_failure_cleans_job_input_and_reports_failure(self):
        client = importer.app.test_client()
        self.upload_csv(client, b"0001\n")
        with patch.object(importer, "load_config", return_value=CONFIG), \
             patch.object(importer, "_build_conn_str", return_value=("mock", None)), \
             patch.object(importer.pyodbc, "connect", side_effect=pyodbc.Error("Cannot connect")):
            snapshot = self.await_job(client, client.post("/import", json={"schema": "lvapdbf"}))
        self.assertEqual(snapshot["state"], "failed")
        self.assertIn("Cannot connect", snapshot["error"])
        self.assertNotIn("sql_download_url", snapshot)
        self.assertGreaterEqual(snapshot["elapsed_seconds"], 0)
        self.assertFalse((Path(importer.UPLOAD_DIR.name) / f"job-{snapshot['job_id']}.csv").exists())

    def test_removed_features_have_no_routes(self):
        client = importer.app.test_client()
        for endpoint in ("/query", "/tables", "/table-data", "/truncate", "/call-batch",
                         "/schemas/list", "/version", "/update"):
            for method in ("GET", "POST", "PATCH", "DELETE"):
                with self.subTest(endpoint=endpoint, method=method):
                    self.assertEqual(client.open(endpoint, method=method).status_code, 404)
        self.assertEqual(client.post("/schemas", json={"name": "new_schema"}).status_code, 405)

    def test_each_workflow_requires_its_own_upload_and_connection(self):
        client = importer.app.test_client()
        with patch.object(importer, "load_config", return_value={}):
            for endpoint in ("/import", "/generate-sql"):
                response = client.post(endpoint, json={})
                self.assertEqual(response.status_code, 400)
                self.assertIn("connection", response.json["error"])
        self.upload_csv(client, b"0001\n", "import")
        with patch.object(importer, "load_config", return_value=CONFIG):
            response = client.post("/generate-sql", json={})
        self.assertEqual(response.status_code, 400)
        self.assertIn("upload", response.json["error"])


    def owned_job(self, client, operation="import"):
        job = importer._new_import_job(self.job_id, ["MDNNOP"], operation)
        job["owner"] = "test-owner"
        importer.IMPORT_JOBS[self.job_id] = job
        with client.session_transaction() as browser_session:
            browser_session["import_owner"] = job["owner"]
        return job

    def assert_cancelled(self, connection, job):
        snapshot = importer._job_snapshot(job)
        self.assertEqual(snapshot["state"], "cancelled")
        self.assertEqual(snapshot["results"][0]["status"], "cancelled")
        self.assertFalse(snapshot["can_cancel"])
        self.assertNotIn("sql_download_url", snapshot)
        self.assertNotIn("sql_path", job)
        self.assertIsNone(connection.committed)
        self.assertEqual(connection.persisted, [("ORIGINAL",)])
        self.assertEqual(connection.pending, [("ORIGINAL",)])
        self.assertFalse((Path(importer.SQL_EXPORT_DIR.name) / f"{self.job_id}.sql").exists())

    def test_http_cancel_during_batch_rolls_back_original_data_and_freezes_time(self):
        client = importer.app.test_client()
        self.upload_csv(client, b"FIRST\nSECOND\nTHIRD\n")
        connection = FakeConnection(existing_rows=[("ORIGINAL",)])
        cursor = connection.cursor()
        original_many = cursor.executemany
        batch_started, release_batch = threading.Event(), threading.Event()

        def blocked_batch(sql, rows):
            original_many(sql, rows)
            batch_started.set()
            if not release_batch.wait(timeout=3):
                raise RuntimeError("Test batch was not released")

        with patch.object(importer, "load_config", return_value=CONFIG), \
             patch.object(importer, "_build_conn_str", return_value=("mock", None)), \
             patch.object(importer.pyodbc, "connect", return_value=connection), \
             patch.object(importer, "IMPORT_BATCH_SIZE", 1), \
             patch.object(cursor, "executemany", side_effect=blocked_batch):
            response = client.post("/import", json={"schema": "lvapdbf"})
            self.assertEqual(response.status_code, 202)
            job_id = response.json["job_id"]
            try:
                self.assertTrue(batch_started.wait(timeout=3))
                cancel_url = f"/import/cancel/{job_id}"
                stranger = importer.app.test_client()
                self.assertEqual(stranger.post(cancel_url).status_code, 404)
                cancelled = client.post(cancel_url)
                self.assertEqual(cancelled.status_code, 202)
                self.assertEqual(cancelled.json["state"], "cancelling")
                self.assertFalse(cancelled.json["can_cancel"])
                self.assertNotIn("sql_download_url", cancelled.json)
                self.assertEqual(client.post(cancel_url).status_code, 202)
                self.assertIsNone(connection.committed)
            finally:
                release_batch.set()
            snapshot = self.await_job(client, response)
        self.assertEqual(snapshot["state"], "cancelled")
        self.assertEqual(snapshot["results"][0]["status"], "cancelled")
        self.assertEqual(snapshot["completed"], 1)
        self.assertEqual(connection.rollbacks, 1)
        self.assertEqual(connection.persisted, [("ORIGINAL",)])
        self.assertEqual(connection.pending, [("ORIGINAL",)])
        self.assertIsNone(connection.committed)
        self.assertEqual(sum(sql.startswith("INSERT") for sql in connection.statements), 1)
        self.assertNotIn("sql_download_url", snapshot)
        self.assertEqual(client.get(f"/sql/{job_id}").status_code, 409)
        self.assertEqual(client.get(response.json["status_url"]).json["elapsed_seconds"], snapshot["elapsed_seconds"])
        self.assertEqual(client.post(cancel_url).json["state"], "cancelled")
        self.assertFalse((Path(importer.UPLOAD_DIR.name) / f"job-{job_id}.csv").exists())
        self.assertFalse((Path(importer.SQL_EXPORT_DIR.name) / f"{job_id}.sql").exists())
        importer.IMPORT_JOBS[job_id]["updated_at"] = 0
        self.assertEqual(client.get(response.json["status_url"]).status_code, 404)

    def test_cancel_before_worker_starts_skips_database_connection(self):
        client = importer.app.test_client()
        self.upload_csv(client, b"FIRST\n")
        original_worker = importer._run_import_job
        release_worker, worker_finished = threading.Event(), threading.Event()

        def delayed_worker(*args):
            try:
                if release_worker.wait(timeout=3):
                    original_worker(*args)
            finally:
                worker_finished.set()

        with patch.object(importer, "load_config", return_value=CONFIG), \
             patch.object(importer, "_build_conn_str", return_value=("mock", None)), \
             patch.object(importer.pyodbc, "connect") as connect, \
             patch.object(importer, "_run_import_job", side_effect=delayed_worker):
            response = client.post("/import", json={"schema": "lvapdbf"})
            self.assertEqual(response.status_code, 202)
            job_id = response.json["job_id"]
            try:
                cancel = client.post(f"/import/cancel/{job_id}")
                self.assertEqual(cancel.status_code, 202)
            finally:
                release_worker.set()
            self.assertTrue(worker_finished.wait(timeout=3))
            snapshot = self.await_job(client, response)
            connect.assert_not_called()
        self.assertEqual(snapshot["state"], "cancelled")
        self.assertFalse((Path(importer.UPLOAD_DIR.name) / f"job-{job_id}.csv").exists())

    def test_cancel_during_metadata_stops_before_clearing_table(self):
        client = importer.app.test_client()
        job = self.owned_job(client)
        connection = FakeConnection(existing_rows=[("ORIGINAL",)])
        cursor = connection.cursor()
        execute = cursor.execute

        def cancel_after_metadata(sql, *args):
            result = execute(sql, *args)
            if sql.startswith("SELECT COLUMN_NAME"):
                self.assertEqual(client.post(f"/import/cancel/{self.job_id}").status_code, 202)
            return result

        with patch.object(cursor, "execute", side_effect=cancel_after_metadata):
            self.run_csv_with_job(connection)
        self.assert_cancelled(connection, job)
        self.assertTrue(all(sql.startswith("SELECT ") for sql in connection.statements))

    def run_csv_with_job(self, connection):
        self.write_csv([["FIRST"], ["SECOND"]])
        importer._run_csv_import(self.job_id, connection, str(self.csv_path), "MDNNOP", ["MDNNOP"], "lvapdbf")
        importer._finish_import_job(self.job_id)

    def test_cancel_during_sql_export_rolls_back_cleared_table_and_discards_partial_file(self):
        client = importer.app.test_client()
        job = self.owned_job(client)
        connection = FakeConnection(existing_rows=[("ORIGINAL",)])
        write_sql = importer._write_import_sql

        def cancel_during_export(path, *args, **kwargs):
            Path(path).write_text("partial SQL")
            self.assertEqual(client.post(f"/import/cancel/{self.job_id}").status_code, 202)
            return write_sql(path, *args, **kwargs)

        with patch.object(importer, "_write_import_sql", side_effect=cancel_during_export):
            self.run_csv_with_job(connection)
        self.assert_cancelled(connection, job)
        self.assertEqual(connection.rollbacks, 1)
        self.assertFalse(any(sql.startswith("INSERT") for sql in connection.statements))

    def test_cancel_during_row_fallback_stops_remaining_rows_and_rolls_back(self):
        client = importer.app.test_client()
        job = self.owned_job(client)
        connection = FakeConnection(existing_rows=[("ORIGINAL",)], fallback=True)
        cursor = connection.cursor()
        execute = cursor.execute

        def cancel_after_first_insert(sql, *args):
            result = execute(sql, *args)
            if sql.startswith("INSERT"):
                self.assertEqual(client.post(f"/import/cancel/{self.job_id}").status_code, 202)
            return result

        with patch.object(cursor, "execute", side_effect=cancel_after_first_insert):
            self.run_csv_with_job(connection)
        self.assert_cancelled(connection, job)
        # One failed bulk INSERT, then the first single-row fallback INSERT.
        self.assertEqual(sum(sql.startswith("INSERT") for sql in connection.statements), 2)

    def test_cancel_immediately_before_commit_prevents_commit(self):
        client = importer.app.test_client()
        job = self.owned_job(client)
        connection = FakeConnection(existing_rows=[("ORIGINAL",)])
        begin_commit = importer._begin_import_commit

        def cancel_at_commit(job_id):
            self.assertEqual(client.post(f"/import/cancel/{job_id}").status_code, 202)
            begin_commit(job_id)

        with patch.object(importer, "_begin_import_commit", side_effect=cancel_at_commit):
            self.run_csv_with_job(connection)
        self.assert_cancelled(connection, job)
        self.assertEqual(connection.rollbacks, 1)

    def test_cancel_is_rejected_once_commit_begins_and_completion_is_preserved(self):
        client = importer.app.test_client()
        job = self.owned_job(client)
        connection = FakeConnection(existing_rows=[("ORIGINAL",)])
        commit = connection.commit

        def cancel_during_commit():
            snapshot = client.get(f"/jobs/{self.job_id}").json
            self.assertEqual(snapshot["state"], "committing")
            self.assertFalse(snapshot["can_cancel"])
            self.assertEqual(client.post(f"/import/cancel/{self.job_id}").status_code, 409)
            commit()

        with patch.object(connection, "commit", side_effect=cancel_during_commit):
            self.run_csv_with_job(connection)
        self.assertEqual(job["state"], "done")
        self.assertIsNotNone(connection.committed)
        self.assertEqual(connection.rollbacks, 0)
        self.assertIn("sql_download_url", importer._job_snapshot(job))
        response = client.post(f"/import/cancel/{self.job_id}")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json["state"], "done")
        self.assertIn("sql_download_url", response.json)

    def test_cancel_route_rejects_unknown_jobs_and_sql_generation(self):
        client = importer.app.test_client()
        self.assertEqual(client.post("/import/cancel/missing").status_code, 404)
        job = self.owned_job(client, "generate")
        self.assertFalse(importer._job_snapshot(job)["can_cancel"])
        response = client.post(f"/import/cancel/{self.job_id}")
        self.assertEqual(response.status_code, 409)
        self.assertEqual(job["state"], "queued")

    def test_rollback_failure_is_reported_as_failure_instead_of_cancelled(self):
        client = importer.app.test_client()
        job = self.owned_job(client)
        connection = FakeConnection(existing_rows=[("ORIGINAL",)])
        self.assertEqual(client.post(f"/import/cancel/{self.job_id}").status_code, 202)
        with patch.object(connection, "rollback", side_effect=pyodbc.Error("Rollback failed")):
            self.run_csv_with_job(connection)
        self.assertEqual(job["state"], "failed")
        self.assertIn("Could not roll back", job["results"][0]["message"])
        self.assertNotIn("sql_download_url", importer._job_snapshot(job))
        self.assertIsNone(connection.committed)


if __name__ == "__main__":
    unittest.main()
