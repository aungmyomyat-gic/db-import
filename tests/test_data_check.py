"""Data Check: every value of a CSV / TSV is checked against its table's columns."""

import contextlib
import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import app as importer
from test_csv_sql_export import FakeConnection

# (name, type, max length, precision, scale, nullable) as INFORMATION_SCHEMA returns them
METADATA = [("CODE", "nvarchar", 3, None, None, "NO"),
            ("QTY", "numeric", None, 8, 0, "NO"),
            ("MEMO", "varchar", 4, None, None, "YES")]
CONFIG = dict(host="example", port=1433, database="test", username="user", password="secret")


class DataCheckTests(unittest.TestCase):
    def check(self, raw, filename, schema="lvapdbf"):
        client = importer.app.test_client()
        with patch.object(importer, "load_config", return_value=CONFIG), \
             patch.object(importer, "_build_conn_str", return_value=("mock", None)), \
             patch.object(importer.pyodbc, "connect", return_value=FakeConnection(metadata=METADATA)):
            return client.post("/data/check", data={"schema": schema, "file": (io.BytesIO(raw), filename)},
                               content_type="multipart/form-data")

    def test_csv_and_tsv_give_the_same_problems(self):
        csv = self.check(b"001,12,ok\n0002,x,toolong\n", "MDNNOP.csv").get_json()
        tsv = self.check(b"001\t12\tok\n0002\tx\ttoolong\n", "MDNNOP.tsv").get_json()
        self.assertEqual((csv["format"], tsv["format"]), ("csv", "tsv"))
        for data in (csv, tsv):
            self.assertEqual((data["rows_checked"], data["bad_rows"], data["error_count"]), (2, 1, 3))
            self.assertEqual([(i["line"], i["column"]) for i in data["issues"]],
                             [(2, "CODE"), (2, "QTY"), (2, "MEMO")])

    def test_tsv_values_may_contain_commas_and_quoted_tabs(self):
        data = self.check(b'001\t1\t"a\tb"\n002\t2\ta,b\n', "MDNNOP.tsv").get_json()
        self.assertEqual(data["error_count"], 0)
        self.assertEqual(data["issue_count"], 0)

    def test_header_row_and_partial_header_warning(self):
        full = self.check(b"QTY\tCODE\n5\t001\n", "MDNNOP.tsv").get_json()
        self.assertTrue(full["header"])
        self.assertEqual(full["header_check"]["status"], "header")
        partial = self.check(b"CODE,QTY,MEMX\n001,5,a\n", "MDNNOP.csv").get_json()
        self.assertEqual((partial["header_check"]["status"], partial["header_check"]["unknown"]),
                         ("partial", ["MEMX"]))
        self.assertTrue(partial["header"])   # same automatic decision as the import
        self.assertEqual(partial["rows_checked"], 1)

    def test_wrong_number_of_values_is_a_warning(self):
        data = self.check(b"001\t1\n", "MDNNOP.tsv").get_json()
        self.assertEqual((data["warning_count"], data["issues"][0]["reason"]),
                         (1, "2 columns, table has 3 (missing columns become NULL/default)"))

    def test_invalid_requests_are_rejected(self):
        self.assertIn(".csv or .tsv", self.check(b"1", "MDNNOP.xlsx").get_json()["error"])
        self.assertIn("TSV file name", self.check(b"1", "123.tsv").get_json()["error"])
        self.assertEqual(self.check(b"1", "MDNNOP.tsv", schema="").status_code, 400)

    def test_tsv_import_logs_the_data_check(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "MDNNOP.tsv"
            path.write_bytes(b"001\tx\tok\n")
            output = io.StringIO()
            with contextlib.redirect_stdout(output):
                importer._datacheck_log(FakeConnection(metadata=METADATA), str(path), "lvapdbf", "MDNNOP", "\t")
        log = output.getvalue()
        self.assertIn("[DATA CHECK] lvapdbf.MDNNOP: 1 problems", log)
        self.assertIn("Column: QTY numeric(8,0)", log)


if __name__ == "__main__":
    unittest.main()
