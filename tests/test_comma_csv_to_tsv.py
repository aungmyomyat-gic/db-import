"""Comma CSV → TSV for MSHHKP, whose file-name columns hold unquoted commas."""

import base64
import io
import unittest
from unittest.mock import patch

import app as importer
from test_csv_sql_export import FakeConnection

FILE_COLUMNS = importer.COMMA_FILE_COLUMNS["MSHHKP"]
# Short stand-in for the 31-column table: two plain columns around the four file columns.
COLUMNS = ["SHKCD", "SHKNM", "SHKTRM", "SHKFL1", "SHKFL2", "SHKFL3", "UPDDT"]


def convert(text, encoding="utf-8"):
    tsv, report = importer.comma_csv_to_tsv(text.encode(encoding), COLUMNS, FILE_COLUMNS)
    return tsv.decode(encoding), report


class CommaCsvToTsvTests(unittest.TestCase):
    def test_commas_stay_only_inside_file_columns(self):
        tsv, report = convert("0001,Name,a,b,c.pdf,x.png,,d,e.jpg,2025-01-01\n"
                              "0002,Plain,t.pdf,1.png,2.png,3.png,2025-01-02\n")
        self.assertEqual(tsv, "0001\tName\ta,b,c.pdf\tx.png\t\td,e.jpg\t2025-01-01\n"
                              "0002\tPlain\tt.pdf\t1.png\t2.png\t3.png\t2025-01-02\n")
        self.assertEqual((report["rows"], report["written"], report["joined"]), (2, 2, 1))
        self.assertEqual(report["errors"], [])

    def test_every_file_column_can_hold_commas_and_blanks_end_a_field(self):
        tsv, _ = convert("1,N,a,1.pdf,b,2.PNG,,c,d,3.xlsx,X\n")
        self.assertEqual(tsv, "1\tN\ta,1.pdf\tb,2.PNG\t\tc,d,3.xlsx\tX\n")

    def test_bad_rows_are_left_out_and_reported_with_line_numbers(self):
        tsv, report = convert("1,N,a.pdf,,,,X\n"
                              "\n"
                              "2,N,a,b,,,,X\n"          # SHKTRM never reaches an extension: ends at blank
                              "3,N,noext,,,\n"         # too few cells
                              "4,N,a.pdf,,,,X,extra\n")
        self.assertEqual(tsv, "1\tN\ta.pdf\t\t\t\tX\n")
        self.assertEqual([e["line"] for e in report["errors"]], [3, 4, 5])
        self.assertEqual(report["errors"][2]["text"], "4,N,a.pdf,,,,X,extra")
        self.assertIn("left over", report["errors"][2]["reason"])
        self.assertEqual((report["rows"], report["written"]), (4, 1))

    def test_line_with_one_cell_per_column_is_used_as_is(self):
        # "memo" has no extension, but the count matches, so nothing is joined.
        tsv, report = convert("1,N,memo,,,,X\n")
        self.assertEqual(tsv, "1\tN\tmemo\t\t\t\tX\n")
        self.assertEqual(report["joined"], 0)

    def test_empty_cells_at_the_end_of_the_line_are_dropped(self):
        # Like the real export: plain rows end with ",,,,", rows with split file names do not.
        tsv, report = convert("1,N, , , , ,X,,,,\n"
                              "2,N, ,試験（1-08,2-08と併用).pdf,報告書 .pdf,図面（1,2).pdf,X\n"
                              "3,N,a,b.pdf,,,,X,,\n")
        self.assertEqual(tsv, "1\tN\t \t \t \t \tX\n"
                              "2\tN\t \t試験（1-08,2-08と併用).pdf\t報告書 .pdf\t図面（1,2).pdf\tX\n"
                              "3\tN\ta,b.pdf\t\t\t\tX\n")
        self.assertEqual((report["written"], report["joined"], report["errors"]), (3, 2, []))

    def test_numbers_after_a_dot_are_not_an_extension(self):
        tsv, _ = convert("1,N,検査 2014.11,結果.pdf,,,,X\n")
        self.assertEqual(tsv, "1\tN\t検査 2014.11,結果.pdf\t\t\t\tX\n")

    def test_only_known_extensions_end_a_file_name(self):
        # ".35GR" is not a file extension, so it is joined with the next cell.
        tsv, report = convert("1,N,a test.pdf,JN75 0.35GR,NV function.pdf,x.Docx,,X\n")
        self.assertEqual(tsv, "1\tN\ta test.pdf\tJN75 0.35GR,NV function.pdf\tx.Docx\t\tX\n")
        self.assertEqual(report["errors"], [])
        self.assertTrue(importer._ends_file_name(" 報告書.PDF "))
        self.assertFalse(importer._ends_file_name("archive.zip"))   # not in FILE_EXTENSIONS

    def test_value_with_a_tab_is_written_in_quotes(self):
        # Like line 5985: a memo with a tab, in a row whose file names also have commas.
        tsv, report = convert("1,N,a,b.pdf,,,,マスター→\t8885→変更\n"
                              "2,N,x.pdf,,,,plain\n")
        self.assertEqual(tsv, '1\tN\ta,b.pdf\t\t\t\t"マスター→\t8885→変更"\n'
                              "2\tN\tx.pdf\t\t\t\tplain\n")
        self.assertEqual((report["written"], report["errors"], report["quoted_lines"]), (2, [], [1]))

    def test_missing_extension_at_end_of_line(self):
        _, report = convert("1,N,a,b,c\n")
        self.assertIn("SHKTRM: no file extension found in 'a,b,c'", report["errors"][0]["reason"])

    def test_header_is_kept_and_crlf_and_shift_jis_are_preserved(self):
        text = ",".join(COLUMNS) + "\r\n1,名前,資料,一.pdf,,,,X\r\n"
        tsv, report = convert(text, "cp932")
        self.assertEqual(tsv, "\t".join(COLUMNS) + "\r\n1\t名前\t資料,一.pdf\t\t\t\tX\r\n")
        self.assertTrue(report["header"])
        self.assertEqual((report["rows"], report["encoding"]), (1, "Shift-JIS"))

    def test_utf8_bom_is_kept_only_when_present(self):
        with_bom, _ = importer.comma_csv_to_tsv("﻿1,N,a.pdf,,,,X\n".encode("utf-8"),
                                                COLUMNS, FILE_COLUMNS)
        without, _ = importer.comma_csv_to_tsv(b"1,N,a.pdf,,,,X\n", COLUMNS, FILE_COLUMNS)
        self.assertTrue(with_bom.startswith(b"\xef\xbb\xbf"))
        self.assertFalse(without.startswith(b"\xef\xbb\xbf"))

    def post(self, filename, connection=None):
        config = dict(host="example", port=1433, database="test", username="user", password="secret")
        with patch.object(importer, "load_config", return_value=config), \
             patch.object(importer, "_build_conn_str", return_value=("mock", None)), \
             patch.object(importer.pyodbc, "connect", return_value=connection or FakeConnection()):
            return importer.app.test_client().post("/convert/comma-csv-to-tsv", data={
                "schema": "lvapdbf",
                "file": (io.BytesIO(b"1,N,a,b,c.pdf,,,,X\n"), filename),
            }, content_type="multipart/form-data")

    def test_route_reads_columns_from_the_table_and_returns_the_tsv(self):
        connection = FakeConnection(metadata=[(name,) for name in COLUMNS])
        response = self.post("MSHHKP (1).csv", connection)
        self.assertEqual(response.status_code, 200)
        data = response.get_json()
        self.assertEqual((data["table"], data["columns"], data["tsv_filename"]),
                         ("MSHHKP", 7, "MSHHKP (1).tsv"))
        self.assertEqual(base64.b64decode(data["tsv_base64"]).decode(),
                         "1\tN\ta,b,c.pdf\t\t\t\tX\n")
        self.assertTrue(connection.closed)

    def test_route_refuses_tables_without_a_rule_or_without_the_columns(self):
        response = self.post("MDNNOP.csv")
        self.assertEqual(response.status_code, 400)
        self.assertIn("No file-name column rule for table MDNNOP", response.get_json()["error"])
        response = self.post("MSHHKP.csv", FakeConnection(metadata=[("SHKCD",), ("SHKTRM",)]))
        self.assertEqual(response.status_code, 400)
        self.assertIn("has no column SHKFL1, SHKFL2, SHKFL3", response.get_json()["error"])


if __name__ == "__main__":
    unittest.main()
