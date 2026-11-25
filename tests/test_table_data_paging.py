import csv
import io
import unittest
from decimal import Decimal
from unittest.mock import patch

import app as importer


CONFIG = {"host": "db", "port": "1433", "database": "test", "username": "u", "password": "p"}


class FakeCursor:
    """Answers the Data Editor's metadata, COUNT and SELECT queries."""

    def __init__(self, columns, primary_keys, rows):
        self.columns = columns
        self.primary_keys = primary_keys
        self.rows = rows
        self.executed = []
        self.result = []
        self.position = 0

    def execute(self, sql, *params):
        self.executed.append((sql, params))
        if "INFORMATION_SCHEMA.COLUMNS" in sql:
            self.result = [(name, dtype, "YES") for name, dtype in self.columns]
        elif "PRIMARY KEY" in sql:
            self.result = [(name,) for name in self.primary_keys]
        elif "COUNT_BIG" in sql:
            self.result = [(len(self.rows),)]
        elif "OFFSET" in sql:
            offset, size = params
            self.result = self.rows[offset:offset + size]
        else:
            self.result = list(self.rows)
        self.position = 0
        return self

    def fetchall(self):
        return list(self.result)

    def fetchone(self):
        return self.result[0]

    def fetchmany(self, size):
        chunk = self.result[self.position:self.position + size]
        self.position += size
        return chunk


class FakeConnection:
    def __init__(self, cursor):
        self._cursor = cursor
        self.closed = False

    def cursor(self):
        return self._cursor

    def close(self):
        self.closed = True


class TableDataPagingTests(unittest.TestCase):
    def setUp(self):
        self.rows = [(i, f"name {i}", Decimal(f"{i}.50"), i % 2 == 0) for i in range(1, 2501)]
        self.cursor = FakeCursor(
            [("ID", "int"), ("NAME", "nvarchar"), ("AMT", "decimal"), ("FLAG", "bit")],
            ["ID"], self.rows)
        self.connection = FakeConnection(self.cursor)
        patches = [
            patch.object(importer, "load_config", return_value=CONFIG),
            patch.object(importer, "_build_conn_str", return_value=("conn", None)),
            patch.object(importer.pyodbc, "connect", return_value=self.connection),
        ]
        for p in patches:
            p.start()
            self.addCleanup(p.stop)
        self.client = importer.app.test_client()

    def select_sql(self):
        return next(sql for sql, _ in self.cursor.executed if sql.startswith("SELECT * FROM"))

    def test_first_page_uses_offset_fetch_and_reports_totals(self):
        data = self.client.get("/table-data?schema=lvapdbf&table=T1").get_json()
        self.assertEqual(len(data["rows"]), 1000)
        self.assertEqual(data["rows"][0][0], 1)
        self.assertEqual((data["page"], data["page_size"], data["total_pages"], data["total_rows"]),
                         (1, 1000, 3, 2500))
        self.assertIn("ORDER BY [ID] OFFSET ? ROWS FETCH NEXT ? ROWS ONLY", self.select_sql())

    def test_last_page_is_partial_and_page_past_the_end_is_clamped(self):
        data = self.client.get("/table-data?schema=lvapdbf&table=T1&page=99&page_size=2000").get_json()
        self.assertEqual(data["page"], 2)
        self.assertEqual(len(data["rows"]), 500)
        self.assertEqual(data["rows"][0][0], 2001)

    def test_sort_column_comes_first_then_primary_key(self):
        data = self.client.get("/table-data?schema=lvapdbf&table=T1&sort=AMT&dir=desc").get_json()
        self.assertEqual((data["sort"], data["dir"]), ("AMT", "desc"))
        self.assertIn("ORDER BY [AMT] DESC, [ID] OFFSET", self.select_sql())

    def test_rejects_unknown_page_size_and_sort_column(self):
        response = self.client.get("/table-data?schema=lvapdbf&table=T1&page_size=3000")
        self.assertEqual(response.status_code, 400)
        response = self.client.get("/table-data?schema=lvapdbf&table=T1&sort=NOPE")
        self.assertEqual(response.status_code, 400)
        self.assertIn("Unknown sort column", response.get_json()["error"])

    def test_table_without_primary_key_orders_by_sortable_columns(self):
        self.cursor.primary_keys = []
        self.cursor.columns = [("A", "int"), ("NOTE", "ntext"), ("B", "varchar")]
        self.client.get("/table-data?schema=lvapdbf&table=T1")
        self.assertIn("ORDER BY [A], [B] OFFSET", self.select_sql())
        response = self.client.get("/table-data?schema=lvapdbf&table=T1&sort=NOTE")
        self.assertEqual(response.status_code, 400)

    def test_download_streams_all_rows_as_csv_in_sort_order(self):
        response = self.client.get("/table-data/download?schema=lvapdbf&table=T1&sort=NAME&dir=desc")
        self.assertEqual(response.status_code, 200)
        self.assertIn("T1 (", response.headers["Content-Disposition"])
        text = response.get_data(as_text=True)
        self.assertTrue(text.startswith("﻿"))
        rows = list(csv.reader(io.StringIO(text.lstrip("﻿"))))
        self.assertEqual(rows[0], ["ID", "NAME", "AMT", "FLAG"])
        self.assertEqual(rows[1], ["1", "name 1", "1.50", "0"])
        self.assertEqual(len(rows), 2501)
        self.assertIn("ORDER BY [NAME] DESC, [ID]", self.select_sql())
        self.assertNotIn("OFFSET", self.select_sql())
        self.assertTrue(self.connection.closed)

    def test_download_without_header_matches_csv_import_format(self):
        response = self.client.get("/table-data/download?schema=lvapdbf&table=T1&header=0")
        rows = list(csv.reader(io.StringIO(response.get_data(as_text=True).lstrip("﻿"))))
        self.assertEqual(len(rows), 2500)
        self.assertEqual(rows[0][0], "1")
        # The CSV import reads the table name from the file name up to the space.
        filename = response.headers["Content-Disposition"].split('filename="')[1].split('"')[0]
        self.assertEqual(importer.csv_table_name(filename), "T1")


if __name__ == "__main__":
    unittest.main()
