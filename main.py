#!/usr/bin/env python3
"""
Excel → MSSQL importer
Scans a sheet for tables named like MSHHNP, MZAIUP, MOKNOP, etc.
and bulk-inserts each into the matching SQL Server table.
"""

import re
import sys
import unicodedata
from getpass import getpass
from pathlib import Path

import pyodbc
import openpyxl

# A table name MUST be followed by a "(description)" — e.g. MZAIHP(倉庫在庫マスタ).
# This distinguishes table-name cells from header cells (DFRNKN, DFRGNO …),
# which are bare uppercase codes with no parentheses.
TABLE_NAME_RE = re.compile(r"^([A-Z][A-Z0-9]{3,9})\s*[\(（]")
HEADER_CODE_RE = re.compile(r"^[A-Z][A-Z0-9_]{1,30}$")

TARGET_SECTION = "実施前テストデータ"   # only import tables under this "■" section

# Invisible characters that can hide at the start of a cell (zero-width, BOM).
_INVISIBLE_RE = re.compile(r"[​‌‍﻿ ]")


def normalize_cell(value: str) -> str:
    """
    Normalize a cell string so full-width letters/digits/parens typed in
    Japanese Excel (e.g. 'ＭＤＮＮＯＰ（…)') become plain ASCII ('MDNNOP(…)'),
    and strip invisible/zero-width characters that break matching.
    """
    s = unicodedata.normalize("NFKC", value)
    s = _INVISIBLE_RE.sub("", s)
    return s.strip()


def is_one_line_marker_row(ws, row_num: int, start_col: int) -> bool:
    values = []
    for col in range(start_col, ws.max_column + 1):
        val = ws.cell(row=row_num, column=col).value
        if val is None:
            continue
        text = normalize_cell(str(val))
        if text:
            values.append(text)
    return len(values) == 1 and not HEADER_CODE_RE.fullmatch(values[0])


# ── Connection ────────────────────────────────────────────────────────────────

def prompt_connection() -> pyodbc.Connection:
    print("── SQL Server Connection ──")
    print("  (Docker DB is on port 11433 — enter: localhost,11433)")
    server   = input("  Server   (e.g. localhost,11433): ").strip()
    database = input("  Database : ").strip()
    username = input("  Username : ").strip()
    password = getpass("  Password : ")

    drivers = [d for d in pyodbc.drivers() if "SQL Server" in d]
    if not drivers:
        sys.exit("ERROR: No SQL Server ODBC driver found. Install 'ODBC Driver 17 for SQL Server'.")
    driver = sorted(drivers)[-1]   # pick the latest version available

    conn_str = (
        f"DRIVER={{{driver}}};"
        f"SERVER={server};"
        f"DATABASE={database};"
        f"UID={username};"
        f"PWD={password};"
        "TrustServerCertificate=yes;"
    )
    try:
        conn = pyodbc.connect(conn_str, timeout=10)
        print("  Connected.\n")
        return conn
    except pyodbc.Error as e:
        sys.exit(f"ERROR connecting: {e}")


# ── Sheet scanning ────────────────────────────────────────────────────────────

def scan_tables(ws: openpyxl.worksheet.worksheet.Worksheet) -> list[dict]:
    """
    Walk every row and look for cells whose value matches TABLE_NAME_RE.
    The row immediately below the table-name cell is treated as the header row.
    Returns a list of dicts: {name, header_row, start_col, name_row}
    """
    found = []
    seen_names = set()
    in_target = None   # None = no marker seen yet (scan everything)

    for row in ws.iter_rows():
        # A section marker is any cell beginning with "■"
        # (e.g. "■ 実施前テストデータ", "■ 予想結果", "■ バッチ").
        marker = next(
            (normalize_cell(c.value) for c in row
             if isinstance(c.value, str) and normalize_cell(c.value).startswith("■")),
            None,
        )
        if marker is not None:
            in_target = TARGET_SECTION in marker
            continue                 # skip the marker row itself

        if in_target is False:       # inside an excluded section → skip
            continue

        for cell in row:
            if not cell.value or not isinstance(cell.value, str):
                continue
            m = TABLE_NAME_RE.match(normalize_cell(cell.value))
            if not m:
                continue
            name = m.group(1)
            if name in seen_names:
                continue
            seen_names.add(name)
            found.append({
                "name":       name,
                "name_row":   cell.row,
                "start_col":  cell.column,
                "header_row": cell.row + 1,
            })

    return found


def extract_table(ws, meta: dict) -> tuple[list, list]:
    """Return (headers, rows) for one table block."""
    hrow      = meta["header_row"]
    start_col = meta["start_col"]
    if is_one_line_marker_row(ws, hrow, start_col):
        hrow += 1

    # Collect column headers (stop at first None)
    headers = []
    col = start_col
    while col <= ws.max_column:
        val = ws.cell(row=hrow, column=col).value
        if val is None:
            break
        # Normalize header codes (may be full-width) so they match DB columns.
        headers.append(normalize_cell(str(val)))
        col += 1

    if not headers:
        return [], []

    n_cols = len(headers)

    # Collect data rows (stop at first all-None row)
    rows = []
    r = hrow + 1
    while r <= ws.max_row:
        row_vals = [ws.cell(row=r, column=start_col + i).value for i in range(n_cols)]
        if all(v is None for v in row_vals):
            break
        rows.append(row_vals)
        r += 1

    return headers, rows


# ── SQL import ────────────────────────────────────────────────────────────────

PREFERRED_SCHEMA = "lvapdbf"   # tables live here, not the login default ('dbo')


def _db_columns(cursor: pyodbc.Cursor, table_name: str, schema: str) -> set[str]:
    cursor.execute(
        "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS "
        "WHERE TABLE_NAME = ? AND TABLE_SCHEMA = ? ORDER BY ORDINAL_POSITION",
        table_name, schema,
    )
    return {r[0].upper() for r in cursor.fetchall()}


def import_table(conn: pyodbc.Connection, table_name: str, headers: list, rows: list) -> None:
    cursor = conn.cursor()

    # Resolve the real schema (e.g. 'lvapdbf'), not the login default.
    cursor.execute(
        "SELECT TABLE_SCHEMA FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = ?",
        table_name,
    )
    schemas = [r[0] for r in cursor.fetchall()]
    if not schemas:
        print(f"  [SKIP]  '{table_name}' — table not found in database")
        return
    schema = PREFERRED_SCHEMA if PREFERRED_SCHEMA in schemas else schemas[0]
    qualified = f"[{schema}].[{table_name}]"

    # Match Excel headers → DB columns (case-insensitive)
    db_cols = _db_columns(cursor, table_name, schema)
    col_map = []   # list of (header_string, col_index_in_row)
    for i, h in enumerate(headers):
        if h.upper() in db_cols:
            col_map.append((h, i))

    if not col_map:
        print(f"  [SKIP]  '{table_name}' — no matching columns between sheet and table")
        return

    valid_headers = [c[0] for c in col_map]
    col_indices   = [c[1] for c in col_map]

    # Truncate existing rows then bulk-insert
    cursor.execute(f"DELETE FROM {qualified}")

    cols_sql     = ", ".join(f"[{h}]" for h in valid_headers)
    placeholders = ", ".join("?" for _ in valid_headers)
    insert_sql   = f"INSERT INTO {qualified} ({cols_sql}) VALUES ({placeholders})"

    inserted = 0
    errors   = 0
    for row in rows:
        try:
            cursor.execute(insert_sql, [row[i] for i in col_indices])
            inserted += 1
        except pyodbc.Error as e:
            errors += 1
            if errors <= 3:
                print(f"    [WARN] {e}")

    conn.commit()
    status = f"{inserted} rows inserted"
    if errors:
        status += f", {errors} errors"
    if len(valid_headers) < len(headers):
        status += f" ({len(headers) - len(valid_headers)} sheet cols skipped — not in DB)"
    print(f"  [OK]    '{table_name}' — {status}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main() -> None:
    # Excel file
    raw_path = input("Excel file path: ").strip().strip('"').strip("'")
    file_path = Path(raw_path)
    if not file_path.exists():
        sys.exit(f"ERROR: File not found — {file_path}")

    try:
        wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
    except Exception as e:
        sys.exit(f"ERROR loading workbook: {e}")

    print(f"\nSheets: {', '.join(wb.sheetnames)}")
    sheet_name = input("Sheet name: ").strip()
    if sheet_name not in wb.sheetnames:
        sys.exit(f"ERROR: Sheet '{sheet_name}' not found.")

    ws = wb[sheet_name]

    print(f"\nScanning '{sheet_name}' …")
    table_metas = scan_tables(ws)

    if not table_metas:
        sys.exit("No tables found. Make sure table names (e.g. MSHHNP) appear in the sheet.")

    # Extract all tables first so we can show a full preview
    extracted = []
    for meta in table_metas:
        headers, rows = extract_table(ws, meta)
        extracted.append((meta, headers, rows))

    # ── Preview ───────────────────────────────────────────────────────────────
    print(f"\n{'─'*52}")
    print(f"  Found {len(extracted)} table(s) in sheet '{sheet_name}':")
    print(f"{'─'*52}")
    print(f"  {'#':<4}  {'Table Name':<16}  {'Columns':>7}  {'Records':>8}")
    print(f"  {'─'*4}  {'─'*16}  {'─'*7}  {'─'*8}")
    for i, (meta, headers, rows) in enumerate(extracted, 1):
        col_count = len(headers) if headers else 0
        row_count = len(rows)
        print(f"  {i:<4}  {meta['name']:<16}  {col_count:>7}  {row_count:>8}")
    print(f"{'─'*52}")
    total_records = sum(len(rows) for _, _, rows in extracted)
    print(f"  Total: {total_records} records across {len(extracted)} tables")
    print(f"{'─'*52}\n")

    answer = input("Proceed with import? [y/N]: ").strip().lower()
    if answer not in ("y", "yes"):
        print("Import cancelled.")
        sys.exit(0)

    # ── Connect & import ──────────────────────────────────────────────────────
    print()
    conn = prompt_connection()

    print("Importing …")
    for meta, headers, rows in extracted:
        if not headers:
            print(f"  [SKIP]  '{meta['name']}' — header row is empty")
            continue
        import_table(conn, meta["name"], headers, rows)

    conn.close()
    print("\nAll done.")


if __name__ == "__main__":
    main()
