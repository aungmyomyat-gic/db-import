import base64
import csv
import io
import re
import os
import json
import shutil
import subprocess
import threading
import time
import unicodedata
import tempfile
import urllib.error
import urllib.parse
import urllib.request
import uuid
from datetime import date, datetime, time as datetime_time
from decimal import Decimal, InvalidOperation
from pathlib import Path

from flask import Flask, render_template, request, jsonify, session, send_file
import openpyxl
from openpyxl.utils import get_column_letter
import pyodbc

app = Flask(__name__)
app.secret_key = os.urandom(24)

# Persisted DB connection (port, username, password, database, host).
# Stored next to the app so it survives restarts; mount it as a volume to
# keep it across container rebuilds.
CONFIG_PATH = Path(os.environ.get("DB_CONFIG_PATH", "/app/data/db_config.json"))


def load_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_config(cfg: dict) -> None:
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")

# A table name MUST be followed by a "(description)" — e.g. MZAIHP(倉庫在庫マスタ).
# This distinguishes table-name cells from header cells (DFRNKN, DFRGNO …),
# which are bare uppercase codes with no parentheses.
TABLE_NAME_RE = re.compile(r"^([A-Z][A-Z0-9]{3,9})\s*[\(（]")
HEADER_CODE_RE = re.compile(r"^[A-Z][A-Z0-9_]{1,30}$")

# Invisible characters that can hide at the start of a cell (zero-width, BOM).
_INVISIBLE_RE = re.compile(r"[​‌‍﻿ ]")


def normalize_cell(value: str) -> str:
    """
    Normalize a cell string so full-width letters/digits/parens typed in
    Japanese Excel (e.g. 'ＭＤＮＮＯＰ（…)') become plain ASCII ('MDNNOP(…)'),
    and strip invisible/zero-width characters that break matching.
    """
    s = unicodedata.normalize("NFKC", value)
    s = _INVISIBLE_RE.sub("", s)
    return s.strip()


# ── Excel helpers ─────────────────────────────────────────────────────────────

TARGET_SECTION = "実施前テストデータ"   # only import tables under this "■" section
MARKER_SCAN_WIDTH = 8

# ── Batch job trigger (derived from the uploaded file name) ────────────────────
BATCH_BASE_URL = os.environ.get("BATCH_BASE_URL", "http://localhost:8080")
BATCH_NAME_RE = re.compile(r"(day|month)[\s_-]*job[\s_-]*(\d+)", re.IGNORECASE)


def derive_batch_job(filename):
    """
    Map an uploaded file name to the batch endpoint it should trigger.
    'day_job_20.xlsx' / 'DayJob20.xlsx'   -> day-job20   / "Call DayJob20"
    'Month Job-1.xlsx' / 'monthjob1.xlsx' -> month-job1  / "Call Month Job-1"
    """
    if not filename:
        return None
    m = BATCH_NAME_RE.search(filename)
    if not m:
        return None
    kind, num = m.group(1).lower(), m.group(2)
    slug = f"{kind}-job{num}"
    label = f"Call DayJob{num}" if kind == "day" else f"Call Month Job-{num}"
    return {"slug": slug, "label": label}
IMPORT_BATCH_SIZE = int(os.environ.get("IMPORT_BATCH_SIZE", "5000"))
IMPORT_JOB_TTL_SECONDS = 3600
IMPORT_JOBS = {}
IMPORT_JOBS_LOCK = threading.Lock()
# Export files live only for this process and follow the import-job expiry.
SQL_EXPORT_DIR = tempfile.TemporaryDirectory(prefix="db-import-sql-")


def is_one_line_marker_row(ws, row_num, start_col):
    values = []
    end_col = min(ws.max_column, start_col + MARKER_SCAN_WIDTH - 1)
    for col in range(start_col, end_col + 1):
        val = ws.cell(row=row_num, column=col).value
        if val is None:
            continue
        text = normalize_cell(str(val))
        if text:
            values.append(text)
            if len(values) > 1:
                return False
    return len(values) == 1 and not HEADER_CODE_RE.fullmatch(values[0])


def resolve_header_row(ws, meta):
    hrow = int(meta["header_row"])
    start_col = int(meta["start_col"])
    marker_skipped = is_one_line_marker_row(ws, hrow, start_col)
    return hrow + 1 if marker_skipped else hrow, marker_skipped


def scan_tables(ws):
    """
    Collect table blocks ONLY while inside the "■ 実施前テストデータ" section.
    Any other "■" section ("予想結果", "バッチ", …) is skipped entirely.
    If the sheet has no "■" markers at all, the whole sheet is scanned.
    """
    found, seen = [], set()
    in_target = None   # None = no marker seen yet (scan everything)
    seen_target = False
    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        # A section marker is any cell beginning with "■"
        marker = next(
            (normalize_cell(value) for value in row
             if isinstance(value, str) and normalize_cell(value).startswith("■")),
            None,
        )
        if marker is not None:
            is_target = TARGET_SECTION in marker
            if seen_target and not is_target:
                break
            in_target = is_target
            seen_target = seen_target or is_target
            continue                 # skip the marker row itself

        if in_target is False:       # inside an excluded section → skip
            continue

        for col_idx, value in enumerate(row, start=1):
            if not value or not isinstance(value, str):
                continue
            m = TABLE_NAME_RE.match(normalize_cell(value))
            if not m:
                continue
            name = m.group(1)
            if name in seen:
                continue
            seen.add(name)
            found.append({
                "name": name,
                "name_row": row_idx,
                "start_col": col_idx,
                "header_row": row_idx + 1,
            })
    return found


def _row_value(row, col_num):
    idx = col_num - 1
    return row[idx] if idx < len(row) else None


def _is_one_line_marker_values(row, start_col):
    values = []
    end_col = start_col + MARKER_SCAN_WIDTH - 1
    for col in range(start_col, end_col + 1):
        val = _row_value(row, col)
        if val is None:
            continue
        text = normalize_cell(str(val))
        if text:
            values.append(text)
            if len(values) > 1:
                return False
    return len(values) == 1 and not HEADER_CODE_RE.fullmatch(values[0])


def _headers_from_values(row, start_col):
    headers, col = [], start_col
    while col <= len(row):
        val = _row_value(row, col)
        if val is None:
            break
        headers.append(normalize_cell(str(val)))
        col += 1
    return headers


def _preview_from_headers(meta, hrow, headers, marker_skipped):
    start_col = int(meta["start_col"])
    data_start_row = hrow + 1
    end_col = start_col + len(headers) - 1 if headers else start_col
    start_letter = get_column_letter(start_col)
    end_letter = get_column_letter(end_col)
    return {
        "name": meta["name"],
        "columns": len(headers),
        "records": 0,
        "name_row": int(meta["name_row"]),
        "header_row": hrow,
        "data_start_row": data_start_row,
        "data_end_row": None,
        "start_col": start_col,
        "end_col": end_col,
        "start_col_letter": start_letter,
        "end_col_letter": end_letter,
        "col_range": f"{start_letter}:{end_letter}",
        "row_range": "-",
        "cell_range": f"{start_letter}{hrow}:{end_letter}...",
        "marker_skipped": marker_skipped,
    }


def _count_active_preview_rows(active_tables, row, row_idx):
    still_active = []
    for table in active_tables:
        n_cols = table["columns"]
        start_col = table["start_col"]
        row_vals = [_row_value(row, start_col + i) for i in range(n_cols)]
        if all(v is None for v in row_vals):
            continue
        table["records"] += 1
        table["data_end_row"] = row_idx
        table["row_range"] = f"{table['data_start_row']}:{row_idx}"
        still_active.append(table)
    return still_active


def scan_table_previews(ws):
    """
    Stream the sheet once and build preview rows with real record counts.
    This avoids slow random access on read-only worksheets during preview.
    """
    tables, active, waiting = [], [], []
    seen = set()
    in_target = None
    seen_target = False

    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        active = _count_active_preview_rows(active, row, row_idx)

        next_waiting = []
        for item in waiting:
            meta = item["meta"]
            start_col = int(meta["start_col"])
            if row_idx == int(meta["header_row"]) and _is_one_line_marker_values(row, start_col):
                item["meta"] = {**meta, "header_row": row_idx + 1}
                item["marker_skipped"] = True
                next_waiting.append(item)
                continue

            headers = _headers_from_values(row, start_col)
            table = _preview_from_headers(meta, row_idx, headers, item["marker_skipped"])
            tables.append(table)
            if headers:
                active.append(table)
        waiting = next_waiting

        marker = next(
            (normalize_cell(value) for value in row
             if isinstance(value, str) and normalize_cell(value).startswith("■")),
            None,
        )
        if marker is not None:
            is_target = TARGET_SECTION in marker
            if seen_target and not is_target:
                break
            in_target = is_target
            seen_target = seen_target or is_target
            continue

        if in_target is False:
            continue

        for col_idx, value in enumerate(row, start=1):
            if not value or not isinstance(value, str):
                continue
            m = TABLE_NAME_RE.match(normalize_cell(value))
            if not m:
                continue
            name = m.group(1)
            if name in seen:
                continue
            seen.add(name)
            waiting.append({
                "meta": {
                    "name": name,
                    "name_row": row_idx,
                    "start_col": col_idx,
                    "header_row": row_idx + 1,
                },
                "marker_skipped": False,
            })

    return tables


def extract_headers(ws, meta):
    hrow, _ = resolve_header_row(ws, meta)
    start_col = int(meta["start_col"])

    headers, col = [], start_col
    while col <= ws.max_column:
        val = ws.cell(row=hrow, column=col).value
        if val is None:
            break
        # Normalize header codes (may be full-width) so they match DB columns.
        headers.append(normalize_cell(str(val)))
        col += 1
    end_col = start_col + len(headers) - 1 if headers else start_col
    return hrow, headers, end_col


def extract_table(ws, meta):
    hrow, headers, _ = extract_headers(ws, meta)
    start_col = int(meta["start_col"])
    if not headers:
        return [], []
    n_cols = len(headers)
    end_col = start_col + n_cols - 1
    rows = []
    # Sequential iter_rows() instead of per-cell ws.cell() lookups: on a
    # read_only workbook, random-access ws.cell() re-walks the sheet's XML
    # from the top on every call, making per-cell extraction O(rows^2).
    for row_vals in ws.iter_rows(
        min_row=hrow + 1, max_row=ws.max_row,
        min_col=start_col, max_col=end_col,
        values_only=True,
    ):
        if all(v is None for v in row_vals):
            break
        rows.append(row_vals)
    return headers, rows


# ── CSV helpers ───────────────────────────────────────────────────────────────
# A CSV upload holds ONE table: the table name comes from the file name
# (MDNNOP.csv, "MDNNOP (1).csv" -> MDNNOP). If the first row holds the table's
# column names it is a header row and columns are matched by name; otherwise
# every row is data and columns are matched to the DB table's columns by position.

CSV_TABLE_NAME_RE = re.compile(r"[A-Z][A-Z0-9_]*")
CSV_ENCODINGS = ("utf-8-sig", "cp932")   # UTF-8 (with/without BOM), then Shift-JIS


def is_csv_filename(filename):
    return bool(filename) and filename.lower().endswith(".csv")


def csv_table_name(filename):
    stem = normalize_cell(Path(filename or "").stem).upper()
    m = CSV_TABLE_NAME_RE.match(stem)
    return m.group(0) if m else None


def is_tsv_filename(filename):
    return bool(filename) and filename.lower().endswith(".tsv")


def parse_delimited(raw, delimiter=","):
    """Decode CSV/TSV bytes; return [(line_number, cells)] for every non-blank row.

    Quoting works the same for both: a value in double quotes may hold the
    separator ("a,b" in a CSV, "a<TAB>b" in a TSV) and the quotes are removed.
    """
    for encoding in CSV_ENCODINGS:
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError("File is not UTF-8 or Shift-JIS text")
    reader = csv.reader(io.StringIO(text, newline=""), delimiter=delimiter)
    return [(reader.line_num, row) for row in reader if any(cell.strip() for cell in row)]


def read_csv_rows(path, delimiter=","):
    """Read every non-blank row; short rows are padded so all rows have the same width."""
    rows = [row for _, row in parse_delimited(Path(path).read_bytes(), delimiter)]
    width = max((len(row) for row in rows), default=0)
    return [row + [""] * (width - len(row)) for row in rows], width


def csv_has_header(first_row, column_names):
    """True when most non-blank cells of first_row are column names of the table."""
    names = {str(name).upper() for name in column_names}
    cells = [normalize_cell(cell).upper() for cell in first_row if cell.strip()]
    return bool(cells) and sum(cell in names for cell in cells) * 2 > len(cells)


def _csv_table_column_names(table_name):
    """Column names of table_name in any schema, or None when the DB can't be reached."""
    cfg = load_config()
    if not cfg:
        return None
    conn_str, err = _build_conn_str(
        cfg["host"], cfg["port"], cfg["database"], cfg["username"], cfg["password"]
    )
    if err:
        return None
    try:
        conn = pyodbc.connect(conn_str, timeout=5)
        try:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT DISTINCT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = ?",
                table_name,
            )
            return [row[0] for row in cursor.fetchall()]
        finally:
            conn.close()
    except pyodbc.Error:
        return None


def _csv_preview(table_name, rows, width, has_header=False):
    """has_header: True / False, or None when it could not be checked (no DB)."""
    records = len(rows) - 1 if has_header and rows else len(rows)
    first = 2 if has_header else 1
    return {
        "name": table_name,
        "columns": width,
        "records": records,
        "header": has_header,
        "col_range": "Header row" if has_header else "No header" if has_header is False
                     else "Checked at import",
        "row_range": f"{first}:{len(rows)}" if records else "-",
    }


# ── TEMP: CSV check against table columns ────────────────────────────────────
# Temporary debugging aid: compares each CSV value with the column it will be
# imported into (matched by position, like the CSV import) using the column
# types, lengths and precision from INFORMATION_SCHEMA. Used by the "CSV Check"
# tab and printed to the container log on every CSV import. Remove this block,
# the /temp/csv-check route and the call in _run_csv_import.
_TEMP_INT_RANGES = {
    "tinyint": (0, 255),
    "smallint": (-2**15, 2**15 - 1),
    "int": (-2**31, 2**31 - 1),
    "bigint": (-2**63, 2**63 - 1),
}
_TEMP_INT_RE = re.compile(r"[+-]?[0-9]+")
_TEMP_DECIMAL_RE = re.compile(r"[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)")
_TEMP_FLOAT_RE = re.compile(r"[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?")


def _temp_table_columns(cursor, schema, table_name):
    cursor.execute(
        "SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, NUMERIC_PRECISION, "
        "NUMERIC_SCALE, IS_NULLABLE FROM INFORMATION_SCHEMA.COLUMNS "
        "WHERE TABLE_SCHEMA = ? AND TABLE_NAME = ? ORDER BY ORDINAL_POSITION",
        schema, table_name,
    )
    columns = []
    for name, dtype, max_len, precision, scale, nullable in cursor.fetchall():
        dtype = (dtype or "").lower()
        if dtype in ("decimal", "numeric"):
            label = f"{dtype}({precision},{scale})"
        elif dtype in STRING_TYPES and dtype not in ("text", "ntext"):
            label = f"{dtype}({'max' if max_len == -1 else max_len})"
        else:
            label = dtype
        columns.append({"name": name, "type": dtype, "label": label,
                        "max_len": max_len, "precision": precision, "scale": scale,
                        "nullable": nullable == "YES"})
    return columns


def _temp_value_problem(value, col):
    """Return (severity, reason) when value does not fit col, else None."""
    dtype = col["type"]
    if dtype not in STRING_TYPES:
        value = value.strip()
    if value == "":
        return None   # blank -> NULL, or the type default for NOT NULL columns
    if dtype in _TEMP_INT_RANGES:
        if not _TEMP_INT_RE.fullmatch(value):
            return "error", "Not an integer"
        low, high = _TEMP_INT_RANGES[dtype]
        if not low <= int(value) <= high:
            return "error", f"Out of range for {dtype} ({low} to {high})"
    elif dtype == "bit":
        if value.lower() not in ("0", "1", "true", "false"):
            return "error", "Not 0/1/true/false"
    elif dtype in ("decimal", "numeric"):
        if not _TEMP_DECIMAL_RE.fullmatch(value):
            return "error", "Not a number"
        whole, _, frac = value.lstrip("+-").partition(".")
        whole = whole.lstrip("0")
        precision, scale = col["precision"] or 18, col["scale"] or 0
        if len(whole) > precision - scale:
            return "error", (f"Too many digits: {len(whole)} before the decimal point, "
                             f"{col['label']} allows {precision - scale}")
        if len(frac.rstrip("0")) > scale:
            return "warning", f"{len(frac.rstrip('0'))} decimals, {col['label']} keeps {scale} (rounded)"
    elif dtype in ("money", "smallmoney", "float", "real"):
        if not _TEMP_FLOAT_RE.fullmatch(value):
            return "error", "Not a number"
    elif dtype in ("char", "varchar", "nchar", "nvarchar") and col["max_len"] not in (None, -1):
        if dtype in ("nchar", "nvarchar"):
            size, unit = len(value.encode("utf-16-le")) // 2, "characters"
        else:
            size, unit = len(value.encode("cp932", errors="replace")), "bytes (Shift-JIS)"
        if size > col["max_len"]:
            return "error", f"Too long: {size} {unit}, {col['label']} allows {col['max_len']}"
    return None


def _temp_csv_issues(raw, columns):
    """Return (rows_checked, issues) for CSV bytes checked against table columns."""
    for encoding in CSV_ENCODINGS:
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError("CSV is not UTF-8 or Shift-JIS text")
    issues = []
    rows_checked = 0
    header = None   # set when the first row is a header: columns are matched by name
    by_name = {c["name"].upper(): c for c in columns}
    positional = columns
    reader = csv.reader(io.StringIO(text, newline=""))
    for row in reader:
        if not any(cell.strip() for cell in row):
            continue   # the import skips blank lines too
        if header is None and rows_checked == 0 and csv_has_header(row, list(by_name)):
            header = [normalize_cell(cell) for cell in row]
            positional = [by_name.get(h.upper()) for h in header]
            continue
        rows_checked += 1
        base = {"row": rows_checked, "line": reader.line_num,
                "record_id": row[0] if row else "", "full_row": row}
        expected = len(header) if header else len(columns)
        if len(row) != expected:
            more = len(row) > expected
            what = "header has" if header else "table has"
            issues.append({**base, "column": None, "type": None, "position": None, "value": None,
                           "severity": "warning",
                           "reason": f"{len(row)} columns, {what} {expected}"
                                     + (" (extra columns are skipped)" if more
                                        else " (missing columns become NULL/default)")})
        for index, (value, col) in enumerate(zip(row, positional)):
            if col is None:
                continue   # header column that is not in the table: the import skips it
            problem = _temp_value_problem(value, col)
            if problem:
                issues.append({**base, "column": col["name"], "type": col["label"],
                               "position": index + 1, "value": value,
                               "severity": problem[0], "reason": problem[1]})
    return rows_checked, issues, header


def _temp_log_csv_check(conn, path, schema, table_name):
    columns = _temp_table_columns(conn.cursor(), schema, table_name)
    if not columns:
        return
    _, issues, _ = _temp_csv_issues(Path(path).read_bytes(), columns)
    print(f"[TEMP CSV CHECK] {schema}.{table_name}: {len(issues)} problems", flush=True)
    for i in issues:
        if i["column"] is None:
            print(f"Row {i['row']} (CSV line {i['line']}): {i['reason']}", flush=True)
            continue
        print(f"{i['severity'].upper()} ROW: {i['row']}\n"
              f"CSV LINE: {i['line']}\n"
              f"Record ID: {i['record_id']}\n"
              f"Column: {i['column']} {i['type']}\n"
              f"Column position: {i['position']}\n"
              f"Invalid value: {i['value']!r}\n"
              f"Problem: {i['reason']}\n"
              f"Full row: {i['full_row']}\n" + "-" * 60, flush=True)
    print(f"[TEMP CSV CHECK] {schema}.{table_name}: done", flush=True)
# ── END TEMP ──────────────────────────────────────────────────────────────────


# ── Comma CSV → TSV for tables with file-name columns ─────────────────────────
# Some systems write CSV without quoting, so a file name such as a,b,c.pdf
# becomes three cells. For the tables below, the listed columns hold file
# names. A line with exactly one cell per table column (after dropping empty
# cells at the end of the line) is used as is. Otherwise the cells are taken
# column by column; in a file-name column, a cell that is neither blank nor
# ending in one of FILE_EXTENSIONS is joined with the next cell until a cell
# with an extension is reached. The result is written as TSV, so only
# the commas inside those file names remain.
COMMA_FILE_COLUMNS = {
    "MSHHKP": ("SHKTRM", "SHKFL1", "SHKFL2", "SHKFL3"),
}
# Only these count (any letter case), so text such as "JN75 0.35GR" or "2014.11"
# is not taken for the end of a file name. They are every extension found in
# the MSHHKP export; a file with another extension shows up as an error row.
FILE_EXTENSIONS = ("pdf", "docx", "xlsx", "jpg", "xls", "doc", "png", "txt", "tif", "jpeg",
                   "pptx", "gif", "ppt")
FILE_EXTENSION_RE = re.compile(r"\.(?:" + "|".join(FILE_EXTENSIONS) + r")$", re.IGNORECASE)


def _ends_file_name(cell):
    cell = cell.strip()
    return cell == "" or bool(FILE_EXTENSION_RE.search(cell))


def _split_comma_row(cells, column_names, file_positions):
    """Return (values, joined, None) with one value per column, or (None, False, reason)."""
    width = len(column_names)
    trimmed = list(cells)
    while len(trimmed) > width and trimmed[-1] == "":
        trimmed.pop()   # empty cells after the last column (",,,," at the end of the line)
    if len(trimmed) == width:
        return trimmed, False, None
    values, i, joined = [], 0, False
    for position, name in enumerate(column_names):
        if i >= len(cells):
            return None, False, (f"{len(cells)} cells: ran out at column {position + 1} {name}, "
                                 f"table has {len(column_names)} columns")
        start = i
        if position in file_positions:
            while not _ends_file_name(cells[i]):
                i += 1
                if i == len(cells):
                    return None, False, (f"{name}: no file extension found in "
                                         f"{','.join(cells[start:])!r}")
        values.append(",".join(cells[start:i + 1]))
        joined = joined or i > start
        i += 1
    if any(cell != "" for cell in cells[i:]):
        return None, False, (f"{len(cells)} cells: {len(cells) - i} left over after the last "
                             f"column, table has {len(column_names)} columns")
    return values, joined, None


def _tsv_line(values):
    """One TSV line. A value holding a tab or a double quote is written in double
    quotes ("a<TAB>b", a quote inside doubled), like CSV quoting, so it stays one
    value; TSV import removes the quotes again."""
    buffer = io.StringIO()
    csv.writer(buffer, delimiter="\t", lineterminator="").writerow(values)
    return buffer.getvalue()


def comma_csv_to_tsv(raw, column_names, file_columns):
    """Convert unquoted comma CSV bytes to TSV bytes in the same encoding.

    Rows that do not come out to exactly one value per column are left out and
    listed in report["errors"] with their line number, for fixing by hand.
    report["quoted_lines"] lists the lines with a value written in quotes
    because it contains a tab.
    """
    for encoding in CSV_ENCODINGS:
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError("CSV is not UTF-8 or Shift-JIS text")
    if encoding == "utf-8-sig" and not raw.startswith(b"\xef\xbb\xbf"):
        encoding = "utf-8"   # keep "no BOM" when the input had none
    wanted = {name.upper() for name in file_columns}
    file_positions = {i for i, name in enumerate(column_names) if name.upper() in wanted}
    newline = "\r\n" if "\r\n" in text else "\n"
    out, errors = [], []
    quoted_lines = []
    report = {"header": False, "rows": 0, "written": 0, "joined": 0, "errors": errors,
              "quoted_lines": quoted_lines,
              "encoding": "Shift-JIS" if encoding == "cp932" else "UTF-8"}
    first = True
    for line_number, line in enumerate(text.split("\n"), 1):
        line = line.rstrip("\r")
        if not line.strip():
            continue
        cells = line.split(",")
        if first:
            first = False
            if csv_has_header(cells, column_names):
                report["header"] = True
                out.append(_tsv_line(cells))
                continue
        report["rows"] += 1
        values, joined, reason = _split_comma_row(cells, column_names, file_positions)
        if values is None:
            errors.append({"line": line_number, "reason": reason, "text": line})
            continue
        out.append(_tsv_line(values))
        if any("\t" in value for value in values):
            quoted_lines.append(line_number)
        report["written"] += 1
        report["joined"] += joined
    tsv = (newline.join(out) + newline) if out else ""
    return tsv.encode(encoding, errors="strict"), report


class RouteError(Exception):
    """A request problem to answer with {"error": message} and an HTTP status."""

    def __init__(self, message, status=400):
        super().__init__(message)
        self.message = message
        self.status = status


def _saved_config():
    cfg = load_config()
    if not cfg:
        raise RouteError("No DB connection saved. Please set it up first.")
    return cfg


def _db_table_columns(cfg, schema, table_name):
    """Column names of schema.table_name in table order."""
    if not schema:
        raise RouteError("Schema is required.")
    conn_str, err = _build_conn_str(
        cfg["host"], cfg["port"], cfg["database"], cfg["username"], cfg["password"]
    )
    if err:
        raise RouteError(err, 500)
    try:
        conn = pyodbc.connect(conn_str, timeout=10)
        try:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS "
                "WHERE TABLE_SCHEMA = ? AND TABLE_NAME = ? ORDER BY ORDINAL_POSITION",
                schema, table_name,
            )
            names = [row[0] for row in cursor.fetchall()]
        finally:
            conn.close()
    except pyodbc.Error as e:
        raise RouteError(f"Could not read table columns: {e}")
    if not names:
        raise RouteError(f"Table {schema}.{table_name} not found.", 404)
    return names


@app.route("/convert/comma-csv-to-tsv", methods=["POST"])
def convert_comma_csv_to_tsv():
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify({"error": "No file selected."}), 400
    if not is_csv_filename(file.filename):
        return jsonify({"error": "Please choose a .csv file."}), 400
    table_name = csv_table_name(file.filename)
    if not table_name:
        return jsonify({"error": "CSV file name must start with the table name, e.g. MSHHKP.csv."}), 400
    file_columns = COMMA_FILE_COLUMNS.get(table_name)
    if not file_columns:
        return jsonify({"error": f"No file-name column rule for table {table_name}. "
                                 f"Supported: {', '.join(sorted(COMMA_FILE_COLUMNS))}."}), 400
    schema = request.form.get("schema", "").strip()
    try:
        column_names = _db_table_columns(_saved_config(), schema, table_name)
    except RouteError as e:
        return jsonify({"error": e.message}), e.status
    missing = [name for name in file_columns
               if name.upper() not in {c.upper() for c in column_names}]
    if missing:
        return jsonify({"error": f"{schema}.{table_name} has no column {', '.join(missing)}."}), 400
    try:
        tsv, report = comma_csv_to_tsv(file.read(), column_names, file_columns)
    except (ValueError, UnicodeError) as e:
        return jsonify({"error": f"Cannot convert file: {e}"}), 400
    return jsonify({
        **report,
        "errors": report["errors"][:2000],
        "error_count": len(report["errors"]),
        "quoted_count": len(report["quoted_lines"]),
        "quoted_lines": report["quoted_lines"][:200],
        "filename": file.filename, "schema": schema, "table": table_name,
        "columns": len(column_names), "file_columns": list(file_columns),
        "tsv_filename": Path(file.filename).stem + ".tsv",
        "tsv_base64": base64.b64encode(tsv).decode("ascii"),
    })


# ── TSV import ────────────────────────────────────────────────────────────────
# Same import as CSV (file name = table name, optional header row, rows replace
# the table's rows, SQL download afterwards), but the values are separated by
# tabs, e.g. the file made by the CSV → TSV tab.

def _tsv_upload():
    """Validate a TSV upload form; return (file, table_name, schema)."""
    file = request.files.get("file")
    if not file or not file.filename:
        raise RouteError("No file selected.")
    if not is_tsv_filename(file.filename):
        raise RouteError("Please choose a .tsv file.")
    table_name = csv_table_name(file.filename)
    if not table_name:
        raise RouteError("TSV file name must start with the table name, e.g. MSHHKP.tsv.")
    return file, table_name, request.form.get("schema", "").strip()


@app.route("/tsv/preview", methods=["POST"])
def tsv_preview():
    """Check a TSV against its table before importing it. Nothing is imported."""
    try:
        file, table_name, schema = _tsv_upload()
        columns = _db_table_columns(_saved_config(), schema, table_name)
        rows = parse_delimited(file.read(), "\t")
    except RouteError as e:
        return jsonify({"error": e.message}), e.status
    except ValueError as e:
        return jsonify({"error": f"Cannot read file: {e}"}), 400
    header = bool(rows) and csv_has_header(rows[0][1], columns)
    data = rows[1:] if header else rows
    expected = len(rows[0][1]) if header else len(columns)
    mismatched = [{"line": line, "cells": len(cells), "first": cells[0] if cells else ""}
                  for line, cells in data if len(cells) != expected]
    known = {c.upper() for c in columns}
    return jsonify({
        "filename": file.filename, "schema": schema, "table": table_name,
        "columns": len(columns), "header": header, "expected": expected,
        "records": len(data),
        "unknown_header_columns": [normalize_cell(c) for c in rows[0][1]
                                   if header and c.strip() and normalize_cell(c).upper() not in known],
        "mismatched_count": len(mismatched),
        "mismatched": mismatched[:200],
    })


@app.route("/tsv/import", methods=["POST"])
def tsv_import():
    """Start a background import of one TSV file; poll /import/status/<job_id>."""
    try:
        file, table_name, schema = _tsv_upload()
        cfg = _saved_config()
        if not schema:
            raise RouteError("Schema is required.")
    except RouteError as e:
        return jsonify({"error": e.message}), e.status
    tmp = Path(tempfile.gettempdir()) / f"{uuid.uuid4()}.tsv"
    file.save(tmp)
    try:
        parse_delimited(tmp.read_bytes(), "\t")
    except ValueError as e:
        tmp.unlink(missing_ok=True)
        return jsonify({"error": f"Cannot read file: {e}"}), 400

    _cleanup_import_jobs()
    job_id = uuid.uuid4().hex
    job = _new_import_job(job_id, [table_name])
    job["owner"] = session.setdefault("import_owner", uuid.uuid4().hex)
    with IMPORT_JOBS_LOCK:
        IMPORT_JOBS[job_id] = job
        snapshot = _job_snapshot(job)
    threading.Thread(target=_run_tsv_import_job, args=(job_id, cfg, str(tmp), table_name, schema),
                     daemon=True).start()
    return jsonify(snapshot), 202


def _run_tsv_import_job(job_id, cfg, tmp_path, table_name, schema):
    try:
        _run_import_job(job_id, cfg, tmp_path, None, [table_name], schema, {}, table_name, "\t")
    finally:
        Path(tmp_path).unlink(missing_ok=True)


def _csv_value(value, dtype):
    """CSV text -> insert value. Blank becomes NULL (NOT NULL columns get a default later)."""
    if dtype not in STRING_TYPES:
        value = value.strip()
    return None if value == "" else value


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.after_request
def disable_browser_cache(response):
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


@app.route("/config", methods=["GET"])
def get_config():
    """Return the saved DB connection so the UI can pre-fill / skip setup."""
    cfg = load_config()
    return jsonify({"configured": bool(cfg), "config": cfg})


@app.route("/config", methods=["POST"])
def set_config():
    """Validate the connection, then persist it to JSON for later reuse."""
    data = request.json or {}
    host     = data.get("host", "").strip()
    port     = str(data.get("port", "")).strip()
    database = data.get("database", "").strip()
    username = data.get("username", "").strip()
    password = data.get("password", "")
    schema   = data.get("schema", "").strip() or "lvapdbf"

    if not all([host, port, database, username]):
        return jsonify({"error": "Host, port, database, and username are required."}), 400

    # Test the connection before saving so we never persist bad credentials.
    err = _test_connection(host, port, database, username, password)
    if err:
        return jsonify({"error": f"Connection failed: {err}"}), 400

    save_config({
        "host": host, "port": port, "database": database,
        "username": username, "password": password, "schema": schema,
    })
    return jsonify({"ok": True})


@app.route("/config", methods=["DELETE"])
def delete_config():
    """Remove the saved DB connection so the user can set up a new one."""
    if CONFIG_PATH.exists():
        CONFIG_PATH.unlink()
    return jsonify({"ok": True})


def _build_conn_str(host, port, database, username, password):
    drivers = [d for d in pyodbc.drivers() if "SQL Server" in d]
    if not drivers:
        return None, "No SQL Server ODBC driver found inside container."
    driver = sorted(drivers)[-1]
    conn_str = (
        f"DRIVER={{{driver}}};"
        f"SERVER={host},{port};"
        f"DATABASE={database};"
        f"UID={username};"
        f"PWD={password};"
        "TrustServerCertificate=yes;"
    )
    return conn_str, None


def _test_connection(host, port, database, username, password):
    conn_str, err = _build_conn_str(host, port, database, username, password)
    if err:
        return err
    try:
        pyodbc.connect(conn_str, timeout=10).close()
        return None
    except pyodbc.Error as e:
        return str(e)


@app.route("/temp/csv-check", methods=["POST"])   # TEMP: remove with the TEMP CSV check block
def temp_csv_check():
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify({"error": "No file selected."}), 400
    if not is_csv_filename(file.filename):
        return jsonify({"error": "Please choose a .csv file."}), 400
    table_name = csv_table_name(file.filename)
    if not table_name:
        return jsonify({"error": "CSV file name must start with the table name, e.g. MDNNOP.csv."}), 400
    schema = request.form.get("schema", "").strip()
    if not schema:
        return jsonify({"error": "Schema is required."}), 400
    cfg = load_config()
    if not cfg:
        return jsonify({"error": "No DB connection saved."}), 400
    conn_str, err = _build_conn_str(
        cfg["host"], cfg["port"], cfg["database"], cfg["username"], cfg["password"]
    )
    if err:
        return jsonify({"error": err}), 500
    try:
        conn = pyodbc.connect(conn_str, timeout=10)
        try:
            columns = _temp_table_columns(conn.cursor(), schema, table_name)
        finally:
            conn.close()
    except pyodbc.Error as e:
        return jsonify({"error": f"Could not read table columns: {e}"}), 400
    if not columns:
        return jsonify({"error": f"Table {schema}.{table_name} not found."}), 404
    try:
        rows_checked, issues, header = _temp_csv_issues(file.read(), columns)
    except Exception as e:
        return jsonify({"error": f"Cannot read file: {e}"}), 400
    return jsonify({
        "filename": file.filename, "schema": schema, "table": table_name,
        "columns": [{"name": c["name"], "type": c["label"]} for c in columns],
        "header": header is not None,
        "unknown_header_columns": [h for h in (header or [])
                                   if h and h.upper() not in {c["name"].upper() for c in columns}],
        "rows_checked": rows_checked,
        "bad_rows": len({i["row"] for i in issues if i["severity"] == "error"}),
        "error_count": sum(i["severity"] == "error" for i in issues),
        "warning_count": sum(i["severity"] == "warning" for i in issues),
        "issue_count": len(issues),
        "issues": issues[:2000],
    })


@app.route("/upload", methods=["POST"])
def upload():
    """Save uploaded file, return sheet names."""
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify({"error": "No file selected."}), 400

    is_csv = is_csv_filename(file.filename)
    table_name = csv_table_name(file.filename) if is_csv else None
    if is_csv and not table_name:
        return jsonify({"error": "CSV file name must start with the table name, e.g. MDNNOP.csv."}), 400

    tmp = Path(tempfile.gettempdir()) / f"{uuid.uuid4()}{'.csv' if is_csv else '.xlsx'}"
    file.save(tmp)

    try:
        if is_csv:
            read_csv_rows(tmp)
            sheets = [table_name]   # a CSV is one "sheet" holding one table
        else:
            wb = openpyxl.load_workbook(tmp, data_only=True, read_only=True)
            sheets = wb.sheetnames
            wb.close()
    except Exception as e:
        tmp.unlink(missing_ok=True)
        return jsonify({"error": f"Cannot open file: {e}"}), 400

    # Clean up previous temp file
    old = session.get("tmp_path")
    if old:
        Path(old).unlink(missing_ok=True)

    session["tmp_path"] = str(tmp)
    session["orig_filename"] = file.filename
    session["file_kind"] = "csv" if is_csv else "excel"
    session["csv_table"] = table_name
    return jsonify({
        "sheets": sheets,
        "kind": session["file_kind"],
        "table": table_name,
        "batch_job": derive_batch_job(file.filename),
    })


@app.route("/scan", methods=["POST"])
def scan():
    """Scan a sheet and return table preview."""
    payload = request.json or {}
    sheet_name = payload.get("sheet", "").strip()
    tmp_path = session.get("tmp_path")

    if not tmp_path or not Path(tmp_path).exists():
        return jsonify({"error": "Session expired — please re-upload the file."}), 400
    if not sheet_name:
        return jsonify({"error": "Sheet name is required."}), 400

    if session.get("file_kind") == "csv":
        try:
            rows, width = read_csv_rows(tmp_path)
        except Exception as e:
            return jsonify({"error": f"Cannot open file: {e}"}), 400
        session["sheet_name"] = sheet_name
        table_name = session.get("csv_table")
        # The import decides with the chosen schema's columns; this is the preview's guess.
        column_names = _csv_table_column_names(table_name) if rows else []
        has_header = None if column_names is None else csv_has_header(rows[0], column_names) if rows else False
        return jsonify({"tables": [_csv_preview(table_name, rows, width, has_header)], "kind": "csv"})

    try:
        wb = openpyxl.load_workbook(tmp_path, data_only=True, read_only=True)
    except Exception as e:
        return jsonify({"error": f"Cannot open file: {e}"}), 400

    if sheet_name not in wb.sheetnames:
        return jsonify({"error": f"Sheet '{sheet_name}' not found."}), 400

    ws = wb[sheet_name]
    tables = scan_table_previews(ws)

    wb.close()
    session["sheet_name"] = sheet_name
    return jsonify({"tables": tables})


@app.route("/schemas", methods=["GET"])
def get_schemas():
    """List schemas that own user tables, so the UI can offer a target dropdown."""
    cfg = load_config()
    if not cfg:
        return jsonify({"error": "No DB connection saved."}), 400
    conn_str, err = _build_conn_str(
        cfg["host"], cfg["port"], cfg["database"], cfg["username"], cfg["password"]
    )
    if err:
        return jsonify({"error": err}), 500
    try:
        conn = pyodbc.connect(conn_str, timeout=10)
        cur = conn.cursor()
        cur.execute(
            "SELECT DISTINCT TABLE_SCHEMA FROM INFORMATION_SCHEMA.TABLES "
            "WHERE TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_SCHEMA"
        )
        schemas = [r[0] for r in cur.fetchall()]
        conn.close()
    except pyodbc.Error as e:
        return jsonify({"error": f"Could not list schemas: {e}"}), 400
    return jsonify({"schemas": schemas, "default": cfg.get("schema") or "lvapdbf"})


SCHEMA_NAME_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_]{0,127}")


@app.route("/schemas/list", methods=["GET"])
def list_all_schemas():
    """List every user schema (including empty ones) for the Schema page."""
    cfg = load_config()
    if not cfg:
        return jsonify({"error": "No DB connection saved."}), 400
    conn_str, err = _build_conn_str(
        cfg["host"], cfg["port"], cfg["database"], cfg["username"], cfg["password"]
    )
    if err:
        return jsonify({"error": err}), 500
    try:
        conn = pyodbc.connect(conn_str, timeout=10)
        cur = conn.cursor()
        # schema_id >= 16384 are the fixed db_* role schemas.
        cur.execute(
            "SELECT s.name, USER_NAME(s.principal_id), "
            "(SELECT COUNT(*) FROM sys.tables t WHERE t.schema_id = s.schema_id) "
            "FROM sys.schemas s "
            "WHERE s.schema_id < 16384 AND s.name NOT IN ('sys', 'INFORMATION_SCHEMA', 'guest') "
            "ORDER BY s.name"
        )
        schemas = [{"name": r[0], "owner": r[1], "tables": r[2]} for r in cur.fetchall()]
        conn.close()
    except pyodbc.Error as e:
        return jsonify({"error": f"Could not list schemas: {e}"}), 400
    return jsonify({
        "schemas": schemas,
        "database": cfg["database"],
        "default": cfg.get("schema") or "lvapdbf",
    })


@app.route("/schemas", methods=["POST"])
def create_schema():
    """Create a new, empty schema in the connected database."""
    cfg = load_config()
    if not cfg:
        return jsonify({"error": "No DB connection saved."}), 400
    name = str((request.json or {}).get("name", "")).strip()
    if not SCHEMA_NAME_RE.fullmatch(name):
        return jsonify({"error": "Schema name must start with a letter or underscore and "
                                 "use only letters, digits, and underscores (max 128)."}), 400

    conn_str, err = _build_conn_str(
        cfg["host"], cfg["port"], cfg["database"], cfg["username"], cfg["password"]
    )
    if err:
        return jsonify({"error": err}), 500

    conn = None
    try:
        conn = pyodbc.connect(conn_str, timeout=10)
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sys.schemas WHERE name = ?", name)
        existing = cursor.fetchone()
        if existing is not None:
            return jsonify({"error": f"Schema '{existing[0]}' already exists."}), 409
        # Name is validated above, so bracket-quoting is safe.
        cursor.execute(f"CREATE SCHEMA [{name}]")
        conn.commit()
        return jsonify({"ok": True, "name": name, "message": f"Schema '{name}' created."})
    except pyodbc.Error as e:
        if conn is not None:
            conn.rollback()
        return jsonify({"error": f"Could not create schema '{name}': {e}"}), 400
    finally:
        if conn is not None:
            conn.close()


@app.route("/tables", methods=["GET"])
def get_tables():
    """List base tables in one schema for the maintenance UI."""
    cfg = load_config()
    if not cfg:
        return jsonify({"error": "No DB connection saved."}), 400
    schema = request.args.get("schema", "").strip()
    if not schema:
        return jsonify({"error": "Schema is required."}), 400

    conn_str, err = _build_conn_str(
        cfg["host"], cfg["port"], cfg["database"], cfg["username"], cfg["password"]
    )
    if err:
        return jsonify({"error": err}), 500
    try:
        conn = pyodbc.connect(conn_str, timeout=10)
        cur = conn.cursor()
        cur.execute(
            "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES "
            "WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_SCHEMA = ? ORDER BY TABLE_NAME",
            schema,
        )
        tables = [row[0] for row in cur.fetchall()]
        conn.close()
    except pyodbc.Error as e:
        return jsonify({"error": f"Could not list tables: {e}"}), 400
    return jsonify({"schema": schema, "tables": tables})


def _table_metadata(cursor, schema, table_name):
    cursor.execute(
        "SELECT COLUMN_NAME, DATA_TYPE, IS_NULLABLE FROM INFORMATION_SCHEMA.COLUMNS "
        "WHERE TABLE_SCHEMA = ? AND TABLE_NAME = ? ORDER BY ORDINAL_POSITION",
        schema, table_name,
    )
    columns = [
        {"name": row[0], "type": row[1], "nullable": row[2] == "YES"}
        for row in cursor.fetchall()
    ]
    if not columns:
        return [], []
    cursor.execute(
        "SELECT k.COLUMN_NAME FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS t "
        "JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE k "
        "ON t.CONSTRAINT_NAME = k.CONSTRAINT_NAME AND t.TABLE_SCHEMA = k.TABLE_SCHEMA "
        "WHERE t.CONSTRAINT_TYPE = 'PRIMARY KEY' AND t.TABLE_SCHEMA = ? AND t.TABLE_NAME = ? "
        "ORDER BY k.ORDINAL_POSITION",
        schema, table_name,
    )
    return columns, [row[0] for row in cursor.fetchall()]


def _json_table_value(value):
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, (date, datetime, datetime_time)):
        return value.isoformat()
    if isinstance(value, bytes):
        return "0x" + value.hex()
    return str(value)


QUERY_MAX_ROWS = 500
QUERY_MAX_LENGTH = 100_000
QUERY_TIMEOUT_SECONDS = 30


@app.route("/query", methods=["POST"])
def run_query():
    """Execute SQL from the query-runner tab and return bounded result sets."""
    cfg = load_config()
    if not cfg:
        return jsonify({"error": "No DB connection saved."}), 400

    payload = request.json or {}
    sql = str(payload.get("sql", "")).strip()
    if not sql:
        return jsonify({"error": "Enter a SQL query to run."}), 400
    if len(sql) > QUERY_MAX_LENGTH:
        return jsonify({"error": f"Query is too long (maximum {QUERY_MAX_LENGTH:,} characters)."}), 400

    conn_str, err = _build_conn_str(
        cfg["host"], cfg["port"], cfg["database"], cfg["username"], cfg["password"]
    )
    if err:
        return jsonify({"error": err}), 500

    conn = None
    started = time.perf_counter()
    try:
        conn = pyodbc.connect(conn_str, timeout=10)
        conn.timeout = QUERY_TIMEOUT_SECONDS
        cursor = conn.cursor()
        cursor.execute(sql)

        result_sets = []
        while True:
            if cursor.description:
                columns = [column[0] for column in cursor.description]
                fetched = cursor.fetchmany(QUERY_MAX_ROWS + 1)
                truncated = len(fetched) > QUERY_MAX_ROWS
                rows = fetched[:QUERY_MAX_ROWS]
                result_sets.append({
                    "columns": columns,
                    "rows": [[_json_table_value(value) for value in row] for row in rows],
                    "row_count": len(rows),
                    "truncated": truncated,
                })
            elif cursor.rowcount != -1:
                result_sets.append({
                    "columns": [],
                    "rows": [],
                    "affected_rows": cursor.rowcount,
                })

            if not cursor.nextset():
                break

        conn.commit()
        elapsed_ms = round((time.perf_counter() - started) * 1000)
        return jsonify({
            "ok": True,
            "message": "Query completed successfully.",
            "result_sets": result_sets,
            "max_rows": QUERY_MAX_ROWS,
            "elapsed_ms": elapsed_ms,
        })
    except pyodbc.Error as e:
        if conn is not None:
            conn.rollback()
        return jsonify({"error": f"Query failed: {e}"}), 400
    except Exception as e:
        if conn is not None:
            conn.rollback()
        app.logger.exception("Unexpected query runner error")
        return jsonify({"error": f"Query runner failed unexpectedly: {e}"}), 500
    finally:
        if conn is not None:
            conn.close()


EDITOR_PAGE_SIZES = (1000, 2000, 5000, 10000)
# SQL Server cannot ORDER BY these types.
UNSORTABLE_TYPES = {"text", "ntext", "image", "xml", "geography", "geometry"}


def _quote_identifier(value):
    return "[" + str(value).replace("]", "]]") + "]"


def _editor_order_sql(columns, primary_keys, sort_column, sort_dir):
    """ORDER BY for paging: the chosen column first, then a stable tie-breaker."""
    direction = "DESC" if sort_dir == "desc" else "ASC"
    parts, used = [], set()
    if sort_column:
        parts.append(f"{_quote_identifier(sort_column)} {direction}")
        used.add(sort_column)
    # Primary key keeps pages stable; without one, order by every sortable column.
    tie_breakers = primary_keys or [c["name"] for c in columns
                                    if (c.get("type") or "").lower() not in UNSORTABLE_TYPES]
    parts += [_quote_identifier(name) for name in tie_breakers if name not in used]
    return " ORDER BY " + (", ".join(parts) if parts else "(SELECT NULL)")


def _editor_request_args(columns):
    """Validate sort/dir query args against the table's columns."""
    sort_column = request.args.get("sort", "").strip() or None
    sort_dir = "desc" if request.args.get("dir", "").lower() == "desc" else "asc"
    if sort_column:
        match = next((c for c in columns if c["name"] == sort_column), None)
        if match is None:
            raise ValueError(f"Unknown sort column '{sort_column}'.")
        if (match.get("type") or "").lower() in UNSORTABLE_TYPES:
            raise ValueError(f"Column '{sort_column}' ({match['type']}) cannot be sorted.")
    return sort_column, sort_dir


@app.route("/table-data", methods=["GET"])
def get_table_data():
    """Return one page of table rows plus metadata for safe row editing."""
    cfg = load_config()
    if not cfg:
        return jsonify({"error": "No DB connection saved."}), 400
    schema = request.args.get("schema", "").strip()
    table_name = request.args.get("table", "").strip()
    if not schema or not table_name:
        return jsonify({"error": "Schema and table are required."}), 400
    try:
        page_size = int(request.args.get("page_size", EDITOR_PAGE_SIZES[0]))
        page = max(1, int(request.args.get("page", 1)))
    except ValueError:
        return jsonify({"error": "page and page_size must be numbers."}), 400
    if page_size not in EDITOR_PAGE_SIZES:
        return jsonify({"error": f"page_size must be one of {', '.join(map(str, EDITOR_PAGE_SIZES))}."}), 400

    conn_str, err = _build_conn_str(
        cfg["host"], cfg["port"], cfg["database"], cfg["username"], cfg["password"]
    )
    if err:
        return jsonify({"error": err}), 500
    conn = None
    try:
        conn = pyodbc.connect(conn_str, timeout=10)
        cursor = conn.cursor()
        columns, primary_keys = _table_metadata(cursor, schema, table_name)
        if not columns:
            return jsonify({"error": f"Table '{schema}.{table_name}' was not found."}), 404
        try:
            sort_column, sort_dir = _editor_request_args(columns)
        except ValueError as e:
            return jsonify({"error": str(e)}), 400
        qualified = f"{_quote_identifier(schema)}.{_quote_identifier(table_name)}"
        cursor.execute(f"SELECT COUNT_BIG(*) FROM {qualified}")
        total_rows = int(cursor.fetchone()[0])
        total_pages = max(1, -(-total_rows // page_size))
        page = min(page, total_pages)   # e.g. rows were deleted since the last load
        order_sql = _editor_order_sql(columns, primary_keys, sort_column, sort_dir)
        cursor.execute(
            f"SELECT * FROM {qualified}{order_sql} OFFSET ? ROWS FETCH NEXT ? ROWS ONLY",
            (page - 1) * page_size, page_size,
        )
        rows = [[_json_table_value(value) for value in row] for row in cursor.fetchall()]
        for column in columns:
            column["sortable"] = (column.get("type") or "").lower() not in UNSORTABLE_TYPES
        return jsonify({
            "schema": schema, "table": table_name, "columns": columns,
            "primary_keys": primary_keys,
            "identifier_columns": primary_keys or [column["name"] for column in columns],
            "rows": rows, "total_rows": total_rows,
            "page": page, "page_size": page_size, "total_pages": total_pages,
            "page_sizes": list(EDITOR_PAGE_SIZES),
            "sort": sort_column, "dir": sort_dir,
        })
    except pyodbc.Error as e:
        return jsonify({"error": f"Could not load table data: {e}"}), 400
    finally:
        if conn is not None:
            conn.close()


def _csv_download_value(value):
    if value is None:
        return ""
    if isinstance(value, bool):
        return "1" if value else "0"
    return str(_json_table_value(value))


@app.route("/table-data/download", methods=["GET"])
def download_table_data():
    """Stream every row of a table as a UTF-8 CSV, in the Data Editor's sort order."""
    cfg = load_config()
    if not cfg:
        return jsonify({"error": "No DB connection saved."}), 400
    schema = request.args.get("schema", "").strip()
    table_name = request.args.get("table", "").strip()
    if not schema or not table_name:
        return jsonify({"error": "Schema and table are required."}), 400
    include_header = request.args.get("header", "1") != "0"

    conn_str, err = _build_conn_str(
        cfg["host"], cfg["port"], cfg["database"], cfg["username"], cfg["password"]
    )
    if err:
        return jsonify({"error": err}), 500
    try:
        conn = pyodbc.connect(conn_str, timeout=10)
    except pyodbc.Error as e:
        return jsonify({"error": f"Could not connect: {e}"}), 400
    try:
        cursor = conn.cursor()
        columns, primary_keys = _table_metadata(cursor, schema, table_name)
        if not columns:
            conn.close()
            return jsonify({"error": f"Table '{schema}.{table_name}' was not found."}), 404
        sort_column, sort_dir = _editor_request_args(columns)
        qualified = f"{_quote_identifier(schema)}.{_quote_identifier(table_name)}"
        cursor.execute(f"SELECT * FROM {qualified}"
                       + _editor_order_sql(columns, primary_keys, sort_column, sort_dir))
    except ValueError as e:
        conn.close()
        return jsonify({"error": str(e)}), 400
    except pyodbc.Error as e:
        conn.close()
        return jsonify({"error": f"Could not load table data: {e}"}), 400

    def generate():
        buffer = io.StringIO()
        writer = csv.writer(buffer)
        try:
            yield "﻿"   # BOM so Excel opens UTF-8 (Japanese text) correctly
            if include_header:
                writer.writerow([c["name"] for c in columns])
            while True:
                chunk = cursor.fetchmany(5000)
                if not chunk:
                    break
                writer.writerows([_csv_download_value(v) for v in row] for row in chunk)
                yield buffer.getvalue()
                buffer.seek(0); buffer.truncate()
            if buffer.tell():
                yield buffer.getvalue()
        finally:
            conn.close()

    # "TABLE (timestamp).csv": the CSV import reads the table name up to the space.
    filename = f"{table_name} ({datetime.now():%Y%m%d_%H%M%S}).csv"
    ascii_name = filename.encode("ascii", "replace").decode().replace('"', "")
    return app.response_class(generate(), mimetype="text/csv", headers={
        "Content-Disposition": f"attachment; filename=\"{ascii_name}\"; "
                               f"filename*=UTF-8''{urllib.parse.quote(filename)}",
    })


@app.route("/table-data", methods=["PATCH"])
def update_table_row():
    """Update one row identified by the table's primary key."""
    cfg = load_config()
    payload = request.json or {}
    schema = str(payload.get("schema", "")).strip()
    table_name = str(payload.get("table", "")).strip()
    keys = payload.get("keys") or {}
    values = payload.get("values") or {}
    if not cfg or not schema or not table_name or not isinstance(keys, dict) or not isinstance(values, dict):
        return jsonify({"error": "Invalid update request."}), 400
    if not values:
        return jsonify({"error": "No changed values were provided."}), 400

    conn_str, err = _build_conn_str(
        cfg["host"], cfg["port"], cfg["database"], cfg["username"], cfg["password"]
    )
    if err:
        return jsonify({"error": err}), 500
    conn = None
    try:
        conn = pyodbc.connect(conn_str, timeout=10)
        cursor = conn.cursor()
        columns, primary_keys = _table_metadata(cursor, schema, table_name)
        column_names = {column["name"] for column in columns}
        if not columns:
            return jsonify({"error": f"Table '{schema}.{table_name}' was not found."}), 404
        identifier_columns = primary_keys or [column["name"] for column in columns]
        if set(keys) != set(identifier_columns):
            return jsonify({"error": "Complete original row values are required to update this row."}), 400
        # Primary-key values may be changed; the original key values in `keys`
        # still identify the row in the WHERE clause.
        changed = [name for name in values if name in column_names]
        if len(changed) != len(values):
            return jsonify({"error": "One or more columns cannot be edited."}), 400

        quote = lambda name: "[" + name.replace("]", "]]") + "]"
        qualified = f"{quote(schema)}.{quote(table_name)}"
        set_sql = ", ".join(f"{quote(name)} = ?" for name in changed)
        where_parts = []
        where_params = []
        for name in identifier_columns:
            if keys[name] is None:
                where_parts.append(f"{quote(name)} IS NULL")
            else:
                where_parts.append(f"{quote(name)} = ?")
                where_params.append(keys[name])
        where_sql = " AND ".join(where_parts)
        params = [values[name] for name in changed] + where_params
        cursor.execute(f"UPDATE {qualified} SET {set_sql} WHERE {where_sql}", params)
        affected = cursor.rowcount
        if affected != 1:
            conn.rollback()
            return jsonify({"error": f"Expected one row, but matched {affected}. Data was not changed."}), 409
        conn.commit()
        return jsonify({"ok": True, "message": "Row updated successfully."})
    except pyodbc.Error as e:
        if conn is not None:
            conn.rollback()
        return jsonify({"error": f"Could not update row: {e}"}), 400
    finally:
        if conn is not None:
            conn.close()


# ── Version / update check ────────────────────────────────────────────────────
# version.json is baked into the image, so it is the version this copy runs.
# `update_url` (or the UPDATE_CHECK_URL env var) points at a published copy of
# the latest version.json; leaving it empty disables the check.
VERSION_PATH = Path(__file__).with_name("version.json")
UPDATE_CACHE_SECONDS = 5 * 60
_update_cache = {"at": 0.0, "data": None}


def _load_version_info() -> dict:
    try:
        return json.loads(VERSION_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"version": "0.0.0"}


def _version_tuple(version: str) -> tuple:
    parts = [int(n) for n in re.findall(r"\d+", str(version))[:3]]
    return tuple(parts + [0] * (3 - len(parts)))


def _fetch_latest_version(url: str) -> dict:
    # GitHub's raw CDN keeps serving an old Gist for a few minutes after an
    # edit; a changing query string makes it return the current content.
    url += ("&" if "?" in url else "?") + f"t={int(time.time())}"
    req = urllib.request.Request(url, headers={"Accept": "application/json", "Cache-Control": "no-cache"})
    with urllib.request.urlopen(req, timeout=5) as resp:
        return json.loads(resp.read().decode("utf-8"))


# ── Self-update (git pull inside the container) ───────────────────────────────
# docker-compose.yml mounts the project folder at /app, so pulling there
# updates the user's own clone; Flask debug mode reloads the code.
APP_DIR = Path(__file__).resolve().parent
DEFAULT_REPO_URL = "https://github.com/aungmyomyat-gic/db-import.git"
UPDATE_BRANCH = "main"
REBUILD_FILES = {"Dockerfile", "requirements.txt"}
_update_lock = threading.Lock()


def _git(*args, timeout=30):
    # Run git as the owner of the project folder so pulled files aren't left
    # owned by root on the host (Linux / WSL bind mounts).
    owner = (APP_DIR / ".git").stat()
    kwargs = {}
    if os.geteuid() == 0 and owner.st_uid != 0:
        kwargs = {"user": owner.st_uid, "group": owner.st_gid}
    env = {**os.environ, "HOME": "/tmp", "GIT_TERMINAL_PROMPT": "0"}
    return subprocess.run(
        ["git", "-c", f"safe.directory={APP_DIR}", "-C", str(APP_DIR), *args],
        capture_output=True, text=True, timeout=timeout, env=env, **kwargs,
    )


def _self_update_status():
    """Return (ok, reason) — whether the in-app Update button can work here."""
    if not shutil.which("git"):
        return False, "git is not installed in the container — rebuild once with docker compose up -d --build"
    if not (APP_DIR / ".git").exists():
        return False, "the app folder is not a git clone"
    branch = _git("rev-parse", "--abbrev-ref", "HEAD")
    if branch.returncode != 0:
        return False, branch.stderr.strip() or "could not read the git branch"
    if branch.stdout.strip() != UPDATE_BRANCH:
        return False, f"this copy is on branch '{branch.stdout.strip()}', updates only apply to '{UPDATE_BRANCH}'"
    return True, ""


@app.route("/update", methods=["POST"])
def run_self_update():
    # JSON-only so a plain cross-site form post can't trigger it
    if not request.is_json:
        return jsonify({"error": "Invalid update request."}), 400
    ok, reason = _self_update_status()
    if not ok:
        return jsonify({"error": f"Update is not available: {reason}."}), 400
    if not _update_lock.acquire(blocking=False):
        return jsonify({"error": "An update is already running."}), 409
    try:
        before = _git("rev-parse", "HEAD").stdout.strip()
        repo_url = _load_version_info().get("repo_url") or DEFAULT_REPO_URL
        pull = _git("pull", "--ff-only", repo_url, UPDATE_BRANCH, timeout=120)
        output = (pull.stdout + pull.stderr).strip()
        if pull.returncode != 0:
            return jsonify({"error": "git pull failed.", "output": output}), 400
        after = _git("rev-parse", "HEAD").stdout.strip()
        changed = []
        if before != after:
            diff = _git("diff", "--name-only", before, after)
            changed = [line for line in diff.stdout.splitlines() if line]
        return jsonify({
            "ok": True, "updated": before != after, "output": output,
            "version": _load_version_info().get("version"),
            "needs_rebuild": any(Path(f).name in REBUILD_FILES for f in changed),
        })
    except (OSError, subprocess.SubprocessError) as e:
        return jsonify({"error": f"Update failed: {e}"}), 500
    finally:
        _update_lock.release()


@app.route("/version", methods=["GET"])
def get_version():
    local = _load_version_info()
    can_update, update_reason = _self_update_status()
    result = {
        "current": local.get("version", "0.0.0"), "date": local.get("date"), "notes": local.get("notes", []),
        "can_self_update": can_update, "self_update_reason": update_reason,
    }
    url = os.environ.get("UPDATE_CHECK_URL") or local.get("update_url") or ""
    if not url:
        return jsonify({**result, "check": "disabled"})

    now = time.time()
    force = request.args.get("force") == "1"
    if force or _update_cache["data"] is None or now - _update_cache["at"] > UPDATE_CACHE_SECONDS:
        try:
            _update_cache["data"] = _fetch_latest_version(url)
            _update_cache["at"] = now
        except (OSError, ValueError, urllib.error.URLError) as e:
            return jsonify({**result, "check": "failed", "error": str(e)})
    latest = _update_cache["data"] or {}
    latest_version = str(latest.get("version", "0.0.0"))
    return jsonify({
        **result, "check": "ok", "latest": latest_version,
        "latest_date": latest.get("date"), "latest_notes": latest.get("notes", []),
        "update_available": _version_tuple(latest_version) > _version_tuple(result["current"]),
    })


@app.route("/table-data", methods=["POST"])
def insert_table_rows():
    """Insert pasted rows into a table in a single all-or-nothing transaction."""
    cfg = load_config()
    payload = request.json or {}
    schema = str(payload.get("schema", "")).strip()
    table_name = str(payload.get("table", "")).strip()
    names = payload.get("columns") or []
    rows = payload.get("rows") or []
    if not cfg or not schema or not table_name or not isinstance(names, list) or not isinstance(rows, list):
        return jsonify({"error": "Invalid insert request."}), 400
    if not names or not rows:
        return jsonify({"error": "No rows to insert."}), 400
    if len(rows) > 5000:
        return jsonify({"error": "Paste at most 5000 rows at a time."}), 400
    if len(set(names)) != len(names) or any(not isinstance(r, list) or len(r) != len(names) for r in rows):
        return jsonify({"error": "Every row must have one value per column."}), 400

    conn_str, err = _build_conn_str(
        cfg["host"], cfg["port"], cfg["database"], cfg["username"], cfg["password"]
    )
    if err:
        return jsonify({"error": err}), 500
    conn = None
    try:
        conn = pyodbc.connect(conn_str, timeout=10)
        cursor = conn.cursor()
        columns, _ = _table_metadata(cursor, schema, table_name)
        if not columns:
            return jsonify({"error": f"Table '{schema}.{table_name}' was not found."}), 404
        by_name = {column["name"]: column for column in columns}
        unknown = [name for name in names if name not in by_name]
        if unknown:
            return jsonify({"error": f"Unknown column(s): {', '.join(unknown)}"}), 400

        quote = lambda name: "[" + name.replace("]", "]]") + "]"
        qualified = f"{quote(schema)}.{quote(table_name)}"
        cursor.execute(
            "SELECT name FROM sys.columns WHERE object_id = OBJECT_ID(?) AND is_identity = 1", qualified
        )
        identity = {row[0] for row in cursor.fetchall()}
        prefix = f"SET IDENTITY_INSERT {qualified} ON; " if identity & set(names) else ""
        suffix = f"; SET IDENTITY_INSERT {qualified} OFF" if prefix else ""
        insert_sql = (
            f"{prefix}INSERT INTO {qualified} ({', '.join(quote(n) for n in names)}) "
            f"VALUES ({', '.join('?' for _ in names)}){suffix}"
        )

        for row_number, row in enumerate(rows, start=1):
            values = []
            for name, value in zip(names, row):
                # Empty Excel cells become NULL, or the type default for NOT NULL columns
                if value is None or (isinstance(value, str) and value == ""):
                    value = None if by_name[name]["nullable"] else _type_default(by_name[name]["type"])
                values.append(value)
            try:
                cursor.execute(insert_sql, values)
            except pyodbc.Error as e:
                conn.rollback()
                return jsonify({"error": f"Row {row_number} failed, nothing was inserted: {e}"}), 400
        conn.commit()
        return jsonify({"ok": True, "inserted": len(rows), "message": f"{len(rows)} row(s) inserted."})
    except pyodbc.Error as e:
        if conn is not None:
            conn.rollback()
        return jsonify({"error": f"Could not insert rows: {e}"}), 400
    finally:
        if conn is not None:
            conn.close()


@app.route("/import", methods=["POST"])
def do_import():
    """Start a background import job and return its initial status."""
    cfg = load_config()
    if not cfg:
        return jsonify({"error": "No DB connection saved. Please set it up first."}), 400

    tmp_path   = session.get("tmp_path")
    sheet_name = session.get("sheet_name")
    payload = request.json or {}

    # Only import the tables the user ticked (None = all, for backward compat).
    selected = payload.get("tables")
    selected_names = None
    if selected is not None:
        selected_names = [str(name) for name in selected if str(name).strip()]
    if selected_names is not None and not selected_names:
        return jsonify({"error": "No tables selected."}), 400

    if not tmp_path or not Path(tmp_path).exists():
        return jsonify({"error": "Session expired — please re-upload the file."}), 400

    conn_str, err = _build_conn_str(
        cfg["host"], cfg["port"], cfg["database"], cfg["username"], cfg["password"]
    )
    if err:
        return jsonify({"error": err}), 500

    # Each table can target its own schema (per-row dropdown in the UI);
    # anything not listed falls back to saved config, then default.
    preferred_schema = str(payload.get("schema", "")).strip() \
        or cfg.get("schema") or "lvapdbf"
    raw_schemas = payload.get("schemas")
    table_schemas = {}
    if isinstance(raw_schemas, dict):
        table_schemas = {str(name): str(schema).strip()
                         for name, schema in raw_schemas.items() if str(schema).strip()}

    _cleanup_import_jobs()
    job_id = uuid.uuid4().hex
    job = _new_import_job(job_id, selected_names or [])
    job["owner"] = session.setdefault("import_owner", uuid.uuid4().hex)
    with IMPORT_JOBS_LOCK:
        IMPORT_JOBS[job_id] = job
        snapshot = _job_snapshot(job)

    worker = threading.Thread(
        target=_run_import_job,
        args=(job_id, cfg, str(tmp_path), sheet_name, selected_names, preferred_schema, table_schemas,
              session.get("csv_table") if session.get("file_kind") == "csv" else None),
        daemon=True,
    )
    worker.start()
    return jsonify(snapshot), 202


@app.route("/import/status/<job_id>", methods=["GET"])
def import_status(job_id):
    _cleanup_import_jobs()
    with IMPORT_JOBS_LOCK:
        job = IMPORT_JOBS.get(job_id)
        if not job or not job.get("owner") or job["owner"] != session.get("import_owner"):
            return jsonify({"error": "Import job not found."}), 404
        return jsonify(_job_snapshot(job))


@app.route("/import/sql/<job_id>", methods=["GET"])
def download_import_sql(job_id):
    _cleanup_import_jobs()
    with IMPORT_JOBS_LOCK:
        job = IMPORT_JOBS.get(job_id)
        if not job or not job.get("owner") or job["owner"] != session.get("import_owner"):
            return jsonify({"error": "Import job not found."}), 404
        if job["state"] != "done" or not job.get("sql_path"):
            return jsonify({"error": "SQL is available only after a successful CSV import."}), 409
        if not Path(job["sql_path"]).is_file():
            return jsonify({"error": "SQL download expired. Please import the CSV again."}), 404
        return send_file(job["sql_path"], as_attachment=True, download_name=job["sql_filename"],
                         mimetype="application/sql", conditional=False, max_age=0)


@app.route("/truncate", methods=["POST"])
def truncate_table():
    """Remove all rows from one validated table in the selected schema."""
    cfg = load_config()
    if not cfg:
        return jsonify({"error": "No DB connection saved. Please set it up first."}), 400

    payload = request.json or {}
    table_name = str(payload.get("table", "")).strip()
    schema = str(payload.get("schema", "")).strip() or cfg.get("schema") or "lvapdbf"
    if not table_name or not schema:
        return jsonify({"error": "Table and schema are required."}), 400

    conn_str, err = _build_conn_str(
        cfg["host"], cfg["port"], cfg["database"], cfg["username"], cfg["password"]
    )
    if err:
        return jsonify({"error": err}), 500

    conn = None
    try:
        conn = pyodbc.connect(conn_str, timeout=10)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT 1 FROM INFORMATION_SCHEMA.TABLES "
            "WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_SCHEMA = ? AND TABLE_NAME = ?",
            schema, table_name,
        )
        if cursor.fetchone() is None:
            return jsonify({"error": f"Table '{schema}.{table_name}' was not found."}), 404

        safe_schema = schema.replace("]", "]]")
        safe_table = table_name.replace("]", "]]")
        cursor.execute(f"TRUNCATE TABLE [{safe_schema}].[{safe_table}]")
        conn.commit()
        return jsonify({"ok": True, "message": f"{schema}.{table_name} truncated successfully."})
    except pyodbc.Error as e:
        if conn is not None:
            conn.rollback()
        return jsonify({"error": f"Could not truncate '{schema}.{table_name}': {e}"}), 400
    finally:
        if conn is not None:
            conn.close()


@app.route("/call-batch", methods=["POST"])
def call_batch():
    """Trigger the batch endpoint matching the uploaded file's name."""
    job = derive_batch_job(session.get("orig_filename"))
    if not job:
        return jsonify({"error": "Could not determine a batch job from the file name."}), 400

    url = f"{BATCH_BASE_URL}/batch/{job['slug']}"
    try:
        with urllib.request.urlopen(url, timeout=15) as resp:
            return jsonify({"ok": True, "slug": job["slug"], "status": resp.status})
    except urllib.error.HTTPError as e:
        return jsonify({"error": f"{url} responded {e.code}: {e.reason}"}), 502
    except urllib.error.URLError as e:
        return jsonify({"error": f"Could not reach {url}: {e.reason}"}), 502


def _new_import_job(job_id, table_names):
    now = time.time()
    return {
        "id": job_id,
        "state": "queued",
        "total": len(table_names),
        "completed": 0,
        "results": [
            {"name": name, "status": "queued", "message": "Waiting", "errors": [],
             "progress": {"done": 0, "total": 0}}
            for name in table_names
        ],
        "error": None,
        "created_at": now,
        "updated_at": now,
    }


def _job_snapshot(job):
    snapshot = {
        "job_id": job["id"],
        "state": job["state"],
        "total": job["total"],
        "completed": job["completed"],
        "results": [dict(result) for result in job["results"]],
        "error": job.get("error"),
    }
    if job["state"] == "done" and job.get("sql_path"):
        snapshot["sql_download_url"] = f"/import/sql/{job['id']}"
        snapshot["sql_filename"] = job["sql_filename"]
    return snapshot


def _cleanup_import_jobs():
    cutoff = time.time() - IMPORT_JOB_TTL_SECONDS
    with IMPORT_JOBS_LOCK:
        for job_id, job in list(IMPORT_JOBS.items()):
            if job["state"] in {"done", "failed"} and job["updated_at"] < cutoff:
                if job.get("sql_path"):
                    try:
                        Path(job["sql_path"]).unlink(missing_ok=True)
                    except OSError:
                        app.logger.warning("Could not remove an expired SQL export.")
                del IMPORT_JOBS[job_id]


def _replace_job_queue(job_id, table_names):
    now = time.time()
    with IMPORT_JOBS_LOCK:
        job = IMPORT_JOBS.get(job_id)
        if not job:
            return
        job["total"] = len(table_names)
        job["completed"] = 0
        job["results"] = [
            {"name": name, "status": "queued", "message": "Waiting", "errors": [],
             "progress": {"done": 0, "total": 0}}
            for name in table_names
        ]
        job["updated_at"] = now


def _set_job_state(job_id, state, error=None):
    with IMPORT_JOBS_LOCK:
        job = IMPORT_JOBS.get(job_id)
        if not job:
            return
        job["state"] = state
        job["error"] = error
        job["updated_at"] = time.time()


def _mark_job_table(job_id, table_name, status, message, errors=None, schema=None):
    final_statuses = {"ok", "warn", "skip", "failed"}
    with IMPORT_JOBS_LOCK:
        job = IMPORT_JOBS.get(job_id)
        if not job:
            return
        existing = next((r for r in job["results"] if r["name"] == table_name), None)
        if existing is None:
            existing = {"name": table_name, "status": "queued", "message": "Waiting", "errors": [],
                        "progress": {"done": 0, "total": 0}}
            job["results"].append(existing)
            job["total"] = max(job["total"], len(job["results"]))

        was_final = existing["status"] in final_statuses
        existing.update({
            "status": status,
            "message": message,
            "errors": errors or [],
        })
        if schema:
            existing["schema"] = schema
        if status in final_statuses and not was_final:
            job["completed"] += 1
        job["updated_at"] = time.time()


def _mark_job_progress(job_id, table_name, done, total):
    with IMPORT_JOBS_LOCK:
        job = IMPORT_JOBS.get(job_id)
        if not job:
            return
        existing = next((r for r in job["results"] if r["name"] == table_name), None)
        if existing is None:
            return
        existing["progress"] = {"done": done, "total": total}
        job["updated_at"] = time.time()


def _finish_import_job(job_id):
    with IMPORT_JOBS_LOCK:
        job = IMPORT_JOBS.get(job_id)
        if not job:
            return
        if job["state"] != "failed":
            job["state"] = "done"
        job["updated_at"] = time.time()


def _fail_import_job(job_id, message):
    with IMPORT_JOBS_LOCK:
        job = IMPORT_JOBS.get(job_id)
        if not job:
            return
        final_statuses = {"ok", "warn", "skip", "failed"}
        for result in job["results"]:
            if result["status"] not in final_statuses:
                result["status"] = "failed"
                result["message"] = message
                result["errors"] = []
                job["completed"] += 1
        job["state"] = "failed"
        job["error"] = message
        job["updated_at"] = time.time()


def _run_import_job(job_id, cfg, tmp_path, sheet_name, selected_names, preferred_schema,
                    table_schemas=None, csv_table=None, delimiter=","):
    table_schemas = table_schemas or {}
    conn = None
    wb = None
    try:
        _set_job_state(job_id, "running")
        conn_str, err = _build_conn_str(
            cfg["host"], cfg["port"], cfg["database"], cfg["username"], cfg["password"]
        )
        if err:
            _fail_import_job(job_id, err)
            return
        conn = pyodbc.connect(conn_str, timeout=10)

        if csv_table:
            _run_csv_import(job_id, conn, tmp_path, csv_table, selected_names,
                            table_schemas.get(csv_table) or preferred_schema, delimiter)
            _finish_import_job(job_id)
            return

        wb = openpyxl.load_workbook(tmp_path, data_only=True, read_only=True)
        if sheet_name not in wb.sheetnames:
            _fail_import_job(job_id, f"Sheet '{sheet_name}' not found.")
            return

        ws = wb[sheet_name]
        metas = scan_tables(ws)
        if selected_names is None:
            selected_names = [meta["name"] for meta in metas]
            _replace_job_queue(job_id, selected_names)

        selected_set = set(selected_names)
        pending = set(selected_names)
        for meta in metas:
            table_name = meta["name"]
            if table_name not in selected_set:
                continue
            pending.discard(table_name)
            schema = table_schemas.get(table_name) or preferred_schema
            _mark_job_table(job_id, table_name, "running", "Importing", schema=schema)
            try:
                headers, rows = extract_table(ws, meta)
                if not headers:
                    result = {"name": table_name, "status": "skip", "message": "Header row empty", "errors": []}
                else:
                    progress_cb = lambda done, total, tn=table_name: _mark_job_progress(job_id, tn, done, total)
                    result = _import_one(conn, table_name, headers, rows, schema, progress_cb)
            except Exception as e:
                try:
                    conn.rollback()
                except Exception:
                    pass
                result = {"name": table_name, "status": "failed", "message": f"Failed: {e}", "errors": []}
            _mark_job_table(
                job_id,
                table_name,
                result.get("status", "failed"),
                result.get("message", ""),
                result.get("errors", []),
                schema=schema,
            )

        for table_name in pending:
            _mark_job_table(job_id, table_name, "skip", "Table not found in sheet", [])
        _finish_import_job(job_id)
    except Exception as e:
        _fail_import_job(job_id, f"Import failed: {e}")
    finally:
        if conn is not None:
            try:
                conn.close()
            except Exception:
                pass
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass


def _run_csv_import(job_id, conn, tmp_path, table_name, selected_names, schema, delimiter=","):
    if selected_names is None:
        _replace_job_queue(job_id, [table_name])
    elif table_name not in selected_names:
        for name in selected_names:
            _mark_job_table(job_id, name, "skip", "Table not found in file", [])
        return
    _mark_job_table(job_id, table_name, "running", "Importing", schema=schema)
    sql_path = Path(SQL_EXPORT_DIR.name) / f"{job_id}.sql"
    if delimiter == ",":
        try:   # TEMP: remove with the TEMP CSV check block
            _temp_log_csv_check(conn, tmp_path, schema, table_name)
        except Exception as e:
            print(f"[TEMP CSV CHECK] failed: {e}", flush=True)
    try:
        rows, _ = read_csv_rows(tmp_path, delimiter)
        progress_cb = lambda done, total: _mark_job_progress(job_id, table_name, done, total)
        result = _import_one(conn, table_name, None, rows, schema, progress_cb,
                             sql_export_path=sql_path)
    except Exception as e:
        try:
            conn.rollback()
        except Exception:
            pass
        result = {"name": table_name, "status": "failed", "message": f"Failed: {e}", "errors": []}
    published = False
    if result.get("status") == "ok" and sql_path.is_file():
        with IMPORT_JOBS_LOCK:
            job = IMPORT_JOBS.get(job_id)
            if job:
                job["sql_path"] = str(sql_path)
                job["sql_filename"] = f"{table_name}_{datetime.now():%Y%m%d_%H%M%S}.sql"
                published = True
    if not published:
        sql_path.unlink(missing_ok=True)
    _mark_job_table(job_id, table_name, result.get("status", "failed"), result.get("message", ""),
                    result.get("errors", []), schema=schema)


def _sql_identifier(value):
    return "[" + str(value).replace("]", "]]") + "]"


SQL_RESERVED_WORDS = frozenset("""
ADD ALL ALTER AND ANY AS ASC AUTHORIZATION BACKUP BEGIN BETWEEN BREAK BROWSE BULK BY
CASCADE CASE CHECK CHECKPOINT CLOSE CLUSTERED COALESCE COLLATE COLUMN COMMIT COMPUTE
CONSTRAINT CONTAINS CONTAINSTABLE CONTINUE CONVERT CREATE CROSS CURRENT CURRENT_DATE
CURRENT_TIME CURRENT_TIMESTAMP CURRENT_USER CURSOR DATABASE DBCC DEALLOCATE DECLARE
DEFAULT DELETE DENY DESC DISK DISTINCT DISTRIBUTED DOUBLE DROP DUMP ELSE END ERRLVL
ESCAPE EXCEPT EXEC EXECUTE EXISTS EXIT EXTERNAL FETCH FILE FILLFACTOR FOR FOREIGN
FREETEXT FREETEXTTABLE FROM FULL FUNCTION GOTO GRANT GROUP HAVING HOLDLOCK IDENTITY
IDENTITY_INSERT IDENTITYCOL IF IN INDEX INNER INSERT INTERSECT INTO IS JOIN KEY KILL
LEFT LIKE LINENO LOAD MERGE NATIONAL NOCHECK NONCLUSTERED NOT NULL NULLIF OF OFF
OFFSETS ON OPEN OPENDATASOURCE OPENQUERY OPENROWSET OPENXML OPTION OR ORDER OUTER
OVER PERCENT PIVOT PLAN PRECISION PRIMARY PRINT PROC PROCEDURE PUBLIC RAISERROR READ
READTEXT RECONFIGURE REFERENCES REPLICATION RESTORE RESTRICT RETURN REVERT REVOKE
RIGHT ROLLBACK ROWCOUNT ROWGUIDCOL RULE SAVE SCHEMA SECURITYAUDIT SELECT
SEMANTICKEYPHRASETABLE SEMANTICSIMILARITYDETAILSTABLE SEMANTICSIMILARITYTABLE
SESSION_USER SET SETUSER SHUTDOWN SOME STATISTICS SYSTEM_USER TABLE TABLESAMPLE
TEXTSIZE THEN TO TOP TRAN TRANSACTION TRIGGER TRUNCATE TRY_CONVERT TSEQUAL UNION
UNIQUE UNPIVOT UPDATE UPDATETEXT USE USER VALUES VARYING VIEW WAITFOR WHEN WHERE
WHILE WITH WITHIN WRITETEXT
""".split())
SQL_NUMERIC_TYPES = {
    "int", "bigint", "smallint", "tinyint", "bit", "decimal", "numeric", "money",
    "smallmoney", "float", "real",
}


def _export_identifier(value):
    # Keep ordinary names readable, e.g. lvapdbf.MSHBRP. Delimit special names
    # and SQL keywords so the generated script still works for those tables.
    if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", value) and value.upper() not in SQL_RESERVED_WORDS:
        return value
    return _sql_identifier(value)


def _sql_literal(value, dtype=None):
    """Use the destination column type to distinguish numbers from text.

    Control characters become NCHAR expressions to keep client batch separators
    such as GO out of standalone lines in the generated script.
    """
    if value is None:
        return "NULL"
    if isinstance(value, str) and dtype in SQL_NUMERIC_TYPES:
        text = value.strip()
        if dtype == "bit" and text.lower() in {"true", "false"}:
            return "1" if text.lower() == "true" else "0"
        if re.fullmatch(r"[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?", text):
            try:
                number = Decimal(text)
            except InvalidOperation:
                number = None
            if number is not None and dtype in {"float", "real"}:
                # Scientific literals keep a multi-row float column from being
                # inferred as int/decimal before it reaches the target table.
                return format(number, "E")
            # Keep extreme inputs as strings for SQL Server to validate, without
            # expanding an arbitrary exponent into a very large Python string.
            if number is not None and -38 <= number.as_tuple().exponent <= 38 and number.adjusted() <= 38:
                return format(number, "f")
        # Invalid numeric inputs retain the importer's existing error handling.
    if isinstance(value, bool):
        return "1" if value else "0"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, Decimal):
        if not value.is_finite():
            raise ValueError("Cannot export a non-finite SQL number.")
        return format(value, "f")
    if isinstance(value, float):
        if value != value or value in (float("inf"), float("-inf")):
            raise ValueError("Cannot export a non-finite SQL number.")
        return repr(value)
    if isinstance(value, (bytes, bytearray, memoryview)):
        return "0x" + bytes(value).hex().upper()
    if isinstance(value, (date, datetime_time)):
        value = value.isoformat()
    text = str(value)
    if re.search(r"[\x00-\x1f]", text):
        parts = re.split(r"([\x00-\x1f])", text)
        expressions = [f"NCHAR({ord(part)})" if len(part) == 1 and ord(part) < 32
                       else "N'" + part.replace("'", "''") + "'" for part in parts]
        # A max-length seed avoids truncation when concatenating long cells.
        return "(CAST(N'' AS nvarchar(max)) + " + " + ".join(expressions) + ")"
    return "N'" + text.replace("'", "''") + "'"


def _write_import_sql(path, schema, table_name, plan, rows, db_cols):
    """Write DELETE followed by multi-row INSERTs in table column order."""
    qualified = f"{_export_identifier(schema)}.{_export_identifier(table_name)}"
    insert_sql = f"INSERT INTO {qualified}"
    # Full-width CSVs use the user's column-free format. A partial mapping needs
    # a column list so omitted columns still receive their database defaults.
    if [column.upper() for column, _ in plan] != list(db_cols):
        columns = ", ".join(_export_identifier(column) for column, _ in plan)
        insert_sql += f" ({columns})"
    types = [db_cols[column.upper()]["type"] for column, _ in plan]
    with Path(path).open("w", encoding="utf-8-sig", newline="\n") as script:
        script.write(f"DELETE FROM {qualified};\n")
        # SQL Server permits at most 1,000 rows in each INSERT ... VALUES.
        for offset in range(0, len(rows), 1000):
            batch = rows[offset:offset + 1000]
            script.write(f"{insert_sql}\nVALUES\n")
            for index, row in enumerate(batch):
                values = ", ".join(_sql_literal(value, dtype) for value, dtype in
                                   zip(_values_for_insert(row, plan, db_cols), types))
                ending = ";\n" if index == len(batch) - 1 else ",\n"
                script.write(f"({values}){ending}")
            if offset + len(batch) < len(rows):
                script.write("\n")


def _values_for_insert(row, plan, db_cols):
    values = []
    for col_name, src in plan:
        if src[0] == "sheet":
            v = row[src[1]]
            if v is None and not db_cols[col_name.upper()]["nullable"]:
                v = _type_default(db_cols[col_name.upper()]["type"])
            values.append(v)
        else:  # "fill"
            values.append(src[1])
    return tuple(values)


def _insert_batch(cursor, insert_sql, batch, error_samples):
    cursor.execute("SAVE TRANSACTION import_batch")
    try:
        cursor.fast_executemany = True
        cursor.executemany(insert_sql, batch)
        return len(batch), 0
    except pyodbc.Error:
        cursor.fast_executemany = False
        cursor.execute("ROLLBACK TRANSACTION import_batch")

    inserted = errors = 0
    for values in batch:
        try:
            cursor.execute(insert_sql, values)
            inserted += 1
        except pyodbc.Error as e:
            errors += 1
            if len(error_samples) < 2:
                error_samples.append(str(e))
    return inserted, errors


def _import_one(conn, table_name, headers, rows, preferred_schema="lvapdbf", progress_cb=None,
                sql_export_path=None):
    """
    Replace the table's rows. headers=None means a CSV: when its first row holds
    the table's column names it is a header (columns matched by name), otherwise
    the row values are matched to the table's columns by position.
    """
    cursor = conn.cursor()

    # Resolve the actual schema (e.g. 'lvapdbf'), not the login default ('dbo').
    cursor.execute(
        "SELECT TABLE_SCHEMA FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = ?",
        table_name,
    )
    schemas = [r[0] for r in cursor.fetchall()]
    if not schemas:
        return {"name": table_name, "status": "skip", "message": "Table not found in DB"}
    # Import into the chosen schema; if the table doesn't exist there, skip it.
    if preferred_schema not in schemas:
        return {"name": table_name, "status": "skip",
                "message": f"Not in schema '{preferred_schema}' (found in: {', '.join(schemas)})"}
    schema = preferred_schema
    qualified = f"{_sql_identifier(schema)}.{_sql_identifier(table_name)}"

    # Full column metadata: type + nullability + whether it has a DB default.
    cursor.execute(
        "SELECT COLUMN_NAME, DATA_TYPE, IS_NULLABLE, COLUMN_DEFAULT "
        "FROM INFORMATION_SCHEMA.COLUMNS "
        "WHERE TABLE_NAME = ? AND TABLE_SCHEMA = ? ORDER BY ORDINAL_POSITION",
        table_name, schema,
    )
    db_cols = {}
    for name, dtype, nullable, default in cursor.fetchall():
        db_cols[name.upper()] = {
            "name": name,
            "type": (dtype or "").lower(),
            "nullable": nullable == "YES",
            "has_default": default is not None,
        }

    extra_csv_cols = 0
    is_csv = headers is None
    csv_header = False
    if is_csv and rows and csv_has_header(rows[0], [info["name"] for info in db_cols.values()]):
        csv_header = True
        headers = [normalize_cell(cell) for cell in rows[0]]
        rows = rows[1:]
        types = [db_cols[h.upper()]["type"] if h.upper() in db_cols else None for h in headers]
        rows = [[_csv_value(v, t) for v, t in zip(row, types)] for row in rows]
    elif is_csv:
        width = max((len(row) for row in rows), default=0)
        ordered = list(db_cols.values())
        headers = [info["name"] for info in ordered[:width]]
        extra_csv_cols = max(0, width - len(ordered))
        types = [info["type"] for info in ordered[:width]]
        rows = [[_csv_value(v, t) for v, t in zip(row, types)] for row in rows]

    # Which sheet header maps to which DB column?
    sheet_idx = {h.upper(): i for i, h in enumerate(headers) if h.upper() in db_cols}
    if not sheet_idx:
        return {"name": table_name, "status": "skip", "message": "No matching columns"}

    # Build the column list to INSERT:
    #  • sheet columns      → take the cell value (NULL→default if NOT NULL)
    #  • missing NOT NULL    → supply a type default (unless the DB has its own default)
    #  • everything else     → omit (let DB default / NULL apply)
    plan = []   # (db_col_name, source)  source = ("sheet", idx) | ("fill", value)
    filled_cols = []
    for up, info in db_cols.items():
        if up in sheet_idx:
            plan.append((info["name"], ("sheet", sheet_idx[up])))
        elif not info["nullable"] and not info["has_default"]:
            plan.append((info["name"], ("fill", _type_default(info["type"]))))
            filled_cols.append(info["name"])

    # TRUNCATE is minimally logged and skips per-row delete cost; fall back to
    # DELETE only if it's blocked (e.g. a foreign key references this table).
    clear_sql = f"TRUNCATE TABLE {qualified}"
    try:
        cursor.execute(clear_sql)
    except pyodbc.Error:
        conn.rollback()
        clear_sql = f"DELETE FROM {qualified}"
        cursor.execute(clear_sql)

    if sql_export_path is not None:
        _write_import_sql(sql_export_path, schema, table_name, plan, rows, db_cols)

    cols_sql     = ", ".join(_sql_identifier(c) for c, _ in plan)
    placeholders = ", ".join("?" for _ in plan)
    insert_sql   = f"INSERT INTO {qualified} ({cols_sql}) VALUES ({placeholders})"

    total_rows = len(rows)
    inserted = errors = 0
    error_samples = []
    batch = []
    if progress_cb:
        progress_cb(0, total_rows)
    for row in rows:
        batch.append(_values_for_insert(row, plan, db_cols))
        if len(batch) >= IMPORT_BATCH_SIZE:
            batch_inserted, batch_errors = _insert_batch(cursor, insert_sql, batch, error_samples)
            inserted += batch_inserted
            errors += batch_errors
            batch = []
            if progress_cb:
                progress_cb(inserted + errors, total_rows)
    if batch:
        batch_inserted, batch_errors = _insert_batch(cursor, insert_sql, batch, error_samples)
        inserted += batch_inserted
        errors += batch_errors

    conn.commit()
    if progress_cb:
        progress_cb(inserted + errors, total_rows)
    # Blank header cells (padding past the last named column) are not counted.
    unmapped = sum(1 for h in headers if h and h.upper() not in db_cols)
    msg = f"{inserted} rows inserted"
    if errors:
        msg += f", {errors} errors"
    if filled_cols:
        msg += f" ({len(filled_cols)} required cols auto-filled: {', '.join(filled_cols)})"
    if csv_header:
        msg += " (header row found — columns matched by name)"
    if unmapped:
        msg += f" ({unmapped} {'CSV' if is_csv else 'sheet'} cols skipped — not in table)"
    if extra_csv_cols:
        msg += f" ({extra_csv_cols} CSV cols skipped — table has only {len(db_cols)} columns)"

    return {
        "name": table_name,
        "status": "ok" if not errors else "warn",
        "message": msg,
        "inserted": inserted,
        "errors": error_samples,
    }


STRING_TYPES = {"char", "varchar", "nchar", "nvarchar", "text", "ntext"}


def _type_default(dtype: str):
    """A safe non-NULL default for a SQL Server column type."""
    numeric = {
        "int", "bigint", "smallint", "tinyint", "bit",
        "decimal", "numeric", "money", "smallmoney", "float", "real",
    }
    dates = {"date", "datetime", "datetime2", "smalldatetime", "datetimeoffset"}
    if dtype in numeric:
        return 0
    if dtype == "time":
        return "00:00:00"
    if dtype in dates:
        return "1900-01-01"
    return ""   # char / varchar / nchar / nvarchar / text → empty string


if __name__ == "__main__":
    debug = os.environ.get("FLASK_DEBUG", "").lower() in {"1", "true", "yes"}
    app.run(host="0.0.0.0", port=5000, debug=debug)
