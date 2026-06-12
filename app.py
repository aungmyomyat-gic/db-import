import re
import os
import json
import unicodedata
import tempfile
import uuid
from pathlib import Path

from flask import Flask, render_template, request, jsonify, session
import openpyxl
from openpyxl.utils import get_column_letter
import pyodbc

app = Flask(__name__)
app.secret_key = os.urandom(24)

# Persisted DB connection (port, username, password, database, host).
# Stored next to the app so it survives restarts; mount it as a volume to
# keep it across container rebuilds.
CONFIG_PATH = Path(os.environ.get("DB_CONFIG_PATH", "/app/data/db_config.json"))
SCAN_CACHE_PATH = Path(os.environ.get(
    "DB_SCAN_CACHE_PATH",
    str(CONFIG_PATH.with_name("scan_cache.json")),
))


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


def load_scan_cache() -> dict:
    if SCAN_CACHE_PATH.exists():
        try:
            return json.loads(SCAN_CACHE_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_scan_cache(cache: dict) -> None:
    try:
        SCAN_CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
        SCAN_CACHE_PATH.write_text(
            json.dumps(cache, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
    except Exception:
        # Cache is an optimization only. Import should still work without it.
        pass

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


def extract_headers(ws, meta):
    hrow, _ = resolve_header_row(ws, meta)
    start_col = int(meta["start_col"])
    cached_end_col = meta.get("end_col")

    headers, col = [], start_col
    while col <= int(cached_end_col or ws.max_column):
        val = ws.cell(row=hrow, column=col).value
        if val is None and cached_end_col is None:
            break
        # Normalize header codes (may be full-width) so they match DB columns.
        headers.append(normalize_cell(str(val)) if val is not None else "")
        col += 1
    end_col = start_col + len(headers) - 1 if headers else int(cached_end_col or start_col)
    return hrow, headers, end_col


def extract_table(ws, meta):
    hrow, headers, _ = extract_headers(ws, meta)
    start_col = int(meta["start_col"])
    if not headers:
        return [], []
    n_cols = len(headers)
    rows, r = [], hrow + 1
    while r <= ws.max_row:
        row_vals = [ws.cell(row=r, column=start_col + i).value for i in range(n_cols)]
        if all(v is None for v in row_vals):
            break
        rows.append(row_vals)
        r += 1
    return headers, rows


def table_preview(ws, meta, from_cache=False):
    hrow, headers, end_col = extract_headers(ws, meta)
    hrow, marker_skipped = resolve_header_row(ws, meta)
    start_col = int(meta["start_col"])
    data_start_row = hrow + 1
    start_letter = get_column_letter(start_col)
    end_letter = get_column_letter(end_col)
    row_range = f"{data_start_row}+"
    cell_range = f"{start_letter}{hrow}:{end_letter}..."
    return {
        "name": meta["name"],
        "columns": len(headers),
        "records": None,
        "name_row": int(meta["name_row"]),
        "header_row": hrow,
        "data_start_row": data_start_row,
        "data_end_row": None,
        "start_col": start_col,
        "end_col": end_col,
        "start_col_letter": start_letter,
        "end_col_letter": end_letter,
        "col_range": f"{start_letter}:{end_letter}",
        "row_range": row_range,
        "cell_range": cell_range,
        "marker_skipped": marker_skipped,
        "from_cache": from_cache,
    }


def cache_meta_from_preview(table):
    return {
        "name": table["name"],
        "name_row": table["name_row"],
        "header_row": table["header_row"] - 1 if table["marker_skipped"] else table["header_row"],
        "start_col": table["start_col"],
        "end_col": table["end_col"],
    }


def cached_metas_for_sheet(ws, sheet_name):
    entry = load_scan_cache().get("sheets", {}).get(sheet_name)
    if not entry:
        return None

    metas = entry.get("tables") or []
    if not metas:
        return None

    for meta in metas:
        try:
            name_row = int(meta["name_row"])
            start_col = int(meta["start_col"])
            expected_name = meta["name"]
        except (KeyError, TypeError, ValueError):
            return None
        if name_row > ws.max_row or start_col > ws.max_column:
            return None
        value = ws.cell(row=name_row, column=start_col).value
        if not isinstance(value, str):
            return None
        match = TABLE_NAME_RE.match(normalize_cell(value))
        if not match or match.group(1) != expected_name:
            return None

    return metas


def save_sheet_scan_cache(sheet_name, tables):
    cache = load_scan_cache()
    cache.setdefault("version", 1)
    cache.setdefault("sheets", {})
    cache["sheets"][sheet_name] = {
        "tables": [cache_meta_from_preview(table) for table in tables],
    }
    save_scan_cache(cache)


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


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


@app.route("/upload", methods=["POST"])
def upload():
    """Save uploaded file, return sheet names."""
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify({"error": "No file selected."}), 400

    tmp = Path(tempfile.gettempdir()) / f"{uuid.uuid4()}.xlsx"
    file.save(tmp)

    try:
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
    return jsonify({"sheets": sheets})


@app.route("/scan", methods=["POST"])
def scan():
    """Scan a sheet and return table preview."""
    payload = request.json or {}
    sheet_name = payload.get("sheet", "").strip()
    refresh_cache = bool(payload.get("refresh"))
    tmp_path = session.get("tmp_path")

    if not tmp_path or not Path(tmp_path).exists():
        return jsonify({"error": "Session expired — please re-upload the file."}), 400
    if not sheet_name:
        return jsonify({"error": "Sheet name is required."}), 400

    try:
        wb = openpyxl.load_workbook(tmp_path, data_only=True, read_only=True)
    except Exception as e:
        return jsonify({"error": f"Cannot open file: {e}"}), 400

    if sheet_name not in wb.sheetnames:
        return jsonify({"error": f"Sheet '{sheet_name}' not found."}), 400

    ws = wb[sheet_name]
    cache_used = False
    metas = None if refresh_cache else cached_metas_for_sheet(ws, sheet_name)
    if metas is not None:
        cache_used = True
    else:
        metas = scan_tables(ws)

    tables = [table_preview(ws, meta, from_cache=cache_used) for meta in metas]
    if cache_used and any(table["columns"] == 0 for table in tables):
        cache_used = False
        metas = scan_tables(ws)
        tables = [table_preview(ws, meta, from_cache=False) for meta in metas]
    if not cache_used:
        save_sheet_scan_cache(sheet_name, tables)

    wb.close()
    session["sheet_name"] = sheet_name
    return jsonify({
        "tables": tables,
        "cache": {
            "used": cache_used,
            "tables": len(tables),
            "path": str(SCAN_CACHE_PATH),
        },
    })


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


@app.route("/import", methods=["POST"])
def do_import():
    """Connect to DB (using saved config) and import all scanned tables."""
    cfg = load_config()
    if not cfg:
        return jsonify({"error": "No DB connection saved. Please set it up first."}), 400

    tmp_path   = session.get("tmp_path")
    sheet_name = session.get("sheet_name")

    # Only import the tables the user ticked (None = all, for backward compat).
    selected = (request.json or {}).get("tables")
    selected_set = set(selected) if selected is not None else None
    if selected_set is not None and not selected_set:
        return jsonify({"error": "No tables selected."}), 400

    if not tmp_path or not Path(tmp_path).exists():
        return jsonify({"error": "Session expired — please re-upload the file."}), 400

    conn_str, err = _build_conn_str(
        cfg["host"], cfg["port"], cfg["database"], cfg["username"], cfg["password"]
    )
    if err:
        return jsonify({"error": err}), 500
    try:
        conn = pyodbc.connect(conn_str, timeout=10)
    except pyodbc.Error as e:
        return jsonify({"error": f"Connection failed: {e}"}), 400

    # Schema chosen in the UI dropdown wins; fall back to saved config, then default.
    preferred_schema = (request.json or {}).get("schema", "").strip() \
        or cfg.get("schema") or "lvapdbf"

    wb = openpyxl.load_workbook(tmp_path, data_only=True, read_only=True)
    ws = wb[sheet_name]
    metas = cached_metas_for_sheet(ws, sheet_name)
    if metas is None:
        metas = scan_tables(ws)

    results = []
    for meta in metas:
        if selected_set is not None and meta["name"] not in selected_set:
            continue
        headers, rows = extract_table(ws, meta)
        if not headers:
            results.append({"name": meta["name"], "status": "skip", "message": "Header row empty"})
            continue
        try:
            results.append(_import_one(conn, meta["name"], headers, rows, preferred_schema))
        except Exception as e:
            conn.rollback()
            results.append({"name": meta["name"], "status": "skip", "message": f"Failed: {e}"})

    conn.close()
    wb.close()
    return jsonify({"results": results})


def _import_one(conn, table_name, headers, rows, preferred_schema="lvapdbf"):
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
    qualified = f"[{schema}].[{table_name}]"

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

    cursor.execute(f"DELETE FROM {qualified}")

    cols_sql     = ", ".join(f"[{c}]" for c, _ in plan)
    placeholders = ", ".join("?" for _ in plan)
    insert_sql   = f"INSERT INTO {qualified} ({cols_sql}) VALUES ({placeholders})"

    inserted = errors = 0
    error_samples = []
    for row in rows:
        values = []
        for col_name, src in plan:
            if src[0] == "sheet":
                v = row[src[1]]
                if v is None and not db_cols[col_name.upper()]["nullable"]:
                    v = _type_default(db_cols[col_name.upper()]["type"])
                values.append(v)
            else:  # "fill"
                values.append(src[1])
        try:
            cursor.execute(insert_sql, values)
            inserted += 1
        except pyodbc.Error as e:
            errors += 1
            if errors <= 2:
                error_samples.append(str(e))

    conn.commit()
    unmapped = len(headers) - len(sheet_idx)
    msg = f"{inserted} rows inserted"
    if errors:
        msg += f", {errors} errors"
    if filled_cols:
        msg += f" ({len(filled_cols)} required cols auto-filled: {', '.join(filled_cols)})"
    if unmapped:
        msg += f" ({unmapped} sheet cols skipped — not in table)"

    return {
        "name": table_name,
        "status": "ok" if not errors else "warn",
        "message": msg,
        "inserted": inserted,
        "errors": error_samples,
    }


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
    app.run(host="0.0.0.0", port=5000, debug=False)
