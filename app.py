import csv
import io
import re
import os
import json
import shutil
import threading
import time
import unicodedata
import tempfile
import uuid
from datetime import date, datetime, time as datetime_time
from decimal import Decimal, InvalidOperation
from pathlib import Path

from flask import Flask, render_template, request, jsonify, session, send_file
import pyodbc

app = Flask(__name__, template_folder=str(Path(__file__).resolve().parent / "templates"))
app.secret_key = os.urandom(24)
CONFIG_PATH = Path(os.environ.get("DB_CONFIG_PATH", Path(__file__).resolve().parent / "data" / "db_config.json"))
IMPORT_BATCH_SIZE = int(os.environ.get("IMPORT_BATCH_SIZE", "5000"))
IMPORT_JOB_TTL_SECONDS = 3600
IMPORT_JOBS = {}
IMPORT_JOBS_LOCK = threading.Lock()
SQL_EXPORT_DIR = tempfile.TemporaryDirectory(prefix="db-import-sql-")
UPLOAD_DIR = tempfile.TemporaryDirectory(prefix="db-import-csv-")
WORKFLOWS = {"import", "generate"}
STRING_TYPES = {"char", "varchar", "nchar", "nvarchar", "text", "ntext"}
_INVISIBLE_RE = re.compile(r"[​‌‍﻿ ]")


def load_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            return cfg if isinstance(cfg, dict) and any(cfg.values()) else {}
        except Exception:
            return {}
    return {}


def save_config(cfg: dict) -> None:
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")


def normalize_cell(value: str) -> str:
    """
    Normalize full-width characters in CSV filenames and strip invisible
    characters that can break table-name matching.
    """
    s = unicodedata.normalize("NFKC", value)
    s = _INVISIBLE_RE.sub("", s)
    return s.strip()


CSV_TABLE_NAME_RE = re.compile(r"[A-Z][A-Z0-9_]*")
CSV_ENCODINGS = ("utf-8-sig", "cp932")   # UTF-8 (with/without BOM), then Shift-JIS


def is_csv_filename(filename):
    return bool(filename) and filename.lower().endswith(".csv")


def csv_table_name(filename):
    stem = normalize_cell(Path(filename or "").stem).upper()
    m = CSV_TABLE_NAME_RE.match(stem)
    return m.group(0) if m else None


def read_csv_rows(path, cancel_cb=None, progress_cb=None):
    """Read every non-blank row; short rows are padded so all rows have the same width."""
    raw = Path(path).read_bytes()
    if cancel_cb:
        cancel_cb()
    for encoding in CSV_ENCODINGS:
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError("CSV is not UTF-8 or Shift-JIS text")
    rows = []
    stream = io.StringIO(text, newline="")
    if progress_cb:
        progress_cb(0, len(text), 0)
    for index, row in enumerate(csv.reader(stream)):
        if cancel_cb and index % 1000 == 0:
            cancel_cb()
        if any(cell.strip() for cell in row):
            rows.append(row)
        if progress_cb and index % 1000 == 0:
            progress_cb(stream.tell(), len(text), len(rows))
    if not rows:
        raise ValueError("CSV contains no data rows.")
    width = max(len(row) for row in rows)
    rows = [row + [""] * (width - len(row)) for row in rows]
    if progress_cb:
        progress_cb(len(text), len(text), len(rows))
    return rows, width


def _csv_preview(table_name, rows, width):
    return {
        "name": table_name,
        "columns": width,
        "records": len(rows),
        "col_range": "No header",
        "row_range": f"1:{len(rows)}" if rows else "-",
    }


def _csv_value(value, dtype):
    """CSV text -> insert value. Blank becomes NULL (NOT NULL columns get a default later)."""
    if dtype not in STRING_TYPES:
        value = value.strip()
    return None if value == "" else value


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
    schema   = data.get("schema", "").strip()

    if not all([host, port, database, username, schema]):
        return jsonify({"error": "Host, port, database, username, and schema are required."}), 400

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


def select_sql_server_driver(drivers):
    """Prefer Driver 18, then 17; avoid Windows' legacy 'SQL Server' driver."""
    available = {name.casefold(): name for name in drivers}
    for version in (18, 17):
        name = f"ODBC Driver {version} for SQL Server"
        if name.casefold() in available:
            return available[name.casefold()]
    return None


def _build_conn_str(host, port, database, username, password):
    driver = select_sql_server_driver(pyodbc.drivers())
    if not driver:
        return None, "No supported SQL Server ODBC driver found. Install Microsoft ODBC Driver 18 or 17."
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
    """Upload a CSV; optionally read its preview in a job with live progress."""
    purpose = request.form.get("purpose", "import")
    if purpose not in WORKFLOWS:
        return jsonify({"error": "Unknown CSV workflow."}), 400
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify({"error": "No file selected."}), 400
    if not is_csv_filename(file.filename):
        return jsonify({"error": "Select a .csv file with no header row."}), 400
    table_name = csv_table_name(file.filename)
    if not table_name or len(table_name) > 128:
        return jsonify({"error": "CSV file name must start with the table name, e.g. MDNNOP.csv."}), 400
    tmp = Path(UPLOAD_DIR.name) / f"{uuid.uuid4().hex}.csv"
    try:
        file.save(tmp)
        if request.form.get("background") == "1":
            return _start_csv_preview(tmp, table_name, file.filename, purpose)
        rows, width = read_csv_rows(tmp)
    except Exception as e:
        tmp.unlink(missing_ok=True)
        return jsonify({"error": f"Cannot open CSV: {e}"}), 400

    uploads = dict(session.get("uploads", {}))
    previous = uploads.get(purpose)
    if previous:
        Path(previous["path"]).unlink(missing_ok=True)
    uploads[purpose] = {"path": str(tmp), "table": table_name, "filename": file.filename}
    session["uploads"] = uploads
    return jsonify({"kind": "csv", "table": table_name,
                    "preview": _csv_preview(table_name, rows, width)})


def _start_csv_preview(tmp, table_name, filename, purpose):
    _cleanup_import_jobs()
    job_id = uuid.uuid4().hex
    ready_path = Path(UPLOAD_DIR.name) / f"ready-{job_id}.csv"
    job = _new_import_job(job_id, [table_name], "preview")
    job["owner"] = session.setdefault("import_owner", uuid.uuid4().hex)
    job["file_size"] = tmp.stat().st_size
    with IMPORT_JOBS_LOCK:
        IMPORT_JOBS[job_id] = job
        snapshot = _job_snapshot(job)
    worker = threading.Thread(target=_run_csv_preview,
                              args=(job_id, tmp, ready_path, table_name), daemon=True)
    try:
        worker.start()
    except Exception:
        tmp.unlink(missing_ok=True)
        _fail_import_job(job_id, "Could not start CSV reading. Please try again.")
        return jsonify({"error": "Could not start CSV reading. Please try again."}), 500
    uploads = dict(session.get("uploads", {}))
    previous = uploads.get(purpose)
    if previous:
        Path(previous["path"]).unlink(missing_ok=True)
    # This path becomes available only after the whole CSV is validated.
    uploads[purpose] = {"path": str(ready_path), "table": table_name, "filename": filename}
    session["uploads"] = uploads
    return jsonify(snapshot), 202


def _run_csv_preview(job_id, tmp, ready_path, table_name):
    def report(done, total, records):
        with IMPORT_JOBS_LOCK:
            job = IMPORT_JOBS.get(job_id)
            if job:
                job["results"][0]["progress"] = {"done": done, "total": total, "records": records}
                job["updated_at"] = time.time()

    try:
        _set_job_state(job_id, "running")
        _mark_job_table(job_id, table_name, "running", "Reading CSV")
        rows, width = read_csv_rows(tmp, progress_cb=report)
        with IMPORT_JOBS_LOCK:
            IMPORT_JOBS[job_id]["preview"] = _csv_preview(table_name, rows, width)
        tmp.replace(ready_path)
        _mark_job_table(job_id, table_name, "ok", "CSV ready")
        _finish_import_job(job_id)
    except Exception as e:
        ready_path.unlink(missing_ok=True)
        _fail_import_job(job_id, f"Cannot open CSV: {e}")
    finally:
        tmp.unlink(missing_ok=True)


@app.route("/scan", methods=["POST"])
def scan():
    """Refresh the preview of the CSV selected for one workflow."""
    purpose = (request.json or {}).get("purpose", "import")
    if purpose not in WORKFLOWS:
        return jsonify({"error": "Unknown CSV workflow."}), 400
    uploaded = session.get("uploads", {}).get(purpose)
    if not uploaded or not Path(uploaded["path"]).is_file():
        return jsonify({"error": "Session expired. Please upload the CSV again."}), 400
    try:
        rows, width = read_csv_rows(uploaded["path"])
    except Exception as e:
        return jsonify({"error": f"Cannot open CSV: {e}"}), 400
    return jsonify({"tables": [_csv_preview(uploaded["table"], rows, width)], "kind": "csv"})


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
    return jsonify({"schemas": schemas, "default": cfg.get("schema") or ""})


@app.route("/import", methods=["POST"])
def do_import():
    return _start_csv_job("import")


@app.route("/generate-sql", methods=["POST"])
def generate_sql():
    """Generate a SQL download from CSV using only database metadata reads."""
    return _start_csv_job("generate")


def _start_csv_job(operation):
    cfg = load_config()
    if not cfg:
        return jsonify({"error": "No DB connection saved. Please set it up first."}), 400
    uploaded = session.get("uploads", {}).get(operation)
    if not uploaded or not Path(uploaded["path"]).is_file():
        return jsonify({"error": "Session expired. Please upload the CSV again."}), 400
    payload = request.json or {}
    selected = payload.get("tables", [uploaded["table"]])
    if selected != [uploaded["table"]]:
        return jsonify({"error": "Select the table from the uploaded CSV."}), 400
    schema = str(payload.get("schema", "")).strip() or cfg.get("schema")
    if not schema:
        return jsonify({"error": "Select a target schema."}), 400
    _, err = _build_conn_str(
        cfg["host"], cfg["port"], cfg["database"], cfg["username"], cfg["password"]
    )
    if err:
        return jsonify({"error": err}), 500

    _cleanup_import_jobs()
    job_id = uuid.uuid4().hex
    # Each worker owns a copy: replacing an upload cannot change an active job.
    job_input = Path(UPLOAD_DIR.name) / f"job-{job_id}.csv"
    try:
        shutil.copyfile(uploaded["path"], job_input)
    except OSError as e:
        job_input.unlink(missing_ok=True)
        return jsonify({"error": f"Cannot prepare CSV: {e}"}), 400
    job = _new_import_job(job_id, [uploaded["table"]], operation)
    job["owner"] = session.setdefault("import_owner", uuid.uuid4().hex)
    with IMPORT_JOBS_LOCK:
        IMPORT_JOBS[job_id] = job
        snapshot = _job_snapshot(job)
    worker = threading.Thread(
        target=_run_import_job,
        args=(job_id, cfg, str(job_input), uploaded["table"], schema, operation),
        daemon=True,
    )
    try:
        worker.start()
    except Exception as e:
        job_input.unlink(missing_ok=True)
        _fail_import_job(job_id, f"Could not start job: {e}")
        return jsonify({"error": "Could not start CSV processing. Please try again."}), 500
    return jsonify(snapshot), 202


@app.route("/jobs/<job_id>", methods=["GET"])
@app.route("/import/status/<job_id>", methods=["GET"])
def import_status(job_id):
    _cleanup_import_jobs()
    with IMPORT_JOBS_LOCK:
        job = IMPORT_JOBS.get(job_id)
        if not job or not job.get("owner") or job["owner"] != session.get("import_owner"):
            return jsonify({"error": "CSV job not found."}), 404
        return jsonify(_job_snapshot(job))


@app.route("/import/cancel/<job_id>", methods=["POST"])
def cancel_import(job_id):
    _cleanup_import_jobs()
    with IMPORT_JOBS_LOCK:
        job = IMPORT_JOBS.get(job_id)
        if not job or not job.get("owner") or job["owner"] != session.get("import_owner"):
            return jsonify({"error": "CSV job not found."}), 404
        if job["operation"] != "import":
            return jsonify({"error": "Only CSV imports can be cancelled."}), 409
        if job["state"] == "committing":
            return jsonify({"error": "The import is saving its changes and can no longer be cancelled."}), 409
        if job["state"] in {"done", "failed", "cancelled"}:
            return jsonify(_job_snapshot(job))
        job["cancel_requested"] = True
        job["state"] = "cancelling"
        job["updated_at"] = time.time()
        return jsonify(_job_snapshot(job)), 202


@app.route("/sql/<job_id>", methods=["GET"])
@app.route("/import/sql/<job_id>", methods=["GET"])
def download_import_sql(job_id):
    _cleanup_import_jobs()
    with IMPORT_JOBS_LOCK:
        job = IMPORT_JOBS.get(job_id)
        if not job or not job.get("owner") or job["owner"] != session.get("import_owner"):
            return jsonify({"error": "CSV job not found."}), 404
        if job["state"] != "done" or not job.get("sql_path"):
            return jsonify({"error": "SQL is available only after successful CSV processing."}), 409
        if not Path(job["sql_path"]).is_file():
            return jsonify({"error": "SQL download expired. Please process the CSV again."}), 404
        return send_file(job["sql_path"], as_attachment=True, download_name=job["sql_filename"],
                         mimetype="application/sql", conditional=False, max_age=0)


def _new_import_job(job_id, table_names, operation="import"):
    now = time.time()
    return {
        "id": job_id,
        "state": "queued",
        "operation": operation,
        "cancel_requested": False,
        "started_at": time.perf_counter(),
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
        "operation": job["operation"],
        "status_url": f"/jobs/{job['id']}",
        "can_cancel": job["operation"] == "import" and job["state"] in {"queued", "running"}
                      and not job["cancel_requested"],
        "elapsed_seconds": round(max(0, job.get("finished_at", time.perf_counter())
                                     - job["started_at"]), 3),
        "total": job["total"],
        "completed": job["completed"],
        "results": [dict(result) for result in job["results"]],
        "error": job.get("error"),
    }
    if job["operation"] == "preview":
        snapshot["file_size"] = job["file_size"]
        if job["state"] == "done":
            snapshot["preview"] = job["preview"]
    if job["state"] == "done" and job.get("sql_path"):
        snapshot["sql_download_url"] = f"/sql/{job['id']}"
        snapshot["sql_filename"] = job["sql_filename"]
    return snapshot


def _cleanup_import_jobs():
    cutoff = time.time() - IMPORT_JOB_TTL_SECONDS
    with IMPORT_JOBS_LOCK:
        for job_id, job in list(IMPORT_JOBS.items()):
            if job["state"] in {"done", "failed", "cancelled"} and job["updated_at"] < cutoff:
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
        job["state"] = "cancelling" if state == "running" and job["cancel_requested"] else state
        job["error"] = error
        job["updated_at"] = time.time()


def _mark_job_table(job_id, table_name, status, message, errors=None, schema=None):
    final_statuses = {"ok", "warn", "skip", "failed", "cancelled"}
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
        if job["state"] not in {"failed", "cancelled"}:
            statuses = {r["status"] for r in job["results"]}
            job["state"] = "failed" if "failed" in statuses else "cancelled" if "cancelled" in statuses else "done"
        job.setdefault("finished_at", time.perf_counter())
        job["updated_at"] = time.time()


def _fail_import_job(job_id, message):
    with IMPORT_JOBS_LOCK:
        job = IMPORT_JOBS.get(job_id)
        if not job:
            return
        final_statuses = {"ok", "warn", "skip", "failed", "cancelled"}
        for result in job["results"]:
            if result["status"] not in final_statuses:
                result["status"] = "failed"
                result["message"] = message
                result["errors"] = []
                job["completed"] += 1
        job["state"] = "failed"
        job.setdefault("finished_at", time.perf_counter())
        job["error"] = message
        job["updated_at"] = time.time()


class ImportCancelled(Exception):
    """Raised at a cancellation checkpoint before the import commits."""


def _check_import_cancelled(job_id):
    with IMPORT_JOBS_LOCK:
        job = IMPORT_JOBS.get(job_id)
        if not job or job["cancel_requested"]:
            raise ImportCancelled()


def _begin_import_commit(job_id):
    # Claim the commit phase under the same lock as cancellation. Once claimed,
    # a cancellation request cannot promise to undo changes being committed.
    with IMPORT_JOBS_LOCK:
        job = IMPORT_JOBS.get(job_id)
        if not job or job["cancel_requested"]:
            raise ImportCancelled()
        job["state"] = "committing"
        for result in job["results"]:
            result["message"] = "Saving changes…"
        job["updated_at"] = time.time()


def _cancel_import_job(job_id):
    with IMPORT_JOBS_LOCK:
        job = IMPORT_JOBS.get(job_id)
        if not job:
            return
        for result in job["results"]:
            result.update(status="cancelled", message="Import cancelled. No changes were saved.", errors=[])
        job["completed"] = len(job["results"])
        job["state"] = "cancelled"
        job.setdefault("finished_at", time.perf_counter())
        job["updated_at"] = time.time()


def _run_import_job(job_id, cfg, tmp_path, table_name, schema, operation="import"):
    conn = None
    failure = None
    cancelled = False
    try:
        if operation == "import":
            _check_import_cancelled(job_id)
        _set_job_state(job_id, "running")
        conn_str, err = _build_conn_str(
            cfg["host"], cfg["port"], cfg["database"], cfg["username"], cfg["password"]
        )
        if err:
            raise RuntimeError(err)
        if operation == "import":
            _check_import_cancelled(job_id)
        conn = pyodbc.connect(conn_str, timeout=10, autocommit=False)
        if operation == "import":
            _check_import_cancelled(job_id)
        _run_csv_import(job_id, conn, tmp_path, table_name, [table_name], schema, operation)
    except ImportCancelled:
        try:
            if conn is not None:
                conn.rollback()
        except Exception as e:
            failure = f"Could not roll back cancelled import: {e}"
        else:
            cancelled = True
    except Exception as e:
        failure = f"CSV processing failed: {e}"
    finally:
        if conn is not None:
            try:
                conn.close()
            except Exception:
                pass
        try:
            Path(tmp_path).unlink(missing_ok=True)
        except OSError:
            app.logger.warning("Could not remove a completed job's input file.")
        # Publish completion only after releasing the connection and input.
        if failure:
            _fail_import_job(job_id, failure)
        elif cancelled:
            _cancel_import_job(job_id)
        else:
            _finish_import_job(job_id)


def _run_csv_import(job_id, conn, tmp_path, table_name, selected_names, schema, operation="import"):
    if selected_names is None:
        _replace_job_queue(job_id, [table_name])
    elif table_name not in selected_names:
        for name in selected_names:
            _mark_job_table(job_id, name, "skip", "Table not found in file", [])
        return
    action = "Importing" if operation == "import" else "Generating SQL"
    _mark_job_table(job_id, table_name, "running", action, schema=schema)
    sql_path = Path(SQL_EXPORT_DIR.name) / f"{job_id}.sql"
    cancel_cb = (lambda: _check_import_cancelled(job_id)) if operation == "import" else None
    try:
        if cancel_cb:
            cancel_cb()
        rows, _ = read_csv_rows(tmp_path, cancel_cb=cancel_cb)
        progress_cb = lambda done, total: _mark_job_progress(job_id, table_name, done, total)
        if operation == "import":
            result = _import_one(conn, table_name, rows, schema, progress_cb,
                                 sql_export_path=sql_path, cancel_cb=cancel_cb,
                                 before_commit_cb=lambda: _begin_import_commit(job_id))
        else:
            prepared = _prepare_csv(conn.cursor(), table_name, rows, schema)
            _write_import_sql(sql_path, schema, table_name, prepared["plan"],
                              prepared["rows"], prepared["columns"])
            progress_cb(len(rows), len(rows))
            result = {"status": "ok", "message": f"SQL generated for {len(rows)} rows"
                      + prepared["message_suffix"], "errors": []}
    except ImportCancelled:
        try:
            conn.rollback()
        except Exception as e:
            result = {"status": "failed", "message": f"Could not roll back cancelled import: {e}", "errors": []}
        else:
            result = {"status": "cancelled", "message": "Import cancelled. No changes were saved.", "errors": []}
    except Exception as e:
        if operation == "import":
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
    # Keep ordinary names readable, e.g. dbo.MSHBRP. Delimit special names
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


def _write_import_sql(path, schema, table_name, plan, rows, db_cols, cancel_cb=None):
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
            if cancel_cb:
                cancel_cb()
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
        if src[0] == "csv":
            v = row[src[1]]
            if v is None and not db_cols[col_name.upper()]["nullable"]:
                v = _type_default(db_cols[col_name.upper()]["type"])
            values.append(v)
        else:  # "fill"
            values.append(src[1])
    return tuple(values)


def _insert_batch(cursor, insert_sql, batch, error_samples, cancel_cb=None):
    if cancel_cb:
        cancel_cb()
    cursor.execute("SAVE TRANSACTION import_batch")
    try:
        cursor.fast_executemany = True
        cursor.executemany(insert_sql, batch)
        if cancel_cb:
            cancel_cb()
        return len(batch), 0
    except pyodbc.Error:
        if cancel_cb:
            cancel_cb()
        cursor.fast_executemany = False
        cursor.execute("ROLLBACK TRANSACTION import_batch")

    inserted = errors = 0
    for values in batch:
        if cancel_cb:
            cancel_cb()
        try:
            cursor.execute(insert_sql, values)
            inserted += 1
        except pyodbc.Error as e:
            errors += 1
            if len(error_samples) < 2:
                error_samples.append(str(e))
    if cancel_cb:
        cancel_cb()
    return inserted, errors


def _prepare_csv(cursor, table_name, rows, schema, cancel_cb=None):
    """Resolve column types and defaults once for both CSV workflows; read only."""
    cursor.execute(
        "SELECT TABLE_SCHEMA FROM INFORMATION_SCHEMA.TABLES "
        "WHERE TABLE_NAME = ? AND TABLE_TYPE = 'BASE TABLE'", table_name,
    )
    schemas = [r[0] for r in cursor.fetchall()]
    if cancel_cb:
        cancel_cb()
    if not schemas:
        raise ValueError("Table not found in DB")
    if schema not in schemas:
        raise ValueError(f"Not in schema '{schema}' (found in: {', '.join(schemas)})")
    cursor.execute(
        "SELECT COLUMN_NAME, DATA_TYPE, IS_NULLABLE, COLUMN_DEFAULT "
        "FROM INFORMATION_SCHEMA.COLUMNS "
        "WHERE TABLE_NAME = ? AND TABLE_SCHEMA = ? ORDER BY ORDINAL_POSITION",
        table_name, schema,
    )
    db_cols = {}
    for name, dtype, nullable, default in cursor.fetchall():
        db_cols[name.upper()] = {
            "name": name, "type": (dtype or "").lower(),
            "nullable": nullable == "YES", "has_default": default is not None,
        }
    if cancel_cb:
        cancel_cb()
    if not db_cols:
        raise ValueError("No matching columns")
    width = max((len(row) for row in rows), default=0)
    if not width:
        raise ValueError("CSV contains no data rows.")
    ordered = list(db_cols.values())
    types = [info["type"] for info in ordered[:width]]
    converted = []
    for index, row in enumerate(rows):
        if cancel_cb and index % 1000 == 0:
            cancel_cb()
        converted.append([_csv_value(value, dtype) for value, dtype in zip(row, types)])
    plan, filled = [], []
    for index, info in enumerate(ordered):
        if index < width:
            plan.append((info["name"], ("csv", index)))
        elif not info["nullable"] and not info["has_default"]:
            plan.append((info["name"], ("fill", _type_default(info["type"]))))
            filled.append(info["name"])
    suffix = ""
    if filled:
        suffix += f" ({len(filled)} required cols auto-filled: {', '.join(filled)})"
    if width > len(ordered):
        suffix += f" ({width - len(ordered)} CSV cols skipped — table has only {len(ordered)} columns)"
    return {"rows": converted, "columns": db_cols, "plan": plan, "message_suffix": suffix}


def _import_one(conn, table_name, rows, preferred_schema, progress_cb=None,
                sql_export_path=None, cancel_cb=None, before_commit_cb=None):
    """Replace one table's rows and write the SQL that reproduces the CSV values."""
    cursor = conn.cursor()
    schema = preferred_schema
    if cancel_cb:
        cancel_cb()
    try:
        prepared = _prepare_csv(cursor, table_name, rows, schema, cancel_cb=cancel_cb)
    except ValueError as e:
        return {"name": table_name, "status": "skip", "message": str(e)}
    rows, db_cols, plan = prepared["rows"], prepared["columns"], prepared["plan"]
    qualified = f"{_sql_identifier(schema)}.{_sql_identifier(table_name)}"
    if cancel_cb:
        cancel_cb()

    # TRUNCATE is minimally logged and skips per-row delete cost; fall back to
    # DELETE only if it's blocked (e.g. a foreign key references this table).
    clear_sql = f"TRUNCATE TABLE {qualified}"
    try:
        cursor.execute(clear_sql)
    except pyodbc.Error:
        conn.rollback()
        if cancel_cb:
            cancel_cb()
        clear_sql = f"DELETE FROM {qualified}"
        cursor.execute(clear_sql)

    if cancel_cb:
        cancel_cb()
    if sql_export_path is not None:
        _write_import_sql(sql_export_path, schema, table_name, plan, rows, db_cols, cancel_cb=cancel_cb)

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
            batch_inserted, batch_errors = _insert_batch(cursor, insert_sql, batch, error_samples, cancel_cb)
            inserted += batch_inserted
            errors += batch_errors
            batch = []
            if progress_cb:
                progress_cb(inserted + errors, total_rows)
    if batch:
        batch_inserted, batch_errors = _insert_batch(cursor, insert_sql, batch, error_samples, cancel_cb)
        inserted += batch_inserted
        errors += batch_errors

    if cancel_cb:
        cancel_cb()
    if before_commit_cb:
        before_commit_cb()
    conn.commit()
    if progress_cb:
        progress_cb(inserted + errors, total_rows)
    msg = f"{inserted} rows inserted"
    if errors:
        msg += f", {errors} errors"
    msg += prepared["message_suffix"]

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
    debug = os.environ.get("FLASK_DEBUG", "").lower() in {"1", "true", "yes"}
    app.run(host="0.0.0.0", port=5000, debug=debug)
