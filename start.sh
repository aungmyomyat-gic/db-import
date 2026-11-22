#!/usr/bin/env bash
# Start the existing app locally from Git Bash, Linux, or macOS.
set -euo pipefail

cd -- "$(dirname -- "${BASH_SOURCE[0]}")"

bootstrap=()
if command -v py >/dev/null 2>&1 && py -3 -c 'import sys; sys.exit(sys.version_info < (3, 10))' >/dev/null 2>&1; then
  bootstrap=(py -3)
else
  for candidate in python3 python; do
    if command -v "$candidate" >/dev/null 2>&1 && "$candidate" -c 'import sys; sys.exit(sys.version_info < (3, 10))' >/dev/null 2>&1; then
      bootstrap=("$candidate")
      break
    fi
  done
fi

if [[ ${#bootstrap[@]} -eq 0 ]]; then
  echo 'Install Python 3.10 or newer, then run bash start.sh again.' >&2
  echo 'Windows: https://www.python.org/downloads/windows/' >&2
  exit 1
fi

if [[ ! -f .venv/Scripts/python.exe && ! -f .venv/bin/python ]]; then
  echo 'Preparing the application (first run only)...'
  "${bootstrap[@]}" -m venv .venv
fi
if [[ -f .venv/Scripts/python.exe ]]; then
  app_python=.venv/Scripts/python.exe
else
  app_python=.venv/bin/python
fi

if ! cmp -s requirements.txt .venv/.requirements-installed; then
  "$app_python" -m pip install --disable-pip-version-check -r requirements.txt
  cp requirements.txt .venv/.requirements-installed
fi

"$app_python" - <<'PY'
import sys
try:
    import pyodbc
    if not any('SQL Server' in driver for driver in pyodbc.drivers()):
        raise RuntimeError('No SQL Server ODBC driver is installed.')
except (ImportError, RuntimeError) as error:
    print(f'ODBC setup required: {error}', file=sys.stderr)
    print('Install Microsoft ODBC Driver 18 for SQL Server, then run bash start.sh again.', file=sys.stderr)
    print('https://learn.microsoft.com/sql/connect/odbc/download-odbc-driver-for-sql-server', file=sys.stderr)
    print('Linux also requires the unixODBC driver manager.', file=sys.stderr)
    sys.exit(1)
PY

# Resolve with Python so Windows receives a native path, even under Git Bash.
export DB_CONFIG_PATH="${DB_CONFIG_PATH:-$("$app_python" -c 'from pathlib import Path; print(Path("data/db_config.json").resolve())')}"

exec "$app_python" - <<'PY'
import os
import threading
import webbrowser

from werkzeug.serving import make_server
from app import app

port = int(os.environ.get('DB_IMPORT_PORT', '5000'))
server = make_server('127.0.0.1', port, app, threaded=True)
url = f'http://127.0.0.1:{server.server_port}'
print(f'\nDB Import: {url}\nKeep this terminal open. Press Ctrl+C to stop.\n', flush=True)

def open_browser():
    try:
        webbrowser.open(url)
    except Exception:
        pass  # The URL above also works when automatic browser launch is unavailable.

timer = threading.Timer(0.5, open_browser)
timer.daemon = True
timer.start()
try:
    server.serve_forever()
except KeyboardInterrupt:
    print('\nStopping DB Import.', flush=True)
finally:
    server.server_close()
PY
