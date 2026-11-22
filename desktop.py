"""Local desktop entry point, also bundled as db-import.exe on Windows."""

import argparse
import ctypes
import json
import os
import sys
import threading
import webbrowser
from pathlib import Path


ODBC_DOWNLOAD_URL = "https://learn.microsoft.com/sql/connect/odbc/download-odbc-driver-for-sql-server"


def config_path():
    """Keep user settings outside the executable and its temporary extraction."""
    override = os.environ.get("DB_CONFIG_PATH")
    if override:
        return Path(override).expanduser().resolve()
    if sys.platform == "win32":
        base = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
    else:
        base = Path(os.environ.get("XDG_DATA_HOME") or Path.home() / ".local" / "share")
    return base / "DBImport" / "db_config.json"


def report_error(message, dialog=True):
    print(f"DB Import could not start:\n{message}", file=sys.stderr, flush=True)
    if dialog and sys.platform == "win32":
        try:
            ctypes.windll.user32.MessageBoxW(None, message, "DB Import", 0x10)
        except (AttributeError, OSError):
            pass


def open_browser(url):
    try:
        webbrowser.open(url)
    except Exception:
        pass  # The console also prints the URL.


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description="Start DB Import and open its local browser interface.")
    parser.add_argument("--port", type=int, default=os.environ.get("DB_IMPORT_PORT", "0"),
                        help="Local port; 0 selects an available port (default).")
    parser.add_argument("--no-browser", action="store_true", help="Print the URL without opening a browser.")
    parser.add_argument("--self-test", action="store_true", help="Check bundled resources and ODBC without connecting to a database.")
    args = parser.parse_args(argv)
    if not 0 <= args.port <= 65535:
        parser.error("port must be between 0 and 65535")
    return args


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="backslashreplace")
    args = parse_args(argv)
    dialog = not (args.no_browser or args.self_test)
    try:
        path = config_path()
        os.environ["DB_CONFIG_PATH"] = str(path)
        # Setting DB_CONFIG_PATH before importing the application also works in
        # a one-file bundle, where __file__ points into temporary extraction.
        import app as importer
        from werkzeug.serving import make_server

        importer.CONFIG_PATH = path
        drivers = importer.pyodbc.drivers()
        driver = importer.select_sql_server_driver(drivers)
        if not driver:
            report_error(
                "Microsoft ODBC Driver 18 or 17 for SQL Server is required.\n"
                "An existing compatible driver can be reused. Install the x64 driver, then reopen DB Import.\n"
                + ODBC_DOWNLOAD_URL, dialog=dialog,
            )
            return 2
        if args.self_test:
            with importer.app.test_client() as client:
                page = client.get("/")
                if page.status_code != 200 or b'id="import-cancel"' not in page.data or b"SQL Generate" not in page.data:
                    raise RuntimeError("The bundled interface could not be loaded.")
            print(json.dumps({"ok": True, "driver": driver, "config_path": str(path)}), flush=True)
            return 0
        path.parent.mkdir(parents=True, exist_ok=True)
        server = make_server("127.0.0.1", args.port, importer.app, threaded=True)
    except ImportError as error:
        report_error(
            f"A required application library could not be loaded:\n{error}\n\n"
            "If Windows reports an application-control policy block, ask your administrator "
            "to approve the application and its libraries, or use an approved Windows build PC. "
            "Installing another ODBC driver will not resolve a policy block on pyodbc.", dialog=dialog,
        )
        return 1
    except (Exception, SystemExit) as error:
        # Werkzeug can raise SystemExit when a requested port is already in use.
        report_error(f"Unable to start the application ({error}).\n"
                     "If the port is in use, reopen DB Import without --port to choose an available port.", dialog=dialog)
        return 1

    url = f"http://127.0.0.1:{server.server_port}"
    print(f"\nDB Import: {url}\nSettings: {path}\n"
          "Keep this window open. Close it or press Ctrl+C to stop DB Import.\n", flush=True)
    timer = None
    if not args.no_browser:
        timer = threading.Timer(0.5, open_browser, args=(url,))
        timer.daemon = True
        timer.start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping DB Import.", flush=True)
    finally:
        if timer:
            timer.cancel()
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
