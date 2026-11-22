"""Desktop configuration, driver selection, and startup checks; no real database."""

import contextlib
import builtins
import io
import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

import app as importer
import desktop


DRIVER_18 = "ODBC Driver 18 for SQL Server"
DRIVER_17 = "ODBC Driver 17 for SQL Server"


class DriverSelectionTests(unittest.TestCase):
    def test_modern_driver_is_preferred_over_windows_legacy_driver(self):
        for drivers in ([DRIVER_17, "SQL Server", DRIVER_18], [DRIVER_18, DRIVER_17, "SQL Server"]):
            with self.subTest(drivers=drivers):
                self.assertEqual(importer.select_sql_server_driver(drivers), DRIVER_18)
                with patch.object(importer.pyodbc, "drivers", return_value=drivers):
                    connection, error = importer._build_conn_str("server", "1433", "db", "user", "password")
                self.assertIsNone(error)
                self.assertIn(f"DRIVER={{{DRIVER_18}}}", connection)

    def test_existing_driver_17_is_reused(self):
        with patch.object(importer.pyodbc, "drivers", return_value=["SQL Server", DRIVER_17]):
            connection, error = importer._build_conn_str("server", "1433", "db", "user", "password")
        self.assertIsNone(error)
        self.assertIn(f"DRIVER={{{DRIVER_17}}}", connection)

    def test_missing_or_legacy_only_driver_has_installation_message(self):
        for drivers in ([], ["SQL Server"], ["ODBC Driver 13 for SQL Server"], ["PostgreSQL Unicode"]):
            with self.subTest(drivers=drivers), patch.object(importer.pyodbc, "drivers", return_value=drivers):
                connection, error = importer._build_conn_str("server", "1433", "db", "user", "password")
                self.assertIsNone(connection)
                self.assertIn("18 or 17", error)


class DesktopTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.directory = Path(temporary.name)
        self.user_config = self.directory / "profile with spaces" / "DBImport" / "db_config.json"
        self.addCleanup(patch.stopall)
        patch.dict(os.environ, {"DB_CONFIG_PATH": str(self.user_config)}, clear=False).start()
        patch.object(importer, "CONFIG_PATH", self.user_config).start()
        self.driver = patch.object(importer.pyodbc, "drivers", return_value=[DRIVER_17]).start()
        self.output = io.StringIO()
        self.addCleanup(self.output.close)

    def test_windows_settings_live_in_user_profile(self):
        profile = self.directory / "Local AppData 日本語"
        with patch.dict(os.environ, {"LOCALAPPDATA": str(profile)}, clear=True), \
             patch.object(desktop.sys, "platform", "win32"):
            self.assertEqual(desktop.config_path(), profile / "DBImport" / "db_config.json")

    def test_windows_settings_fallback_still_uses_user_home(self):
        with patch.dict(os.environ, {}, clear=True), \
             patch.object(desktop.sys, "platform", "win32"), \
             patch.object(Path, "home", return_value=self.directory):
            self.assertEqual(desktop.config_path(), self.directory / "AppData" / "Local" / "DBImport" / "db_config.json")

    def test_explicit_settings_path_takes_precedence(self):
        self.assertEqual(desktop.config_path(), self.user_config)

    def test_self_test_finds_template_from_an_unrelated_working_directory(self):
        previous = Path.cwd()
        try:
            os.chdir(self.directory)
            with contextlib.redirect_stdout(self.output), patch.object(importer.pyodbc, "connect") as connect:
                self.assertEqual(desktop.main(["--self-test"]), 0)
                connect.assert_not_called()
        finally:
            os.chdir(previous)
        report = json.loads(self.output.getvalue())
        self.assertTrue(report["ok"])
        self.assertEqual(report["driver"], DRIVER_17)
        self.assertEqual(Path(report["config_path"]), self.user_config)
        self.assertFalse(self.user_config.exists())

    def test_missing_driver_reports_help_without_opening_a_server(self):
        self.driver.return_value = ["SQL Server"]
        with patch.object(desktop, "report_error") as report, \
             patch("werkzeug.serving.make_server") as make_server:
            self.assertEqual(desktop.main(["--no-browser"]), 2)
            make_server.assert_not_called()
        self.assertIn(desktop.ODBC_DOWNLOAD_URL, report.call_args.args[0])
        self.assertFalse(report.call_args.kwargs["dialog"])

    def test_launcher_binds_only_to_loopback_and_preserves_user_settings(self):
        self.user_config.parent.mkdir(parents=True)
        config = {"host": "test.invalid", "database": "saved", "username": "test"}
        self.user_config.write_text(json.dumps(config), encoding="utf-8")
        server = MagicMock(server_port=54321)
        server.serve_forever.side_effect = KeyboardInterrupt
        with patch("werkzeug.serving.make_server", return_value=server) as make_server, \
             patch.object(desktop.webbrowser, "open") as browser, contextlib.redirect_stdout(self.output):
            self.assertEqual(desktop.main(["--no-browser", "--port", "0"]), 0)
        make_server.assert_called_once_with("127.0.0.1", 0, importer.app, threaded=True)
        browser.assert_not_called()
        server.server_close.assert_called_once()
        self.assertIn("http://127.0.0.1:54321", self.output.getvalue())
        self.assertEqual(importer.load_config(), config)
        self.assertEqual(json.loads(self.user_config.read_text(encoding="utf-8")), config)

    def test_browser_uses_actual_server_port_and_timer_is_cancelled_on_exit(self):
        server = MagicMock(server_port=54321)
        server.serve_forever.side_effect = KeyboardInterrupt
        with patch("werkzeug.serving.make_server", return_value=server), \
             patch.object(desktop.threading, "Timer") as timer, contextlib.redirect_stdout(self.output):
            self.assertEqual(desktop.main(["--port", "0"]), 0)
        timer.assert_called_once_with(0.5, desktop.open_browser, args=("http://127.0.0.1:54321",))
        self.assertTrue(timer.return_value.daemon)
        timer.return_value.start.assert_called_once()
        timer.return_value.cancel.assert_called_once()
        server.server_close.assert_called_once()

    def test_bind_failure_is_reported_instead_of_an_unhandled_exit(self):
        with patch("werkzeug.serving.make_server", side_effect=SystemExit(1)), \
             patch.object(desktop, "report_error") as report:
            self.assertEqual(desktop.main(["--no-browser", "--port", "5000"]), 1)
        self.assertIn("without --port", report.call_args.args[0])

    def test_blocked_native_library_reports_policy_help_instead_of_port_help(self):
        original_import = builtins.__import__

        def blocked_import(name, *args, **kwargs):
            if name == "app":
                raise ImportError("DLL load failed while importing pyodbc: application control policy blocked this file")
            return original_import(name, *args, **kwargs)

        with patch.object(builtins, "__import__", side_effect=blocked_import), \
             patch.object(desktop, "report_error") as report:
            self.assertEqual(desktop.main(["--no-browser"]), 1)
        message = report.call_args.args[0]
        self.assertIn("administrator", message)
        self.assertIn("pyodbc", message)
        self.assertNotIn("port is in use", message)

    def test_unwritable_settings_parent_is_reported_before_server_start(self):
        self.user_config.parent.parent.mkdir(parents=True)
        self.user_config.parent.write_text("this is a file", encoding="utf-8")
        with patch.object(desktop, "report_error") as report, patch("werkzeug.serving.make_server") as make_server:
            self.assertEqual(desktop.main(["--no-browser"]), 1)
            make_server.assert_not_called()
        self.assertIn("Unable to start", report.call_args.args[0])

    def test_invalid_ports_are_rejected(self):
        for port in ("-1", "65536", "text"):
            with self.subTest(port=port), contextlib.redirect_stderr(self.output), self.assertRaises(SystemExit) as error:
                desktop.parse_args(["--port", port])
            self.assertEqual(error.exception.code, 2)


if __name__ == "__main__":
    unittest.main()
