"""Exercise a built executable without connecting to or modifying a database."""

import argparse
import json
import os
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request
from pathlib import Path


def read_json(url):
    with urllib.request.urlopen(url, timeout=5) as response:
        return json.load(response)


def command_for(path):
    return [sys.executable, str(path)] if path.suffix == ".py" else [str(path)]


def self_test(command, cwd, environment, expected_config):
    result = subprocess.run(command + ["--self-test"], cwd=cwd, env=environment,
                            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=45)
    if result.returncode:
        raise RuntimeError(f"Executable self-test failed:\n{result.stdout}\n{result.stderr}")
    report = json.loads(result.stdout.strip())
    assert report["ok"] is True, report
    assert report["driver"] in {"ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server"}, report
    assert Path(report["config_path"]) == expected_config, report


def smoke_test(path):
    path = path.resolve(strict=True)
    with tempfile.TemporaryDirectory(prefix="dbimport-smoke-") as temporary:
        directory = Path(temporary)
        # Moving the single-file executable proves it does not need the checkout.
        relocated = directory / "App 日本語"
        relocated.mkdir()
        if path.suffix != ".py":
            copied = relocated / path.name
            shutil.copyfile(path, copied)
            copied.chmod(0o755)
            command = command_for(copied)
        else:
            command = command_for(path)
        environment = os.environ.copy()
        environment.pop("DB_CONFIG_PATH", None)
        environment.pop("DB_IMPORT_PORT", None)
        profile = directory / "User settings 日本語"
        environment["LOCALAPPDATA"] = str(profile)
        environment["XDG_DATA_HOME"] = str(profile)
        environment["PYTHONUTF8"] = "1"
        expected_config = profile / "DBImport" / "db_config.json"
        self_test(command, relocated, environment, expected_config)
        process = subprocess.Popen(command + ["--no-browser", "--port", "0"], cwd=relocated,
                                   env=environment, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                   text=True, encoding="utf-8", errors="replace", bufsize=1)
        lines = queue.Queue()

        def read_output():
            for line in process.stdout:
                lines.put(line)

        reader = threading.Thread(target=read_output, daemon=True)
        reader.start()
        output = []
        try:
            deadline = time.monotonic() + 45
            url = None
            while time.monotonic() < deadline:
                try:
                    line = lines.get(timeout=0.2)
                except queue.Empty:
                    if process.poll() is not None:
                        break
                    continue
                output.append(line)
                match = re.search(r"DB Import: (http://127\.0\.0\.1:\d+)", line)
                if match:
                    url = match.group(1)
                    break
            if not url:
                raise RuntimeError("Executable did not start:\n" + "".join(output))
            with urllib.request.urlopen(url, timeout=5) as response:
                html = response.read().decode("utf-8")
                assert response.status == 200
                for control in ('id="nav-import"', 'id="nav-generate"', 'id="import-cancel"'):
                    assert control in html, f"Missing interface control: {control}"
            assert read_json(url + "/config")["configured"] is False, "The bundle included saved connection settings"
            boundary = "dbimport-smoke-boundary"
            body = (f"--{boundary}\r\nContent-Disposition: form-data; name=\"file\"; filename=\"MDNNOP.csv\"\r\n"
                    "Content-Type: text/csv\r\n\r\n0001,日本語\r\n"
                    f"--{boundary}--\r\n").encode("utf-8")
            request = urllib.request.Request(url + "/upload", data=body,
                                             headers={"Content-Type": f"multipart/form-data; boundary={boundary}"})
            with urllib.request.urlopen(request, timeout=5) as response:
                preview = json.load(response)["preview"]
                assert (preview["name"], preview["records"], preview["columns"]) == ("MDNNOP", 1, 2)
            # Settings must be read from the user profile, independently of the
            # executable location. These are intentionally fake; no DB is used.
            config = {"host": "test.invalid", "port": "1433", "database": "smoke-test", "username": "test", "password": "", "schema": "dbo"}
            expected_config.write_text(json.dumps(config), encoding="utf-8")
            assert read_json(url + "/config")["config"] == config
        finally:
            if process.poll() is None:
                if sys.platform == "win32":
                    # A one-file bundle has a bootloader parent and Python child.
                    subprocess.run(["taskkill", "/PID", str(process.pid), "/T", "/F"],
                                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
                else:
                    process.terminate()
            try:
                process.wait(timeout=10)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait(timeout=5)
            reader.join(timeout=2)
            process.stdout.close()
        # Reopening the bundle keeps settings from the previous run.
        self_test(command, relocated, environment, expected_config)
        assert json.loads(expected_config.read_text(encoding="utf-8")) == config
        print("Executable smoke test passed: relocated app, Unicode paths, bundled UI, CSV parsing, separate settings, and restart persistence.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("executable", type=Path)
    smoke_test(parser.parse_args().executable)
