# DB Import for Windows

## Run the application 

1. Extract the downloaded `db-import-Windows-x64` ZIP to a folder.
2. Double-click **db-import.exe**. It opens the app in your default browser.
3. Keep the console window open while using the app. Close it or press **Ctrl+C**
   when finished. Closing the browser tab leaves the application running.
4. Set up **Connection**, then use **CSV Import** or **SQL Generate**. CSV Import
   includes progress, elapsed time, cancellation, and a SQL download on success.

The executable includes Python and its packages. The receiving PC does not need
Python, Git, Docker, or this source repository. Use 64-bit Windows 10/11.
An existing compatible **64-bit Microsoft ODBC Driver 18 or 17 for SQL Server**
is reused; Driver 18 is preferred. The driver must be installed on the PC running
db-import.exe, even when the database server is on another computer.

If no supported driver is found, the launcher shows the installation instructions.
Install [Microsoft ODBC Driver 18 for SQL Server (x64)](https://learn.microsoft.com/sql/connect/odbc/download-odbc-driver-for-sql-server)
and reopen the app. Older driver installers may also require the Visual C++
runtime described on that page. Installing SQL Server itself does not confirm
that a compatible ODBC driver is available.

Connection settings, including the password, are stored in:

```text
%LOCALAPPDATA%\DBImport\db_config.json
```

Replacing or moving the executable keeps those settings. Each recipient enters
their own connection; all connection fields start blank, and the executable
contains no saved database credentials.
CSV uploads and SQL downloads are temporary and expire when the app closes.

The app listens only on `127.0.0.1`. It selects an available local port each time,
and prints the URL if the browser does not open automatically. Advanced options:

```powershell
.\db-import.exe --port 5000
.\db-import.exe --no-browser
.\db-import.exe --self-test
```

`--self-test` checks the ODBC runtime and bundled interface without connecting to
a database. To change the settings location, set `DB_CONFIG_PATH` before starting.
Starting from source with `python desktop.py` uses the same user settings location.

## Build on a Windows PC

Builds require **64-bit Python 3.10 or newer** (3.12 recommended), internet access
for the first package installation, and the ODBC driver used by the smoke test.
The receiving PC only needs the executable and ODBC driver.

From this source folder, double-click **build_windows.bat**, or run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build_windows.ps1
```

The build creates `.venv-build`, installs `requirements-build.txt`, runs the
automated tests, packages `DBImport.spec`, and smoke-tests the actual executable
from a separate folder with Unicode characters in its path. No database is used
by these checks. The output is:

```text
dist\db-import.exe
dist\db-import.exe.sha256
dist\WINDOWS.md
```

Share those files. The smoke test uses temporary, fake settings and checks that
saved credentials from the source checkout are absent from the package. The
SHA-256 file lets recipients verify the executable downloaded correctly.
The executable is unsigned; organizational policy may require your organization
to sign or approve it before distribution.

If tests or startup report **DLL load failed while importing pyodbc** followed
by an **application control policy** message, Windows has blocked the Python
database library. This is separate from an absent ODBC driver. Ask the Windows
administrator to approve the application and libraries, or build on an approved
Windows PC or GitHub Actions runner. The build stops when verification fails.

## Build with GitHub Actions

Push the `feat/windows-exe` branch to the repository. Its **Build Windows
executable** workflow runs on Windows, tests the app, builds the executable,
and uploads the **db-import-Windows-x64** artifact. Open the completed run in
**Actions** and download that artifact. Artifacts are kept for 14 days.

After the workflow exists on the default branch, it can also be started manually
with **Actions → Build Windows executable → Run workflow** and the desired branch.
PyInstaller produces executables for the operating system on which it runs;
use this Windows build rather than renaming a Linux executable to `.exe`.
