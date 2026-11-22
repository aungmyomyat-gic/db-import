# DB Import — CSV & SQL

This branch contains two tools: **CSV Import** and **SQL Generate**, with shared
SQL Server connection settings. This branch contains the Windows application
and its build tools.

## Windows executable

Double-click **db-import.exe** to open the interface in your browser. The receiving
PC needs a compatible 64-bit SQL Server ODBC Driver 18 or 17. Python is bundled;
Git and Docker are unnecessary for the executable. Saved connection settings use
`%LOCALAPPDATA%\DBImport\db_config.json`. On first launch, all connection fields
are blank; enter your own details and click **Test & Save connection**.

Build on Windows by running **build_windows.bat**. It runs the tests, packages a
single executable, and smoke-tests it. The output is **dist\db-import.exe**.
The `feat/windows-exe` branch also includes a Windows GitHub Actions build.
See [WINDOWS.md](WINDOWS.md) for build, download, and sharing instructions.

## CSV Import

1. Save your SQL Server details under **Connection**.
2. Open **CSV Import** and choose a CSV with **no header row**. The file name
   determines the table: `MDNNOP.csv` or `MDNNOP (1).csv` → `MDNNOP`.
3. Check the table, row count, and column count, then choose the target schema.
4. Click **Import CSV & Generate SQL**. The import replaces the existing rows
   in that table. Progress and elapsed time appear while it runs; the completed
   duration stays in the result.
   Use **Cancel import** to stop before changes are saved. Cancellation waits
   for the current operation to finish, then rolls back the import and keeps
   the original table rows. The result shows **Cancelled** and its elapsed time.
   Cancellation is unavailable during the final save. Cancelled imports do not
   offer a SQL download.
5. After a successful import, click **Download SQL** to save the matching script.
   Imports with row errors do not offer a SQL download.

## SQL Generate

1. Open **SQL Generate** and choose a headerless CSV named after its table.
2. Choose the schema and click **Generate SQL**.
3. Click **Download SQL**. Generation reads column metadata from the connected
   database without importing, deleting, or changing database rows.

Each tab keeps its own CSV selection. Both tools need an existing matching
SQL Server table to determine column order, types, nullability, and defaults.
Selecting a CSV shows its file size, upload percentage, then reading percentage
and rows read. The reading percentage follows the characters parsed from the
file; it reaches 100% when validation finishes and the CSV is ready to use.

## CSV and SQL format

- UTF-8 (with or without BOM) and Shift-JIS CSVs are supported. Empty files and
  files containing only blank rows are rejected.
- Columns map by position. Short rows are padded with blanks. Extra CSV columns
  are skipped and reported in the result.
- Blank values become NULL, or the type's default for NOT NULL columns. Missing
  required columns without database defaults are filled automatically; omitted
  nullable columns and columns with database defaults are left to SQL Server.
- The UTF-8 SQL file starts with `DELETE FROM schema.table;`, followed by
  `INSERT INTO` and a `VALUES` block. Inserts are split into at most 1,000 rows.
- Numeric values are unquoted; text uses `N'...'`, preserving leading zeros,
  Japanese text, and escaped quotes. Partial mappings include a column list so
  omitted columns still receive database defaults.
- Running a generated script replaces the selected table's rows. Database
  defaults and triggers run again. Generation uses metadata to format the
  values; SQL Server validates constraints when the script is executed.
- Downloads belong to the browser session and expire one hour after completion
  or when the app restarts. Timing covers server processing, including the
  connection, CSV conversion, import where applicable, and SQL generation.

## Run from source

With Python 3.10 or newer and Microsoft ODBC Driver 18 or 17 installed:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe desktop.py
```

The launcher opens the interface on an available local port and saves connection
settings under `%LOCALAPPDATA%\DBImport`. Keep the console open while using the
app. Close it or press **Ctrl+C** to stop. See [WINDOWS.md](WINDOWS.md) for
executable builds and sharing instructions.

## Development

Install `requirements.txt` into a Python environment, then run:

```bash
python -m unittest discover -s tests -v
```

The tests use a simulated database and do not change SQL Server data. They cover
CSV/SQL value parity, generation without database writes, timing, upload
isolation, failure handling, download access and expiry, and removed routes.

| File | Purpose |
| --- | --- |
| `app.py` | Flask backend, CSV conversion, background jobs, SQL exports |
| `templates/index.html` | CSV Import, SQL Generate, and Connection UI |
| `desktop.py` | Executable launcher and user-profile settings |
| `DBImport.spec` / `build_windows.ps1` | Single-file Windows build and verification |
| `.github/workflows/windows-exe.yml` | Windows CI build and downloadable artifact |
