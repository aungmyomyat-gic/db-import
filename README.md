# DB Import — CSV & SQL

This branch contains two tools: **CSV Import** and **SQL Generate**, with shared
SQL Server connection settings. It runs locally or in Docker at
http://localhost:5000.

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

## Run locally

Install Python 3.10 or newer and Microsoft ODBC Driver 18 for SQL Server.
On Windows, use Git Bash. Linux/macOS also need the ODBC driver manager and
Python `venv` support.

From this project folder, run:

```bash
bash start.sh
```

The launcher creates `.venv`, installs dependencies on the first run, checks the
ODBC driver, and opens http://127.0.0.1:5000. Keep the terminal open; press
**Ctrl+C** to stop. To choose another port:

```bash
DB_IMPORT_PORT=5055 bash start.sh
```

Connection settings, including the password, are saved in `data/db_config.json`.
Use the actual SQL Server hostname/IP, or `localhost` for a local SQL Server.

## Run in Docker

With Docker installed, run from this branch:

```bash
docker compose up -d --build
```

Open http://localhost:5000. The supplied compose file uses the
`icl-lv-core-system_sazaby-network` network; start the corresponding database
stack first, or adapt the network for your SQL Server. For a database running
on the Docker host, use `host.docker.internal` where available.

The source and `data` folder are mounted into the container. Rebuild when
`requirements.txt` or `Dockerfile` changes. Stop with `docker compose down`.

## Updating this branch

Run `./update.sh` (Linux/macOS/Git Bash) or `update.bat` (Windows) after this
branch is available on the remote. They pull the **current branch** and rebuild
the container, keeping saved connection settings. This branch has no in-app
updater or release-merging script.

See [SHARE.md](SHARE.md) for sharing this CSV and SQL version.

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
| `start.sh` | Local setup and launcher |
| `update.sh` / `update.bat` | Update the current branch and restart Docker |
