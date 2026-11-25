# DB Import

A small web app for loading Excel test data into SQL Server and working with
the tables afterwards. It runs in Docker and opens at http://localhost:5000.

## Features

- **Import** — upload an Excel file, pick a sheet, and the app finds every table
  under the **■ 実施前テストデータ** section (names like `MZAIHP(倉庫在庫マスタ)`).
  Preview them, choose which to import, and insert them into the target schema.
  If the file name contains `day job 12` / `month job 3`, the matching batch job
  can be triggered after the import.
- **CSV import** — choose a `.csv` file and the file name is used as the table
  name (`MDNNOP.csv` or `MDNNOP (1).csv` → `MDNNOP`). The file may have a header
  row or not. If more than half of the first row's cells are column names of
  the table, that row is the header and columns are matched **by name** (any
  order; header columns not in the table are skipped). Otherwise every row is
  data and columns are matched **by position**. The preview shows which one was
  found. Blank values become NULL, or the type's default for NOT NULL columns.
  UTF-8 and Shift-JIS files both work.
  After a successful CSV import, **Download SQL** appears in Import Results.
  The UTF-8 `.sql` file starts with `DELETE FROM schema.table;`, followed by
  `INSERT INTO schema.table` and a `VALUES` block with one tuple per row.
  Numeric columns use unquoted numbers; text uses `N'...'` literals, preserving
  leading zeros, Japanese text, and escaped quotes. NULLs and required-column
  defaults match the import. Full-width mappings omit the column list; partial
  mappings include it so database defaults still apply to omitted columns.
  Large exports split inserts into groups of at most 1,000 rows. Run the script
  against an existing matching table to replace its rows. Database defaults and
  triggers run again. Imports with row errors do not offer a SQL download.
  Downloads expire one hour after completion, or when the app restarts.
- **CSV → TSV** — for CSV files from systems that do not quote values, where a
  file name such as `a,b,c.pdf` is split into several cells. The table comes
  from the file name and its columns are read from the database. In the
  table's file-name columns (`MSHHKP`: `SHKTRM`, `SHKFL1`, `SHKFL2`, `SHKFL3`;
  more tables can be added to `COMMA_FILE_COLUMNS` in `app.py`). A line with
  one cell per table column (empty cells at the end of the line are dropped)
  just has its commas changed to tabs. Otherwise, in the file-name columns, a
  cell that is neither blank nor ending in one of the known extensions
  (`FILE_EXTENSIONS` in `app.py`: pdf, docx, xlsx, jpg, xls, doc, png, txt,
  tif, jpeg, pptx, gif, ppt, any letter case) is joined with the next cell, so
  text like `JN75 0.35GR` is not taken for the end of a file name.
  Every other comma becomes a tab. A value that itself contains a tab (e.g. a
  memo pasted from Excel) is written in double quotes, `"a<TAB>b"`, so it stays
  one value; TSV Import removes the quotes and keeps the tab. A header row is kept, and
  the encoding (UTF-8 or Shift-JIS) and line endings stay the same. Rows that
  do not come out to exactly one value per column are left out of the TSV and
  listed with their CSV line number so they can be fixed by hand.
- **TSV Import** — import a tab-separated `.tsv` file, such as the one made by
  CSV → TSV. It works like CSV import (file name = table name, header row
  matched by name or values matched by position, blank → NULL/default, the
  table's rows are replaced, Download SQL afterwards). Tabs separate values;
  a value in double quotes may contain tabs, and the quotes are removed
  (`""` inside stands for one `"`), as CSV → TSV writes it. Before the
  import, a preview shows the row count, whether a header was found, and the
  lines that do not have one value per column; the import asks for
  confirmation.
- **Truncate Table** — empty one or more tables in a schema.
- **Data Editor** — browse every record in any table:
  - pages of 1,000, 2,000, 5,000 or 10,000 rows (the choice is remembered)
  - row numbers, and a summary of primary keys, total rows and columns
  - sort by clicking a column header (sorted by the database, across all
    pages); search columns by name
  - edit rows in place (Edit mode)
  - copy the current page to Excel, with or without the header
  - download every row as a UTF-8 CSV, in the current sort order. The
    file without a header can be imported again with CSV import
  - paste rows copied from Excel (Ctrl+V), review them, then insert
- **SQL Runner** — run all or selected SQL with schema-aware table-name and SQL
  keyword suggestions, multiple result grids, write-query confirmation, and
  one-click copying for the first result.
- **Update notification** — when a new version is released, the app shows an
  "Update available" popup with the changes.

## Install

You need **Git** and **Docker Desktop**. The repo is private, so you need read
access on GitHub.

```bash
git clone https://github.com/aungmyomyat-gic/db-import.git
cd db-import
docker compose up -d --build
```

`docker-compose.yml` joins the `icl-lv-core-system_sazaby-network` Docker
network, so start the ICL core system first. The project folder is mounted
into the container, so a `git pull` updates the app straight away.

Open http://localhost:5000, go to **Connection**, and enter your SQL Server
details. They are saved in `./data` and kept across updates.

> If SQL Server runs on your own computer (not in Docker), use
> `host.docker.internal` as the host.

## Run locally with a shell script

On a Windows customer PC, install **Git for Windows (Git Bash)**, **Python
3.10 or newer**, and **Microsoft ODBC Driver 18 for SQL Server** once.
The driver is available from [Microsoft](https://learn.microsoft.com/sql/connect/odbc/download-odbc-driver-for-sql-server).
For Linux/macOS, install the platform's SQL Server driver and ODBC driver
manager as well as Python with `venv` support.

Extract or copy this project folder to the PC. Open Git Bash in the folder
(or a terminal on Linux/macOS), then run:

```bash
bash start.sh
```

The script creates a local `.venv`, installs the Python packages on the first
run (internet access required), checks the ODBC driver, and opens the existing
UI at http://127.0.0.1:5000. Later runs reuse the installed packages. Keep the
terminal open while using the app; press **Ctrl+C** to stop it.

Enter the database connection under **Connection**, select a headerless CSV
named after its table, import it, then click **Download SQL** in the results.
Connection settings use `data/db_config.json`, including the password, just
as in the Docker version. Use the actual database server hostname/IP, or
`localhost` for SQL Server on that PC. To change the local app port:

```bash
DB_IMPORT_PORT=5055 bash start.sh
```

`setup.sh` continues to start the Docker version.

## Update

When the app shows **Update available**, double-click `update.bat` (Windows) or
run `./update.sh` (Mac / Linux) in the `db-import` folder. Both run:

```bash
git pull --ff-only origin main
docker compose up -d --build
```

Then reload the page. The `--build` step is quick when nothing changed; it only
really rebuilds when `Dockerfile` or `requirements.txt` changed.

## Stop / start

```bash
docker compose down     # stop
docker compose up -d    # start again
```

## Troubleshooting

| Problem | Fix |
| --- | --- |
| `No SQL Server ODBC driver found inside container` | Rebuild: `docker compose up -d --build` |
| Can't connect to a database on your own PC | Use `host.docker.internal` as the host and check SQL Server listens on port 1433 |
| Port 5000 is already in use | Change `"5000:5000"` to e.g. `"5050:5000"` in `docker-compose.yml`, then open http://localhost:5050 |
| `git pull` fails in the update script | You have local changes in the folder. Run `git status`, then `git stash` or discard them, and run the update again |

More detail on sharing, the database host and releasing is in [SHARE.md](SHARE.md).

---

## Development (maintainer)

Run the CSV import/SQL download tests with the project's dependencies installed:

```bash
python -m unittest discover -s tests -v
```

These tests use a simulated database and do not change any SQL Server data.

```bash
docker compose up --build
```

`docker-compose.yml` mounts the source into the container with Flask debug on,
so code and template changes apply without a rebuild. Rebuild only when
`Dockerfile` or `requirements.txt` changes.

| File | What it is |
| --- | --- |
| `app.py` | Flask backend: Excel scanning, import jobs, truncate, data editor API, version check |
| `templates/index.html` | The whole UI (HTML, CSS and JS in one file) |
| `version.json` | Current version, release notes and the update-check URL |
| `release.sh` | Release script (see below) |
| `update.sh` / `update.bat` | Update scripts for users |

### Releasing

Work on `dev`. When it's steady, commit everything and run:

```bash
./release.sh minor --dry-run   # preview
./release.sh minor             # 1.1.0 → 1.2.0  new features
./release.sh patch             # 1.1.0 → 1.1.1  bug fixes
```

The script builds release notes from commit messages since the last `v*` tag,
bumps `version.json`, merges `dev` → `main`, tags and pushes. Finally, paste the
printed `version.json` into the Gist that `update_url` points to — running apps
check it (at most every 5 minutes) and show the update popup.
