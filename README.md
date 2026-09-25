# DB Import

A small web app for loading Excel test data into SQL Server and working with
the tables afterwards. It runs in Docker and opens at http://localhost:5000.

## Features

- **Import** — upload an Excel file, pick a sheet, and the app finds every table
  under the **■ 実施前テストデータ** section (names like `MZAIHP(倉庫在庫マスタ)`).
  Preview them, choose which to import, and insert them into the target schema.
  If the file name contains `day job 12` / `month job 3`, the matching batch job
  can be triggered after the import.
- **Truncate Table** — empty one or more tables in a schema.
- **Data Editor** — browse up to 100 rows of any table:
  - row numbers, and a summary of primary keys, total rows and columns
  - sort by clicking a column header; search columns by name
  - edit rows in place (Edit mode)
  - copy rows to Excel, with or without the header
  - paste rows copied from Excel (Ctrl+V), review them, then insert
- **Update notification** — when a new version is released, the app shows an
  "Update available" popup with the changes.

## Install

You need **Git** and **Docker Desktop**. The repo is private, so you need read
access on GitHub.

```bash
git clone https://github.com/aungmyomyat-gic/db-import.git
cd db-import
docker compose -f docker-compose.share.yml up -d --build
```

Open http://localhost:5000, go to **Connection**, and enter your SQL Server
details. They are saved in `./data` and kept across updates.

> If SQL Server runs on your own computer (not in Docker), use
> `host.docker.internal` as the host.

## Update

When the app shows **Update available**, double-click `update.bat` (Windows) or
run `./update.sh` (Mac / Linux) in the `db-import` folder. Both run:

```bash
git pull --ff-only origin main
docker compose -f docker-compose.share.yml up -d --build
```

Then reload the page.

## Stop / start

```bash
docker compose -f docker-compose.share.yml down     # stop
docker compose -f docker-compose.share.yml up -d    # start again
```

## Troubleshooting

| Problem | Fix |
| --- | --- |
| `No SQL Server ODBC driver found inside container` | Rebuild: `docker compose -f docker-compose.share.yml up -d --build` |
| Can't connect to a database on your own PC | Use `host.docker.internal` as the host and check SQL Server listens on port 1433 |
| Port 5000 is already in use | Change `"5000:5000"` to e.g. `"5050:5000"` in `docker-compose.share.yml`, then open http://localhost:5050 |
| `git pull` fails in the update script | You have local changes in the folder. Run `git status`, then `git stash` or discard them, and run the update again |

More sharing options (e.g. sending a built image to someone without internet
access to the package servers) are in [SHARE.md](SHARE.md).

---

## Development (maintainer)

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
| `main.py` | Older command-line version of the importer |
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
check it (at most every 30 minutes) and show the update popup.
