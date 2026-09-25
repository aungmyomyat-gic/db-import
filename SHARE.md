# Sharing This App

The Docker image includes the Python packages and the Microsoft SQL Server ODBC
driver. Your friend only needs Docker, unless they choose to run the app outside
Docker.

## Share Source Code (recommended — gets update notifications)

Your friend clones the repo over HTTPS once (they need read access to the
private repo):

```bash
git clone https://github.com/aungmyomyat-gic/db-import.git
cd db-import
docker compose up -d --build
```

Open:

```text
http://localhost:5000
```

## Database Host

If SQL Server is running on your friend's computer, use this host in the app:

```text
host.docker.internal
```

If SQL Server is another Docker container, both containers must be on the same
Docker network, or the database must publish port `1433`.

## Driver Error

This app uses ODBC through `pyodbc`, not JDBC. If your friend sees:

```text
No SQL Server ODBC driver found inside container.
```

they are probably running an old/bad image or running Python outside Docker.
Rebuild with `docker compose up -d --build`.

## When A Rebuild Is Needed

No rebuild is needed for normal app changes in development mode.

A rebuild is needed when `Dockerfile` or `requirements.txt` changes.

## Updating

When a new version is released, the app shows an **Update available** popup.
In the `db-import` folder, double-click `update.bat` (Windows) or run
`./update.sh` (Mac / Linux). Both do:

```bash
git pull --ff-only origin main
docker compose up -d --build
```

The saved DB connection in `./data` is kept.

## Releasing A New Version (maintainer)

Work on the `dev` branch. When it's steady, commit everything and run:

```bash
./release.sh patch            # 1.1.0 → 1.1.1  bug fixes
./release.sh minor            # 1.1.0 → 1.2.0  new features
./release.sh major            # 1.1.0 → 2.0.0  big changes
./release.sh minor --dry-run  # preview only
```

It builds release notes from your commit messages since the last `v*` tag
(you can edit them before confirming), bumps `version.json`, merges
`dev` → `main`, tags `vX.Y.Z` and pushes. At the end it prints the new
`version.json` — paste it into the Gist that `update_url` points to.

Each running app compares its own `version.json` (the one in its project folder) with
the file at `update_url` (or the `UPDATE_CHECK_URL` env var), at most every
30 minutes. If `update_url` is empty, the check is off. The repo is private,
so `update_url` must point to a copy the container can read without logging in.
