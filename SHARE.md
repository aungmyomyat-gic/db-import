# Sharing the CSV & SQL version

Share **DBImport.exe** for the Windows version with CSV Import, SQL Generate,
progress, elapsed time, and cancellation. Build it on Windows with
`build_windows.bat`, or download **DBImport-Windows-x64** from a successful
Windows GitHub Actions run. See [WINDOWS.md](WINDOWS.md) for the complete steps.

Recipients can double-click the executable with a compatible 64-bit ODBC Driver
18 or 17 installed. They enter their own connection settings; your saved
credentials are excluded from the package. Share `dist/DBImport.exe`, its SHA-256
file, and `dist/WINDOWS.md`. Replacing the executable keeps their saved settings.

## Share source

With read access to the repository:

```bash
git clone --branch feat/windows-exe --single-branch https://github.com/aungmyomyat-gic/db-import.git
cd db-import
```

Run `bash start.sh` with Python and the SQL Server ODBC driver installed, or
run `docker compose up -d --build` with Docker. Open http://localhost:5000.
The Docker image includes Python packages and the Microsoft ODBC driver.

The supplied Docker setup uses the `icl-lv-core-system_sazaby-network` network.
Use that database stack or adapt the compose network for the target database.
Use `host.docker.internal` where available for SQL Server on the Docker host,
or a hostname/IP reachable from the container for a remote database.

## Share a Docker image

A source checkout is not needed on the receiving PC when sharing a built image:

```bash
docker build -t db-import-csv-sql .
docker save -o db-import-csv-sql.tar db-import-csv-sql
```

Copy the archive to the receiving PC, then run:

```bash
docker load -i db-import-csv-sql.tar
docker run -d --name db-import-csv-sql -p 5000:5000 -v db-import-csv-data:/app/data db-import-csv-sql
```

Configure a database host reachable from that container. Attach the container
to the database's Docker network if necessary.

## Updating source installations

Use `./update.sh` or `update.bat` to pull the current branch and rebuild Docker.
Saved connection settings stay in `data/db_config.json`. Each recipient should
configure their own connection. Do not include your `data` folder when sharing
project files.
