# Sharing the CSV & SQL version

Share the `feat/csv-import-sql-generator` branch to give someone the version with
CSV Import and standalone SQL Generate. The branch must be pushed before it can
be cloned from GitHub.

## Share source

With read access to the repository:

```bash
git clone --branch feat/csv-import-sql-generator --single-branch https://github.com/aungmyomyat-gic/db-import.git
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
