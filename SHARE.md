# Sharing This App

The Docker image includes the Python packages and the Microsoft SQL Server ODBC
driver. Your friend only needs Docker, unless they choose to run the app outside
Docker.

## Share Source Code

Send the project folder, then your friend can run:

```bash
docker compose -f docker-compose.share.yml up --build
```

Open:

```text
http://localhost:5000
```

## Share A Built Image

Build and export the image on your machine:

```bash
docker build -t db-connect:latest .
docker save db-connect:latest -o db-connect.tar
```

Send `db-connect.tar` and `docker-compose.image.yml`. Your friend runs:

```bash
docker load -i db-connect.tar
docker compose -f docker-compose.image.yml up
```

This option is useful when your friend cannot build the image because their
network cannot download the Microsoft ODBC driver packages.

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
Rebuild from this Dockerfile, or send them the built image with
`docker save` / `docker load`.

## When A Rebuild Is Needed

No rebuild is needed for normal app changes in development mode.

A rebuild is needed when `Dockerfile` or `requirements.txt` changes.
