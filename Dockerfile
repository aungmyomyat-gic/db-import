FROM python:3.12-slim

# System deps + unixODBC
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl gnupg2 apt-transport-https unixodbc-dev libgssapi-krb5-2 git \
    && rm -rf /var/lib/apt/lists/*

# Microsoft ODBC Driver 18 for SQL Server
RUN curl -fsSL https://packages.microsoft.com/keys/microsoft.asc \
      | gpg --dearmor -o /usr/share/keyrings/microsoft-prod.gpg \
    && curl -fsSL https://packages.microsoft.com/config/debian/12/prod.list \
      > /etc/apt/sources.list.d/mssql-release.list \
    && apt-get update \
    && ACCEPT_EULA=Y apt-get install -y --no-install-recommends msodbcsql18 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Fail the image build early if pyodbc cannot see the SQL Server driver.
RUN python -c "import pyodbc; drivers = pyodbc.drivers(); print(drivers); assert any('SQL Server' in d for d in drivers)"

COPY . .

EXPOSE 5000
CMD ["python", "app.py"]
