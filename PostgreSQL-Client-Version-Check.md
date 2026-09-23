# PostgreSQL Client Version Compatibility Check

## Overview

When taking a backup using `pg_dump`, the client version should match or be newer than the PostgreSQL server version. Using an older `pg_dump` client against a newer PostgreSQL server results in a version mismatch error.

## Environment

- OS: Ubuntu 20.04 LTS
- PostgreSQL Server: 14.22 (AWS RDS)
- Installed Client: PostgreSQL 12.22

## Error

```text
pg_dump: error: server version: 14.22
pg_dump: error: pg_dump version: 12.22
pg_dump: error: aborting because of server version mismatch
```

## Checks Performed

### Check PostgreSQL Client Version

```bash
pg_dump --version
psql --version
```

### Check Current Binary Path

```bash
which pg_dump
which psql
```

### Check All Installed PostgreSQL Client Versions

```bash
find /usr -type f \( -name psql -o -name pg_dump \) \
-exec sh -c 'echo -n "{} : "; "{}" --version' \;
```

### Check Installed PostgreSQL Packages

```bash
dpkg -l | grep postgresql
```

### Check PostgreSQL Installation Directories

```bash
ls -l /usr/lib/postgresql/
```

## Solution

Install the PostgreSQL 14 client:

```bash
sudo apt update
sudo apt install postgresql-client-14
```

Verify:

```bash
/usr/lib/postgresql/14/bin/pg_dump --version
```

Run backup using the PostgreSQL 14 binary:

```bash
/usr/lib/postgresql/14/bin/pg_dump \
-h <hostname> \
-p 5432 \
-U <username> \
-d <database> \
-t <schema>.<table> \
-F c \
-f backup.sql
```

## Notes

- `pg_dump` should be the same major version or newer than the PostgreSQL server.
- Multiple PostgreSQL client versions can coexist on the same Linux server.
- There is no need to uninstall older client versions unless they are no longer required.
