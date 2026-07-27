# Understanding `pg_lsclusters`

`pg_lsclusters` is a **Debian/Ubuntu utility** that lists all PostgreSQL clusters installed on the system.

A **cluster** in PostgreSQL is a complete database server instance, including:
- Its own data directory
- Configuration files
- Port
- Log file
- Status (online/down)

A single server can host multiple PostgreSQL clusters, even with different PostgreSQL versions.

---

## List PostgreSQL Clusters

```bash
pg_lsclusters
```

### Example Output

```text
Ver Cluster Port Status Owner    Data directory              Log file
18  main    5432 online postgres /var/lib/postgresql/18/main /var/log/postgresql/postgresql-18-main.log
```

---

## Column Descriptions

| Column | Description |
|---------|-------------|
| **Ver** | PostgreSQL version (e.g., 18) |
| **Cluster** | Cluster name (usually `main`) |
| **Port** | TCP port used by the PostgreSQL cluster (default: `5432`) |
| **Status** | Cluster status (`online` or `down`) |
| **Owner** | Operating system user owning the cluster (usually `postgres`) |
| **Data directory** | Location where PostgreSQL stores database files |
| **Log file** | PostgreSQL server log file location |

---

# Common Cluster Management Commands

## List all clusters

```bash
pg_lsclusters
```

---

## Start a cluster

```bash
sudo pg_ctlcluster 18 main start
```

---

## Stop a cluster

```bash
sudo pg_ctlcluster 18 main stop
```

---

## Restart a cluster

```bash
sudo pg_ctlcluster 18 main restart
```

---

## Check cluster status

```bash
sudo pg_ctlcluster 18 main status
```

---

## View PostgreSQL log

```bash
sudo tail -100 /var/log/postgresql/postgresql-18-main.log
```

---

## Why Ubuntu Uses Clusters

Unlike some operating systems, Debian/Ubuntu packages allow multiple PostgreSQL versions to coexist on the same machine.

Example:

```text
Ver Cluster Port Status
16  main    5432 online
17  test    5433 online
18  dev     5434 online
```

Each cluster is an independent PostgreSQL server with its own:

- Databases
- Users and roles
- Configuration files
- Data directory
- WAL files
- Log files
- Network port

This design makes it easy to run development, testing, and production environments simultaneously on the same server.

---

# Related Commands

## Show PostgreSQL version

```bash
psql --version
```

---

## List databases

```bash
sudo -u postgres psql -c "\l"
```

---

## Connect to PostgreSQL

```bash
sudo -u postgres psql
```

---

## Connect to a specific database

```bash
sudo -u postgres psql -d dbname
```

---

## Check if PostgreSQL is listening on port 5432

```bash
ss -ltnp | grep 5432
```

---

## View running PostgreSQL processes

```bash
ps -ef | grep postgres
```

---

## Check PostgreSQL service status

```bash
sudo systemctl status postgresql
```
