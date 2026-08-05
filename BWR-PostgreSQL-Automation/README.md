# 360tf PostgreSQL Biweekly Report

This program collects PostgreSQL monitoring data and produces a Word (`.docx`) biweekly DBA report, including a centered cover page, a S. No./Title contents table, overview tables, maintenance observations, and headings/placeholders for console graphs. The completed report is then sent as an email attachment.

## Install

```bash
cd outputs
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
```

Edit `.env` with the database and SMTP credentials. The credential file is deliberately excluded from the script; protect it with `chmod 600 .env`. Email is sent to `postgresqlsupport@geopits.com`; configure the SMTP account used as the sender.

## Run

```bash
cd outputs
python3 generate_360tf_biweekly_report.py
```

The report is saved under `REPORT_OUTPUT_DIR` (default: `./reports`) and is emailed as a `.docx` attachment after it is generated.

## Growth comparison

The first successful report stores raw database and `app360tf` table sizes in `360tf_size_snapshot.json` beside the reports. The first report shows `N/A` in **Size (Last Time)**. Every later report automatically uses the previous snapshot as **Size (Last Time)** and calculates the percentage difference. Keep this snapshot file to preserve the comparison history.

## Database permissions

Use a dedicated read-only reporting role. At a minimum, it needs `CONNECT` on each database and enough catalog/statistics visibility to read `pg_database`, `pg_roles`, `pg_stat_database`, `pg_stat_user_tables`, `pg_statio_user_tables`, `pg_tablespace`, `information_schema`, and `pg_settings`. Missing permissions do not stop the report; the relevant section shows an observation instead.

For accurate relation sizes and table statistics across every database, list them in `PGDATABASES`; otherwise the script discovers only databases visible to the supplied account.

## Scope note

The “bloat percentage” is an estimated dead-tuple ratio (`n_dead_tup / (n_live_tup + n_dead_tup) × 100`), not an exact on-disk bloat calculation. Exact bloat needs the optional `pgstattuple` extension and elevated privileges.
