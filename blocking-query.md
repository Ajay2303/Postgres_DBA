```
SELECT pid,
       usename,
       now()-query_start AS blocking_time,
       left(query,120)
FROM pg_stat_activity
WHERE pid IN (
    SELECT unnest(pg_blocking_pids(pid))
    FROM pg_stat_activity
)
AND now()-query_start > interval '30 minutes'
ORDER BY blocking_time DESC;
