# Count Partitions with Primary Key

This query counts total partitions of a partitioned table
and verifies how many of them have a primary key index.
```
SELECT
    parent.relname                AS table_name,
    COUNT(DISTINCT child.relname) AS total_partitions,
    COUNT(DISTINCT pk.indrelid)   AS partitions_with_pk
FROM pg_inherits inh
JOIN pg_class parent
    ON parent.oid = inh.inhparent
JOIN pg_class child
    ON child.oid = inh.inhrelid
LEFT JOIN pg_index pk
    ON pk.indrelid = child.oid
   AND pk.indisprimary
WHERE parent.relname IN (
    'table_name_1',
    'table_name_2',
    'table_name_n'
)
```
GROUP BY parent.relname
ORDER BY parent.relname;
