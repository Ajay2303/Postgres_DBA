## Query to filter ONLY Primary-Key Indexes on Partitions
```
SELECT
    child.relname AS partition_name,
    idx.relname   AS pk_index_name,
    string_agg(att.attname, ', ' ORDER BY x.n) AS pk_columns
FROM pg_inherits inh
JOIN pg_class parent ON parent.oid = inh.inhparent
JOIN pg_class child  ON child.oid  = inh.inhrelid
JOIN pg_index ind    ON ind.indrelid = child.oid
JOIN pg_class idx    ON idx.oid = ind.indexrelid
JOIN LATERAL unnest(ind.indkey) WITH ORDINALITY AS x(attnum, n) ON true
JOIN pg_attribute att
  ON att.attrelid = child.oid
 AND att.attnum = x.attnum
WHERE parent.relname = 'table_name' 
  AND ind.indisprimary
GROUP BY child.relname, idx.relname
ORDER BY child.relname;
```
