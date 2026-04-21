-- ============================================================================
-- slow_query_tuner.sql
-- Database Performance Diagnostics
--
-- PURPOSE
-- -------
-- Provides a reusable library of queries any DBA / backend dev can run on a
-- production OLTP database to find the real performance bottlenecks:
--
--   - Top N slowest queries by mean & total time
--   - Missing index recommendations (SQL Server DMV style)
--   - Unused indexes (costing writes but never read)
--   - Deadlock and blocking sessions
--   - Table bloat / autovacuum candidates (PostgreSQL)
--   - TempDB / spill-to-disk queries
--
-- WHY THIS IS NOT NATIVE EXCEL / ONEDRIVE
-- ---------------------------------------
-- Nothing in Excel / OneDrive can introspect a database's live performance
-- counters. This is DMV / pg_stat_statements / information_schema territory.
-- Without these, "the app is slow" turns into a 2-day detective hunt.
--
-- USE CASE
-- --------
-- Every infra / platform team has a few hero queries like these saved in
-- a wiki. This file is a portable version — drop it into any engagement.
--
-- TARGET DIALECTS: SQL Server + PostgreSQL. Sections are labeled.
-- ============================================================================


-- ====================
-- SQL SERVER
-- ====================

-- Top 20 slowest stored procedures by mean duration
SELECT TOP 20
    OBJECT_NAME(ps.object_id, ps.database_id) AS procedure_name,
    DB_NAME(ps.database_id)                    AS database_name,
    ps.execution_count,
    ps.total_elapsed_time / 1000               AS total_ms,
    ps.total_elapsed_time / NULLIF(ps.execution_count, 0) / 1000 AS mean_ms,
    ps.total_logical_reads,
    ps.total_physical_reads,
    ps.last_execution_time
FROM sys.dm_exec_procedure_stats ps
WHERE ps.database_id = DB_ID()
ORDER BY mean_ms DESC;


-- Top 20 expensive individual statements (across the plan cache)
SELECT TOP 20
    qs.execution_count,
    qs.total_elapsed_time / 1000                                  AS total_ms,
    (qs.total_elapsed_time / NULLIF(qs.execution_count, 0)) / 1000 AS mean_ms,
    qs.total_worker_time / 1000                                   AS total_cpu_ms,
    qs.total_logical_reads,
    SUBSTRING(qt.text,
              qs.statement_start_offset / 2 + 1,
              (CASE WHEN qs.statement_end_offset = -1
                    THEN DATALENGTH(qt.text)
                    ELSE qs.statement_end_offset END
               - qs.statement_start_offset) / 2 + 1)              AS statement_text,
    qs.last_execution_time
FROM sys.dm_exec_query_stats qs
CROSS APPLY sys.dm_exec_sql_text(qs.sql_handle) qt
ORDER BY mean_ms DESC;


-- Missing index recommendations (use judgement; don't blindly add)
SELECT
    DB_NAME(mid.database_id) AS database_name,
    OBJECT_NAME(mid.object_id, mid.database_id) AS table_name,
    migs.avg_user_impact,
    migs.user_seeks,
    migs.user_scans,
    'CREATE INDEX IX_' + REPLACE(OBJECT_NAME(mid.object_id, mid.database_id), ' ', '_')
      + '_auto ON ' + mid.statement + ' ('
      + ISNULL(mid.equality_columns, '')
      + CASE WHEN mid.equality_columns IS NOT NULL
                  AND mid.inequality_columns IS NOT NULL
             THEN ', ' ELSE '' END
      + ISNULL(mid.inequality_columns, '') + ')'
      + CASE WHEN mid.included_columns IS NOT NULL
             THEN ' INCLUDE (' + mid.included_columns + ')' ELSE '' END AS create_statement
FROM sys.dm_db_missing_index_details mid
JOIN sys.dm_db_missing_index_groups mig ON mig.index_handle = mid.index_handle
JOIN sys.dm_db_missing_index_group_stats migs ON migs.group_handle = mig.index_group_handle
WHERE mid.database_id = DB_ID()
ORDER BY migs.avg_user_impact DESC;


-- Unused indexes (zero reads, nonzero writes) - reclaim candidates
SELECT
    OBJECT_NAME(s.object_id) AS table_name,
    i.name                   AS index_name,
    s.user_seeks, s.user_scans, s.user_lookups,
    s.user_updates
FROM sys.dm_db_index_usage_stats s
JOIN sys.indexes i ON i.object_id = s.object_id AND i.index_id = s.index_id
WHERE s.database_id = DB_ID()
  AND OBJECTPROPERTY(s.object_id, 'IsUserTable') = 1
  AND i.is_primary_key = 0 AND i.is_unique_constraint = 0
  AND s.user_seeks + s.user_scans + s.user_lookups = 0
  AND s.user_updates > 0
ORDER BY s.user_updates DESC;


-- Active blocking (who's waiting on whom right now)
SELECT
    s.session_id, s.login_name, s.host_name,
    r.blocking_session_id,
    r.status, r.wait_type, r.wait_time,
    t.text AS current_statement
FROM sys.dm_exec_requests r
JOIN sys.dm_exec_sessions s ON s.session_id = r.session_id
CROSS APPLY sys.dm_exec_sql_text(r.sql_handle) t
WHERE r.blocking_session_id <> 0;


-- ====================
-- POSTGRESQL
-- ====================

-- Enable the stats extension once:
-- CREATE EXTENSION IF NOT EXISTS pg_stat_statements;

-- Top 20 slowest statements by mean duration
SELECT
    ROUND(total_exec_time::numeric, 0)      AS total_ms,
    ROUND(mean_exec_time::numeric, 2)       AS mean_ms,
    calls,
    rows,
    LEFT(query, 200)                        AS statement
FROM pg_stat_statements
ORDER BY mean_exec_time DESC
LIMIT 20;


-- Unused indexes
SELECT schemaname, relname AS table, indexrelname AS index,
       idx_scan AS times_used, pg_size_pretty(pg_relation_size(indexrelid)) AS size
FROM pg_stat_user_indexes
WHERE idx_scan = 0
ORDER BY pg_relation_size(indexrelid) DESC;


-- Table bloat (needs pgstattuple extension) - simplified heuristic:
SELECT schemaname, relname,
       n_dead_tup, n_live_tup,
       ROUND(n_dead_tup::numeric / NULLIF(n_live_tup, 0) * 100, 1) AS pct_dead,
       last_autovacuum
FROM pg_stat_user_tables
WHERE n_dead_tup > 10000
ORDER BY pct_dead DESC NULLS LAST;


-- Active blocking on Postgres
SELECT blocked.pid AS blocked_pid,
       blocked.usename AS blocked_user,
       blocking.pid AS blocking_pid,
       blocking.usename AS blocking_user,
       blocking.query AS blocking_statement,
       blocked.query AS blocked_statement
FROM pg_stat_activity blocked
JOIN pg_stat_activity blocking
     ON blocking.pid = ANY(pg_blocking_pids(blocked.pid));
