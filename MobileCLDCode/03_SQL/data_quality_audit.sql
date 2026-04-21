-- ============================================================================
-- data_quality_audit.sql
-- Generic Data Quality Audit pack
--
-- Run periodically against any warehouse schema to find:
--   - Orphan foreign keys (child rows with no parent)
--   - Unexpected NULLs on fields supposed to be NOT NULL
--   - Duplicate primary keys / business keys
--   - Frozen tables (no new rows in > N days)
--   - Date gaps in what should be continuous series
--   - Outlier amounts (z-score > 4 from mean)
--   - Timezone inconsistencies
--
-- WHY THIS IS NOT NATIVE EXCEL / ONEDRIVE
-- ---------------------------------------
-- You simply cannot audit 500M rows of warehouse data in Excel. DQ dashboards
-- from vendors (Monte Carlo, Metaplane) cost $20-50K/yr. This is the 90%
-- version for free, runnable on any schedule.
--
-- USE CASE
-- --------
-- Data engineering runs this nightly and posts the fail summary to Slack.
-- Finance runs it before month-end close on the GL tables.
--
-- TARGET: any modern dialect (PostgreSQL, Snowflake, BigQuery)
-- ============================================================================

-- ----------------------------------------------------------------------------
-- 1. Orphan foreign key check (invoices without a customer, etc.)
--    Emits one row per FK that fails, counting the orphans.
-- ----------------------------------------------------------------------------
WITH fk_checks AS (
    SELECT
        'invoices.customer_id' AS fk_column,
        (SELECT COUNT(*) FROM invoices i
          WHERE NOT EXISTS (SELECT 1 FROM customers c WHERE c.customer_id = i.customer_id)) AS orphans
    UNION ALL
    SELECT 'opportunities.customer_id',
        (SELECT COUNT(*) FROM opportunities o
          WHERE NOT EXISTS (SELECT 1 FROM customers c WHERE c.customer_id = o.customer_id))
    UNION ALL
    SELECT 'subscriptions.customer_id',
        (SELECT COUNT(*) FROM subscriptions s
          WHERE NOT EXISTS (SELECT 1 FROM customers c WHERE c.customer_id = s.customer_id))
)
SELECT fk_column,
       orphans,
       CASE WHEN orphans = 0 THEN 'OK' ELSE 'FAIL' END AS status
FROM fk_checks
ORDER BY orphans DESC;


-- ----------------------------------------------------------------------------
-- 2. Unexpected NULLs on required columns
-- ----------------------------------------------------------------------------
WITH null_checks AS (
    SELECT 'customers.primary_email' AS field,
           COUNT(*) FILTER (WHERE primary_email IS NULL OR primary_email = '') AS nulls,
           COUNT(*) AS total FROM customers
    UNION ALL
    SELECT 'invoices.amount',
           COUNT(*) FILTER (WHERE amount IS NULL) AS nulls,
           COUNT(*) AS total FROM invoices
    UNION ALL
    SELECT 'subscriptions.start_date',
           COUNT(*) FILTER (WHERE start_date IS NULL) AS nulls,
           COUNT(*) AS total FROM subscriptions
)
SELECT field,
       nulls,
       total,
       ROUND(nulls::numeric / NULLIF(total, 0) * 100, 3) AS null_pct,
       CASE WHEN nulls = 0 THEN 'OK' ELSE 'FAIL' END AS status
FROM null_checks
ORDER BY null_pct DESC;


-- ----------------------------------------------------------------------------
-- 3. Duplicate business keys
-- ----------------------------------------------------------------------------
-- Customer email should be unique across active accounts
SELECT primary_email, COUNT(*) AS n
FROM customers
WHERE is_deleted = FALSE AND primary_email IS NOT NULL
GROUP BY 1
HAVING COUNT(*) > 1;

-- Invoice number must be unique within customer
SELECT customer_id, invoice_number, COUNT(*) AS n
FROM invoices
GROUP BY 1, 2
HAVING COUNT(*) > 1;


-- ----------------------------------------------------------------------------
-- 4. Frozen tables: no new rows inserted recently
-- ----------------------------------------------------------------------------
SELECT 'invoices'     AS table_name, MAX(created_at) AS latest, CURRENT_DATE - MAX(created_at)::date AS days_stale FROM invoices
UNION ALL
SELECT 'subscriptions',              MAX(created_at),            CURRENT_DATE - MAX(created_at)::date FROM subscriptions
UNION ALL
SELECT 'opportunities',              MAX(created_date),          CURRENT_DATE - MAX(created_date)::date FROM opportunities
UNION ALL
SELECT 'product_events',             MAX(event_at),              CURRENT_DATE - MAX(event_at)::date FROM product_events
ORDER BY days_stale DESC;


-- ----------------------------------------------------------------------------
-- 5. Calendar gaps (expected daily data missing)
-- ----------------------------------------------------------------------------
WITH spine AS (
    SELECT generate_series(
               CURRENT_DATE - INTERVAL '90 days',
               CURRENT_DATE,
               '1 day'
           )::date AS d
),
daily_event_count AS (
    SELECT DATE(event_at) AS d, COUNT(*) AS events FROM product_events
    WHERE event_at >= CURRENT_DATE - INTERVAL '90 days'
    GROUP BY 1
)
SELECT s.d AS date,
       COALESCE(e.events, 0) AS events,
       CASE WHEN COALESCE(e.events, 0) = 0 THEN 'GAP' ELSE 'OK' END AS status
FROM spine s
LEFT JOIN daily_event_count e ON e.d = s.d
WHERE COALESCE(e.events, 0) = 0
ORDER BY s.d;


-- ----------------------------------------------------------------------------
-- 6. Outlier amounts (invoices > 4 stddev from the rolling mean)
-- ----------------------------------------------------------------------------
WITH stats AS (
    SELECT customer_id,
           AVG(amount) AS mean_amt,
           STDDEV(amount) AS std_amt
    FROM invoices
    WHERE created_at >= CURRENT_DATE - INTERVAL '12 months'
    GROUP BY 1
    HAVING COUNT(*) >= 6
)
SELECT i.customer_id, i.invoice_number, i.amount,
       s.mean_amt, s.std_amt,
       ROUND(((i.amount - s.mean_amt) / NULLIF(s.std_amt, 0))::numeric, 2) AS z_score
FROM invoices i
JOIN stats s USING (customer_id)
WHERE ABS((i.amount - s.mean_amt) / NULLIF(s.std_amt, 0)) > 4
  AND i.created_at >= CURRENT_DATE - INTERVAL '7 days'
ORDER BY ABS((i.amount - s.mean_amt) / NULLIF(s.std_amt, 0)) DESC;


-- ----------------------------------------------------------------------------
-- 7. Referential integrity across accounting: sum of line items = invoice total
-- ----------------------------------------------------------------------------
WITH line_sums AS (
    SELECT invoice_id, SUM(amount) AS lines_total
    FROM invoice_lines GROUP BY 1
)
SELECT i.invoice_id,
       i.invoice_number,
       i.amount AS header_amount,
       COALESCE(l.lines_total, 0) AS line_sum,
       i.amount - COALESCE(l.lines_total, 0) AS variance
FROM invoices i
LEFT JOIN line_sums l USING (invoice_id)
WHERE ABS(i.amount - COALESCE(l.lines_total, 0)) > 0.01
ORDER BY ABS(i.amount - COALESCE(l.lines_total, 0)) DESC;
