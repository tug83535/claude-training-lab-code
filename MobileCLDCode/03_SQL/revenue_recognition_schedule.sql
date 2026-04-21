-- ============================================================================
-- revenue_recognition_schedule.sql
-- ASC 606 revenue recognition schedule in pure SQL
--
-- Produces:
--   - Per-customer, per-month recognized revenue for any time range
--   - Deferred revenue rollforward (opening + billed - recognized = closing)
--   - Exceptions: over-recognized balances, orphan billings, past-term billings
--
-- WHY THIS IS NOT NATIVE EXCEL / ONEDRIVE
-- ---------------------------------------
-- The Python version of this sits in ../02_Python/revenue_recognition_engine.py
-- This pure-SQL version runs directly in the warehouse so BI tools and
-- scheduled jobs can consume the numbers without re-running a Python script.
--
-- USE CASE
-- --------
-- Snowflake task runs this nightly; the finance team's Looker dashboard pulls
-- straight from revrec_monthly and revrec_rollforward.
--
-- ASSUMED TABLES:
--   contracts     (contract_id, customer_id, start_date, end_date, total_value,
--                  recognition_pattern ('Ratable' | 'PointInTime'), delivered_date)
--   billings      (bill_id, contract_id, bill_date, amount, status)
-- ============================================================================

-- ----------------------------------------------------------------------------
-- Recognized revenue by (contract, month). Straight-line proration by day.
-- ----------------------------------------------------------------------------
CREATE OR REPLACE VIEW revrec_monthly AS
WITH months AS (
    SELECT generate_series('2022-01-01'::date, CURRENT_DATE + INTERVAL '24 months', '1 month') AS m
),
month_bounds AS (
    SELECT m AS month_start,
           (m + INTERVAL '1 month' - INTERVAL '1 day')::date AS month_end
    FROM months
),
ratable AS (
    SELECT
        c.contract_id,
        c.customer_id,
        DATE_TRUNC('month', mb.month_start) AS period,
        c.total_value * GREATEST(0, LEAST(c.end_date, mb.month_end)::date
                                  - GREATEST(c.start_date, mb.month_start)::date + 1)
                      / (c.end_date - c.start_date + 1) AS recognized
    FROM contracts c
    JOIN month_bounds mb
      ON c.start_date <= mb.month_end AND c.end_date >= mb.month_start
    WHERE c.recognition_pattern = 'Ratable'
),
point_in_time AS (
    SELECT
        c.contract_id,
        c.customer_id,
        DATE_TRUNC('month', c.delivered_date) AS period,
        c.total_value AS recognized
    FROM contracts c
    WHERE c.recognition_pattern = 'PointInTime'
      AND c.delivered_date IS NOT NULL
)
SELECT contract_id, customer_id, period, ROUND(recognized::numeric, 2) AS recognized_revenue
FROM ratable
WHERE recognized > 0
UNION ALL
SELECT contract_id, customer_id, period, ROUND(recognized::numeric, 2)
FROM point_in_time;


-- ----------------------------------------------------------------------------
-- Deferred revenue rollforward by month
-- ----------------------------------------------------------------------------
CREATE OR REPLACE VIEW revrec_rollforward AS
WITH all_months AS (
    SELECT DISTINCT period FROM revrec_monthly
),
billed_by_month AS (
    SELECT DATE_TRUNC('month', bill_date) AS period, contract_id, SUM(amount) AS billed
    FROM billings
    GROUP BY 1, 2
),
rec_by_month AS (
    SELECT period, contract_id, SUM(recognized_revenue) AS recognized
    FROM revrec_monthly
    GROUP BY 1, 2
),
joined AS (
    SELECT COALESCE(b.period, r.period) AS period,
           COALESCE(b.contract_id, r.contract_id) AS contract_id,
           COALESCE(b.billed, 0) AS billed,
           COALESCE(r.recognized, 0) AS recognized
    FROM billed_by_month b
    FULL OUTER JOIN rec_by_month r
      ON r.period = b.period AND r.contract_id = b.contract_id
)
SELECT
    period,
    contract_id,
    billed,
    recognized,
    SUM(billed - recognized) OVER (
        PARTITION BY contract_id
        ORDER BY period
        ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW
    ) AS running_deferred_balance
FROM joined
ORDER BY contract_id, period;


-- ----------------------------------------------------------------------------
-- Period rollforward, aggregated across contracts (for board-ready reporting)
-- ----------------------------------------------------------------------------
CREATE OR REPLACE VIEW revrec_period_summary AS
WITH per_period AS (
    SELECT period,
           SUM(billed)     AS billed_in_period,
           SUM(recognized) AS recognized_in_period
    FROM revrec_rollforward
    GROUP BY 1
),
running AS (
    SELECT
        period,
        SUM(billed_in_period) OVER (ORDER BY period) AS cumulative_billed,
        SUM(recognized_in_period) OVER (ORDER BY period) AS cumulative_recognized
    FROM per_period
)
SELECT
    period,
    COALESCE(LAG(cumulative_billed - cumulative_recognized) OVER (ORDER BY period), 0)
        AS opening_deferred,
    (SELECT billed_in_period FROM per_period p WHERE p.period = running.period)
        AS billed_this_period,
    (SELECT recognized_in_period FROM per_period p WHERE p.period = running.period)
        AS recognized_this_period,
    cumulative_billed - cumulative_recognized AS closing_deferred
FROM running
ORDER BY period;


-- ----------------------------------------------------------------------------
-- Exceptions report
-- ----------------------------------------------------------------------------
CREATE OR REPLACE VIEW revrec_exceptions AS
-- Over-recognized (negative deferred)
SELECT 'Over-recognized' AS exception_type,
       contract_id,
       period,
       running_deferred_balance AS detail
FROM revrec_rollforward
WHERE running_deferred_balance < -0.01
UNION ALL
-- Orphan billings
SELECT 'Orphan billing',
       b.contract_id,
       DATE_TRUNC('month', b.bill_date),
       b.amount
FROM billings b
LEFT JOIN contracts c USING (contract_id)
WHERE c.contract_id IS NULL
UNION ALL
-- Billed after contract end
SELECT 'Post-term billing',
       b.contract_id,
       DATE_TRUNC('month', b.bill_date),
       b.amount
FROM billings b
JOIN contracts c USING (contract_id)
WHERE b.bill_date > c.end_date;
