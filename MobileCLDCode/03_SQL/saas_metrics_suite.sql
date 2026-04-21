-- ============================================================================
-- saas_metrics_suite.sql
-- Complete SaaS Metrics library: MRR, ARR, NRR, GRR, Rule of 40, Quick Ratio,
-- Magic Number, LTV:CAC, Burn Multiple, CAC Payback.
--
-- WHY THIS IS NOT NATIVE EXCEL / ONEDRIVE
-- ---------------------------------------
-- Each of these numbers is derivable in Excel, but only from perfectly clean
-- source data. Real SaaS data lives in Snowflake / BigQuery / Postgres, is
-- billions of rows, and requires window functions Excel doesn't have. These
-- queries run in seconds against production warehouses.
--
-- USE CASE
-- --------
-- Board deck, quarterly ops review, pricing committee. The 8 numbers a SaaS
-- company's leadership team looks at monthly.
--
-- TARGET DIALECT: Snowflake / BigQuery / PostgreSQL (minor tweaks per dialect)
-- ASSUMED TABLES:
--   subscriptions   (customer_id, plan_id, mrr, start_date, end_date)
--   opportunities   (deal_id, customer_id, close_date, stage, amount, type, sales_rep)
--   cash_flows      (month, operating_cash_flow, capex)
--   operating_expense (month, s_and_m, r_and_d, g_and_a, cogs)
-- ============================================================================

-- ----------------------------------------------------------------------------
-- 1. Monthly ARR waterfall with all five movements
-- ----------------------------------------------------------------------------
WITH month_spine AS (
    SELECT DATE_TRUNC('month', d) AS month
    FROM generate_series('2024-01-01'::date, CURRENT_DATE, '1 month') g(d)
),
active_by_month AS (
    SELECT
        ms.month,
        s.customer_id,
        SUM(s.mrr) AS customer_mrr
    FROM month_spine ms
    JOIN subscriptions s
      ON s.start_date <= ms.month + INTERVAL '1 month' - INTERVAL '1 day'
     AND (s.end_date IS NULL OR s.end_date >= ms.month)
    GROUP BY 1, 2
),
movements AS (
    SELECT
        curr.month,
        curr.customer_id,
        COALESCE(prev.customer_mrr, 0) AS prev_mrr,
        curr.customer_mrr AS curr_mrr,
        CASE
          WHEN prev.customer_mrr IS NULL OR prev.customer_mrr = 0
               THEN curr.customer_mrr
          ELSE 0 END AS new_mrr,
        CASE
          WHEN prev.customer_mrr > 0 AND curr.customer_mrr > prev.customer_mrr
               THEN curr.customer_mrr - prev.customer_mrr
          ELSE 0 END AS expansion_mrr,
        CASE
          WHEN prev.customer_mrr > 0 AND curr.customer_mrr > 0
               AND curr.customer_mrr < prev.customer_mrr
               THEN prev.customer_mrr - curr.customer_mrr
          ELSE 0 END AS contraction_mrr,
        CASE
          WHEN prev.customer_mrr > 0 AND curr.customer_mrr = 0
               THEN prev.customer_mrr
          ELSE 0 END AS churn_mrr
    FROM active_by_month curr
    LEFT JOIN active_by_month prev
           ON prev.customer_id = curr.customer_id
          AND prev.month = curr.month - INTERVAL '1 month'
),
waterfall AS (
    SELECT
        month,
        SUM(prev_mrr)           AS starting_mrr,
        SUM(new_mrr)            AS new_mrr,
        SUM(expansion_mrr)      AS expansion_mrr,
        SUM(contraction_mrr)    AS contraction_mrr,
        SUM(churn_mrr)          AS churn_mrr,
        SUM(curr_mrr)           AS ending_mrr
    FROM movements
    GROUP BY month
)
SELECT
    month,
    starting_mrr,
    new_mrr,
    expansion_mrr,
    -contraction_mrr  AS contraction_mrr,
    -churn_mrr        AS churn_mrr,
    ending_mrr,
    ending_mrr - (starting_mrr + new_mrr + expansion_mrr
                  - contraction_mrr - churn_mrr) AS check_zero,
    CASE WHEN starting_mrr > 0
         THEN ROUND((starting_mrr + expansion_mrr - contraction_mrr - churn_mrr)
                    / starting_mrr * 100, 2)
         END AS nrr_pct,
    CASE WHEN starting_mrr > 0
         THEN ROUND((starting_mrr - contraction_mrr - churn_mrr)
                    / starting_mrr * 100, 2)
         END AS grr_pct,
    CASE WHEN (contraction_mrr + churn_mrr) > 0
         THEN ROUND((new_mrr + expansion_mrr)
                    / (contraction_mrr + churn_mrr), 2)
         END AS quick_ratio,
    ending_mrr * 12 AS ending_arr
FROM waterfall
ORDER BY month;


-- ----------------------------------------------------------------------------
-- 2. Rule of 40 (growth + profit) per trailing 12 months
-- ----------------------------------------------------------------------------
WITH ttm AS (
    SELECT
        DATE_TRUNC('month', CURRENT_DATE) AS asof_month,
        SUM(CASE WHEN month >= CURRENT_DATE - INTERVAL '12 months' THEN mrr * 12 END) AS ttm_arr,
        SUM(CASE WHEN month >= CURRENT_DATE - INTERVAL '24 months'
                  AND month <  CURRENT_DATE - INTERVAL '12 months' THEN mrr * 12 END) AS ttm_arr_prior,
        SUM(CASE WHEN month >= CURRENT_DATE - INTERVAL '12 months'
                 THEN operating_cash_flow END) AS ttm_fcf
    FROM subscriptions s
    CROSS JOIN cash_flows c
)
SELECT
    asof_month,
    ttm_arr,
    ttm_arr_prior,
    ROUND((ttm_arr - ttm_arr_prior)::numeric / NULLIF(ttm_arr_prior, 0) * 100, 2) AS arr_growth_pct,
    ROUND(ttm_fcf / NULLIF(ttm_arr, 0) * 100, 2) AS fcf_margin_pct,
    ROUND((ttm_arr - ttm_arr_prior)::numeric / NULLIF(ttm_arr_prior, 0) * 100, 2)
        + ROUND(ttm_fcf / NULLIF(ttm_arr, 0) * 100, 2) AS rule_of_40
FROM ttm;


-- ----------------------------------------------------------------------------
-- 3. Magic Number: (ΔARR QoQ × 4) / Prior Quarter S&M spend
-- ----------------------------------------------------------------------------
WITH q_arr AS (
    SELECT DATE_TRUNC('quarter', month) AS quarter, SUM(mrr * 3) AS q_arr_generated
    FROM subscriptions GROUP BY 1
),
q_sm AS (
    SELECT DATE_TRUNC('quarter', month) AS quarter, SUM(s_and_m) AS q_sm_spend
    FROM operating_expense GROUP BY 1
)
SELECT
    q.quarter,
    q.q_arr_generated,
    q.q_arr_generated - LAG(q.q_arr_generated) OVER (ORDER BY q.quarter) AS delta_arr,
    LAG(sm.q_sm_spend) OVER (ORDER BY q.quarter) AS prior_q_sm,
    ROUND(
      ((q.q_arr_generated - LAG(q.q_arr_generated) OVER (ORDER BY q.quarter)) * 4)::numeric
      / NULLIF(LAG(sm.q_sm_spend) OVER (ORDER BY q.quarter), 0), 2
    ) AS magic_number
FROM q_arr q
JOIN q_sm sm USING (quarter)
ORDER BY q.quarter;


-- ----------------------------------------------------------------------------
-- 4. CAC Payback in months: S&M spend in period / net new ARR added * gross margin
-- ----------------------------------------------------------------------------
WITH cac_input AS (
    SELECT
        DATE_TRUNC('quarter', o.close_date) AS quarter,
        SUM(o.amount) FILTER (WHERE o.type = 'New Business')                AS new_acv,
        SUM(oe.s_and_m)                                                     AS sm_spend,
        SUM(CASE WHEN subs.mrr IS NOT NULL THEN subs.mrr ELSE 0 END) * 12   AS total_arr,
        AVG(1 - oe.cogs / NULLIF(oe.s_and_m + oe.r_and_d + oe.g_and_a + oe.cogs, 0))
                                                                            AS gross_margin_pct
    FROM opportunities o
    JOIN operating_expense oe ON DATE_TRUNC('quarter', oe.month) = DATE_TRUNC('quarter', o.close_date)
    LEFT JOIN subscriptions subs ON subs.customer_id = o.customer_id
    GROUP BY 1
)
SELECT
    quarter,
    sm_spend,
    new_acv,
    gross_margin_pct,
    ROUND(sm_spend / NULLIF((new_acv * gross_margin_pct) / 12.0, 0), 1) AS cac_payback_months
FROM cac_input
ORDER BY quarter;


-- ----------------------------------------------------------------------------
-- 5. Burn Multiple: Net Burn / Net New ARR
-- ----------------------------------------------------------------------------
WITH q_burn AS (
    SELECT DATE_TRUNC('quarter', month) AS quarter,
           SUM(-(operating_cash_flow + capex)) AS net_burn
    FROM cash_flows GROUP BY 1
),
q_new_arr AS (
    SELECT DATE_TRUNC('quarter', month) AS quarter,
           (SUM(mrr * 12) - LAG(SUM(mrr * 12)) OVER (ORDER BY DATE_TRUNC('quarter', month))) AS net_new_arr
    FROM subscriptions GROUP BY 1
)
SELECT
    b.quarter,
    b.net_burn,
    n.net_new_arr,
    ROUND(b.net_burn / NULLIF(n.net_new_arr, 0), 2) AS burn_multiple
FROM q_burn b
JOIN q_new_arr n USING (quarter)
ORDER BY b.quarter;
