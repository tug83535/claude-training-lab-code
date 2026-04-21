-- ============================================================================
-- sales_pipeline_velocity.sql
-- Full Sales Pipeline Velocity Analytics
--
-- Computes stage-by-stage conversion rates, time-in-stage, pipeline velocity
-- per segment, and the "4 drivers" of sales velocity:
--
--     Velocity = (# Deals x Average Deal Size x Win Rate) / Sales Cycle Length
--
-- WHY THIS IS NOT NATIVE EXCEL / ONEDRIVE
-- ---------------------------------------
-- Sales teams live in CRM. Excel exports of CRM lose stage-progression data.
-- Calculating time-in-stage requires the audit history table, which Excel
-- cannot reshape. Window functions are native to SQL and trivial here.
--
-- USE CASE
-- --------
-- Sales ops weekly review. Shows which segment/rep/stage is the bottleneck.
--
-- ASSUMED TABLES:
--   opportunities      (deal_id, customer_id, amount, close_date, stage, segment,
--                       owner_id, created_date, won_flag)
--   stage_history      (deal_id, stage, stage_entered_at, stage_exited_at)
-- ============================================================================

-- ----------------------------------------------------------------------------
-- 1. Stage-by-stage funnel: how many entered, how many advanced, avg time
-- ----------------------------------------------------------------------------
WITH stage_durations AS (
    SELECT
        deal_id,
        stage,
        stage_entered_at,
        COALESCE(stage_exited_at, CURRENT_TIMESTAMP) AS stage_exited_at,
        EXTRACT(EPOCH FROM (COALESCE(stage_exited_at, CURRENT_TIMESTAMP) - stage_entered_at))/86400
            AS days_in_stage
    FROM stage_history
),
funnel AS (
    SELECT
        stage,
        COUNT(DISTINCT deal_id) AS deals_entered,
        ROUND(AVG(days_in_stage)::numeric, 1) AS avg_days_in_stage,
        ROUND(PERCENTILE_CONT(0.5) WITHIN GROUP (ORDER BY days_in_stage)::numeric, 1) AS median_days
    FROM stage_durations
    GROUP BY 1
)
SELECT
    stage,
    deals_entered,
    avg_days_in_stage,
    median_days,
    LAG(deals_entered) OVER (ORDER BY stage_order(stage)) AS prior_stage_deals,
    ROUND(deals_entered::numeric / NULLIF(LAG(deals_entered) OVER (ORDER BY stage_order(stage)), 0) * 100, 1)
        AS conversion_from_prior_pct
FROM funnel
ORDER BY stage_order(stage);


-- Helper: custom sort order for stages (Snowflake syntax; adapt per dialect).
-- CREATE OR REPLACE FUNCTION stage_order(s VARCHAR) RETURNS INT AS $$
--     CASE s WHEN 'Prospect' THEN 1 WHEN 'Qualified' THEN 2
--            WHEN 'Demo' THEN 3 WHEN 'Proposal' THEN 4
--            WHEN 'Negotiation' THEN 5 WHEN 'Closed Won' THEN 6
--            WHEN 'Closed Lost' THEN 7 ELSE 99 END
-- $$;


-- ----------------------------------------------------------------------------
-- 2. Quarterly velocity by segment (the "4 drivers" equation)
-- ----------------------------------------------------------------------------
WITH closed_q AS (
    SELECT
        DATE_TRUNC('quarter', close_date) AS quarter,
        segment,
        COUNT(*) FILTER (WHERE won_flag)  AS wins,
        COUNT(*)                          AS deals_closed,
        AVG(CASE WHEN won_flag THEN amount END) AS avg_win_size,
        AVG(EXTRACT(EPOCH FROM (close_date - created_date)) / 86400) AS avg_cycle_days
    FROM opportunities
    WHERE stage IN ('Closed Won', 'Closed Lost')
    GROUP BY 1, 2
)
SELECT
    quarter,
    segment,
    deals_closed,
    wins,
    ROUND(wins::numeric / NULLIF(deals_closed, 0) * 100, 1) AS win_rate_pct,
    ROUND(avg_win_size, 0)                                  AS avg_deal_size,
    ROUND(avg_cycle_days::numeric, 1)                       AS avg_cycle_days,
    ROUND(
        (deals_closed * avg_win_size * (wins::numeric / NULLIF(deals_closed, 0)))
        / NULLIF(avg_cycle_days, 0)
    , 0) AS pipeline_velocity_per_day
FROM closed_q
ORDER BY quarter, segment;


-- ----------------------------------------------------------------------------
-- 3. Stuck deals: open deals that have sat in one stage longer than the P75
-- ----------------------------------------------------------------------------
WITH current_stage AS (
    SELECT o.deal_id, o.amount, o.segment, o.owner_id, o.stage,
           MAX(sh.stage_entered_at) AS entered_current_stage_at
    FROM opportunities o
    JOIN stage_history sh ON sh.deal_id = o.deal_id AND sh.stage = o.stage
    WHERE o.stage NOT IN ('Closed Won', 'Closed Lost')
    GROUP BY 1, 2, 3, 4, 5
),
p75 AS (
    SELECT stage,
           PERCENTILE_CONT(0.75) WITHIN GROUP (
               ORDER BY EXTRACT(EPOCH FROM (stage_exited_at - stage_entered_at))/86400
           ) AS p75_days
    FROM stage_history
    WHERE stage_exited_at IS NOT NULL
    GROUP BY 1
)
SELECT
    c.deal_id,
    c.segment,
    c.owner_id,
    c.stage,
    ROUND(EXTRACT(EPOCH FROM (CURRENT_TIMESTAMP - c.entered_current_stage_at))/86400::numeric, 1)
        AS days_in_current_stage,
    p.p75_days AS stage_p75_days,
    c.amount
FROM current_stage c
JOIN p75 p ON p.stage = c.stage
WHERE EXTRACT(EPOCH FROM (CURRENT_TIMESTAMP - c.entered_current_stage_at))/86400 > p.p75_days
ORDER BY c.amount DESC, days_in_current_stage DESC;


-- ----------------------------------------------------------------------------
-- 4. "Commit model" forecast: weighted pipeline this quarter
-- ----------------------------------------------------------------------------
WITH stage_win_probs AS (
    -- Historical win rate per stage
    SELECT stage,
           COUNT(*) FILTER (WHERE won_flag)::numeric / NULLIF(COUNT(*), 0) AS historical_p_win
    FROM opportunities
    WHERE stage IN ('Closed Won', 'Closed Lost')
    GROUP BY 1
)
SELECT
    o.segment,
    o.owner_id,
    COUNT(*)                                              AS open_deals,
    SUM(o.amount)                                         AS unweighted_pipeline,
    SUM(o.amount * COALESCE(p.historical_p_win, 0.1))     AS weighted_forecast,
    SUM(CASE WHEN o.stage IN ('Negotiation', 'Verbal Commit')
             THEN o.amount END)                           AS commit_stage_pipeline
FROM opportunities o
LEFT JOIN stage_win_probs p USING (stage)
WHERE o.stage NOT IN ('Closed Won', 'Closed Lost')
  AND DATE_TRUNC('quarter', o.close_date) = DATE_TRUNC('quarter', CURRENT_DATE)
GROUP BY 1, 2
ORDER BY weighted_forecast DESC;
