-- ============================================================================
-- customer_360_view.sql
-- Unified Customer 360 View: one row per customer joining CRM, billing,
-- support, product-usage, and success signals. The single source of truth
-- for executive, sales, CS, and finance reporting.
--
-- WHY THIS IS NOT NATIVE EXCEL / ONEDRIVE
-- ---------------------------------------
-- Pulling "everything we know about one customer" is the #1 ask of every
-- BI team. Excel can VLOOKUP maybe 3 sources; a proper 360 joins 8-12.
-- This view lives in the warehouse, refreshes nightly, and powers Tableau
-- / Power BI / Looker dashboards AND Excel pivots downstream.
--
-- USE CASE
-- --------
-- - CSM quarterly business review deck builder
-- - Account executive pre-call research
-- - Finance credit-risk evaluation
-- - Marketing targeted campaign segmentation
--
-- TARGET DIALECT: Snowflake / BigQuery-flavored SQL
-- ============================================================================

CREATE OR REPLACE VIEW customer_360 AS
WITH base AS (
    SELECT
        c.customer_id,
        c.account_name,
        c.industry,
        c.region,
        c.segment,                          -- SMB / Mid / Enterprise
        c.signup_date,
        c.logo_url,
        c.account_manager_id,
        c.primary_email,
        c.hq_country
    FROM customers c
    WHERE c.is_deleted = FALSE
),

-- -----------------------------------------------------------------
-- A. Commercial profile (subscriptions + billing)
-- -----------------------------------------------------------------
commercial AS (
    SELECT
        s.customer_id,
        SUM(CASE WHEN s.end_date IS NULL OR s.end_date >= CURRENT_DATE
                 THEN s.mrr END) * 12                              AS current_arr,
        MAX(CASE WHEN s.end_date IS NULL OR s.end_date >= CURRENT_DATE
                 THEN s.plan END)                                  AS current_plan,
        MIN(s.start_date)                                          AS first_contract_date,
        COUNT(DISTINCT s.plan)                                     AS plans_ever_purchased,
        SUM(b.amount)                                              AS lifetime_billed,
        SUM(CASE WHEN b.status = 'PAID' THEN b.amount END)         AS lifetime_paid,
        SUM(CASE WHEN b.status = 'OVERDUE' THEN b.amount END)      AS current_overdue,
        MAX(b.bill_date)                                           AS last_bill_date
    FROM subscriptions s
    LEFT JOIN billing b USING (customer_id)
    GROUP BY 1
),

-- -----------------------------------------------------------------
-- B. Product usage signals (30 and 90 day)
-- -----------------------------------------------------------------
usage AS (
    SELECT
        p.customer_id,
        COUNT(DISTINCT p.user_id)
            FILTER (WHERE p.event_at >= CURRENT_DATE - INTERVAL '30 days')  AS active_users_30d,
        COUNT(DISTINCT p.user_id)
            FILTER (WHERE p.event_at >= CURRENT_DATE - INTERVAL '90 days')  AS active_users_90d,
        COUNT(*)
            FILTER (WHERE p.event_at >= CURRENT_DATE - INTERVAL '30 days')  AS events_30d,
        MAX(p.event_at)                                                     AS last_event_at,
        DATE_PART('day', CURRENT_DATE - MAX(p.event_at))                    AS days_since_last_event,
        COUNT(DISTINCT p.feature_area)
            FILTER (WHERE p.event_at >= CURRENT_DATE - INTERVAL '90 days')  AS feature_breadth_90d
    FROM product_events p
    GROUP BY 1
),

-- -----------------------------------------------------------------
-- C. Support signals
-- -----------------------------------------------------------------
support AS (
    SELECT
        t.customer_id,
        COUNT(*) FILTER (WHERE t.created_at >= CURRENT_DATE - INTERVAL '90 days') AS tickets_90d,
        COUNT(*) FILTER (WHERE t.priority IN ('Urgent', 'High')
                           AND t.created_at >= CURRENT_DATE - INTERVAL '90 days') AS hi_prio_90d,
        AVG(EXTRACT(EPOCH FROM (t.first_response_at - t.created_at))/3600)         AS avg_first_resp_hours,
        AVG(CASE WHEN t.csat_score IS NOT NULL THEN t.csat_score END)              AS avg_csat_score
    FROM support_tickets t
    GROUP BY 1
),

-- -----------------------------------------------------------------
-- D. CS signals (NPS + health scores)
-- -----------------------------------------------------------------
cs AS (
    SELECT
        customer_id,
        AVG(score) FILTER (WHERE survey_type = 'NPS'
                             AND answered_at >= CURRENT_DATE - INTERVAL '180 days') AS nps_180d,
        MAX(health_score)                                                            AS latest_health_score,
        MAX(stage)                                                                   AS lifecycle_stage
    FROM cs_signals
    GROUP BY 1
),

-- -----------------------------------------------------------------
-- E. Open opportunities (pipeline)
-- -----------------------------------------------------------------
pipeline AS (
    SELECT
        customer_id,
        SUM(amount) FILTER (WHERE stage NOT IN ('Closed Won', 'Closed Lost')) AS open_pipeline_amt,
        COUNT(*)    FILTER (WHERE stage NOT IN ('Closed Won', 'Closed Lost')) AS open_pipeline_count,
        MAX(close_date) FILTER (WHERE stage NOT IN ('Closed Won', 'Closed Lost')) AS next_close_date
    FROM opportunities
    GROUP BY 1
)

-- -----------------------------------------------------------------
-- Assemble
-- -----------------------------------------------------------------
SELECT
    b.customer_id,
    b.account_name,
    b.industry,
    b.region,
    b.segment,
    b.signup_date,
    DATE_PART('day', CURRENT_DATE - b.signup_date) / 30.44 AS tenure_months,
    b.account_manager_id,

    -- Commercial
    COALESCE(c.current_arr, 0)          AS arr,
    c.current_plan,
    c.first_contract_date,
    c.plans_ever_purchased,
    c.lifetime_billed,
    c.lifetime_paid,
    c.current_overdue,
    c.last_bill_date,

    -- Usage
    COALESCE(u.active_users_30d, 0)     AS active_users_30d,
    COALESCE(u.active_users_90d, 0)     AS active_users_90d,
    u.last_event_at,
    u.days_since_last_event,
    u.feature_breadth_90d,

    -- Support
    COALESCE(sup.tickets_90d, 0)        AS tickets_90d,
    COALESCE(sup.hi_prio_90d, 0)        AS hi_prio_tickets_90d,
    sup.avg_first_resp_hours,
    sup.avg_csat_score,

    -- CS
    cs.nps_180d,
    cs.latest_health_score,
    cs.lifecycle_stage,

    -- Pipeline
    COALESCE(p.open_pipeline_amt, 0)    AS open_pipeline,
    p.open_pipeline_count,
    p.next_close_date,

    -- Derived traffic-light
    CASE
        WHEN cs.latest_health_score < 40 OR u.days_since_last_event > 45 OR c.current_overdue > 0
             THEN 'RED'
        WHEN cs.latest_health_score < 70 OR u.days_since_last_event > 14 OR sup.hi_prio_90d >= 3
             THEN 'YELLOW'
        ELSE 'GREEN'
    END AS health_tier,

    CURRENT_TIMESTAMP AS snapshot_at

FROM base b
LEFT JOIN commercial c ON c.customer_id = b.customer_id
LEFT JOIN usage      u ON u.customer_id = b.customer_id
LEFT JOIN support  sup ON sup.customer_id = b.customer_id
LEFT JOIN cs         cs ON cs.customer_id = b.customer_id
LEFT JOIN pipeline   p ON p.customer_id = b.customer_id;

COMMENT ON VIEW customer_360 IS
   'Single-row-per-customer unified view. Refreshed by nightly warehouse job.';
