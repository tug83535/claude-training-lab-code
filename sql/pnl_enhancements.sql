-- =============================================================================
-- pnl_enhancements.sql
-- Keystone BenefitTech P&L — SQL Enhancements (S1–S5)
-- =============================================================================
--
-- FILE:    pnl_enhancements.sql
-- PURPOSE: 5 new SQL additions recommended by the Code Audit:
--            S1. Budget vs Actual tables + queries
--            S2. Allocation audit trail with triggers
--            S3. Rolling 12-month P&L summary view
--            S4. Vendor contract renewal calendar
--            S5. Allocation reconciliation queries
--
-- USAGE:   sqlite3 keystone_pnl.db < pnl_enhancements.sql
-- PREREQ:  Run pnl_create_tables.sql first
--
-- =============================================================================


-- ═════════════════════════════════════════════════════════════════════════════
-- S1: BUDGET VS ACTUAL — Tables + Queries
-- ═════════════════════════════════════════════════════════════════════════════

-- Budget dimension table — supports multiple budget versions (v1, reforecast)
CREATE TABLE IF NOT EXISTS dim_budget (
    budget_id       INTEGER PRIMARY KEY AUTOINCREMENT,
    fiscal_year     TEXT    NOT NULL,
    month           INTEGER NOT NULL,
    product         TEXT    NOT NULL,
    department      TEXT    NOT NULL,
    line_item       TEXT    NOT NULL DEFAULT 'Total Expense',
    budget_amount   REAL    NOT NULL DEFAULT 0,
    budget_version  TEXT    NOT NULL DEFAULT 'v1',
    notes           TEXT,
    created_at      TEXT    NOT NULL DEFAULT (datetime('now')),
    updated_at      TEXT    NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_budget_lookup
ON dim_budget(fiscal_year, month, product, department, budget_version);

-- Seed example budget data (replace with actuals)
INSERT OR IGNORE INTO dim_budget (fiscal_year, month, product, department, line_item, budget_amount, budget_version)
VALUES
    -- January budgets by product × department
    ('FY2025', 1, 'iGO',         'NetOps',     'Total Expense', 15000, 'v1'),
    ('FY2025', 1, 'iGO',         'Security',   'Total Expense',  5000, 'v1'),
    ('FY2025', 1, 'iGO',         'Support',    'Total Expense',  8000, 'v1'),
    ('FY2025', 1, 'iGO',         'R&D',        'Total Expense', 12000, 'v1'),
    ('FY2025', 1, 'Affirm',      'NetOps',     'Total Expense',  8000, 'v1'),
    ('FY2025', 1, 'Affirm',      'Security',   'Total Expense',  3000, 'v1'),
    ('FY2025', 1, 'Affirm',      'Support',    'Total Expense',  5000, 'v1'),
    ('FY2025', 1, 'Affirm',      'R&D',        'Total Expense',  7000, 'v1'),
    ('FY2025', 1, 'InsureSight',  'NetOps',     'Total Expense',  5000, 'v1'),
    ('FY2025', 1, 'InsureSight',  'R&D',        'Total Expense',  6000, 'v1'),
    ('FY2025', 1, 'DocFast',      'NetOps',     'Total Expense',  2000, 'v1'),
    ('FY2025', 1, 'DocFast',      'R&D',        'Total Expense',  3000, 'v1');


-- Q-BVA1: Budget vs Actual by Product × Month
SELECT
    b.fiscal_year,
    b.month,
    b.product,
    b.department,
    b.budget_amount,
    COALESCE(a.actual_spend, 0)     AS actual_spend,
    b.budget_amount - COALESCE(a.actual_spend, 0) AS variance_dollar,
    CASE
        WHEN b.budget_amount != 0
        THEN ROUND((COALESCE(a.actual_spend, 0) - b.budget_amount) * 100.0 / ABS(b.budget_amount), 2)
        ELSE 0
    END AS variance_pct,
    CASE
        WHEN COALESCE(a.actual_spend, 0) > b.budget_amount * 1.10 THEN 'OVER BUDGET'
        WHEN COALESCE(a.actual_spend, 0) < b.budget_amount * 0.90 THEN 'UNDER BUDGET'
        ELSE 'ON TRACK'
    END AS status
FROM dim_budget b
LEFT JOIN (
    SELECT product, department, month,
           SUM(abs_amount) AS actual_spend
    FROM fact_gl
    WHERE product != '' AND department != ''
    GROUP BY product, department, month
) a ON b.product = a.product AND b.department = a.department AND b.month = a.month
WHERE b.budget_version = 'v1'
ORDER BY b.month, b.product, b.department;


-- Q-BVA2: Budget vs Actual Summary by Product
SELECT
    b.product,
    SUM(b.budget_amount)                        AS total_budget,
    SUM(COALESCE(a.actual_spend, 0))            AS total_actual,
    SUM(b.budget_amount) - SUM(COALESCE(a.actual_spend, 0)) AS total_variance,
    CASE
        WHEN SUM(b.budget_amount) != 0
        THEN ROUND(
            (SUM(COALESCE(a.actual_spend, 0)) - SUM(b.budget_amount)) * 100.0
            / ABS(SUM(b.budget_amount)), 2
        )
        ELSE 0
    END AS variance_pct
FROM dim_budget b
LEFT JOIN (
    SELECT product, department, month,
           SUM(abs_amount) AS actual_spend
    FROM fact_gl
    WHERE product != '' AND department != ''
    GROUP BY product, department, month
) a ON b.product = a.product AND b.department = a.department AND b.month = a.month
WHERE b.budget_version = 'v1'
GROUP BY b.product
ORDER BY total_actual DESC;


-- ═════════════════════════════════════════════════════════════════════════════
-- S2: ALLOCATION AUDIT TRAIL — Table + Triggers
-- ═════════════════════════════════════════════════════════════════════════════

CREATE TABLE IF NOT EXISTS allocation_audit (
    audit_id    INTEGER PRIMARY KEY AUTOINCREMENT,
    table_name  TEXT    NOT NULL,
    record_id   INTEGER NOT NULL,
    column_name TEXT    NOT NULL,
    old_value   TEXT,
    new_value   TEXT,
    changed_by  TEXT    NOT NULL DEFAULT 'system',
    changed_at  TEXT    NOT NULL DEFAULT (datetime('now')),
    change_type TEXT    NOT NULL DEFAULT 'UPDATE'  -- INSERT, UPDATE, DELETE
);

CREATE INDEX IF NOT EXISTS idx_audit_table ON allocation_audit(table_name, record_id);
CREATE INDEX IF NOT EXISTS idx_audit_date ON allocation_audit(changed_at);

-- Trigger: Track revenue share changes
DROP TRIGGER IF EXISTS tr_product_revenue_share;
CREATE TRIGGER tr_product_revenue_share
AFTER UPDATE OF revenue_share ON dim_product
WHEN OLD.revenue_share != NEW.revenue_share
BEGIN
    INSERT INTO allocation_audit (table_name, record_id, column_name, old_value, new_value)
    VALUES ('dim_product', NEW.product_id, 'revenue_share',
            CAST(OLD.revenue_share AS TEXT), CAST(NEW.revenue_share AS TEXT));
END;

-- Trigger: Track AWS compute share changes
DROP TRIGGER IF EXISTS tr_product_aws_share;
CREATE TRIGGER tr_product_aws_share
AFTER UPDATE OF aws_compute_share ON dim_product
WHEN OLD.aws_compute_share != NEW.aws_compute_share
BEGIN
    INSERT INTO allocation_audit (table_name, record_id, column_name, old_value, new_value)
    VALUES ('dim_product', NEW.product_id, 'aws_compute_share',
            CAST(OLD.aws_compute_share AS TEXT), CAST(NEW.aws_compute_share AS TEXT));
END;

-- Trigger: Track department allocation method changes
DROP TRIGGER IF EXISTS tr_dept_allocation;
CREATE TRIGGER tr_dept_allocation
AFTER UPDATE OF allocation_method ON dim_department
WHEN OLD.allocation_method != NEW.allocation_method
BEGIN
    INSERT INTO allocation_audit (table_name, record_id, column_name, old_value, new_value)
    VALUES ('dim_department', NEW.department_id, 'allocation_method',
            OLD.allocation_method, NEW.allocation_method);
END;

-- Q-AUDIT1: View recent allocation changes
-- SELECT * FROM allocation_audit ORDER BY changed_at DESC LIMIT 20;


-- ═════════════════════════════════════════════════════════════════════════════
-- S3: ROLLING 12-MONTH P&L SUMMARY VIEW
-- ═════════════════════════════════════════════════════════════════════════════

DROP VIEW IF EXISTS v_rolling_12m_product;
CREATE VIEW v_rolling_12m_product AS
WITH max_month AS (
    SELECT MAX(month) AS latest FROM fact_gl WHERE month IS NOT NULL
),
ttm AS (
    SELECT
        f.product,
        SUM(f.amount)       AS ttm_spend,
        SUM(f.abs_amount)   AS ttm_abs_spend,
        COUNT(*)            AS ttm_txns,
        COUNT(DISTINCT f.vendor)      AS ttm_vendors,
        COUNT(DISTINCT f.department)  AS ttm_departments
    FROM fact_gl f, max_month m
    WHERE f.product != ''
      AND f.month BETWEEN (m.latest - 11) AND m.latest
    GROUP BY f.product
),
prior_ttm AS (
    SELECT
        f.product,
        SUM(f.abs_amount) AS prior_abs_spend
    FROM fact_gl f, max_month m
    WHERE f.product != ''
      AND f.month BETWEEN (m.latest - 23) AND (m.latest - 12)
    GROUP BY f.product
)
SELECT
    t.product,
    t.ttm_abs_spend,
    t.ttm_spend,
    t.ttm_txns,
    t.ttm_vendors,
    COALESCE(p.prior_abs_spend, 0) AS prior_ttm_abs_spend,
    CASE
        WHEN COALESCE(p.prior_abs_spend, 0) != 0
        THEN ROUND((t.ttm_abs_spend - p.prior_abs_spend) * 100.0 / p.prior_abs_spend, 2)
        ELSE NULL
    END AS yoy_change_pct,
    ROUND(t.ttm_abs_spend * 100.0 / NULLIF(
        (SELECT SUM(abs_amount) FROM fact_gl f2, max_month m2
         WHERE f2.product != '' AND f2.month BETWEEN (m2.latest - 11) AND m2.latest), 0
    ), 2) AS pct_of_total
FROM ttm t
LEFT JOIN prior_ttm p ON t.product = p.product
ORDER BY t.ttm_abs_spend DESC;

-- Similar view by department
DROP VIEW IF EXISTS v_rolling_12m_department;
CREATE VIEW v_rolling_12m_department AS
WITH max_month AS (
    SELECT MAX(month) AS latest FROM fact_gl WHERE month IS NOT NULL
)
SELECT
    f.department,
    SUM(f.abs_amount)               AS ttm_abs_spend,
    SUM(f.amount)                   AS ttm_net_spend,
    COUNT(*)                        AS ttm_txns,
    COUNT(DISTINCT f.vendor)        AS ttm_vendors,
    COUNT(DISTINCT f.product)       AS products_served,
    ROUND(SUM(f.abs_amount) * 100.0 / NULLIF(
        (SELECT SUM(abs_amount) FROM fact_gl f2, max_month m2
         WHERE f2.department != '' AND f2.month BETWEEN (m2.latest - 11) AND m2.latest), 0
    ), 2) AS pct_of_total
FROM fact_gl f, max_month m
WHERE f.department != ''
  AND f.month BETWEEN (m.latest - 11) AND m.latest
GROUP BY f.department
ORDER BY ttm_abs_spend DESC;


-- ═════════════════════════════════════════════════════════════════════════════
-- S4: VENDOR CONTRACT RENEWAL CALENDAR
-- ═════════════════════════════════════════════════════════════════════════════

-- Identifies vendors with consistent monthly spending (likely fixed contracts)
-- and estimates renewal timing

DROP VIEW IF EXISTS v_vendor_contracts;
CREATE VIEW v_vendor_contracts AS
WITH vendor_monthly AS (
    SELECT
        vendor,
        month,
        SUM(abs_amount) AS monthly_spend
    FROM fact_gl
    WHERE vendor != '' AND month IS NOT NULL
    GROUP BY vendor, month
),
vendor_stats AS (
    SELECT
        vendor,
        COUNT(DISTINCT month)   AS months_active,
        AVG(monthly_spend)      AS avg_monthly,
        MIN(monthly_spend)      AS min_monthly,
        MAX(monthly_spend)      AS max_monthly,
        -- Coefficient of variation: lower = more consistent = likely contract
        CASE
            WHEN AVG(monthly_spend) != 0
            THEN ROUND(
                (MAX(monthly_spend) - MIN(monthly_spend)) * 100.0 / AVG(monthly_spend), 2
            )
            ELSE 0
        END AS spend_variability_pct,
        SUM(monthly_spend) AS total_spend
    FROM vendor_monthly
    GROUP BY vendor
    HAVING COUNT(DISTINCT month) >= 3  -- At least 3 months of data
)
SELECT
    vs.vendor,
    vs.months_active,
    ROUND(vs.avg_monthly, 2)        AS avg_monthly_spend,
    ROUND(vs.total_spend, 2)        AS total_spend,
    vs.spend_variability_pct,
    CASE
        WHEN vs.spend_variability_pct < 15  THEN 'FIXED CONTRACT (likely)'
        WHEN vs.spend_variability_pct < 40  THEN 'SEMI-VARIABLE'
        ELSE 'VARIABLE / PROJECT-BASED'
    END AS contract_type,
    dv.is_aws,
    dv.is_software,
    dv.last_seen_date,
    CASE
        WHEN vs.months_active >= 12 THEN 'ANNUAL RENEWAL DUE'
        WHEN vs.months_active >= 6  THEN 'MID-TERM'
        ELSE 'NEW VENDOR'
    END AS renewal_status
FROM vendor_stats vs
LEFT JOIN dim_vendor dv ON vs.vendor = dv.vendor_name
ORDER BY total_spend DESC;


-- ═════════════════════════════════════════════════════════════════════════════
-- S5: ALLOCATION RECONCILIATION QUERIES
-- ═════════════════════════════════════════════════════════════════════════════

-- Q-RECON1: Do revenue shares sum to 100%?
SELECT
    'Revenue Shares' AS check_name,
    SUM(revenue_share) AS actual_sum,
    1.0 AS expected,
    ROUND(ABS(SUM(revenue_share) - 1.0), 6) AS difference,
    CASE
        WHEN ABS(SUM(revenue_share) - 1.0) < 0.001 THEN 'PASS'
        ELSE 'FAIL'
    END AS status
FROM dim_product
WHERE is_active = 1;

-- Q-RECON2: Do AWS compute shares sum to 100%?
SELECT
    'AWS Compute Shares' AS check_name,
    SUM(aws_compute_share) AS actual_sum,
    1.0 AS expected,
    ROUND(ABS(SUM(aws_compute_share) - 1.0), 6) AS difference,
    CASE
        WHEN ABS(SUM(aws_compute_share) - 1.0) < 0.001 THEN 'PASS'
        ELSE 'FAIL'
    END AS status
FROM dim_product
WHERE is_active = 1;

-- Q-RECON3: Do allocated costs per product equal department totals?
SELECT
    f.department,
    dd.allocation_method,
    SUM(f.abs_amount) AS total_dept_spend,
    SUM(CASE WHEN f.product = 'iGO'         THEN f.abs_amount ELSE 0 END) AS igo_alloc,
    SUM(CASE WHEN f.product = 'Affirm'       THEN f.abs_amount ELSE 0 END) AS affirm_alloc,
    SUM(CASE WHEN f.product = 'InsureSight'   THEN f.abs_amount ELSE 0 END) AS insuresight_alloc,
    SUM(CASE WHEN f.product = 'DocFast'       THEN f.abs_amount ELSE 0 END) AS docfast_alloc,
    -- Check: sum of product allocations = department total
    ROUND(
        SUM(CASE WHEN f.product IN ('iGO','Affirm','InsureSight','DocFast') THEN f.abs_amount ELSE 0 END)
        - SUM(f.abs_amount), 2
    ) AS allocation_gap,
    CASE
        WHEN ABS(
            SUM(CASE WHEN f.product IN ('iGO','Affirm','InsureSight','DocFast') THEN f.abs_amount ELSE 0 END)
            - SUM(f.abs_amount)
        ) < 1.0 THEN 'PASS'
        ELSE 'FAIL'
    END AS status
FROM fact_gl f
LEFT JOIN dim_department dd ON f.department_id = dd.department_id
WHERE f.department != ''
GROUP BY f.department, dd.allocation_method
ORDER BY total_dept_spend DESC;

-- Q-RECON4: Product spend share vs configured revenue share
SELECT
    dp.product_name,
    dp.revenue_share                AS configured_share,
    ROUND(
        SUM(f.abs_amount) * 100.0 / NULLIF(
            (SELECT SUM(abs_amount) FROM fact_gl WHERE product != ''), 0
        ), 2
    )                               AS actual_spend_pct,
    ROUND(
        SUM(f.abs_amount) * 100.0 / NULLIF(
            (SELECT SUM(abs_amount) FROM fact_gl WHERE product != ''), 0
        ) - dp.revenue_share * 100, 2
    )                               AS gap_pct_points,
    CASE
        WHEN ABS(
            SUM(f.abs_amount) * 100.0 / NULLIF(
                (SELECT SUM(abs_amount) FROM fact_gl WHERE product != ''), 0
            ) - dp.revenue_share * 100
        ) < 5 THEN 'ALIGNED'
        WHEN SUM(f.abs_amount) * 100.0 / NULLIF(
            (SELECT SUM(abs_amount) FROM fact_gl WHERE product != ''), 0
        ) > dp.revenue_share * 100 THEN 'OVER-ALLOCATED'
        ELSE 'UNDER-ALLOCATED'
    END AS alignment_status
FROM fact_gl f
LEFT JOIN dim_product dp ON f.product_id = dp.product_id
WHERE f.product != ''
GROUP BY dp.product_name, dp.revenue_share
ORDER BY actual_spend_pct DESC;


-- =============================================================================
-- VERIFY
-- =============================================================================

SELECT '✓ SQL Enhancements loaded — S1 through S5' AS status;
SELECT name, type FROM sqlite_master
WHERE name LIKE '%budget%' OR name LIKE '%audit%' OR name LIKE '%rolling%'
   OR name LIKE '%contract%'
ORDER BY type, name;
