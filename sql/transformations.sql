-- =============================================================================
-- transformations.sql
-- Keystone BenefitTech P&L — Allocation Pivot and Summary Views
-- =============================================================================
--
-- ENGINE:  SQLite 3.x
-- USAGE:   sqlite3 keystone_pnl.db < transformations.sql
-- PREREQ:  Run staging.sql first (creates fact_gl, dim_* tables)
-- PURPOSE: Allocation pivots, product/dept summaries, MoM variance calculations
--
-- =============================================================================


-- ─────────────────────────────────────────────────────────────────────────────
-- 1. ALLOCATION SHARES TABLE
-- ─────────────────────────────────────────────────────────────────────────────

CREATE TABLE IF NOT EXISTS allocation_shares (
    share_id      INTEGER PRIMARY KEY AUTOINCREMENT,
    share_type    TEXT NOT NULL,    -- 'revenue', 'aws_compute', 'headcount'
    product_name  TEXT NOT NULL REFERENCES dim_product(product_name),
    share_pct     REAL NOT NULL,
    effective_date TEXT NOT NULL DEFAULT (date('now')),
    UNIQUE(share_type, product_name, effective_date)
);

-- Default revenue shares (must sum to 1.0)
INSERT OR IGNORE INTO allocation_shares (share_type, product_name, share_pct) VALUES
    ('revenue', 'iGO',         0.50),
    ('revenue', 'Affirm',      0.25),
    ('revenue', 'InsureSight',  0.15),
    ('revenue', 'DocFast',     0.10);

-- Default AWS compute shares
INSERT OR IGNORE INTO allocation_shares (share_type, product_name, share_pct) VALUES
    ('aws_compute', 'iGO',         0.55),
    ('aws_compute', 'Affirm',      0.20),
    ('aws_compute', 'InsureSight',  0.15),
    ('aws_compute', 'DocFast',     0.10);

-- Default headcount shares
INSERT OR IGNORE INTO allocation_shares (share_type, product_name, share_pct) VALUES
    ('headcount', 'iGO',         0.40),
    ('headcount', 'Affirm',      0.25),
    ('headcount', 'InsureSight',  0.20),
    ('headcount', 'DocFast',     0.15);


-- ─────────────────────────────────────────────────────────────────────────────
-- 2. DEPARTMENT x PRODUCT ALLOCATION PIVOT
-- ─────────────────────────────────────────────────────────────────────────────

-- Actual spend by department and product (from GL)
CREATE VIEW IF NOT EXISTS v_dept_product_actual AS
SELECT
    dd.dept_name,
    dp.product_name,
    dt.month_abbrev,
    dt.month,
    dt.quarter,
    SUM(f.amount) AS total_amount,
    COUNT(f.fact_id) AS txn_count,
    AVG(f.amount) AS avg_amount
FROM fact_gl f
JOIN dim_department dd ON dd.dept_id = f.dept_id
JOIN dim_product dp ON dp.product_id = f.product_id
LEFT JOIN dim_date dt ON dt.date_key = f.date_key
GROUP BY dd.dept_name, dp.product_name, dt.month
ORDER BY dd.dept_name, dp.product_name, dt.month;


-- Allocated spend using revenue shares (for shared costs)
CREATE VIEW IF NOT EXISTS v_dept_product_allocated AS
SELECT
    dd.dept_name,
    dp.product_name,
    dt.month,
    dt.month_abbrev,
    -- Actual direct spend
    SUM(f.amount) AS direct_spend,
    -- Allocated share based on revenue shares
    SUM(f.amount) * als.share_pct AS allocated_spend,
    als.share_pct AS revenue_share
FROM fact_gl f
JOIN dim_department dd ON dd.dept_id = f.dept_id
JOIN dim_product dp ON dp.product_id = f.product_id
LEFT JOIN dim_date dt ON dt.date_key = f.date_key
LEFT JOIN allocation_shares als
    ON als.product_name = dp.product_name
    AND als.share_type = 'revenue'
GROUP BY dd.dept_name, dp.product_name, dt.month
ORDER BY dd.dept_name, dp.product_name, dt.month;


-- ─────────────────────────────────────────────────────────────────────────────
-- 3. PRODUCT LINE SUMMARY
-- ─────────────────────────────────────────────────────────────────────────────

CREATE VIEW IF NOT EXISTS v_product_summary AS
SELECT
    dp.product_name,
    dt.month,
    dt.month_abbrev,
    dt.quarter,
    SUM(f.amount) AS net_spend,
    SUM(f.abs_amount) AS gross_spend,
    COUNT(f.fact_id) AS txn_count,
    COUNT(DISTINCT f.vendor) AS vendor_count,
    AVG(f.amount) AS avg_txn,
    MAX(f.abs_amount) AS max_txn,
    -- Revenue share for contribution margin calculation
    als.share_pct AS revenue_share
FROM fact_gl f
JOIN dim_product dp ON dp.product_id = f.product_id
LEFT JOIN dim_date dt ON dt.date_key = f.date_key
LEFT JOIN allocation_shares als
    ON als.product_name = dp.product_name
    AND als.share_type = 'revenue'
GROUP BY dp.product_name, dt.month
ORDER BY dp.product_name, dt.month;


-- FY totals by product
CREATE VIEW IF NOT EXISTS v_product_fy_total AS
SELECT
    dp.product_name,
    SUM(f.amount) AS fy_net_spend,
    SUM(f.abs_amount) AS fy_gross_spend,
    COUNT(f.fact_id) AS fy_txn_count,
    COUNT(DISTINCT f.vendor) AS fy_vendor_count,
    ROUND(SUM(f.amount) * 1.0 /
        (SELECT SUM(amount) FROM fact_gl), 4) AS spend_share
FROM fact_gl f
JOIN dim_product dp ON dp.product_id = f.product_id
GROUP BY dp.product_name
ORDER BY fy_net_spend DESC;


-- ─────────────────────────────────────────────────────────────────────────────
-- 4. DEPARTMENT SUMMARY
-- ─────────────────────────────────────────────────────────────────────────────

CREATE VIEW IF NOT EXISTS v_department_summary AS
SELECT
    dd.dept_name,
    dt.month,
    dt.month_abbrev,
    SUM(f.amount) AS net_spend,
    SUM(f.abs_amount) AS gross_spend,
    COUNT(f.fact_id) AS txn_count,
    COUNT(DISTINCT f.vendor) AS vendor_count
FROM fact_gl f
JOIN dim_department dd ON dd.dept_id = f.dept_id
LEFT JOIN dim_date dt ON dt.date_key = f.date_key
GROUP BY dd.dept_name, dt.month
ORDER BY dd.dept_name, dt.month;


-- FY totals by department
CREATE VIEW IF NOT EXISTS v_department_fy_total AS
SELECT
    dd.dept_name,
    SUM(f.amount) AS fy_net_spend,
    COUNT(f.fact_id) AS fy_txn_count,
    ROUND(SUM(f.amount) * 1.0 /
        (SELECT SUM(amount) FROM fact_gl), 4) AS spend_share
FROM fact_gl f
JOIN dim_department dd ON dd.dept_id = f.dept_id
GROUP BY dd.dept_name
ORDER BY fy_net_spend DESC;


-- ─────────────────────────────────────────────────────────────────────────────
-- 5. MONTH-OVER-MONTH VARIANCE
-- ─────────────────────────────────────────────────────────────────────────────

-- Product-level MoM variance
CREATE VIEW IF NOT EXISTS v_product_mom_variance AS
SELECT
    curr.product_name,
    curr.month AS current_month,
    curr.month_abbrev AS current_month_name,
    curr.net_spend AS current_spend,
    prev.net_spend AS prior_spend,
    curr.net_spend - COALESCE(prev.net_spend, 0) AS dollar_change,
    CASE
        WHEN COALESCE(prev.net_spend, 0) = 0 THEN NULL
        ELSE ROUND((curr.net_spend - prev.net_spend) * 100.0 / ABS(prev.net_spend), 2)
    END AS pct_change,
    CASE
        WHEN prev.net_spend IS NULL THEN 'NEW'
        WHEN ABS(curr.net_spend - prev.net_spend) * 100.0 / ABS(prev.net_spend) > 15 THEN 'FLAG'
        ELSE 'OK'
    END AS variance_status
FROM v_product_summary curr
LEFT JOIN v_product_summary prev
    ON prev.product_name = curr.product_name
    AND prev.month = curr.month - 1
WHERE curr.month IS NOT NULL
ORDER BY curr.product_name, curr.month;


-- Department-level MoM variance
CREATE VIEW IF NOT EXISTS v_department_mom_variance AS
SELECT
    curr.dept_name,
    curr.month AS current_month,
    curr.month_abbrev AS current_month_name,
    curr.net_spend AS current_spend,
    prev.net_spend AS prior_spend,
    curr.net_spend - COALESCE(prev.net_spend, 0) AS dollar_change,
    CASE
        WHEN COALESCE(prev.net_spend, 0) = 0 THEN NULL
        ELSE ROUND((curr.net_spend - prev.net_spend) * 100.0 / ABS(prev.net_spend), 2)
    END AS pct_change,
    CASE
        WHEN prev.net_spend IS NULL THEN 'NEW'
        WHEN ABS(curr.net_spend - prev.net_spend) * 100.0 / ABS(prev.net_spend) > 15 THEN 'FLAG'
        ELSE 'OK'
    END AS variance_status
FROM v_department_summary curr
LEFT JOIN v_department_summary prev
    ON prev.dept_name = curr.dept_name
    AND prev.month = curr.month - 1
WHERE curr.month IS NOT NULL
ORDER BY curr.dept_name, curr.month;


-- ─────────────────────────────────────────────────────────────────────────────
-- 6. EXPENSE CATEGORY MIX
-- ─────────────────────────────────────────────────────────────────────────────

CREATE VIEW IF NOT EXISTS v_category_mix AS
SELECT
    ec.category_name,
    dt.month,
    dt.month_abbrev,
    SUM(f.amount) AS net_spend,
    COUNT(f.fact_id) AS txn_count,
    ROUND(SUM(f.amount) * 100.0 /
        (SELECT SUM(amount) FROM fact_gl f2
         LEFT JOIN dim_date dt2 ON dt2.date_key = f2.date_key
         WHERE dt2.month = dt.month), 2) AS pct_of_month
FROM fact_gl f
JOIN dim_expense_category ec ON ec.category_id = f.category_id
LEFT JOIN dim_date dt ON dt.date_key = f.date_key
GROUP BY ec.category_name, dt.month
ORDER BY ec.category_name, dt.month;


-- =============================================================================
-- POWER QUERY M EQUIVALENT — Allocation Pivot
-- =============================================================================
--
-- let
--     GLSource = Excel.CurrentWorkbook(){[Name="CrossfireHiddenWorksheet"]}[Content],
--     Typed = Table.TransformColumnTypes(GLSource, {
--         {"Department", type text}, {"Product", type text}, {"Amount", type number}
--     }),
--     Grouped = Table.Group(Typed, {"Department", "Product"}, {
--         {"TotalAmount", each List.Sum([Amount]), type number},
--         {"TxnCount", each Table.RowCount(_), Int64.Type}
--     }),
--     Pivoted = Table.Pivot(Grouped, List.Distinct(Grouped[Product]),
--         "Product", "TotalAmount", List.Sum),
--     -- Add revenue share allocation
--     ShareTable = #table({"Product","Share"}, {
--         {"iGO",0.50}, {"Affirm",0.25}, {"InsureSight",0.15}, {"DocFast",0.10}
--     }),
--     Merged = Table.NestedJoin(Pivoted, {"Department"}, Pivoted, {"Department"},
--         "Self", JoinKind.Inner)
-- in
--     Pivoted
--
-- =============================================================================
