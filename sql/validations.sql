-- =============================================================================
-- validations.sql
-- Keystone BenefitTech P&L — Data Validation and Integrity Checks
-- =============================================================================
--
-- ENGINE:  SQLite 3.x
-- USAGE:   sqlite3 keystone_pnl.db < validations.sql
-- PREREQ:  Run staging.sql and transformations.sql first
-- PURPOSE: Referential integrity, orphan detection, completeness,
--          balance validation, allocation reconciliation
--
-- =============================================================================


-- ─────────────────────────────────────────────────────────────────────────────
-- 1. REFERENTIAL INTEGRITY CHECKS
-- ─────────────────────────────────────────────────────────────────────────────

-- GL rows with unknown products (not in dim_product)
CREATE VIEW IF NOT EXISTS v_check_unknown_products AS
SELECT
    'Unknown Product' AS check_name,
    s.gl_id,
    s.product AS unknown_value,
    s.amount,
    s.gl_date,
    CASE WHEN COUNT(*) > 0 THEN 'FAIL' ELSE 'PASS' END AS status
FROM stg_gl_raw s
LEFT JOIN dim_product dp ON dp.product_name = s.product
WHERE dp.product_id IS NULL AND s.product IS NOT NULL
GROUP BY s.product;


-- GL rows with unknown departments
CREATE VIEW IF NOT EXISTS v_check_unknown_departments AS
SELECT
    'Unknown Department' AS check_name,
    s.gl_id,
    s.department AS unknown_value,
    s.amount,
    s.gl_date,
    CASE WHEN COUNT(*) > 0 THEN 'FAIL' ELSE 'PASS' END AS status
FROM stg_gl_raw s
LEFT JOIN dim_department dd ON dd.dept_name = s.department
WHERE dd.dept_id IS NULL AND s.department IS NOT NULL
GROUP BY s.department;


-- GL rows with unknown expense categories
CREATE VIEW IF NOT EXISTS v_check_unknown_categories AS
SELECT
    'Unknown Expense Category' AS check_name,
    s.gl_id,
    s.expense_category AS unknown_value,
    s.amount,
    s.gl_date
FROM stg_gl_raw s
LEFT JOIN dim_expense_category ec ON ec.category_name = s.expense_category
WHERE ec.category_id IS NULL AND s.expense_category IS NOT NULL;


-- ─────────────────────────────────────────────────────────────────────────────
-- 2. ORPHAN DETECTION
-- ─────────────────────────────────────────────────────────────────────────────

-- Fact rows with NULL foreign keys (failed to join during ETL)
CREATE VIEW IF NOT EXISTS v_check_orphan_facts AS
SELECT
    'Orphaned GL Rows' AS check_name,
    f.fact_id,
    f.gl_id,
    CASE
        WHEN f.dept_id IS NULL THEN 'Missing department'
        WHEN f.product_id IS NULL THEN 'Missing product'
        WHEN f.category_id IS NULL THEN 'Missing category'
        WHEN f.date_key IS NULL THEN 'Missing date'
        ELSE 'Unknown'
    END AS orphan_reason,
    f.amount
FROM fact_gl f
WHERE f.dept_id IS NULL
   OR f.product_id IS NULL
   OR f.category_id IS NULL
   OR f.date_key IS NULL;


-- Products in allocation_shares but not in fact_gl
CREATE VIEW IF NOT EXISTS v_check_unallocated_products AS
SELECT
    'Product with Shares but No Transactions' AS check_name,
    als.product_name,
    als.share_type,
    als.share_pct
FROM allocation_shares als
LEFT JOIN (
    SELECT DISTINCT dp.product_name
    FROM fact_gl f
    JOIN dim_product dp ON dp.product_id = f.product_id
) actual ON actual.product_name = als.product_name
WHERE actual.product_name IS NULL;


-- ─────────────────────────────────────────────────────────────────────────────
-- 3. COMPLETENESS CHECKS
-- ─────────────────────────────────────────────────────────────────────────────

-- Month coverage: which months have data?
CREATE VIEW IF NOT EXISTS v_check_month_coverage AS
SELECT
    dt.month,
    dt.month_abbrev,
    COUNT(f.fact_id) AS txn_count,
    SUM(f.amount) AS total_amount,
    CASE WHEN COUNT(f.fact_id) > 0 THEN 'PASS' ELSE 'MISSING' END AS status
FROM dim_date dt
LEFT JOIN fact_gl f ON f.date_key = dt.date_key
WHERE dt.month BETWEEN 1 AND 12
GROUP BY dt.month, dt.month_abbrev
ORDER BY dt.month;


-- Product coverage per month
CREATE VIEW IF NOT EXISTS v_check_product_month_coverage AS
SELECT
    dt.month,
    dt.month_abbrev,
    dp.product_name,
    COUNT(f.fact_id) AS txn_count,
    CASE WHEN COUNT(f.fact_id) > 0 THEN 'PASS' ELSE 'GAP' END AS status
FROM dim_date dt
CROSS JOIN dim_product dp
LEFT JOIN fact_gl f
    ON f.date_key = dt.date_key
    AND f.product_id = dp.product_id
WHERE dt.month BETWEEN 1 AND 12
GROUP BY dt.month, dp.product_name
ORDER BY dt.month, dp.product_name;


-- Department coverage per month
CREATE VIEW IF NOT EXISTS v_check_dept_month_coverage AS
SELECT
    dt.month,
    dt.month_abbrev,
    dd.dept_name,
    COUNT(f.fact_id) AS txn_count,
    CASE WHEN COUNT(f.fact_id) > 0 THEN 'PASS' ELSE 'GAP' END AS status
FROM dim_date dt
CROSS JOIN dim_department dd
LEFT JOIN fact_gl f
    ON f.date_key = dt.date_key
    AND f.dept_id = dd.dept_id
WHERE dt.month BETWEEN 1 AND 12
GROUP BY dt.month, dd.dept_name
ORDER BY dt.month, dd.dept_name;


-- ─────────────────────────────────────────────────────────────────────────────
-- 4. BALANCE VALIDATION
-- ─────────────────────────────────────────────────────────────────────────────

-- Allocation shares must sum to 1.0 per share type
CREATE VIEW IF NOT EXISTS v_check_share_balance AS
SELECT
    share_type,
    ROUND(SUM(share_pct), 4) AS total_share,
    CASE
        WHEN ABS(SUM(share_pct) - 1.0) < 0.001 THEN 'PASS'
        ELSE 'FAIL'
    END AS status,
    CASE
        WHEN ABS(SUM(share_pct) - 1.0) < 0.001 THEN NULL
        ELSE 'Shares sum to ' || ROUND(SUM(share_pct), 4) || ', expected 1.0000'
    END AS detail
FROM allocation_shares
GROUP BY share_type;


-- GL staging row count vs fact row count
CREATE VIEW IF NOT EXISTS v_check_etl_completeness AS
SELECT
    'ETL Completeness' AS check_name,
    (SELECT COUNT(*) FROM stg_gl_raw WHERE amount IS NOT NULL) AS staging_rows,
    (SELECT COUNT(*) FROM fact_gl) AS fact_rows,
    (SELECT COUNT(*) FROM stg_gl_raw WHERE amount IS NOT NULL) -
        (SELECT COUNT(*) FROM fact_gl) AS delta,
    CASE
        WHEN (SELECT COUNT(*) FROM stg_gl_raw WHERE amount IS NOT NULL) =
             (SELECT COUNT(*) FROM fact_gl)
        THEN 'PASS'
        ELSE 'WARN'
    END AS status;


-- Cross-check: fact_gl total vs stg_gl_raw total
CREATE VIEW IF NOT EXISTS v_check_amount_reconciliation AS
SELECT
    'Amount Reconciliation' AS check_name,
    (SELECT ROUND(SUM(amount), 2) FROM stg_gl_raw) AS staging_total,
    (SELECT ROUND(SUM(amount), 2) FROM fact_gl) AS fact_total,
    ROUND(
        (SELECT SUM(amount) FROM stg_gl_raw) -
        (SELECT SUM(amount) FROM fact_gl), 2
    ) AS difference,
    CASE
        WHEN ABS(
            (SELECT SUM(amount) FROM stg_gl_raw) -
            (SELECT SUM(amount) FROM fact_gl)
        ) < 0.01 THEN 'PASS'
        ELSE 'FAIL'
    END AS status;


-- ─────────────────────────────────────────────────────────────────────────────
-- 5. DATA QUALITY CHECKS
-- ─────────────────────────────────────────────────────────────────────────────

-- Blank required fields
CREATE VIEW IF NOT EXISTS v_check_blank_fields AS
SELECT
    'Blank GL ID' AS check_name,
    COUNT(*) AS count,
    CASE WHEN COUNT(*) = 0 THEN 'PASS' ELSE 'FAIL' END AS status
FROM stg_gl_raw WHERE gl_id IS NULL OR TRIM(gl_id) = ''
UNION ALL
SELECT
    'Blank Department',
    COUNT(*),
    CASE WHEN COUNT(*) = 0 THEN 'PASS' ELSE 'FAIL' END
FROM stg_gl_raw WHERE department IS NULL OR TRIM(department) = ''
UNION ALL
SELECT
    'Blank Product',
    COUNT(*),
    CASE WHEN COUNT(*) = 0 THEN 'PASS' ELSE 'FAIL' END
FROM stg_gl_raw WHERE product IS NULL OR TRIM(product) = ''
UNION ALL
SELECT
    'Blank Amount',
    COUNT(*),
    CASE WHEN COUNT(*) = 0 THEN 'PASS' ELSE 'FAIL' END
FROM stg_gl_raw WHERE amount IS NULL;


-- Outlier detection (Z-score > 3)
CREATE VIEW IF NOT EXISTS v_check_outliers AS
SELECT
    f.fact_id,
    f.gl_id,
    f.amount,
    stats.avg_amount,
    stats.std_amount,
    ROUND((f.abs_amount - stats.avg_amount) / stats.std_amount, 2) AS z_score
FROM fact_gl f
CROSS JOIN (
    SELECT
        AVG(abs_amount) AS avg_amount,
        -- SQLite doesn't have STDDEV, so compute manually
        SQRT(AVG(abs_amount * abs_amount) - AVG(abs_amount) * AVG(abs_amount)) AS std_amount
    FROM fact_gl
) stats
WHERE stats.std_amount > 0
  AND (f.abs_amount - stats.avg_amount) / stats.std_amount > 3
ORDER BY z_score DESC;


-- Zero-amount transactions
CREATE VIEW IF NOT EXISTS v_check_zero_amounts AS
SELECT
    'Zero Amount Transactions' AS check_name,
    COUNT(*) AS count,
    CASE WHEN COUNT(*) = 0 THEN 'PASS' ELSE 'WARN' END AS status
FROM fact_gl
WHERE amount = 0;


-- ─────────────────────────────────────────────────────────────────────────────
-- 6. CONSOLIDATED VALIDATION REPORT
-- ─────────────────────────────────────────────────────────────────────────────

-- Run this query to get a full validation summary
CREATE VIEW IF NOT EXISTS v_validation_summary AS
-- Share balances
SELECT check_name, status, detail FROM (
    SELECT share_type || ' Shares Balance' AS check_name, status, detail
    FROM v_check_share_balance
)
UNION ALL
-- ETL completeness
SELECT check_name, status,
    'Staging: ' || staging_rows || ', Fact: ' || fact_rows || ', Delta: ' || delta
FROM v_check_etl_completeness
UNION ALL
-- Amount reconciliation
SELECT check_name, status,
    'Staging total: ' || staging_total || ', Fact total: ' || fact_total || ', Diff: ' || difference
FROM v_check_amount_reconciliation
UNION ALL
-- Blank fields
SELECT check_name, status, 'Count: ' || count
FROM v_check_blank_fields
UNION ALL
-- Zero amounts
SELECT check_name, status, 'Count: ' || count
FROM v_check_zero_amounts
ORDER BY
    CASE status WHEN 'FAIL' THEN 1 WHEN 'WARN' THEN 2 ELSE 3 END,
    check_name;


-- =============================================================================
-- POWER QUERY M EQUIVALENT — Validation Checks
-- =============================================================================
--
-- let
--     GLSource = Excel.CurrentWorkbook(){[Name="CrossfireHiddenWorksheet"]}[Content],
--     Typed = Table.TransformColumnTypes(GLSource, {
--         {"ID", type text}, {"Department", type text}, {"Product", type text},
--         {"Expense Category", type text}, {"Amount", type number}
--     }),
--
--     // Check 1: Blank fields
--     BlankIDs = Table.RowCount(Table.SelectRows(Typed,
--         each [ID] = null or Text.Trim([ID]) = "")),
--     BlankDepts = Table.RowCount(Table.SelectRows(Typed,
--         each [Department] = null or Text.Trim([Department]) = "")),
--     BlankProducts = Table.RowCount(Table.SelectRows(Typed,
--         each [Product] = null or Text.Trim([Product]) = "")),
--
--     // Check 2: Unknown products
--     ValidProducts = {"iGO", "Affirm", "InsureSight", "DocFast"},
--     UnknownProducts = Table.SelectRows(Typed,
--         each not List.Contains(ValidProducts, [Product])),
--
--     // Check 3: Unknown departments
--     ValidDepts = {"NetOps","Security","Support","Partners",
--                   "Content","R&D","Product Management"},
--     UnknownDepts = Table.SelectRows(Typed,
--         each not List.Contains(ValidDepts, [Department])),
--
--     // Check 4: Allocation share balance
--     Shares = #table({"Product","Share"}, {
--         {"iGO",0.50},{"Affirm",0.25},{"InsureSight",0.15},{"DocFast",0.10}
--     }),
--     ShareSum = List.Sum(Shares[Share]),
--     ShareCheck = if ShareSum = 1.0 then "PASS" else "FAIL",
--
--     // Build results table
--     Results = #table({"Check","Count","Status"}, {
--         {"Blank IDs", BlankIDs, if BlankIDs = 0 then "PASS" else "FAIL"},
--         {"Blank Departments", BlankDepts, if BlankDepts = 0 then "PASS" else "FAIL"},
--         {"Blank Products", BlankProducts, if BlankProducts = 0 then "PASS" else "FAIL"},
--         {"Unknown Products", Table.RowCount(UnknownProducts),
--             if Table.RowCount(UnknownProducts) = 0 then "PASS" else "FAIL"},
--         {"Unknown Departments", Table.RowCount(UnknownDepts),
--             if Table.RowCount(UnknownDepts) = 0 then "PASS" else "FAIL"},
--         {"Share Balance", ShareSum, ShareCheck}
--     })
-- in
--     Results
--
-- =============================================================================
