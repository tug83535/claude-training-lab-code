-- =============================================================================
-- staging.sql
-- Keystone BenefitTech P&L — GL Staging and Normalization
-- =============================================================================
--
-- ENGINE:  SQLite 3.x (portable, zero-install)
-- USAGE:   sqlite3 keystone_pnl.db < staging.sql
-- PURPOSE: Import raw GL data, normalize dimensions, deduplicate
--
-- =============================================================================


-- ─────────────────────────────────────────────────────────────────────────────
-- 1. DIMENSION TABLES
-- ─────────────────────────────────────────────────────────────────────────────

CREATE TABLE IF NOT EXISTS dim_product (
    product_id   INTEGER PRIMARY KEY AUTOINCREMENT,
    product_name TEXT    NOT NULL UNIQUE,
    is_active    INTEGER NOT NULL DEFAULT 1,
    created_at   TEXT    NOT NULL DEFAULT (datetime('now'))
);

INSERT OR IGNORE INTO dim_product (product_name) VALUES
    ('iGO'), ('Affirm'), ('InsureSight'), ('DocFast');


CREATE TABLE IF NOT EXISTS dim_department (
    dept_id    INTEGER PRIMARY KEY AUTOINCREMENT,
    dept_name  TEXT    NOT NULL UNIQUE,
    is_active  INTEGER NOT NULL DEFAULT 1,
    created_at TEXT    NOT NULL DEFAULT (datetime('now'))
);

INSERT OR IGNORE INTO dim_department (dept_name) VALUES
    ('NetOps'), ('Security'), ('Support'), ('Partners'),
    ('Content'), ('R&D'), ('Product Management');


CREATE TABLE IF NOT EXISTS dim_expense_category (
    category_id   INTEGER PRIMARY KEY AUTOINCREMENT,
    category_name TEXT    NOT NULL UNIQUE,
    is_active     INTEGER NOT NULL DEFAULT 1
);

INSERT OR IGNORE INTO dim_expense_category (category_name) VALUES
    ('AWS'), ('Employee Expenses'), ('Software Licenses'),
    ('Professional Services'), ('Rent & Facilities'),
    ('Travel & Entertainment'), ('Marketing'),
    ('Hardware & Equipment'), ('Telecom'),
    ('Insurance'), ('Depreciation'), ('Other');


CREATE TABLE IF NOT EXISTS dim_date (
    date_key     TEXT PRIMARY KEY,   -- YYYY-MM-DD
    year         INTEGER NOT NULL,
    month        INTEGER NOT NULL,
    month_abbrev TEXT    NOT NULL,   -- Jan, Feb, ...
    quarter      INTEGER NOT NULL,
    fiscal_year  TEXT    NOT NULL    -- FY2025
);

-- Populate FY2025 dates (Jan-Dec 2025)
INSERT OR IGNORE INTO dim_date (date_key, year, month, month_abbrev, quarter, fiscal_year)
SELECT
    date(printf('%04d-%02d-%02d', 2025, m.n, d.n)) AS date_key,
    2025 AS year,
    m.n AS month,
    CASE m.n
        WHEN 1 THEN 'Jan' WHEN 2 THEN 'Feb' WHEN 3 THEN 'Mar'
        WHEN 4 THEN 'Apr' WHEN 5 THEN 'May' WHEN 6 THEN 'Jun'
        WHEN 7 THEN 'Jul' WHEN 8 THEN 'Aug' WHEN 9 THEN 'Sep'
        WHEN 10 THEN 'Oct' WHEN 11 THEN 'Nov' WHEN 12 THEN 'Dec'
    END AS month_abbrev,
    CASE WHEN m.n <= 3 THEN 1 WHEN m.n <= 6 THEN 2
         WHEN m.n <= 9 THEN 3 ELSE 4 END AS quarter,
    'FY2025' AS fiscal_year
FROM
    (SELECT 1 AS n UNION ALL SELECT 2 UNION ALL SELECT 3 UNION ALL
     SELECT 4 UNION ALL SELECT 5 UNION ALL SELECT 6 UNION ALL
     SELECT 7 UNION ALL SELECT 8 UNION ALL SELECT 9 UNION ALL
     SELECT 10 UNION ALL SELECT 11 UNION ALL SELECT 12) m,
    (SELECT 1 AS n UNION ALL SELECT 2 UNION ALL SELECT 3 UNION ALL
     SELECT 4 UNION ALL SELECT 5 UNION ALL SELECT 6 UNION ALL
     SELECT 7 UNION ALL SELECT 8 UNION ALL SELECT 9 UNION ALL
     SELECT 10 UNION ALL SELECT 11 UNION ALL SELECT 12 UNION ALL
     SELECT 13 UNION ALL SELECT 14 UNION ALL SELECT 15 UNION ALL
     SELECT 16 UNION ALL SELECT 17 UNION ALL SELECT 18 UNION ALL
     SELECT 19 UNION ALL SELECT 20 UNION ALL SELECT 21 UNION ALL
     SELECT 22 UNION ALL SELECT 23 UNION ALL SELECT 24 UNION ALL
     SELECT 25 UNION ALL SELECT 26 UNION ALL SELECT 27 UNION ALL
     SELECT 28) d
WHERE d.n <= CASE
    WHEN m.n IN (1,3,5,7,8,10,12) THEN 31
    WHEN m.n IN (4,6,9,11) THEN 30
    WHEN m.n = 2 THEN 28
END;


-- ─────────────────────────────────────────────────────────────────────────────
-- 2. GL STAGING TABLE
-- ─────────────────────────────────────────────────────────────────────────────

CREATE TABLE IF NOT EXISTS stg_gl_raw (
    row_num          INTEGER PRIMARY KEY AUTOINCREMENT,
    gl_id            TEXT,
    gl_date          TEXT,
    department       TEXT,
    product          TEXT,
    expense_category TEXT,
    vendor           TEXT,
    amount           REAL,
    load_timestamp   TEXT NOT NULL DEFAULT (datetime('now')),
    source_file      TEXT
);

-- Load from CSV (run from CLI):
-- .mode csv
-- .import gl_extract.csv stg_gl_raw


-- ─────────────────────────────────────────────────────────────────────────────
-- 3. NORMALIZED FACT TABLE
-- ─────────────────────────────────────────────────────────────────────────────

CREATE TABLE IF NOT EXISTS fact_gl (
    fact_id          INTEGER PRIMARY KEY AUTOINCREMENT,
    gl_id            TEXT,
    date_key         TEXT    REFERENCES dim_date(date_key),
    dept_id          INTEGER REFERENCES dim_department(dept_id),
    product_id       INTEGER REFERENCES dim_product(product_id),
    category_id      INTEGER REFERENCES dim_expense_category(category_id),
    vendor           TEXT,
    amount           REAL    NOT NULL,
    abs_amount       REAL    GENERATED ALWAYS AS (ABS(amount)) STORED,
    is_positive      INTEGER GENERATED ALWAYS AS (CASE WHEN amount > 0 THEN 1 ELSE 0 END) STORED,
    load_timestamp   TEXT    NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_gl_date ON fact_gl(date_key);
CREATE INDEX IF NOT EXISTS idx_gl_product ON fact_gl(product_id);
CREATE INDEX IF NOT EXISTS idx_gl_dept ON fact_gl(dept_id);


-- ─────────────────────────────────────────────────────────────────────────────
-- 4. STAGING → FACT ETL
-- ─────────────────────────────────────────────────────────────────────────────

-- Normalize and load from staging to fact table
INSERT INTO fact_gl (gl_id, date_key, dept_id, product_id, category_id, vendor, amount)
SELECT
    s.gl_id,
    CASE
        WHEN s.gl_date IS NOT NULL AND s.gl_date != ''
        THEN date(s.gl_date)
        ELSE NULL
    END AS date_key,
    d.dept_id,
    p.product_id,
    c.category_id,
    s.vendor,
    s.amount
FROM stg_gl_raw s
LEFT JOIN dim_department d ON d.dept_name = s.department
LEFT JOIN dim_product p ON p.product_name = s.product
LEFT JOIN dim_expense_category c ON c.category_name = s.expense_category
WHERE s.amount IS NOT NULL;


-- ─────────────────────────────────────────────────────────────────────────────
-- 5. DEDUPLICATION
-- ─────────────────────────────────────────────────────────────────────────────

-- Identify exact duplicates (same GL ID, date, dept, product, amount)
CREATE VIEW IF NOT EXISTS v_duplicate_candidates AS
SELECT
    gl_id, date_key, dept_id, product_id, amount,
    COUNT(*) AS occurrence_count,
    GROUP_CONCAT(fact_id) AS fact_ids
FROM fact_gl
GROUP BY gl_id, date_key, dept_id, product_id, amount
HAVING COUNT(*) > 1;

-- To remove duplicates (keep lowest fact_id):
-- DELETE FROM fact_gl
-- WHERE fact_id NOT IN (
--     SELECT MIN(fact_id) FROM fact_gl
--     GROUP BY gl_id, date_key, dept_id, product_id, amount
-- );


-- =============================================================================
-- POWER QUERY M EQUIVALENT — GL Staging
-- =============================================================================
--
-- let
--     Source = Excel.CurrentWorkbook(){[Name="CrossfireHiddenWorksheet"]}[Content],
--     PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
--     TypedColumns = Table.TransformColumnTypes(PromotedHeaders, {
--         {"ID", type text},
--         {"Date", type date},
--         {"Department", type text},
--         {"Product", type text},
--         {"Expense Category", type text},
--         {"Vendor", type text},
--         {"Amount", type number}
--     }),
--     AddMonth = Table.AddColumn(TypedColumns, "Month", each Date.Month([Date]), Int64.Type),
--     AddMonthAbbrev = Table.AddColumn(AddMonth, "MonthAbbrev",
--         each Date.ToText([Date], "MMM"), type text),
--     AddQuarter = Table.AddColumn(AddMonthAbbrev, "Quarter",
--         each Date.QuarterOfYear([Date]), Int64.Type),
--     AddYear = Table.AddColumn(AddQuarter, "Year", each Date.Year([Date]), Int64.Type),
--     AddAbsAmount = Table.AddColumn(AddYear, "AbsAmount",
--         each Number.Abs([Amount]), type number),
--     RemoveDuplicates = Table.Distinct(AddAbsAmount, {"ID", "Date", "Department",
--         "Product", "Amount"})
-- in
--     RemoveDuplicates
--
-- =============================================================================
