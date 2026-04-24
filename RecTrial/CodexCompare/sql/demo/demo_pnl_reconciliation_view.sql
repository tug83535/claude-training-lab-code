-- Demo P&L Reconciliation View
-- Purpose: create a normalized source for reconciliation checks.

CREATE OR ALTER VIEW demo.vw_pnl_reconciliation_source AS
SELECT
    gl.transaction_id,
    CAST(gl.posting_date AS DATE)                AS posting_date,
    gl.department_code                            AS department,
    gl.product_code                               AS product,
    gl.account_name                               AS account_name,
    gl.amount_usd                                 AS amount_usd,
    CASE WHEN gl.amount_usd IS NULL THEN 1 ELSE 0 END AS is_amount_null,
    CASE WHEN gl.posting_date IS NULL THEN 1 ELSE 0 END AS is_date_null
FROM finance.gl_transactions gl
WHERE gl.is_reversal = 0;
GO
