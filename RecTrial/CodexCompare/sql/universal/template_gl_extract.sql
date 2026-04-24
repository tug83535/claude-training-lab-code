-- Universal GL Extract Template
-- Replace parameter placeholders for your environment.

DECLARE @StartDate DATE = '2026-01-01';
DECLARE @EndDate   DATE = '2026-03-31';
DECLARE @EntityCode NVARCHAR(50) = 'ENTITY-001';

SELECT
    gl.transaction_id      AS TransactionID,
    gl.posting_date        AS PostingDate,
    gl.department_code     AS Department,
    gl.product_code        AS Product,
    gl.account_name        AS AccountName,
    gl.vendor_name         AS Vendor,
    gl.amount_usd          AS AmountUSD,
    gl.currency_code       AS CurrencyCode,
    gl.reference_number    AS ReferenceNumber
FROM finance.gl_transactions gl
WHERE gl.entity_code = @EntityCode
  AND gl.posting_date BETWEEN @StartDate AND @EndDate
  AND gl.is_reversal = 0
ORDER BY gl.posting_date, gl.transaction_id;
