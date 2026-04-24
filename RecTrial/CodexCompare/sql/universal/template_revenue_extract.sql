-- Universal Revenue Extract Template
-- Replace parameter placeholders for your environment.

DECLARE @StartDate DATE = '2026-01-01';
DECLARE @EndDate   DATE = '2026-03-31';
DECLARE @Region NVARCHAR(50) = NULL;

SELECT
    rev.revenue_id             AS RevenueID,
    rev.transaction_date       AS TransactionDate,
    rev.region_name            AS Region,
    rev.sales_rep_name         AS SalesRep,
    rev.product_name           AS Product,
    rev.customer_name          AS Customer,
    rev.status_name            AS Status,
    rev.amount_usd             AS AmountUSD,
    rev.commission_pct         AS CommissionPct
FROM finance.revenue_transactions rev
WHERE rev.transaction_date BETWEEN @StartDate AND @EndDate
  AND (@Region IS NULL OR rev.region_name = @Region)
ORDER BY rev.transaction_date, rev.revenue_id;
