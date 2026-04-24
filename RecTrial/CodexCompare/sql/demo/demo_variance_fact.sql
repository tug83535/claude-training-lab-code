-- Demo Variance Fact Query
-- Purpose: produce period-over-period variance rows for narrative generation.

WITH base AS (
    SELECT
        metric_name,
        period_name,
        value_amount,
        ROW_NUMBER() OVER (PARTITION BY metric_name ORDER BY period_sort) AS period_rank,
        COUNT(*) OVER (PARTITION BY metric_name) AS period_count
    FROM demo.pnl_metric_values
)
SELECT
    b1.metric_name,
    b1.period_name AS first_period,
    b2.period_name AS latest_period,
    b1.value_amount AS first_value,
    b2.value_amount AS latest_value,
    b2.value_amount - b1.value_amount AS delta_value,
    CASE
        WHEN b1.value_amount = 0 THEN NULL
        ELSE (b2.value_amount - b1.value_amount) / NULLIF(b1.value_amount, 0)
    END AS delta_pct
FROM base b1
JOIN base b2
  ON b1.metric_name = b2.metric_name
WHERE b1.period_rank = 1
  AND b2.period_rank = b2.period_count;
