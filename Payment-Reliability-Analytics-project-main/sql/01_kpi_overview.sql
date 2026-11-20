-- ---------------------------------------------------------------------------
-- 01. Leadership scorecard
-- Answers: how reliable is the payment platform overall, and is it improving?
-- Volume, value, reliability and latency in one row, plus month-over-month
-- movement so a stakeholder can see direction rather than a single snapshot.
-- ---------------------------------------------------------------------------

DROP TABLE IF EXISTS kpi_overview;
CREATE TABLE kpi_overview AS
SELECT
    COUNT(*)                                                                      AS transactions,
    SUM(CASE WHEN transaction_status = 'Success' THEN 1 ELSE 0 END)                AS successful_transactions,
    SUM(CASE WHEN transaction_status = 'Failed'  THEN 1 ELSE 0 END)                AS failed_transactions,
    SUM(amount_inr)                                                               AS payment_value_inr,
    SUM(CASE WHEN transaction_status = 'Success' THEN amount_inr ELSE 0 END)       AS successful_value_inr,
    SUM(CASE WHEN transaction_status = 'Failed'  THEN amount_inr ELSE 0 END)       AS failed_value_inr,
    -- Attempt-weighted reliability: the headline KPI.
    AVG(CASE WHEN transaction_status = 'Success' THEN 1.0 ELSE 0 END)              AS success_rate,
    AVG(CASE WHEN transaction_status = 'Failed'  THEN 1.0 ELSE 0 END)              AS failure_rate,
    -- Value-weighted reliability: shows whether failures skew to larger payments.
    SUM(CASE WHEN transaction_status = 'Success' THEN amount_inr ELSE 0 END)
        / NULLIF(SUM(amount_inr), 0)                                              AS value_success_rate,
    AVG(processing_time_seconds)                                                  AS avg_processing_time_seconds,
    AVG(retry_flag)                                                               AS retry_rate,
    -- Safe division: NULLIF protects the KPI if a filtered slice has no retries.
    SUM(CASE WHEN final_retry_outcome = 'Recovered' THEN 1 ELSE 0 END) * 1.0
        / NULLIF(SUM(retry_flag), 0)                                              AS retry_recovery_rate
FROM transactions;


-- Month-over-month movement. LAG turns a static table into a trend statement.
DROP TABLE IF EXISTS monthly_kpis;
CREATE TABLE monthly_kpis AS
WITH monthly AS (
    SELECT
        transaction_month                                                  AS month,
        COUNT(*)                                                           AS transactions,
        SUM(amount_inr)                                                    AS payment_value_inr,
        SUM(CASE WHEN transaction_status = 'Failed' THEN 1 ELSE 0 END)     AS failures,
        AVG(CASE WHEN transaction_status = 'Success' THEN 1.0 ELSE 0 END)  AS success_rate,
        AVG(processing_time_seconds)                                       AS avg_processing_seconds
    FROM transaction_enriched
    GROUP BY transaction_month
)
SELECT
    month,
    transactions,
    payment_value_inr,
    failures,
    success_rate,
    avg_processing_seconds,
    success_rate - LAG(success_rate) OVER (ORDER BY month)   AS success_rate_pop_change,
    transactions - LAG(transactions) OVER (ORDER BY month)   AS transaction_pop_change
FROM monthly
ORDER BY month;
