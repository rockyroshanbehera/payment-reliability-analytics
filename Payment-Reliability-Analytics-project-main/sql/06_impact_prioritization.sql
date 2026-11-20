-- ---------------------------------------------------------------------------
-- 06. Incident evidence and impact baselines
-- Answers: was the March dip a real, contained event, and how large is the pool
-- of failures a fix could realistically address?
-- The before/during/after cut is normalised per day because the three windows
-- differ in length; comparing raw counts across them would be misleading.
-- ---------------------------------------------------------------------------

DROP TABLE IF EXISTS incident_analysis;
CREATE TABLE incident_analysis AS
WITH periods AS (
    SELECT
        incident_period,
        COUNT(*)                                                           AS transactions,
        COUNT(DISTINCT transaction_date)                                   AS days,
        SUM(CASE WHEN transaction_status = 'Failed' THEN 1 ELSE 0 END)     AS failures,
        AVG(CASE WHEN transaction_status = 'Success' THEN 1.0 ELSE 0 END)  AS success_rate,
        SUM(CASE WHEN bank_group = 'Crescent Bank' AND failure_reason = 'Bank Timeout'
                 THEN 1 ELSE 0 END)                                        AS crescent_timeout_failures
    FROM transaction_enriched
    GROUP BY incident_period
)
SELECT
    CASE incident_period WHEN 'Before' THEN '1. Before (1 Jan - 9 Mar)'
                         WHEN 'During' THEN '2. During (10 - 20 Mar)'
                         ELSE '3. After (21 Mar - 30 Jun)' END AS period,
    transactions,
    days,
    failures,
    success_rate,
    crescent_timeout_failures,
    -- Per-day normalisation makes windows of different length comparable.
    crescent_timeout_failures * 1.0 / NULLIF(days, 0)          AS crescent_timeouts_per_day,
    failures * 1.0 / NULLIF(days, 0)                           AS failures_per_day
FROM periods
ORDER BY CASE incident_period WHEN 'Before' THEN 1 WHEN 'During' THEN 2 ELSE 3 END;


-- The decisive test: did the affected bank move differently from every other
-- bank in the same peak hours over the same three windows?
DROP TABLE IF EXISTS incident_bank_comparison;
CREATE TABLE incident_bank_comparison AS
SELECT
    bank_group,
    CASE incident_period WHEN 'Before' THEN '1. Before (1 Jan - 9 Mar)'
                         WHEN 'During' THEN '2. During (10 - 20 Mar)'
                         ELSE '3. After (21 Mar - 30 Jun)' END             AS period,
    COUNT(*)                                                               AS transactions,
    SUM(CASE WHEN transaction_status = 'Failed' THEN 1 ELSE 0 END)         AS failures,
    AVG(CASE WHEN transaction_status = 'Success' THEN 1.0 ELSE 0 END)      AS success_rate,
    SUM(CASE WHEN failure_reason = 'Bank Timeout' THEN 1 ELSE 0 END)       AS timeout_failures,
    SUM(CASE WHEN failure_reason = 'Bank Timeout' THEN 1.0 ELSE 0 END)
        / NULLIF(COUNT(*), 0)                                              AS timeout_rate
FROM transaction_enriched
WHERE peak_period = 'Peak'   -- the incident was engineered into evening peak only
GROUP BY bank_group, incident_period
ORDER BY bank_group DESC, period;


-- Addressable failure pools for the scenario simulator.
DROP TABLE IF EXISTS impact_baseline;
CREATE TABLE impact_baseline AS
SELECT
    failure_reason,
    COUNT(*)                                   AS failures,
    SUM(amount_inr)                            AS failed_value_inr,
    AVG(amount_inr)                            AS avg_failed_value_inr,
    SUM(CASE WHEN peak_period = 'Peak' THEN 1 ELSE 0 END) * 1.0 / NULLIF(COUNT(*), 0) AS peak_share
FROM transaction_enriched
WHERE failure_reason IN ('Bank Timeout', 'Network Error')
GROUP BY failure_reason
ORDER BY failures DESC;


-- Every number the Excel Impact Simulator needs, in one auditable table.
-- Pools are mutually exclusive so the scenario levers cannot double-count a
-- transaction: the retry lever deliberately excludes the two targeted causes.
DROP TABLE IF EXISTS simulator_baseline;
CREATE TABLE simulator_baseline AS
SELECT 1 AS sort_order, 'Total transactions' AS pool,
       COUNT(*) AS transactions, SUM(amount_inr) AS value_inr,
       AVG(amount_inr) AS avg_value_inr
FROM transactions
UNION ALL
SELECT 2, 'Successful transactions',
       SUM(CASE WHEN transaction_status = 'Success' THEN 1 ELSE 0 END),
       SUM(CASE WHEN transaction_status = 'Success' THEN amount_inr ELSE 0 END),
       SUM(CASE WHEN transaction_status = 'Success' THEN amount_inr ELSE 0 END)
         / NULLIF(SUM(CASE WHEN transaction_status = 'Success' THEN 1 ELSE 0 END), 0)
FROM transactions
UNION ALL
SELECT 3, 'Bank Timeout failures', COUNT(*), SUM(amount_inr), AVG(amount_inr)
FROM transactions WHERE failure_reason = 'Bank Timeout'
UNION ALL
SELECT 4, 'Network Error failures', COUNT(*), SUM(amount_inr), AVG(amount_inr)
FROM transactions WHERE failure_reason = 'Network Error'
UNION ALL
SELECT 5, 'Retried, still failed (other causes)', COUNT(*), SUM(amount_inr), AVG(amount_inr)
FROM transactions
WHERE final_retry_outcome = 'Still Failed'
  AND failure_reason NOT IN ('Bank Timeout', 'Network Error');
