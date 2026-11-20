-- ---------------------------------------------------------------------------
-- 05. Retry effectiveness
-- Answers: is the retry policy recovering the failures that can be recovered?
-- Retry is cheap for infrastructure faults and pointless for customer-side
-- declines, so recovery rate is reported per root cause rather than in total.
-- ---------------------------------------------------------------------------

DROP TABLE IF EXISTS retry_analysis;
CREATE TABLE retry_analysis AS
WITH by_reason AS (
    SELECT
        failure_reason,
        COUNT(*)                                                                        AS failures,
        SUM(retry_flag)                                                                 AS retried_transactions,
        SUM(CASE WHEN final_retry_outcome = 'Recovered'   THEN 1 ELSE 0 END)            AS recovered_transactions,
        SUM(CASE WHEN final_retry_outcome = 'Still Failed' THEN 1 ELSE 0 END)           AS still_failed_transactions,
        SUM(CASE WHEN final_retry_outcome = 'Recovered' THEN amount_inr ELSE 0 END)     AS recovered_value_inr,
        SUM(CASE WHEN final_retry_outcome = 'Still Failed' THEN amount_inr ELSE 0 END)  AS still_failed_value_inr
    FROM transactions
    WHERE transaction_status = 'Failed'
    GROUP BY failure_reason
)
SELECT
    failure_reason,
    failures,
    retried_transactions,
    recovered_transactions,
    still_failed_transactions,
    still_failed_value_inr,
    retried_transactions * 1.0 / NULLIF(failures, 0)              AS retry_rate_of_failures,
    recovered_transactions * 1.0 / NULLIF(retried_transactions, 0) AS recovery_rate,
    RANK() OVER (ORDER BY recovered_transactions * 1.0 / NULLIF(retried_transactions, 0) DESC) AS recovery_rank
FROM by_reason
ORDER BY retried_transactions DESC;


-- Does a second or third attempt still pay for itself?
DROP TABLE IF EXISTS retry_by_attempt;
CREATE TABLE retry_by_attempt AS
SELECT
    retry_attempt,
    COUNT(*)                                                             AS retried_transactions,
    SUM(CASE WHEN final_retry_outcome = 'Recovered' THEN 1 ELSE 0 END)   AS recovered_transactions,
    SUM(CASE WHEN final_retry_outcome = 'Recovered' THEN 1 ELSE 0 END) * 1.0
        / NULLIF(COUNT(*), 0)                                            AS recovery_rate,
    AVG(amount_inr)                                                      AS avg_amount_inr
FROM transactions
WHERE retry_flag = 1
GROUP BY retry_attempt
ORDER BY retry_attempt;
