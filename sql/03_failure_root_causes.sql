-- ---------------------------------------------------------------------------
-- 03. Failure root causes
-- Answers: which failure causes are worth engineering time, and when do they hit?
-- A Pareto view concentrates attention on the few causes that carry most of the
-- volume; the peak/hourly views test whether failures are load-driven.
-- ---------------------------------------------------------------------------

DROP TABLE IF EXISTS failure_pareto;
CREATE TABLE failure_pareto AS
WITH reasons AS (
    SELECT
        failure_reason,
        COUNT(*)         AS failures,
        SUM(amount_inr)  AS failed_value_inr,
        AVG(amount_inr)  AS avg_failed_value_inr
    FROM transactions
    WHERE transaction_status = 'Failed'
    GROUP BY failure_reason
), ranked AS (
    SELECT
        failure_reason,
        failures,
        failed_value_inr,
        avg_failed_value_inr,
        RANK() OVER (ORDER BY failures DESC)                    AS failure_rank,
        failures * 1.0 / SUM(failures) OVER ()                  AS failure_contribution
    FROM reasons
)
SELECT
    failure_reason,
    failures,
    failed_value_inr,
    avg_failed_value_inr,
    failure_rank,
    failure_contribution,
    -- Running share: how much of all failures the top N causes explain.
    SUM(failure_contribution) OVER (ORDER BY failures DESC ROWS UNBOUNDED PRECEDING)
                                                                AS cumulative_contribution
FROM ranked
ORDER BY failures DESC;


-- Evening load window versus the rest of the day.
DROP TABLE IF EXISTS peak_comparison;
CREATE TABLE peak_comparison AS
SELECT
    peak_period,
    COUNT(*)                                                           AS transactions,
    SUM(CASE WHEN transaction_status = 'Failed' THEN 1 ELSE 0 END)     AS failures,
    AVG(CASE WHEN transaction_status = 'Success' THEN 1.0 ELSE 0 END)  AS success_rate,
    AVG(processing_time_seconds)                                       AS avg_processing_seconds,
    SUM(CASE WHEN failure_reason = 'Bank Timeout' THEN 1 ELSE 0 END) * 1.0
        / NULLIF(SUM(CASE WHEN transaction_status = 'Failed' THEN 1 ELSE 0 END), 0)
                                                                       AS timeout_share_of_failures
FROM transaction_enriched
GROUP BY peak_period
ORDER BY peak_period DESC;


-- Hour-of-day profile: volume and reliability plotted on the same axis of time.
DROP TABLE IF EXISTS hourly_performance;
CREATE TABLE hourly_performance AS
WITH hourly AS (
    SELECT
        transaction_hour                                                   AS hour,
        peak_period,
        COUNT(*)                                                           AS transactions,
        SUM(CASE WHEN transaction_status = 'Failed' THEN 1 ELSE 0 END)     AS failures,
        AVG(CASE WHEN transaction_status = 'Success' THEN 1.0 ELSE 0 END)  AS success_rate,
        SUM(CASE WHEN failure_reason = 'Bank Timeout' THEN 1 ELSE 0 END)   AS timeout_failures
    FROM transaction_enriched
    GROUP BY transaction_hour, peak_period
)
SELECT
    hour,
    peak_period,
    transactions,
    failures,
    success_rate,
    timeout_failures,
    RANK() OVER (ORDER BY success_rate ASC) AS worst_hour_rank
FROM hourly
ORDER BY hour;
