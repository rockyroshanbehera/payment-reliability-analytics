-- ---------------------------------------------------------------------------
-- 02. Reliability trend
-- Answers: is the decline a one-off day, a recurring pattern, or a lasting shift?
-- Daily series carries a 7-day moving average so single-day noise does not get
-- mistaken for a trend; the weekly series flags partial weeks so the first and
-- last reporting weeks cannot create a fake dip in a chart.
-- ---------------------------------------------------------------------------

DROP TABLE IF EXISTS daily_trend;
CREATE TABLE daily_trend AS
WITH daily AS (
    SELECT
        transaction_date                                                   AS date,
        COUNT(*)                                                           AS transactions,
        SUM(CASE WHEN transaction_status = 'Failed' THEN 1 ELSE 0 END)     AS failures,
        AVG(CASE WHEN transaction_status = 'Success' THEN 1.0 ELSE 0 END)  AS success_rate,
        AVG(processing_time_seconds)                                       AS avg_processing_seconds
    FROM transaction_enriched
    GROUP BY transaction_date
)
SELECT
    date,
    transactions,
    failures,
    success_rate,
    avg_processing_seconds,
    AVG(success_rate) OVER (ORDER BY date ROWS BETWEEN 6 PRECEDING AND CURRENT ROW)
                                                             AS success_rate_7d_moving_avg,
    success_rate - LAG(success_rate) OVER (ORDER BY date)     AS success_rate_dod_change
FROM daily
ORDER BY date;


DROP TABLE IF EXISTS weekly_trend;
CREATE TABLE weekly_trend AS
WITH weekly AS (
    SELECT
        week_start_date                                                    AS week_start,
        COUNT(*)                                                           AS transactions,
        COUNT(DISTINCT transaction_date)                                   AS days_covered,
        SUM(amount_inr)                                                    AS payment_value_inr,
        SUM(CASE WHEN transaction_status = 'Failed' THEN 1 ELSE 0 END)     AS failures,
        AVG(CASE WHEN transaction_status = 'Success' THEN 1.0 ELSE 0 END)  AS success_rate
    FROM transaction_enriched
    GROUP BY week_start_date
)
SELECT
    week_start,
    transactions,
    days_covered,
    -- Only whole weeks are comparable, so charts filter on this flag.
    CASE WHEN days_covered = 7 THEN 1 ELSE 0 END              AS is_complete_week,
    payment_value_inr,
    failures,
    success_rate,
    success_rate - LAG(success_rate) OVER (ORDER BY week_start) AS success_rate_wow_change
FROM weekly
ORDER BY week_start;
