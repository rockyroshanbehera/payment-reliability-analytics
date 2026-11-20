-- ---------------------------------------------------------------------------
-- 04. Segment scorecards
-- Answers: is the problem concentrated in a bank, channel, device or network?
-- Every scorecard reports the rate AND the absolute failure count, plus a 95%
-- confidence half-width, so a small segment with a noisy rate is never mistaken
-- for a real operational problem.
-- ---------------------------------------------------------------------------

DROP TABLE IF EXISTS bank_performance;
CREATE TABLE bank_performance AS
WITH by_bank AS (
    SELECT
        issuing_bank                                                       AS segment,
        COUNT(*)                                                           AS transactions,
        SUM(CASE WHEN transaction_status = 'Failed' THEN 1 ELSE 0 END)     AS failures,
        AVG(CASE WHEN transaction_status = 'Success' THEN 1.0 ELSE 0 END)  AS success_rate,
        SUM(CASE WHEN failure_reason = 'Bank Timeout' THEN 1 ELSE 0 END)   AS timeout_failures
    FROM transactions
    GROUP BY issuing_bank
)
SELECT
    segment,
    transactions,
    failures,
    success_rate,
    timeout_failures,
    failures * 1.0 / SUM(failures) OVER ()                     AS failure_contribution,
    -- 95% normal-approximation half-width on the success rate. Bands that
    -- overlap mean the observed gap is not distinguishable from noise.
    1.96 * sqrt(success_rate * (1 - success_rate) / transactions) AS success_rate_ci95_halfwidth,
    success_rate - (SELECT AVG(CASE WHEN transaction_status = 'Success' THEN 1.0 ELSE 0 END) FROM transactions)
                                                               AS gap_to_portfolio_rate
FROM by_bank
ORDER BY failures DESC;


-- One long, reusable scorecard across every categorical dimension, so the
-- workbook and the report read the same shape of table for each cut.
DROP TABLE IF EXISTS segment_performance;
CREATE TABLE segment_performance AS
WITH unioned AS (
    SELECT 'Payment Method' AS dimension, payment_method AS segment, 1 AS sort_order,
           transaction_status, amount_inr FROM transaction_enriched
    UNION ALL SELECT 'Region', region, 1, transaction_status, amount_inr FROM transaction_enriched
    UNION ALL SELECT 'State', state, 1, transaction_status, amount_inr FROM transaction_enriched
    UNION ALL SELECT 'Merchant Category', merchant_category, 1, transaction_status, amount_inr FROM transaction_enriched
    UNION ALL SELECT 'Device', device_type, 1, transaction_status, amount_inr FROM transaction_enriched
    UNION ALL SELECT 'Network', network_type, 1, transaction_status, amount_inr FROM transaction_enriched
    UNION ALL SELECT 'Customer Tenure', customer_tenure_group, 1, transaction_status, amount_inr FROM transaction_enriched
    UNION ALL SELECT 'Amount Band', amount_band, amount_band_order, transaction_status, amount_inr FROM transaction_enriched
), aggregated AS (
    SELECT
        dimension,
        segment,
        MIN(sort_order)                                                    AS sort_order,
        COUNT(*)                                                           AS transactions,
        SUM(CASE WHEN transaction_status = 'Failed' THEN 1 ELSE 0 END)     AS failures,
        AVG(CASE WHEN transaction_status = 'Success' THEN 1.0 ELSE 0 END)  AS success_rate,
        SUM(amount_inr)                                                    AS payment_value_inr
    FROM unioned
    GROUP BY dimension, segment
)
SELECT
    dimension,
    segment,
    transactions,
    failures,
    success_rate,
    payment_value_inr,
    -- Share of the dimension's failures carried by this segment.
    failures * 1.0 / NULLIF(SUM(failures) OVER (PARTITION BY dimension), 0) AS failure_contribution,
    success_rate - (SELECT AVG(CASE WHEN transaction_status = 'Success' THEN 1.0 ELSE 0 END) FROM transactions)
                                                                            AS gap_to_portfolio_rate,
    RANK() OVER (PARTITION BY dimension ORDER BY success_rate ASC)          AS worst_rank_in_dimension,
    -- Segments below 1,000 attempts are flagged rather than dropped: they stay
    -- visible, but no recommendation is built on them.
    CASE WHEN transactions < 1000 THEN 1 ELSE 0 END                         AS low_sample_flag
FROM aggregated
ORDER BY dimension, sort_order, success_rate;


-- Bank x hour combinations carrying the most absolute failures. Prioritising by
-- volume rather than rate points engineering at the largest recoverable pool.
DROP TABLE IF EXISTS bank_hour_impact;
CREATE TABLE bank_hour_impact AS
WITH combos AS (
    SELECT
        issuing_bank,
        transaction_hour                                                   AS hour,
        peak_period,
        COUNT(*)                                                           AS transactions,
        SUM(CASE WHEN transaction_status = 'Failed' THEN 1 ELSE 0 END)     AS failures,
        AVG(CASE WHEN transaction_status = 'Success' THEN 1.0 ELSE 0 END)  AS success_rate
    FROM transaction_enriched
    GROUP BY issuing_bank, transaction_hour, peak_period
    HAVING COUNT(*) >= 300   -- ignore thin cells that cannot support a decision
)
SELECT
    issuing_bank,
    hour,
    peak_period,
    transactions,
    failures,
    success_rate,
    RANK() OVER (ORDER BY failures DESC, success_rate ASC) AS impact_rank
FROM combos
ORDER BY impact_rank
LIMIT 20;
