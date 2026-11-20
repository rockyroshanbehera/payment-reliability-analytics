"""Reconciliation tests: SQL results, the Python summary and the raw data must agree."""
from __future__ import annotations

import sqlite3
from contextlib import closing

import pandas as pd
import pytest

import config
from conftest import ANALYSIS_TABLES, require

TOLERANCE = 1e-9


def test_every_analysis_table_is_exported(tables: dict[str, pd.DataFrame]):
    for name in ANALYSIS_TABLES:
        assert not tables[name].empty, f"{name}.csv is empty"


def test_summary_contains_every_documented_metric(summary: dict):
    expected = {
        "period_start", "period_end", "transactions", "successful_transactions", "failed_transactions",
        "payment_value_inr", "success_rate", "failure_rate", "value_success_rate",
        "avg_processing_seconds", "p95_processing_seconds", "retry_rate", "retry_recovery_rate",
        "top_failure_reason", "top_failure_contribution", "peak_success_rate", "non_peak_success_rate",
        "peak_gap_pp", "incident_bank", "incident_bank_peak_during", "other_banks_peak_during",
        "bank_spread_pp", "bank_max_ci95_pp", "scenarios", "expected_recovered_transactions",
        "expected_recovered_value_inr", "expected_success_rate_uplift_pp",
    }
    assert expected <= set(summary), sorted(expected - set(summary))
    assert "synthetic" in summary["disclaimer"].lower()


def test_kpi_table_matches_summary(tables: dict[str, pd.DataFrame], summary: dict):
    kpi = tables["kpi_overview"].iloc[0]
    assert int(kpi.transactions) == summary["transactions"]
    assert kpi.success_rate == pytest.approx(summary["success_rate"], abs=TOLERANCE)
    assert kpi.failure_rate == pytest.approx(summary["failure_rate"], abs=TOLERANCE)
    assert kpi.payment_value_inr == pytest.approx(summary["payment_value_inr"], rel=1e-12)
    assert kpi.retry_recovery_rate == pytest.approx(summary["retry_recovery_rate"], abs=TOLERANCE)


def test_sql_kpis_match_an_independent_pandas_calculation(transactions: pd.DataFrame, summary: dict):
    """Recompute the headline KPIs in pandas and compare to what SQL produced."""
    success = transactions["transaction_status"].eq("Success")
    assert len(transactions) == summary["transactions"]
    assert success.mean() == pytest.approx(summary["success_rate"], abs=TOLERANCE)
    assert (~success).sum() == summary["failed_transactions"]
    assert transactions["amount_inr"].sum() == pytest.approx(summary["payment_value_inr"], rel=1e-9)
    assert (transactions.loc[success, "amount_inr"].sum() / transactions["amount_inr"].sum()
            == pytest.approx(summary["value_success_rate"], abs=TOLERANCE))
    assert (transactions["processing_time_seconds"].quantile(0.95)
            == pytest.approx(summary["p95_processing_seconds"], abs=1e-6))
    retried = transactions["retry_flag"].eq(1)
    recovered = transactions["final_retry_outcome"].eq("Recovered")
    assert retried.mean() == pytest.approx(summary["retry_rate"], abs=TOLERANCE)
    assert recovered.sum() / retried.sum() == pytest.approx(summary["retry_recovery_rate"], abs=TOLERANCE)


def test_peak_definition_matches_config(transactions: pd.DataFrame, summary: dict):
    hour = pd.to_datetime(transactions["transaction_timestamp"]).dt.hour
    peak = hour.between(config.PEAK_START_HOUR, config.PEAK_END_HOUR)
    success = transactions["transaction_status"].eq("Success")
    assert success[peak].mean() == pytest.approx(summary["peak_success_rate"], abs=TOLERANCE)
    assert success[~peak].mean() == pytest.approx(summary["non_peak_success_rate"], abs=TOLERANCE)
    assert summary["peak_definition"] == f"{config.PEAK_START_HOUR:02d}:00-{config.PEAK_END_HOUR:02d}:59"


def test_failure_pareto_is_internally_consistent(tables: dict[str, pd.DataFrame], summary: dict):
    pareto = tables["failure_pareto"]
    assert pareto.failures.sum() == summary["failed_transactions"]
    assert pareto.failure_contribution.sum() == pytest.approx(1.0, abs=1e-9)
    assert pareto.cumulative_contribution.iloc[-1] == pytest.approx(1.0, abs=1e-9)
    assert pareto.failures.is_monotonic_decreasing
    assert pareto.iloc[0].failure_reason == summary["top_failure_reason"]
    assert pareto.iloc[1].cumulative_contribution == pytest.approx(summary["top_two_contribution"], abs=TOLERANCE)


def test_segment_contributions_sum_to_one_within_each_dimension(tables: dict[str, pd.DataFrame]):
    grouped = tables["segment_performance"].groupby("dimension").failure_contribution.sum()
    assert grouped.round(9).eq(1.0).all(), grouped[~grouped.round(9).eq(1.0)]


def test_every_dimension_reconciles_to_the_total_transaction_count(tables, summary):
    totals = tables["segment_performance"].groupby("dimension").transactions.sum()
    assert totals.eq(summary["transactions"]).all(), totals[~totals.eq(summary["transactions"])]


def test_bank_confidence_intervals_are_reported_and_overlapping(tables, summary):
    """The six-month bank ranking must be presented as inconclusive, not as a finding."""
    banks = tables["bank_performance"]
    assert (banks.success_rate_ci95_halfwidth > 0).all()
    assert summary["bank_spread_pp"] < 2 * summary["bank_max_ci95_pp"]


def test_incident_window_is_visible_in_the_bank_comparison(tables, summary):
    comparison = tables["incident_bank_comparison"]
    during = comparison[comparison.period.str.startswith("2.")].set_index("bank_group")
    assert during.loc[config.INCIDENT_BANK, "success_rate"] < during.loc["All other banks", "success_rate"]
    assert during.loc[config.INCIDENT_BANK, "timeout_rate"] > 2 * during.loc["All other banks", "timeout_rate"]
    assert summary["incident_bank_peak_during"] < summary["incident_bank_peak_before"]
    assert summary["incident_bank_peak_during"] < summary["incident_bank_peak_after"]


def test_complete_weeks_really_cover_seven_days(tables: dict[str, pd.DataFrame]):
    weekly = tables["weekly_trend"]
    assert weekly.loc[weekly.is_complete_week.eq(1), "days_covered"].eq(7).all()
    assert weekly.transactions.sum() == weekly.transactions.sum()  # no rows dropped by the window function
    assert weekly.is_complete_week.eq(1).sum() >= 20


def test_daily_moving_average_is_a_seven_day_window(tables: dict[str, pd.DataFrame]):
    daily = tables["daily_trend"]
    expected = daily.success_rate.rolling(7, min_periods=1).mean()
    pd.testing.assert_series_equal(daily.success_rate_7d_moving_avg, expected, check_names=False, atol=1e-12)


def test_retry_analysis_reconciles_with_the_raw_data(tables, transactions):
    retry = tables["retry_analysis"].set_index("failure_reason")
    failed = transactions[transactions.transaction_status.eq("Failed")]
    observed = failed.groupby("failure_reason").retry_flag.sum()
    for reason, retried in observed.items():
        assert retry.loc[reason, "retried_transactions"] == retried
    assert retry.recovered_transactions.sum() == transactions.final_retry_outcome.eq("Recovered").sum()


def test_simulator_pools_are_mutually_exclusive(tables, transactions, summary):
    pools = tables["simulator_baseline"].set_index("pool")
    timeout = transactions.failure_reason.eq("Bank Timeout")
    network = transactions.failure_reason.eq("Network Error")
    retry_pool = (transactions.final_retry_outcome.eq("Still Failed")
                  & ~transactions.failure_reason.isin(config.TARGET_FAILURE_REASONS))
    assert pools.loc["Bank Timeout failures", "transactions"] == timeout.sum()
    assert pools.loc["Network Error failures", "transactions"] == network.sum()
    assert pools.loc["Retried, still failed (other causes)", "transactions"] == retry_pool.sum()
    assert (timeout & retry_pool).sum() == 0
    assert (network & retry_pool).sum() == 0
    assert summary["retry_pool_transactions"] == retry_pool.sum()


def test_scenario_arithmetic_is_reproducible(summary: dict):
    for name, scenario in summary["scenarios"].items():
        recovered = round(summary["timeout_failures"] * scenario["timeout_reduction"]
                          + summary["network_failures"] * scenario["network_reduction"]
                          + summary["retry_pool_transactions"] * scenario["retry_uplift"])
        assert scenario["recovered_transactions"] == recovered, name
        assert scenario["success_rate_uplift_pp"] == pytest.approx(recovered / summary["transactions"] * 100)
        assert scenario["projected_success_rate"] == pytest.approx(
            (summary["successful_transactions"] + recovered) / summary["transactions"])
        assert 0 < scenario["success_rate_uplift_pp"] < 5, "an uplift above 5 pp would be implausible"


def test_scenarios_increase_monotonically(summary: dict):
    order = ["Conservative", "Expected", "Ambitious"]
    recovered = [summary["scenarios"][name]["recovered_transactions"] for name in order]
    assert recovered == sorted(recovered)


def test_database_row_count_reconciles(summary: dict):
    require(config.DB_PATH, "SQLite database")
    with closing(sqlite3.connect(config.DB_PATH)) as conn:
        rows = conn.execute("SELECT COUNT(*) FROM transactions").fetchone()[0]
        view_rows = conn.execute("SELECT COUNT(*) FROM transaction_enriched").fetchone()[0]
        peak_hours = conn.execute(
            "SELECT MIN(transaction_hour), MAX(transaction_hour) FROM transaction_enriched WHERE peak_period='Peak'"
        ).fetchone()
    assert rows == view_rows == summary["transactions"]
    assert peak_hours == (config.PEAK_START_HOUR, config.PEAK_END_HOUR)


def test_sql_files_demonstrate_the_documented_techniques():
    combined = "\n".join(path.read_text(encoding="utf-8").upper() for path in sorted(config.SQL_DIR.glob("*.sql")))
    for technique in ["WITH ", "CASE WHEN", "OVER (", "LAG(", "RANK()", "NULLIF(", "UNION ALL", "GROUP BY", "HAVING"]:
        assert technique in combined, f"SQL layer no longer demonstrates {technique.strip()}"
