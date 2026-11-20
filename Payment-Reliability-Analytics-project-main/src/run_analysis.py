"""Execute the SQL analysis layer and export every result the project publishes.

Each file in sql/ creates named result tables; this script runs them in order,
writes one CSV per table to outputs/analysis/, and then assembles summary.json -
the single set of numbers that the Excel workbook, the README and the executive
summary all quote, so the three can never disagree.
"""
from __future__ import annotations

import json
import math
import sqlite3
from contextlib import closing

import pandas as pd

from config import (
    ANALYSIS_DIR, DB_PATH, DISCLAIMER, HEADLINE_SCENARIO, INCIDENT_BANK, INCIDENT_END,
    INCIDENT_START, PEAK_END_HOUR, PEAK_START_HOUR, SCENARIOS, SQL_DIR, SUMMARY_PATH,
    TARGET_FAILURE_REASONS, configure_logging,
)

logger = configure_logging()

EXPORT_TABLES = [
    "kpi_overview", "monthly_kpis", "daily_trend", "weekly_trend", "failure_pareto",
    "peak_comparison", "hourly_performance", "bank_performance", "segment_performance",
    "bank_hour_impact", "retry_analysis", "retry_by_attempt", "incident_analysis",
    "incident_bank_comparison", "impact_baseline", "simulator_baseline",
]


def safe_divide(numerator: float, denominator: float, default: float = 0.0) -> float:
    """Divide without ever raising on a zero denominator."""
    return float(numerator) / float(denominator) if denominator else default


def run_sql(conn: sqlite3.Connection) -> None:
    """Execute every numbered SQL file in order."""
    for sql_file in sorted(SQL_DIR.glob("*.sql")):
        logger.info("Running %s", sql_file.name)
        conn.executescript(sql_file.read_text(encoding="utf-8"))
    conn.commit()


def build_scenarios(pools: dict[str, dict[str, float]], transactions: int, successes: int) -> dict[str, dict[str, float]]:
    """Convert assumption sets into recovered volume, value and success-rate uplift.

    The three levers act on mutually exclusive pools, so no transaction is
    counted twice. Outputs are scenario estimates, not forecasts.
    """
    scenarios: dict[str, dict[str, float]] = {}
    for name, assumptions in SCENARIOS.items():
        timeout = pools["Bank Timeout failures"]["transactions"] * assumptions["timeout_reduction"]
        network = pools["Network Error failures"]["transactions"] * assumptions["network_reduction"]
        retry = pools["Retried, still failed (other causes)"]["transactions"] * assumptions["retry_uplift"]
        recovered = round(timeout + network + retry)
        value = (timeout * pools["Bank Timeout failures"]["avg_value_inr"]
                 + network * pools["Network Error failures"]["avg_value_inr"]
                 + retry * pools["Retried, still failed (other causes)"]["avg_value_inr"])
        scenarios[name] = {
            **assumptions,
            "recovered_transactions": int(recovered),
            "recovered_value_inr": float(value),
            "success_rate_uplift_pp": safe_divide(recovered, transactions) * 100,
            "projected_success_rate": safe_divide(successes + recovered, transactions),
        }
    return scenarios


def main() -> None:
    ANALYSIS_DIR.mkdir(parents=True, exist_ok=True)
    with closing(sqlite3.connect(DB_PATH)) as conn:
        # Guarantees sqrt() regardless of how the local SQLite was compiled.
        conn.create_function("sqrt", 1, lambda x: math.sqrt(x) if x is not None and x >= 0 else None)
        run_sql(conn)
        tables = {name: pd.read_sql_query(f"SELECT * FROM {name}", conn) for name in EXPORT_TABLES}
        latency = pd.read_sql_query("SELECT processing_time_seconds FROM transactions", conn)
        period = pd.read_sql_query(
            "SELECT MIN(date(transaction_timestamp)) AS start, MAX(date(transaction_timestamp)) AS end,"
            " COUNT(DISTINCT date(transaction_timestamp)) AS days FROM transactions", conn).iloc[0]

    for name, frame in tables.items():
        frame.to_csv(ANALYSIS_DIR / f"{name}.csv", index=False)
    logger.info("Exported %d analysis tables to %s", len(tables), ANALYSIS_DIR.name)

    kpi = tables["kpi_overview"].iloc[0]
    pareto = tables["failure_pareto"]
    peak = tables["peak_comparison"].set_index("peak_period")
    banks = tables["bank_performance"]
    months = tables["monthly_kpis"]
    incident = tables["incident_analysis"]
    incident_banks = tables["incident_bank_comparison"]
    retry = tables["retry_analysis"]
    segments = tables["segment_performance"]
    pools = tables["simulator_baseline"].set_index("pool").to_dict("index")

    transactions = int(kpi.transactions)
    successes = int(kpi.successful_transactions)
    networks = segments[segments.dimension.eq("Network")].sort_values("success_rate")
    worst_network = networks.iloc[0]
    crescent_peak = incident_banks[incident_banks.bank_group.eq(INCIDENT_BANK)].set_index("period")
    others_peak = incident_banks[~incident_banks.bank_group.eq(INCIDENT_BANK)].set_index("period")
    during, before, after = "2. During (10 - 20 Mar)", "1. Before (1 Jan - 9 Mar)", "3. After (21 Mar - 30 Jun)"

    summary = {
        "disclaimer": DISCLAIMER,
        "period_start": str(period["start"]),
        "period_end": str(period["end"]),
        "period_days": int(period["days"]),
        "transactions": transactions,
        "successful_transactions": successes,
        "failed_transactions": int(kpi.failed_transactions),
        "payment_value_inr": float(kpi.payment_value_inr),
        "successful_value_inr": float(kpi.successful_value_inr),
        "failed_value_inr": float(kpi.failed_value_inr),
        "success_rate": float(kpi.success_rate),
        "failure_rate": float(kpi.failure_rate),
        "value_success_rate": float(kpi.value_success_rate),
        "avg_processing_seconds": float(kpi.avg_processing_time_seconds),
        "p95_processing_seconds": float(latency.processing_time_seconds.quantile(0.95)),
        "retry_rate": float(kpi.retry_rate),
        "retry_recovery_rate": float(kpi.retry_recovery_rate),
        "retry_rate_of_failures": safe_divide(retry.retried_transactions.sum(), retry.failures.sum()),

        # Trend
        "first_month": str(months.iloc[0].month),
        "last_month": str(months.iloc[-1].month),
        "first_month_success_rate": float(months.iloc[0].success_rate),
        "last_month_success_rate": float(months.iloc[-1].success_rate),
        "last_month_change_pp": float(months.iloc[-1].success_rate_pop_change) * 100,
        "worst_month": str(months.sort_values("success_rate").iloc[0].month),
        "worst_month_success_rate": float(months.sort_values("success_rate").iloc[0].success_rate),
        "monthly_range_pp": float(months.success_rate.max() - months.success_rate.min()) * 100,

        # Failure concentration
        "top_failure_reason": str(pareto.iloc[0].failure_reason),
        "top_failure_count": int(pareto.iloc[0].failures),
        "top_failure_contribution": float(pareto.iloc[0].failure_contribution),
        "second_failure_reason": str(pareto.iloc[1].failure_reason),
        "second_failure_count": int(pareto.iloc[1].failures),
        "second_failure_contribution": float(pareto.iloc[1].failure_contribution),
        "top_two_contribution": float(pareto.iloc[1].cumulative_contribution),
        "top_two_failures": int(pareto.iloc[0].failures + pareto.iloc[1].failures),

        # Load and network
        "peak_definition": f"{PEAK_START_HOUR:02d}:00-{PEAK_END_HOUR:02d}:59",
        "peak_success_rate": float(peak.loc["Peak", "success_rate"]),
        "non_peak_success_rate": float(peak.loc["Non-peak", "success_rate"]),
        "peak_gap_pp": float(peak.loc["Non-peak", "success_rate"] - peak.loc["Peak", "success_rate"]) * 100,
        "peak_transactions": int(peak.loc["Peak", "transactions"]),
        "peak_share_of_volume": safe_divide(peak.loc["Peak", "transactions"], transactions),
        "peak_share_of_failures": safe_divide(peak.loc["Peak", "failures"], kpi.failed_transactions),
        "peak_timeout_share_of_failures": float(peak.loc["Peak", "timeout_share_of_failures"]),
        "worst_network": str(worst_network.segment),
        "worst_network_success_rate": float(worst_network.success_rate),
        "worst_network_transactions": int(worst_network.transactions),

        # Bank comparison: reported with its uncertainty on purpose.
        "lowest_bank": str(banks.sort_values("success_rate").iloc[0].segment),
        "lowest_bank_success_rate": float(banks.sort_values("success_rate").iloc[0].success_rate),
        "highest_bank": str(banks.sort_values("success_rate").iloc[-1].segment),
        "highest_bank_success_rate": float(banks.sort_values("success_rate").iloc[-1].success_rate),
        "bank_spread_pp": float(banks.success_rate.max() - banks.success_rate.min()) * 100,
        "bank_max_ci95_pp": float(banks.success_rate_ci95_halfwidth.max()) * 100,

        # Engineered incident
        "incident_bank": INCIDENT_BANK,
        "incident_start": INCIDENT_START,
        "incident_end": INCIDENT_END,
        "incident_before_success_rate": float(incident.iloc[0].success_rate),
        "incident_during_success_rate": float(incident.iloc[1].success_rate),
        "incident_after_success_rate": float(incident.iloc[2].success_rate),
        "incident_bank_peak_before": float(crescent_peak.loc[before, "success_rate"]),
        "incident_bank_peak_during": float(crescent_peak.loc[during, "success_rate"]),
        "incident_bank_peak_after": float(crescent_peak.loc[after, "success_rate"]),
        "other_banks_peak_during": float(others_peak.loc[during, "success_rate"]),
        "incident_bank_timeout_rate_before": float(crescent_peak.loc[before, "timeout_rate"]),
        "incident_bank_timeout_rate_during": float(crescent_peak.loc[during, "timeout_rate"]),
        "incident_gap_during_pp": float(others_peak.loc[during, "success_rate"] - crescent_peak.loc[during, "success_rate"]) * 100,

        # Simulator pools
        "target_failure_reasons": TARGET_FAILURE_REASONS,
        "timeout_failures": int(pools["Bank Timeout failures"]["transactions"]),
        "timeout_avg_value_inr": float(pools["Bank Timeout failures"]["avg_value_inr"]),
        "network_failures": int(pools["Network Error failures"]["transactions"]),
        "network_avg_value_inr": float(pools["Network Error failures"]["avg_value_inr"]),
        "retry_pool_transactions": int(pools["Retried, still failed (other causes)"]["transactions"]),
        "retry_pool_avg_value_inr": float(pools["Retried, still failed (other causes)"]["avg_value_inr"]),
    }
    summary["scenarios"] = build_scenarios(pools, transactions, successes)
    headline = summary["scenarios"][HEADLINE_SCENARIO]
    summary.update({
        "headline_scenario": HEADLINE_SCENARIO,
        "expected_recovered_transactions": headline["recovered_transactions"],
        "expected_recovered_value_inr": headline["recovered_value_inr"],
        "expected_success_rate_uplift_pp": headline["success_rate_uplift_pp"],
        "expected_projected_success_rate": headline["projected_success_rate"],
    })

    SUMMARY_PATH.write_text(json.dumps(summary, indent=2), encoding="utf-8")
    logger.info("Success rate %.4f%% | failure rate %.4f%% | p95 %.2fs",
                summary["success_rate"] * 100, summary["failure_rate"] * 100, summary["p95_processing_seconds"])
    logger.info("Top causes: %s (%.1f%%) and %s (%.1f%%); peak gap %.2f pp",
                summary["top_failure_reason"], summary["top_failure_contribution"] * 100,
                summary["second_failure_reason"], summary["second_failure_contribution"] * 100,
                summary["peak_gap_pp"])
    # Three decimals: the uplift can land exactly on a 2-decimal midpoint, and
    # printing it rounded would disagree with how the workbook displays it.
    logger.info("%s scenario: +%s transactions, INR %s, +%.3f pp",
                HEADLINE_SCENARIO, f"{headline['recovered_transactions']:,}",
                f"{headline['recovered_value_inr']:,.0f}", headline["success_rate_uplift_pp"])


if __name__ == "__main__":
    main()
