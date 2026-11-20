"""Automated data-quality gate for the generated transaction dataset.

The pipeline stops here if any check fails, so no downstream SQL, workbook or
documentation is ever built on data that violates the documented contract.
Checks are expressed as (name, passed, detail) so the run prints an auditable
report rather than a single pass/fail.
"""
from __future__ import annotations

import pandas as pd

from config import (
    END_DATE, FAILURE_REASONS, FULL_DATA_PATH, MAX_RETRY_ATTEMPTS, MAX_SUCCESS_RATE,
    MIN_SUCCESS_RATE, RETRY_OUTCOMES, START_DATE, STATUSES, configure_logging,
)

logger = configure_logging()

REQUIRED_COLUMNS = [
    "transaction_id", "transaction_timestamp", "customer_id", "merchant_id", "amount_inr",
    "payment_method", "issuing_bank", "merchant_category", "city", "state", "region",
    "device_type", "network_type", "customer_tenure_group", "transaction_status",
    "failure_reason", "processing_time_seconds", "retry_flag", "retry_attempt",
    "final_retry_outcome",
]


def run_checks(df: pd.DataFrame) -> list[tuple[str, bool, str]]:
    """Return one (name, passed, detail) row per data-quality rule."""
    missing_columns = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing_columns:
        return [("Schema", False, f"missing columns: {missing_columns}")]

    ts = pd.to_datetime(df["transaction_timestamp"], errors="coerce")
    is_success = df["transaction_status"].eq("Success")
    is_failed = df["transaction_status"].eq("Failed")
    reason = df["failure_reason"].fillna("")
    retried = df["retry_flag"].eq(1)
    success_rate = float(is_success.mean())
    # failure_reason is blank by design on successful rows, so it is excluded
    # from the completeness check rather than treated as missing data.
    completeness = df.drop(columns=["failure_reason"]).isna().mean().max()

    return [
        ("Schema", True, f"{len(df.columns)} expected columns present"),
        ("Row count", len(df) > 0, f"{len(df):,} rows"),
        ("Unique transaction ids", bool(df["transaction_id"].is_unique), f"{df['transaction_id'].nunique():,} distinct ids"),
        ("Positive amounts", bool((df["amount_inr"] > 0).all()), f"min INR {df['amount_inr'].min():,.2f}"),
        ("Non-negative processing time", bool((df["processing_time_seconds"] > 0).all()), f"min {df['processing_time_seconds'].min():.3f}s"),
        ("Valid timestamps", bool(ts.notna().all()), "all timestamps parse"),
        ("Timestamps inside reporting period",
         bool(ts.min() >= pd.Timestamp(START_DATE) and ts.max() < pd.Timestamp(END_DATE) + pd.Timedelta(days=1)),
         f"{ts.min():%Y-%m-%d} to {ts.max():%Y-%m-%d}"),
        ("Allowed statuses", bool(df["transaction_status"].isin(STATUSES).all()), f"{sorted(df['transaction_status'].unique())}"),
        ("Allowed failure reasons", bool(reason[is_failed].isin(FAILURE_REASONS).all()), f"{reason[is_failed].nunique()} distinct causes"),
        ("Failures always carry a cause", bool(reason[is_failed].ne("").all()), f"{int(is_failed.sum()):,} failed rows"),
        ("Successes never carry a cause", bool(reason[is_success].eq("").all()), f"{int(is_success.sum()):,} successful rows"),
        ("Retries only follow failures", bool(retried.loc[is_success].sum() == 0), f"{int(retried.sum()):,} retried rows, all failed"),
        ("Retry flag matches attempt count", bool((retried == df["retry_attempt"].gt(0)).all()), "flag and ordinal agree"),
        ("Retry attempts within bounds", bool(df["retry_attempt"].between(0, MAX_RETRY_ATTEMPTS).all()), f"max {int(df['retry_attempt'].max())}"),
        ("Allowed retry outcomes", bool(df["final_retry_outcome"].isin(RETRY_OUTCOMES).all()), f"{sorted(df['final_retry_outcome'].unique())}"),
        ("Non-retried rows marked Not Retried", bool(df.loc[~retried, "final_retry_outcome"].eq("Not Retried").all()), "no orphan outcomes"),
        ("Retried rows resolved", bool(df.loc[retried, "final_retry_outcome"].isin(["Recovered", "Still Failed"]).all()), "every retry resolved"),
        ("Completeness of required fields", bool(completeness <= 0.02), f"worst column {completeness:.2%} missing"),
        ("Success rate plausible", bool(MIN_SUCCESS_RATE <= success_rate <= MAX_SUCCESS_RATE),
         f"{success_rate:.2%} within {MIN_SUCCESS_RATE:.0%}-{MAX_SUCCESS_RATE:.0%}"),
    ]


def validate(df: pd.DataFrame) -> list[str]:
    """Return a list of failure messages; empty means the dataset is valid."""
    return [f"{name}: {detail}" for name, passed, detail in run_checks(df) if not passed]


def main() -> None:
    df = pd.read_csv(FULL_DATA_PATH)
    checks = run_checks(df)
    for name, passed, detail in checks:
        logger.info("%s %-36s %s", "PASS" if passed else "FAIL", name, detail)
    failures = [f"{name}: {detail}" for name, passed, detail in checks if not passed]
    if failures:
        raise SystemExit("VALIDATION FAILED\n- " + "\n- ".join(failures))
    logger.info("Validation passed: %d/%d checks on %s rows", len(checks), len(checks), f"{len(df):,}")


if __name__ == "__main__":
    main()
