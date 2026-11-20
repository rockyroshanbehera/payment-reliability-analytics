"""Tests for the synthetic data generator and its data-quality contract."""
from __future__ import annotations

import pandas as pd
import pytest

import config
from generate_data import generate
from validate_data import REQUIRED_COLUMNS, run_checks, validate

SMALL = 4_000


@pytest.fixture(scope="module")
def small() -> pd.DataFrame:
    return generate(SMALL, config.SEED)


def test_generator_is_deterministic_for_a_given_seed():
    first = generate(SMALL, config.SEED)
    second = generate(SMALL, config.SEED)
    pd.testing.assert_frame_equal(first, second)


def test_generator_changes_with_the_seed():
    baseline = generate(SMALL, config.SEED)
    other = generate(SMALL, config.SEED + 1)
    assert not baseline["transaction_status"].equals(other["transaction_status"])


def test_row_and_column_counts(small: pd.DataFrame):
    assert len(small) == SMALL
    assert list(small.columns) == REQUIRED_COLUMNS


def test_transaction_ids_are_unique(small: pd.DataFrame):
    assert small["transaction_id"].is_unique
    assert small["transaction_id"].str.match(r"^TXN\d{9}$").all()


def test_amounts_and_latencies_are_positive(small: pd.DataFrame):
    assert (small["amount_inr"] > 0).all()
    assert (small["processing_time_seconds"] > 0).all()


def test_only_allowed_categorical_values(small: pd.DataFrame):
    assert set(small["transaction_status"]) <= set(config.STATUSES)
    assert set(small["issuing_bank"]) <= set(config.BANKS)
    assert set(small["payment_method"]) <= set(config.PAYMENT_METHODS)
    assert set(small["network_type"]) <= set(config.NETWORKS)
    assert set(small["final_retry_outcome"]) <= set(config.RETRY_OUTCOMES)
    failed_reasons = small.loc[small.transaction_status.eq("Failed"), "failure_reason"]
    assert set(failed_reasons) <= set(config.FAILURE_REASONS)


def test_status_and_failure_reason_are_consistent(small: pd.DataFrame):
    reason = small["failure_reason"].fillna("")
    assert reason[small.transaction_status.eq("Success")].eq("").all()
    assert reason[small.transaction_status.eq("Failed")].ne("").all()


def test_retry_fields_are_internally_consistent(small: pd.DataFrame):
    retried = small["retry_flag"].eq(1)
    assert retried.loc[small.transaction_status.eq("Success")].sum() == 0
    assert (retried == small["retry_attempt"].gt(0)).all()
    assert small["retry_attempt"].between(0, config.MAX_RETRY_ATTEMPTS).all()
    assert small.loc[~retried, "final_retry_outcome"].eq("Not Retried").all()
    assert small.loc[retried, "final_retry_outcome"].isin(["Recovered", "Still Failed"]).all()


def test_success_rate_is_within_documented_bounds(small: pd.DataFrame):
    rate = small["transaction_status"].eq("Success").mean()
    assert config.MIN_SUCCESS_RATE <= rate <= config.MAX_SUCCESS_RATE


def test_timestamps_stay_inside_the_reporting_period(small: pd.DataFrame):
    stamps = pd.to_datetime(small["transaction_timestamp"])
    assert stamps.min() >= pd.Timestamp(config.START_DATE)
    assert stamps.max() < pd.Timestamp(config.END_DATE) + pd.Timedelta(days=1)


def test_validation_passes_on_generated_data(small: pd.DataFrame):
    assert validate(small) == []
    assert all(passed for _name, passed, _detail in run_checks(small))


@pytest.mark.parametrize("corrupt,expected", [
    (lambda df: df.assign(transaction_id=["TXN000000001"] * len(df)), "Unique transaction ids"),
    (lambda df: df.assign(amount_inr=df.amount_inr * -1), "Positive amounts"),
    (lambda df: df.assign(transaction_status="Pending"), "Allowed statuses"),
    (lambda df: df.assign(failure_reason="Bank Timeout"), "Successes never carry a cause"),
    (lambda df: df.assign(retry_flag=1), "Retries only follow failures"),
    (lambda df: df.assign(final_retry_outcome="Maybe"), "Allowed retry outcomes"),
])
def test_validation_catches_broken_data(small: pd.DataFrame, corrupt, expected: str):
    failures = validate(corrupt(small.copy()))
    assert any(message.startswith(expected) for message in failures), failures


def test_engineered_peak_pattern_is_present(small: pd.DataFrame):
    """Peak hours must fail measurably more often - the analysis depends on it."""
    hour = pd.to_datetime(small["transaction_timestamp"]).dt.hour
    peak = hour.between(config.PEAK_START_HOUR, config.PEAK_END_HOUR)
    failed = small["transaction_status"].eq("Failed")
    assert failed[peak].mean() > failed[~peak].mean()


def test_engineered_incident_is_recoverable_from_the_full_dataset(transactions: pd.DataFrame):
    """The affected bank must fail more in peak hours inside its incident window."""
    stamps = pd.to_datetime(transactions["transaction_timestamp"])
    peak = stamps.dt.hour.between(config.PEAK_START_HOUR, config.PEAK_END_HOUR)
    bank = transactions["issuing_bank"].eq(config.INCIDENT_BANK)
    window = (stamps >= config.INCIDENT_START) & (stamps < config.INCIDENT_END)
    failed = transactions["transaction_status"].eq("Failed")

    during = failed[bank & peak & window].mean()
    outside = failed[bank & peak & ~window].mean()
    others_during = failed[~bank & peak & window].mean()

    assert during > outside * 1.5, "incident window should show a clear failure uplift"
    assert during > others_during * 1.5, "the uplift should be specific to the affected bank"
