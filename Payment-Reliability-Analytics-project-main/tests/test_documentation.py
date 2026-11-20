"""Pin the figures quoted in the written deliverables to the analysis outputs.

Documentation drifting away from the numbers it describes is the most common
way an analysis project becomes wrong after the fact. These tests fail the
build if the README, the stakeholder report or the portfolio notes quote a
figure the pipeline no longer produces.
"""
from __future__ import annotations

from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import pytest

import config


def pct(value: float, places: int = 2) -> str:
    return f"{value:.{places}%}"


def pp(value: float) -> str:
    """Percentage points, rounded the way Excel displays them.

    The expected-case uplift lands exactly on a two-decimal midpoint (0.605),
    where Python's binary float rounds down and Excel rounds half away from
    zero. Going through str() recovers the intended decimal first, so the
    documentation and the workbook agree.
    """
    return f"{Decimal(str(value)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)} pp"


def read(path: Path) -> str:
    assert path.exists(), f"{path.name} is missing"
    return path.read_text(encoding="utf-8")


@pytest.fixture(scope="module")
def readme() -> str:
    return read(config.ROOT / "README.md")


@pytest.fixture(scope="module")
def report() -> str:
    return read(config.ROOT / "reports" / "executive_summary.md")


@pytest.fixture(scope="module")
def notes() -> str:
    return read(config.ROOT / "docs" / "portfolio_notes.md")


def headline_figures(summary: dict) -> dict[str, str]:
    """Every figure the written documents quote, formatted as they quote it."""
    headline = summary["scenarios"][config.HEADLINE_SCENARIO]
    return {
        "transactions": f"{summary['transactions']:,}",
        "failed transactions": f"{summary['failed_transactions']:,}",
        "success rate": pct(summary["success_rate"]),
        "failure rate": pct(summary["failure_rate"]),
        "value success rate": pct(summary["value_success_rate"]),
        "p95 latency": f"{summary['p95_processing_seconds']:.2f} s",
        "average latency": f"{summary['avg_processing_seconds']:.2f} s",
        "retry recovery rate": pct(summary["retry_recovery_rate"]),
        "top failure count": f"{summary['top_failure_count']:,}",
        "top failure share": pct(summary["top_failure_contribution"]),
        "second failure count": f"{summary['second_failure_count']:,}",
        "second failure share": pct(summary["second_failure_contribution"]),
        "top two share": pct(summary["top_two_contribution"]),
        "peak share of volume": pct(summary["peak_share_of_volume"]),
        "peak share of failures": pct(summary["peak_share_of_failures"]),
        "peak success rate": pct(summary["peak_success_rate"]),
        "non-peak success rate": pct(summary["non_peak_success_rate"]),
        "peak gap": pp(summary["peak_gap_pp"]),
        "worst network rate": pct(summary["worst_network_success_rate"]),
        "worst network volume": f"{summary['worst_network_transactions']:,}",
        "bank spread": pp(summary["bank_spread_pp"]),
        "incident bank during": pct(summary["incident_bank_peak_during"]),
        "other banks during": pct(summary["other_banks_peak_during"]),
        "expected recovered": f"{headline['recovered_transactions']:,}",
        "expected uplift": pp(headline["success_rate_uplift_pp"]),
        "expected projected rate": pct(headline["projected_success_rate"]),
    }


def test_readme_quotes_current_figures(readme: str, summary: dict):
    missing = {name: value for name, value in headline_figures(summary).items() if value not in readme}
    assert not missing, f"README.md no longer matches the analysis: {missing}"


def test_executive_summary_quotes_current_figures(report: str, summary: dict):
    """The report is written in prose for a non-technical reader.

    It rounds shares to one decimal and value to millions, so it is checked
    against the figures in the form it actually quotes them.
    """
    headline = summary["scenarios"][config.HEADLINE_SCENARIO]
    expected = {
        "transactions": f"{summary['transactions']:,}",
        "failed transactions": f"{summary['failed_transactions']:,}",
        "success rate": pct(summary["success_rate"]),
        "failure rate": pct(summary["failure_rate"]),
        "value success rate": pct(summary["value_success_rate"]),
        "p95 latency": f"{summary['p95_processing_seconds']:.2f}",
        "top failure count": f"{summary['top_failure_count']:,}",
        "top failure share": pct(summary["top_failure_contribution"], 1),
        "second failure count": f"{summary['second_failure_count']:,}",
        "top two share": pct(summary["top_two_contribution"], 1),
        "peak share of volume": pct(summary["peak_share_of_volume"], 1),
        "peak share of failures": pct(summary["peak_share_of_failures"], 1),
        "peak success rate": pct(summary["peak_success_rate"]),
        "non-peak success rate": pct(summary["non_peak_success_rate"]),
        "worst network rate": pct(summary["worst_network_success_rate"]),
        "incident bank during": pct(summary["incident_bank_peak_during"]),
        "other banks during": pct(summary["other_banks_peak_during"]),
        "incident timeout rate before": pct(summary["incident_bank_timeout_rate_before"]),
        "incident timeout rate during": pct(summary["incident_bank_timeout_rate_during"]),
        "expected recovered": f"{headline['recovered_transactions']:,}",
        "expected projected rate": pct(headline["projected_success_rate"]),
    }
    missing = {name: value for name, value in expected.items() if value not in report}
    assert not missing, f"executive_summary.md no longer matches the analysis: {missing}"


def test_scenario_table_matches_the_model(readme: str, summary: dict):
    for name, scenario in summary["scenarios"].items():
        assert f"{scenario['recovered_transactions']:,}" in readme, name
        assert f"₹{scenario['recovered_value_inr']:,.0f}" in readme, name
        assert pct(scenario["projected_success_rate"]) in readme, name


def test_every_deliverable_carries_the_synthetic_disclosure(readme: str, report: str, notes: str):
    for name, text in [("README.md", readme), ("executive_summary.md", report), ("portfolio_notes.md", notes)]:
        assert "synthetic" in text.lower(), name
        assert "fictional" in text.lower(), name


def test_no_deliverable_implies_real_company_data(readme: str, report: str, notes: str):
    """Any mention of a real company must be an explicit denial, never a claim."""
    for name, text in [("README.md", readme), ("executive_summary.md", report), ("portfolio_notes.md", notes)]:
        for line in text.splitlines():
            if "phonepe" in line.lower():
                assert any(token in line.lower() for token in ["not use", "no phonepe", "does not", "nothing here"]), \
                    f"{name} mentions a real company without denying its use: {line.strip()}"


def test_portfolio_notes_avoid_claiming_realised_impact(notes: str):
    for verb in ["identified", "estimated", "simulated", "demonstrated"]:
        assert verb in notes.lower(), f"portfolio notes should use '{verb}'"
    for forbidden in ["saved the company", "increased revenue by", "delivered a ", "resulting in savings"]:
        assert forbidden not in notes.lower(), f"portfolio notes claim realised impact: {forbidden!r}"


def test_readme_links_the_committed_preview_image(readme: str):
    assert "![Dashboard preview](outputs/dashboard_preview.png)" in readme
    assert config.PREVIEW_PATH.exists()


def test_readme_documents_the_one_command_reproduction(readme: str):
    assert "python src/run_pipeline.py" in readme
    assert (config.ROOT / "src" / "run_pipeline.py").exists()
