"""Tests over the delivered Excel workbook, read straight from its XML."""
from __future__ import annotations

import re

import pytest

import config
from conftest import Workbook

EXPECTED_SHEETS = [
    "Executive Dashboard", "Trend Analysis", "Failure Analysis", "Bank and Segment Analysis",
    "Impact Simulator", "KPI Definitions", "Data Dictionary", "Analysis Data", "Read Me",
]
ERROR_LITERALS = ["#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NULL!", "#NUM!"]
# Cell references of the six KPI cards, in the order they appear on the dashboard.
KPI_CARDS = {
    "A5": "transactions",
    "D5": "payment_value_inr",
    "G5": "success_rate",
    "J5": "failure_rate",
    "M5": "p95_processing_seconds",
    "P5": "retry_recovery_rate",
}


def find_row(workbook: Workbook, sheet: str, column: str, text: str) -> int:
    """Return the 1-based row where `column` holds `text`."""
    for ref, cell in workbook.cells(sheet).items():
        if ref.rstrip("0123456789") == column and cell["value"] == text:
            return int(re.sub(r"[A-Z]+", "", ref))
    raise AssertionError(f"{text!r} not found in column {column} of {sheet}")


def test_workbook_file_exists():
    assert config.WORKBOOK_PATH.exists()
    assert config.WORKBOOK_PATH.stat().st_size > 50_000


def test_sheets_are_present_in_the_required_order(workbook: Workbook):
    assert workbook.sheet_names == EXPECTED_SHEETS


def test_no_sheet_is_empty(workbook: Workbook):
    for name in EXPECTED_SHEETS:
        assert len(workbook.cells(name)) > 10, f"{name} looks empty"


def test_no_cell_holds_a_formula_error(workbook: Workbook):
    for name in EXPECTED_SHEETS:
        errors = {ref: cell["value"] for ref, cell in workbook.cells(name).items() if cell["type"] == "e"}
        assert not errors, f"{name} contains error cells: {errors}"


def test_no_error_literal_appears_in_any_sheet_xml(workbook: Workbook):
    for name in EXPECTED_SHEETS:
        xml = workbook.sheet_xml(name)
        found = [literal for literal in ERROR_LITERALS if literal in xml]
        assert not found, f"{name} contains {found}"


def test_dashboard_kpis_are_cross_sheet_formulas(workbook: Workbook):
    cells = workbook.cells("Executive Dashboard")
    for ref in KPI_CARDS:
        formula = cells[ref]["formula"]
        assert formula, f"{ref} should be a formula, not a typed value"
        assert "Analysis Data" in formula, f"{ref} should read from the Analysis Data sheet"


def test_dashboard_kpis_carry_the_correct_cached_values(workbook: Workbook, summary: dict):
    for ref, metric in KPI_CARDS.items():
        assert workbook.number("Executive Dashboard", ref) == pytest.approx(summary[metric], rel=1e-9)


def test_dashboard_states_the_synthetic_disclosure(workbook: Workbook):
    xml = workbook.sheet_xml("Executive Dashboard") + "".join(workbook.shared_strings)
    assert "SYNTHETIC DATA" in xml.upper()


def test_simulator_inputs_are_editable_numbers_not_formulas(workbook: Workbook, summary: dict):
    cells = workbook.cells("Impact Simulator")
    for name, scenario in summary["scenarios"].items():
        row = find_row(workbook, "Impact Simulator", "A", name)
        for column, key in [("B", "timeout_reduction"), ("C", "network_reduction"), ("D", "retry_uplift")]:
            cell = cells[f"{column}{row}"]
            assert cell["formula"] is None, f"{column}{row} must stay an editable input"
            assert float(cell["value"]) == pytest.approx(scenario[key])


def test_simulator_outputs_are_formulas_over_the_baselines(workbook: Workbook, summary: dict):
    cells = workbook.cells("Impact Simulator")
    row = find_row(workbook, "Impact Simulator", "A", config.HEADLINE_SCENARIO)
    recovered, value, uplift, projected = (cells[f"{c}{row}"] for c in "EFGH")

    assert recovered["formula"].startswith("ROUND(")
    assert f"B{row}" in recovered["formula"] and f"C{row}" in recovered["formula"] and f"D{row}" in recovered["formula"]
    assert "IFERROR(" in uplift["formula"], "the uplift formula must divide safely"
    assert "IFERROR(" in projected["formula"], "the projected rate must divide safely"
    assert f"E{row}" in uplift["formula"] and f"E{row}" in projected["formula"]

    headline = summary["scenarios"][config.HEADLINE_SCENARIO]
    assert float(recovered["value"]) == pytest.approx(headline["recovered_transactions"])
    assert float(value["value"]) == pytest.approx(headline["recovered_value_inr"], rel=1e-9)
    assert float(uplift["value"]) == pytest.approx(headline["success_rate_uplift_pp"], rel=1e-9)
    assert float(projected["value"]) == pytest.approx(headline["projected_success_rate"], rel=1e-9)


def test_simulator_baselines_point_at_the_analysis_sheet(workbook: Workbook):
    cells = workbook.cells("Impact Simulator")
    row = find_row(workbook, "Impact Simulator", "A", "Bank Timeout failures")
    assert "Analysis Data" in cells[f"B{row}"]["formula"]


def test_simulator_inputs_are_range_validated(workbook: Workbook):
    """The yellow cells must reject anything outside 0%-100%.

    OOXML treats "between" as the default operator, so it is absent from the
    XML; the pair of bounds is what proves the range is enforced.
    """
    xml = workbook.sheet_xml("Impact Simulator")
    assert "<dataValidation" in xml, "the yellow input cells need validation"
    assert 'type="decimal"' in xml
    assert "<formula1>0</formula1>" in xml and "<formula2>1</formula2>" in xml

    sqref = re.search(r'<dataValidation[^>]*sqref="([^"]+)"', xml)
    assert sqref, "the validation must be anchored to the input cells"
    first_row = find_row(workbook, "Impact Simulator", "A", "Conservative")
    last_row = find_row(workbook, "Impact Simulator", "A", "Ambitious")
    assert sqref.group(1) == f"B{first_row}:D{last_row}"


def test_simulator_carries_an_explicit_disclaimer(workbook: Workbook):
    text = " ".join(workbook.shared_strings).lower()
    assert "scenario estimate" in text
    assert "not a forecast" in text or "not guaranteed" in text


def test_workbook_contains_every_expected_chart(workbook: Workbook):
    # 4 dashboard + 2 trend + 2 failure + 1 incident + 1 simulator, some of which
    # are combined charts stored as an extra chart part.
    assert workbook.chart_count() >= 10


def test_presentation_sheets_hide_gridlines_and_freeze_headers(workbook: Workbook):
    for name in EXPECTED_SHEETS:
        assert 'showGridLines="0"' in workbook.sheet_xml(name), f"{name} still shows gridlines"
    for name in ["Trend Analysis", "Failure Analysis", "Bank and Segment Analysis", "Analysis Data"]:
        assert "<pane" in workbook.sheet_xml(name), f"{name} should freeze its header rows"
