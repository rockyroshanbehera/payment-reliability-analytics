"""Build the PayPulse Excel workbook from the exported analysis outputs.

Design rules followed here:
  * Every presentation number is a formula pointing at the Analysis Data sheet,
    so the workbook reconciles to the SQL/Python outputs by construction.
  * Formulas are written with their computed result cached, so the workbook
    shows correct values in any viewer, not only after a recalculation.
  * Charts read from named blocks on Analysis Data rather than from duplicated
    copies of the same series.
  * Nothing is fabricated: text panels are rendered from summary.json.

Implementation note: an earlier version of this builder targeted a proprietary
JavaScript spreadsheet runtime that is not installable from a public checkout.
It was replaced with XlsxWriter so the workbook can be rebuilt by anyone who
clones this repository. See docs/methodology.md.
"""
from __future__ import annotations

import datetime
import json
from decimal import ROUND_HALF_UP, Decimal

import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name, xl_range_abs, xl_rowcol_to_cell

from config import (
    ANALYSIS_DIR, COMPANY, DISCLAIMER, HEADLINE_SCENARIO, OUTPUT_DIR, SUMMARY_PATH,
    WORKBOOK_PATH, configure_logging,
)

logger = configure_logging()

DATA_SHEET = "Analysis Data"
# Fixed document-creation date, so rebuilding the workbook is reproducible.
WORKBOOK_CREATED = datetime.datetime(2025, 7, 1)
SHEET_ORDER = [
    "Executive Dashboard", "Trend Analysis", "Failure Analysis",
    "Bank and Segment Analysis", "Impact Simulator", "KPI Definitions",
    "Data Dictionary", DATA_SHEET, "Read Me",
]

C = {
    "navy": "#102A43", "blue": "#147D92", "teal": "#1FA7A0", "green": "#2E8B57",
    "red": "#C94C4C", "amber": "#B7791F", "ink": "#243B53", "muted": "#627D98",
    "pale": "#EAF4F4", "light": "#F4F7FA", "white": "#FFFFFF", "line": "#D9E2EC",
    "input": "#FFF2CC", "inputborder": "#E8A838",
}
FONT = "Calibri"

# The [$-409] locale prefix pins digit grouping to the international convention,
# so the workbook renders identically for a reviewer in any Excel locale.
MONEY = "[$₹-409]#,##0"
MONEY_M = '[$₹-409]#,##0.0,,"M"'
NUM = "[$-409]#,##0"
PCT2 = "0.00%"
PCT1 = "0.0%"
DEC2 = "0.00"
PP = '0.00" pp"'
SEC = '0.00" s"'


# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------
TABLES = [
    "kpi_overview", "monthly_kpis", "daily_trend", "weekly_trend", "failure_pareto",
    "peak_comparison", "hourly_performance", "bank_performance", "segment_performance",
    "bank_hour_impact", "retry_analysis", "retry_by_attempt", "incident_analysis",
    "incident_bank_comparison", "impact_baseline", "simulator_baseline",
]


def load_analysis() -> tuple[dict[str, pd.DataFrame], dict]:
    """Read every exported analysis table plus the summary metrics."""
    tables = {name: pd.read_csv(ANALYSIS_DIR / f"{name}.csv") for name in TABLES}
    summary = json.loads(SUMMARY_PATH.read_text(encoding="utf-8"))
    tables["weekly_complete"] = tables["weekly_trend"].query("is_complete_week == 1").reset_index(drop=True)
    return tables, summary


class Block:
    """A rectangular table written to the Analysis Data sheet."""

    def __init__(self, frame: pd.DataFrame, header_row: int, first_col: int) -> None:
        self.frame = frame
        self.header_row = header_row
        self.first_row = header_row + 1
        self.last_row = header_row + len(frame)
        self.first_col = first_col
        self.columns = {name: first_col + i for i, name in enumerate(frame.columns)}

    def col(self, name: str) -> int:
        return self.columns[name]

    def series(self, name: str, rows: int | None = None) -> list:
        """Chart series reference in list form (handles the space in the sheet name)."""
        last = self.last_row if rows is None else self.first_row + rows - 1
        return [DATA_SHEET, self.first_row, self.col(name), last, self.col(name)]

    def ref(self, name: str) -> str:
        """Absolute A1 range for the column, for formulas and error bars."""
        col = self.col(name)
        return f"='{DATA_SHEET}'!{xl_range_abs(self.first_row, col, self.last_row, col)}"


def sheet_cell(row: int, col: int) -> str:
    """Absolute cross-sheet reference to a cell on the Analysis Data sheet."""
    return f"'{DATA_SHEET}'!${xl_col_to_name(col)}${row + 1}"


def pp(value: float, places: int = 2) -> str:
    """Format percentage points exactly as Excel renders them.

    Python rounds a binary float at a two-decimal midpoint downwards while
    Excel rounds half away from zero, which would otherwise print +0.60 pp in a
    text panel next to a formula cell showing 0.61 pp. Going through str()
    recovers the intended decimal before rounding.
    """
    quantum = Decimal(1).scaleb(-places)
    return f"{Decimal(str(value)).quantize(quantum, rounding=ROUND_HALF_UP)}"


# ---------------------------------------------------------------------------
# Formats
# ---------------------------------------------------------------------------
def build_formats(wb: xlsxwriter.Workbook) -> dict:
    """Create every cell format once so styling stays consistent."""
    base = {"font_name": FONT}

    def fmt(**kwargs):
        return wb.add_format({**base, **kwargs})

    return {
        "title": fmt(bold=True, font_size=20, font_color=C["white"], bg_color=C["navy"], align="left", valign="vcenter", indent=1),
        "subtitle": fmt(font_size=10, italic=True, font_color=C["ink"], bg_color=C["pale"], align="left", valign="vcenter", indent=1, text_wrap=True),
        "section": fmt(bold=True, font_size=11, font_color=C["white"], bg_color=C["blue"], align="left", valign="vcenter", indent=1),
        "card_label": fmt(bold=True, font_size=9, font_color=C["muted"], bg_color=C["white"], align="left", valign="vcenter", indent=1,
                          top=1, left=1, right=1, top_color=C["line"], left_color=C["line"], right_color=C["line"]),
        "card_note": fmt(font_size=8, font_color=C["muted"], bg_color=C["white"], align="left", valign="top", indent=1,
                         bottom=1, left=1, right=1, bottom_color=C["line"], left_color=C["line"], right_color=C["line"]),
        "card_label_hl": fmt(bold=True, font_size=9, font_color=C["muted"], bg_color=C["pale"], align="left", valign="vcenter", indent=1,
                             top=1, left=1, right=1, top_color=C["line"], left_color=C["line"], right_color=C["line"]),
        "card_note_hl": fmt(font_size=8, font_color=C["muted"], bg_color=C["pale"], align="left", valign="top", indent=1,
                            bottom=1, left=1, right=1, bottom_color=C["line"], left_color=C["line"], right_color=C["line"]),
        "panel": fmt(font_size=10, font_color=C["ink"], bg_color=C["light"], align="left", valign="top", text_wrap=True, indent=1, border=1, border_color=C["line"]),
        "panel_amber": fmt(font_size=10, font_color=C["ink"], bg_color="#FDF6E3", align="left", valign="top", text_wrap=True, indent=1, border=1, border_color=C["inputborder"]),
        "note": fmt(font_size=9, font_color=C["muted"], italic=True, align="left", valign="vcenter", text_wrap=True),
        "header": fmt(bold=True, font_size=10, font_color=C["white"], bg_color=C["navy"], align="center", valign="vcenter", text_wrap=True, border=1, border_color=C["navy"]),
        "block_title": fmt(bold=True, font_size=11, font_color=C["navy"], bottom=2, bottom_color=C["teal"]),
        "text": fmt(font_size=10, font_color=C["ink"], border=1, border_color=C["line"], valign="vcenter"),
        "text_wrap": fmt(font_size=10, font_color=C["ink"], border=1, border_color=C["line"], valign="top", text_wrap=True),
        "text_bold": fmt(bold=True, font_size=10, font_color=C["navy"], border=1, border_color=C["line"], valign="vcenter"),
        "num": fmt(font_size=10, font_color=C["ink"], num_format=NUM, border=1, border_color=C["line"]),
        "dec2": fmt(font_size=10, font_color=C["ink"], num_format=DEC2, border=1, border_color=C["line"]),
        "money": fmt(font_size=10, font_color=C["ink"], num_format=MONEY, border=1, border_color=C["line"]),
        "pct2": fmt(font_size=10, font_color=C["ink"], num_format=PCT2, border=1, border_color=C["line"]),
        "pct1": fmt(font_size=10, font_color=C["ink"], num_format=PCT1, border=1, border_color=C["line"]),
        "pp": fmt(font_size=10, font_color=C["ink"], num_format=PP, border=1, border_color=C["line"]),
        "sec": fmt(font_size=10, font_color=C["ink"], num_format=SEC, border=1, border_color=C["line"]),
        "date": fmt(font_size=10, font_color=C["ink"], num_format="yyyy-mm-dd", border=1, border_color=C["line"]),
        "kpi_num": fmt(bold=True, font_size=18, font_color=C["navy"], bg_color=C["white"], num_format=NUM, align="left", valign="vcenter", indent=1, left=1, right=1, left_color=C["line"], right_color=C["line"]),
        "input": fmt(bold=True, font_size=11, font_color="#7A5C00", bg_color=C["input"], num_format="0%", align="center", valign="vcenter", border=1, border_color=C["inputborder"]),
        "out_num": fmt(bold=True, font_size=11, font_color=C["navy"], num_format=NUM, border=1, border_color=C["line"], bg_color=C["light"]),
        "out_money": fmt(bold=True, font_size=11, font_color=C["navy"], num_format=MONEY, border=1, border_color=C["line"], bg_color=C["light"]),
        "out_pp": fmt(bold=True, font_size=11, font_color=C["green"], num_format=PP, border=1, border_color=C["line"], bg_color=C["light"]),
        "out_pct": fmt(bold=True, font_size=11, font_color=C["navy"], num_format=PCT2, border=1, border_color=C["line"], bg_color=C["light"]),
        "label": fmt(font_size=10, font_color=C["ink"], bold=True, valign="vcenter"),
        "small": fmt(font_size=9, font_color=C["muted"], valign="vcenter"),
    }


def kpi_format(wb: xlsxwriter.Workbook, num_format: str, highlight: bool = False) -> object:
    return wb.add_format({
        "font_name": FONT, "bold": True, "font_size": 18, "font_color": C["navy"],
        "bg_color": C["pale"] if highlight else C["white"], "num_format": num_format,
        "align": "left", "valign": "vcenter", "indent": 1,
        "left": 1, "right": 1, "left_color": C["line"], "right_color": C["line"],
    })


# ---------------------------------------------------------------------------
# Shared writers
# ---------------------------------------------------------------------------
def write_banner(ws, F, title: str, subtitle: str, last_col: int) -> None:
    ws.set_row(0, 38)
    ws.set_row(1, 26)
    ws.merge_range(0, 0, 0, last_col, title, F["title"])
    ws.merge_range(1, 0, 1, last_col, subtitle, F["subtitle"])


def write_section(ws, F, row: int, first_col: int, last_col: int, text: str) -> None:
    ws.set_row(row, 20)
    ws.merge_range(row, first_col, row, last_col, text, F["section"])


def column_formats(frame: pd.DataFrame, F) -> list:
    """Pick a number format per column from its name and dtype.

    Order matters: anything expressed as a proportion is matched first, because
    columns such as timeout_share_of_failures or success_rate_pop_change would
    otherwise fall through to an integer format and display as zero.
    """
    chosen = []
    for name in frame.columns:
        lowered = name.lower()
        is_ratio = any(token in lowered for token in ("rate", "share", "contribution", "halfwidth"))
        if is_ratio and not lowered.endswith("_rank"):
            chosen.append(F["pct2"])
        elif lowered.endswith("_inr"):
            chosen.append(F["money"])
        elif "seconds" in lowered or "per_day" in lowered:
            chosen.append(F["dec2"])
        elif frame[name].dtype.kind in "if":
            chosen.append(F["num"])
        else:
            chosen.append(F["text"])
    return chosen


def write_table(ws, F, frame: pd.DataFrame, row: int, col: int, widths: dict[str, int] | None = None,
                header_labels: dict[str, str] | None = None) -> Block:
    """Write a dataframe as a formatted table and return its block reference."""
    labels = header_labels or {}
    formats = column_formats(frame, F)
    for j, name in enumerate(frame.columns):
        ws.write(row, col + j, labels.get(name, name.replace("_", " ").strip().capitalize()), F["header"])
    for i, record in enumerate(frame.itertuples(index=False), start=1):
        for j, value in enumerate(record):
            cell_format = formats[j]
            if pd.isna(value):
                ws.write_blank(row + i, col + j, None, cell_format)
            else:
                ws.write(row + i, col + j, value, cell_format)
    ws.set_row(row, 30)
    if widths:
        for name, width in widths.items():
            ws.set_column(col + list(frame.columns).index(name), col + list(frame.columns).index(name), width)
    return Block(frame, row, col)


# ---------------------------------------------------------------------------
# Analysis Data sheet
# ---------------------------------------------------------------------------
def scalar_format(key: str, value, F) -> object:
    """Choose a format for a summary metric from its unit suffix and its value.

    Names alone are not reliable (incident_bank_peak_during is a rate, and
    peak_share_of_volume ends in "volume"), so any unlabelled float between 0
    and 1 is treated as a proportion.
    """
    if isinstance(value, str):
        return F["text"]
    if isinstance(value, bool) or isinstance(value, int):
        return F["num"]
    if key.endswith("_pp"):
        return F["pp"]
    if key.endswith("_inr"):
        return F["money"]
    if "seconds" in key:
        return F["sec"]
    if 0 <= value <= 1:
        return F["pct2"]
    return F["dec2"]


def write_analysis_data(wb, ws, F, tables: dict[str, pd.DataFrame], summary: dict) -> tuple[dict[str, Block], dict[str, str]]:
    """Write the reconciliation layer and return block and metric-cell registries."""
    ws.hide_gridlines(2)
    ws.set_tab_color(C["muted"])
    write_banner(ws, F, "ANALYSIS DATA", "Every figure on the presentation sheets is a formula pointing here. "
                                         "Tables are the exact CSV outputs of the SQL analysis. " + DISCLAIMER, 7)
    ws.set_column(0, 0, 42)
    ws.set_column(1, 12, 16)
    ws.freeze_panes(3, 0)

    row = 3
    ws.write(row, 0, "Headline metrics (outputs/analysis/summary.json)", F["block_title"])
    row += 1
    ws.write(row, 0, "Metric", F["header"])
    ws.write(row, 1, "Value", F["header"])
    metric_cells: dict[str, str] = {}
    row += 1
    for key, value in summary.items():
        if isinstance(value, (dict, list)):
            continue
        ws.write(row, 0, key, F["text"])
        ws.write(row, 1, value, scalar_format(key, value, F))
        metric_cells[key] = sheet_cell(row, 1)
        row += 1

    # Scenario grid, so the simulator can point at auditable assumptions.
    row += 1
    ws.write(row, 0, "Scenario assumptions and modelled outcomes", F["block_title"])
    row += 1
    scenario_frame = pd.DataFrame([
        {"scenario": name, **values} for name, values in summary["scenarios"].items()
    ])
    scenario_block = write_table(ws, F, scenario_frame, row, 0)
    row = scenario_block.last_row + 3

    blocks: dict[str, Block] = {"scenarios": scenario_block}
    for name in [*TABLES, "weekly_complete"]:
        frame = tables[name]
        ws.write(row, 0, name, F["block_title"])
        row += 1
        blocks[name] = write_table(ws, F, frame, row, 0)
        row = blocks[name].last_row + 3

    ws.print_area(0, 0, row, 7)
    return blocks, metric_cells


# ---------------------------------------------------------------------------
# Executive Dashboard
# ---------------------------------------------------------------------------
def write_dashboard(wb, ws, F, blocks, cells, summary) -> None:
    ws.hide_gridlines(2)
    ws.set_tab_color(C["navy"])
    last_col = 17  # A..R
    ws.set_column(0, last_col, 9.5)
    write_banner(ws, F, "PAYMENT RELIABILITY ANALYTICS",
                 f"{COMPANY} (fictional) · Transaction success-rate and failure root-cause analysis · "
                 f"Reporting period {summary['period_start']} to {summary['period_end']} · "
                 f"SYNTHETIC DATA, NOT A REAL COMPANY", last_col)

    recovered_retries = round(summary["retry_recovery_rate"] * summary["retry_rate"] * summary["transactions"])
    retried = round(summary["retry_rate"] * summary["transactions"])
    cards = [
        ("TRANSACTIONS", cells["transactions"], summary["transactions"], NUM,
         f"{summary['period_days']} days of attempts", False),
        ("PAYMENT VALUE", cells["payment_value_inr"], summary["payment_value_inr"], MONEY_M,
         f"₹{summary['successful_value_inr'] / 1e6:.1f}M settled successfully", False),
        ("SUCCESS RATE", cells["success_rate"], summary["success_rate"], PCT2,
         f"Value-weighted {summary['value_success_rate']:.2%}", True),
        ("FAILURE RATE", cells["failure_rate"], summary["failure_rate"], PCT2,
         f"{summary['failed_transactions']:,} failed attempts", False),
        ("P95 PROCESSING TIME", cells["p95_processing_seconds"], summary["p95_processing_seconds"], SEC,
         f"Average {summary['avg_processing_seconds']:.2f} s", False),
        ("RETRY RECOVERY", cells["retry_recovery_rate"], summary["retry_recovery_rate"], PCT1,
         f"{recovered_retries:,} of {retried:,} retries", True),
    ]
    ws.set_row(3, 18)
    ws.set_row(4, 30)
    ws.set_row(5, 16)
    for i, (label, ref, value, number_format, note, highlight) in enumerate(cards):
        first, last = i * 3, i * 3 + 2
        ws.merge_range(3, first, 3, last, label, F["card_label_hl" if highlight else "card_label"])
        ws.merge_range(4, first, 4, last, "", kpi_format(wb, number_format, highlight))
        ws.write_formula(4, first, f"={ref}", kpi_format(wb, number_format, highlight), value)
        ws.merge_range(5, first, 5, last, note, F["card_note_hl" if highlight else "card_note"])

    # --- Charts -------------------------------------------------------------
    weekly = blocks["weekly_complete"]
    pareto = blocks["failure_pareto"]
    banks = blocks["bank_performance"]
    peak = blocks["peak_comparison"]

    write_section(ws, F, 7, 0, 8, "Weekly transaction success rate (complete weeks only)")
    write_section(ws, F, 7, 9, last_col, "Failure concentration (Pareto)")

    trend = wb.add_chart({"type": "line"})
    trend.add_series({
        "name": "Success rate",
        "categories": weekly.series("week_start"),
        "values": weekly.series("success_rate"),
        "line": {"color": C["blue"], "width": 2.25},
        "marker": {"type": "circle", "size": 5, "border": {"color": C["blue"]}, "fill": {"color": C["white"]}},
    })
    trend.set_y_axis({"min": 0.94, "max": 0.965, "num_format": PCT1, "major_gridlines": {"visible": True, "line": {"color": C["line"]}},
                      "line": {"none": True}, "num_font": {"name": FONT, "size": 9}})
    trend.set_x_axis({"num_font": {"name": FONT, "size": 8, "rotation": -45}, "line": {"color": C["line"]}})
    trend.set_title({"none": True})   # the teal section bar above is the title
    trend.set_legend({"none": True})
    trend.set_chartarea({"border": {"color": C["line"]}, "fill": {"color": C["white"]}})
    trend.set_size({"width": 640, "height": 290})
    ws.insert_chart(8, 0, trend, {"x_offset": 4, "y_offset": 4})

    pareto_chart = wb.add_chart({"type": "column"})
    pareto_chart.add_series({
        "name": "Failed transactions",
        "categories": pareto.series("failure_reason"),
        "values": pareto.series("failures"),
        "fill": {"color": C["blue"]},
        "gap": 40,
    })
    cumulative = wb.add_chart({"type": "line"})
    cumulative.add_series({
        "name": "Cumulative share",
        "categories": pareto.series("failure_reason"),
        "values": pareto.series("cumulative_contribution"),
        "line": {"color": C["amber"], "width": 2},
        "marker": {"type": "circle", "size": 5, "fill": {"color": C["amber"]}},
        "y2_axis": True,
    })
    pareto_chart.combine(cumulative)
    pareto_chart.set_y_axis({"num_format": NUM, "major_gridlines": {"visible": True, "line": {"color": C["line"]}},
                             "num_font": {"name": FONT, "size": 9}})
    # In a combined chart the secondary axis belongs to the secondary chart object.
    cumulative.set_y2_axis({"num_format": PCT1, "min": 0, "max": 1, "num_font": {"name": FONT, "size": 9}})
    pareto_chart.set_title({"none": True})
    pareto_chart.set_x_axis({"num_font": {"name": FONT, "size": 8, "rotation": -30}})
    pareto_chart.set_legend({"position": "bottom", "font": {"name": FONT, "size": 9}})
    pareto_chart.set_chartarea({"border": {"color": C["line"]}, "fill": {"color": C["white"]}})
    pareto_chart.set_size({"width": 640, "height": 290})
    ws.insert_chart(8, 9, pareto_chart, {"x_offset": 4, "y_offset": 4})

    write_section(ws, F, 24, 0, 8, "Issuing-bank success rate with 95% confidence interval")
    write_section(ws, F, 24, 9, last_col, "Peak versus non-peak reliability")

    bank_chart = wb.add_chart({"type": "column"})
    bank_chart.add_series({
        "name": "Success rate",
        "categories": banks.series("segment"),
        "values": banks.series("success_rate"),
        "fill": {"color": C["teal"]},
        "gap": 55,
        "y_error_bars": {
            "type": "custom",
            "plus_values": banks.ref("success_rate_ci95_halfwidth"),
            "minus_values": banks.ref("success_rate_ci95_halfwidth"),
            "line": {"color": C["ink"], "width": 1},
        },
    })
    bank_chart.set_y_axis({"min": 0.945, "max": 0.962, "num_format": PCT1,
                           "major_gridlines": {"visible": True, "line": {"color": C["line"]}},
                           "num_font": {"name": FONT, "size": 9}})
    bank_chart.set_x_axis({"num_font": {"name": FONT, "size": 9, "rotation": -20}})
    bank_chart.set_title({"none": True})
    bank_chart.set_legend({"none": True})
    bank_chart.set_chartarea({"border": {"color": C["line"]}, "fill": {"color": C["white"]}})
    bank_chart.set_size({"width": 640, "height": 270})
    ws.insert_chart(25, 0, bank_chart, {"x_offset": 4, "y_offset": 4})

    peak_chart = wb.add_chart({"type": "column"})
    peak_chart.add_series({
        "name": "Success rate",
        "categories": peak.series("peak_period"),
        "values": peak.series("success_rate"),
        "fill": {"color": C["blue"]},
        "gap": 120,
        "points": [{"fill": {"color": C["red"]}}, {"fill": {"color": C["green"]}}],
        "data_labels": {"value": True, "num_format": PCT2, "font": {"name": FONT, "size": 10, "bold": True}},
    })
    peak_chart.set_y_axis({"min": 0.94, "max": 0.965, "num_format": PCT1,
                           "major_gridlines": {"visible": True, "line": {"color": C["line"]}},
                           "num_font": {"name": FONT, "size": 9}})
    peak_chart.set_x_axis({"num_font": {"name": FONT, "size": 10}})
    peak_chart.set_title({"none": True})
    peak_chart.set_legend({"none": True})
    peak_chart.set_chartarea({"border": {"color": C["line"]}, "fill": {"color": C["white"]}})
    peak_chart.set_size({"width": 640, "height": 270})
    ws.insert_chart(25, 9, peak_chart, {"x_offset": 4, "y_offset": 4})

    # --- Narrative ----------------------------------------------------------
    findings = (
        f"1.  Two causes carry half of all failures. {summary['top_failure_reason']} "
        f"({summary['top_failure_count']:,} failures, {summary['top_failure_contribution']:.1%}) and "
        f"{summary['second_failure_reason']} ({summary['second_failure_count']:,}, "
        f"{summary['second_failure_contribution']:.1%}) together account for "
        f"{summary['top_two_contribution']:.1%} of the {summary['failed_transactions']:,} failed attempts.\n\n"
        f"2.  Reliability degrades under evening load. The {summary['peak_definition']} window carries "
        f"{summary['peak_share_of_volume']:.1%} of volume but {summary['peak_share_of_failures']:.1%} of failures: "
        f"{summary['peak_success_rate']:.2%} success versus {summary['non_peak_success_rate']:.2%} off-peak, a gap of "
        f"{pp(summary['peak_gap_pp'])} pp. Timeouts are {summary['peak_timeout_share_of_failures']:.1%} of peak failures.\n\n"
        f"3.  One bank had a contained 11-day outage. In peak hours between {summary['incident_start']} and 20 Mar 2025, "
        f"{summary['incident_bank']} fell to {summary['incident_bank_peak_during']:.2%} while every other bank held at "
        f"{summary['other_banks_peak_during']:.2%} — a {pp(summary['incident_gap_during_pp'], 1)} pp gap that closed afterwards. "
        f"Across the whole period bank success rates differ by only {pp(summary['bank_spread_pp'])} pp, within the "
        f"±{pp(summary['bank_max_ci95_pp'])} pp confidence band, so no bank should be judged on the six-month average alone."
    )
    actions = (
        "1.  Cap and circuit-break slow bank calls during the evening peak.\n"
        f"     Targets the {summary['top_failure_count']:,} {summary['top_failure_reason']} failures that concentrate in "
        f"{summary['peak_definition']}, the largest single recoverable pool.\n\n"
        "2.  Harden the low-bandwidth payment path.\n"
        f"     {summary['worst_network']} sessions succeed at {summary['worst_network_success_rate']:.2%} across "
        f"{summary['worst_network_transactions']:,} attempts; lighter payloads, longer client timeouts and offline "
        f"confirmation address {summary['second_failure_reason']} failures.\n\n"
        "3.  Make retry policy cause-aware.\n"
        f"     Retries recover {summary['retry_recovery_rate']:.1%} overall, but recovery differs sharply by cause. "
        "Retry infrastructure faults harder, stop retrying customer-side declines, and route around a bank while it is "
        "in an active incident.\n\n"
        "4.  Alert on bank × peak-hour success rate.\n"
        "     The March incident is only visible when the data is cut by bank and hour together; a daily portfolio "
        "average moved by well under one point."
    )
    headline = summary["scenarios"][HEADLINE_SCENARIO]
    impact = (
        f"{HEADLINE_SCENARIO} case — cut {summary['top_failure_reason']} and {summary['second_failure_reason']} "
        f"failures by {headline['timeout_reduction']:.0%} and lift retry recovery on other causes by "
        f"{headline['retry_uplift']:.0%}:\n\n"
        f"    •  {headline['recovered_transactions']:,} additional successful transactions\n"
        f"    •  ₹{headline['recovered_value_inr'] / 1e6:.2f}M of payment value recovered\n"
        f"    •  +{pp(headline['success_rate_uplift_pp'])} pp success rate → {headline['projected_success_rate']:.2%}\n\n"
        f"Conservative case: {summary['scenarios']['Conservative']['recovered_transactions']:,} transactions / "
        f"+{pp(summary['scenarios']['Conservative']['success_rate_uplift_pp'])} pp.\n"
        f"Ambitious case: {summary['scenarios']['Ambitious']['recovered_transactions']:,} transactions / "
        f"+{pp(summary['scenarios']['Ambitious']['success_rate_uplift_pp'])} pp.\n\n"
        "These are scenario estimates on synthetic data, not guaranteed benefits. Assumptions are editable on the "
        "Impact Simulator sheet."
    )

    write_section(ws, F, 41, 0, 5, "What the data shows")
    write_section(ws, F, 41, 6, 11, "Recommended actions")
    write_section(ws, F, 41, 12, last_col, f"Scenario impact ({HEADLINE_SCENARIO.lower()} case)")
    for r in range(42, 55):
        ws.set_row(r, 15)
    ws.merge_range(42, 0, 54, 5, findings, F["panel"])
    ws.merge_range(42, 6, 54, 11, actions, F["panel"])
    ws.merge_range(42, 12, 54, last_col, impact, F["panel_amber"])

    ws.set_row(56, 26)
    ws.merge_range(56, 0, 56, last_col, DISCLAIMER + " Built from a seeded generator; rebuilding the pipeline reproduces "
                                        "every figure on this sheet.", F["note"])

    ws.set_landscape()
    ws.set_paper(9)
    ws.fit_to_pages(1, 1)
    ws.set_margins(0.3, 0.3, 0.4, 0.3)
    ws.print_area(0, 0, 56, last_col)
    ws.set_header(f"&L{COMPANY} Payment Reliability Analytics&RSynthetic portfolio project")


# ---------------------------------------------------------------------------
# Trend Analysis
# ---------------------------------------------------------------------------
def write_trend_sheet(wb, ws, F, blocks, tables, summary) -> None:
    ws.hide_gridlines(2)
    ws.set_tab_color(C["blue"])
    ws.set_column(0, 0, 14)
    ws.set_column(1, 8, 15)
    write_banner(ws, F, "TREND ANALYSIS",
                 "Is the platform getting better or worse? Daily noise is smoothed with a 7-day moving average and "
                 "partial reporting weeks are excluded from the weekly chart. " + DISCLAIMER, 11)
    ws.freeze_panes(3, 0)

    write_section(ws, F, 3, 0, 11, "Monthly scorecard with month-over-month movement")
    monthly = tables["monthly_kpis"].copy()
    monthly_block = write_table(ws, F, monthly, 4, 0, header_labels={
        "month": "Month", "transactions": "Transactions", "payment_value_inr": "Payment value",
        "failures": "Failures", "success_rate": "Success rate", "avg_processing_seconds": "Avg processing (s)",
        "success_rate_pop_change": "MoM change in success rate", "transaction_pop_change": "Volume change",
    })
    ws.set_column(monthly_block.col("payment_value_inr"), monthly_block.col("payment_value_inr"), 16)
    ws.set_column(monthly_block.col("avg_processing_seconds"), monthly_block.col("avg_processing_seconds"), 17)
    ws.set_column(monthly_block.col("success_rate_pop_change"), monthly_block.col("success_rate_pop_change"), 16)
    ws.conditional_format(monthly_block.first_row, monthly_block.col("success_rate_pop_change"),
                          monthly_block.last_row, monthly_block.col("success_rate_pop_change"),
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": wb.add_format({"font_name": FONT, "font_size": 10, "font_color": C["red"],
                                                    "num_format": PCT2, "border": 1, "border_color": C["line"]})})

    row = monthly_block.last_row + 2
    ws.merge_range(row, 0, row, 11,
                   f"Monthly success rate moves inside a {pp(summary['monthly_range_pp'])} pp band. "
                   f"The weakest month is {summary['worst_month']} at {summary['worst_month_success_rate']:.2%}, which is "
                   f"where the bank incident falls; the portfolio average never moves by a full point, which is exactly why "
                   f"segment-level monitoring matters.", F["note"])
    ws.set_row(row, 30)

    chart_row = row + 2
    write_section(ws, F, chart_row, 0, 11, "Weekly success rate (complete weeks) and daily rate with 7-day moving average")
    weekly = blocks["weekly_complete"]
    daily = blocks["daily_trend"]

    weekly_chart = wb.add_chart({"type": "line"})
    weekly_chart.add_series({
        "name": "Weekly success rate",
        "categories": weekly.series("week_start"),
        "values": weekly.series("success_rate"),
        "line": {"color": C["blue"], "width": 2.25},
        "marker": {"type": "circle", "size": 5, "border": {"color": C["blue"]}, "fill": {"color": C["white"]}},
    })
    weekly_chart.set_title({"name": "Weekly transaction success rate", "name_font": {"name": FONT, "size": 11, "color": C["navy"]}})
    weekly_chart.set_y_axis({"min": 0.94, "max": 0.965, "num_format": PCT1, "num_font": {"name": FONT, "size": 9},
                             "major_gridlines": {"visible": True, "line": {"color": C["line"]}}})
    weekly_chart.set_x_axis({"num_font": {"name": FONT, "size": 8, "rotation": -45}})
    weekly_chart.set_legend({"none": True})
    weekly_chart.set_chartarea({"border": {"color": C["line"]}})
    weekly_chart.set_size({"width": 620, "height": 300})
    ws.insert_chart(chart_row + 1, 0, weekly_chart, {"x_offset": 4, "y_offset": 4})

    daily_chart = wb.add_chart({"type": "line"})
    daily_chart.add_series({
        "name": "Daily success rate",
        "categories": daily.series("date"),
        "values": daily.series("success_rate"),
        "line": {"color": C["line"], "width": 1},
    })
    daily_chart.add_series({
        "name": "7-day moving average",
        "categories": daily.series("date"),
        "values": daily.series("success_rate_7d_moving_avg"),
        "line": {"color": C["navy"], "width": 2.25},
    })
    daily_chart.set_title({"name": "Daily success rate and 7-day moving average", "name_font": {"name": FONT, "size": 11, "color": C["navy"]}})
    daily_chart.set_y_axis({"min": 0.92, "max": 0.98, "num_format": PCT1, "num_font": {"name": FONT, "size": 9},
                            "major_gridlines": {"visible": True, "line": {"color": C["line"]}}})
    daily_chart.set_x_axis({"num_font": {"name": FONT, "size": 8}, "interval_unit": 21, "interval_tick": 21})
    daily_chart.set_legend({"position": "bottom", "font": {"name": FONT, "size": 9}})
    daily_chart.set_chartarea({"border": {"color": C["line"]}})
    daily_chart.set_size({"width": 660, "height": 300})
    ws.insert_chart(chart_row + 1, 6, daily_chart, {"x_offset": 4, "y_offset": 4})

    table_row = chart_row + 17
    write_section(ws, F, table_row, 0, 11, "Weekly detail (partial first and last reporting weeks flagged)")
    weekly_table = write_table(ws, F, tables["weekly_trend"], table_row + 1, 0, header_labels={
        "week_start": "Week starting", "transactions": "Transactions", "days_covered": "Days covered",
        "is_complete_week": "Complete week", "payment_value_inr": "Payment value", "failures": "Failures",
        "success_rate": "Success rate", "success_rate_wow_change": "WoW change in success rate",
    })
    ws.set_column(weekly_table.col("payment_value_inr"), weekly_table.col("payment_value_inr"), 16)
    ws.set_column(weekly_table.col("success_rate_wow_change"), weekly_table.col("success_rate_wow_change"), 16)
    ws.set_landscape()
    ws.set_paper(9)
    ws.fit_to_pages(1, 0)
    ws.repeat_rows(table_row + 1)


# ---------------------------------------------------------------------------
# Failure Analysis
# ---------------------------------------------------------------------------
def write_failure_sheet(wb, ws, F, blocks, tables, summary) -> None:
    ws.hide_gridlines(2)
    ws.set_tab_color(C["red"])
    ws.set_column(0, 0, 22)
    ws.set_column(1, 8, 15)
    write_banner(ws, F, "FAILURE ANALYSIS",
                 "Which causes are worth engineering time, and when do they strike? Volume, not rate, drives "
                 "prioritisation here. " + DISCLAIMER, 11)
    ws.freeze_panes(3, 0)

    write_section(ws, F, 3, 0, 11, "Failure Pareto: two causes carry half of all failures")
    pareto_table = write_table(ws, F, tables["failure_pareto"], 4, 0, header_labels={
        "failure_reason": "Failure reason", "failures": "Failures", "failed_value_inr": "Failed value",
        "avg_failed_value_inr": "Avg failed value", "failure_rank": "Rank",
        "failure_contribution": "Share of failures", "cumulative_contribution": "Cumulative share",
    })
    ws.set_column(pareto_table.col("failed_value_inr"), pareto_table.col("avg_failed_value_inr"), 16)
    ws.set_column(pareto_table.col("failure_rank"), pareto_table.col("failure_rank"), 8)

    pareto = blocks["failure_pareto"]
    chart = wb.add_chart({"type": "column"})
    chart.add_series({
        "name": "Failed transactions", "categories": pareto.series("failure_reason"),
        "values": pareto.series("failures"), "fill": {"color": C["blue"]}, "gap": 40,
        "data_labels": {"value": True, "num_format": NUM, "font": {"name": FONT, "size": 9}},
    })
    line = wb.add_chart({"type": "line"})
    line.add_series({
        "name": "Cumulative share", "categories": pareto.series("failure_reason"),
        "values": pareto.series("cumulative_contribution"), "y2_axis": True,
        "line": {"color": C["amber"], "width": 2}, "marker": {"type": "circle", "size": 5, "fill": {"color": C["amber"]}},
    })
    chart.combine(line)
    chart.set_title({"name": "Failure Pareto", "name_font": {"name": FONT, "size": 11, "color": C["navy"]}})
    chart.set_y_axis({"num_format": NUM, "num_font": {"name": FONT, "size": 9},
                      "major_gridlines": {"visible": True, "line": {"color": C["line"]}}})
    line.set_y2_axis({"num_format": PCT1, "min": 0, "max": 1, "num_font": {"name": FONT, "size": 9}})
    chart.set_x_axis({"num_font": {"name": FONT, "size": 9, "rotation": -30}})
    chart.set_legend({"position": "bottom", "font": {"name": FONT, "size": 9}})
    chart.set_chartarea({"border": {"color": C["line"]}})
    chart.set_size({"width": 700, "height": 300})
    ws.insert_chart(13, 0, chart, {"x_offset": 4, "y_offset": 8})

    row = 30
    write_section(ws, F, row, 0, 11, "Peak versus non-peak")
    peak_table = write_table(ws, F, tables["peak_comparison"], row + 1, 0, header_labels={
        "peak_period": "Period", "transactions": "Transactions", "failures": "Failures",
        "success_rate": "Success rate", "avg_processing_seconds": "Avg processing (s)",
        "timeout_share_of_failures": "Timeout share of failures",
    })
    ws.set_column(peak_table.col("avg_processing_seconds"), peak_table.col("timeout_share_of_failures"), 20)
    row = peak_table.last_row + 2
    ws.set_row(row, 30)
    ws.merge_range(row, 0, row, 11,
                   f"The {summary['peak_definition']} window carries {summary['peak_share_of_volume']:.1%} of volume but "
                   f"{summary['peak_share_of_failures']:.1%} of failures, and timeouts rise from "
                   f"{tables['peak_comparison'].set_index('peak_period').loc['Non-peak', 'timeout_share_of_failures']:.1%} "
                   f"of off-peak failures to {summary['peak_timeout_share_of_failures']:.1%} of peak failures — the "
                   f"signature of capacity pressure rather than customer behaviour.", F["note"])

    row += 2
    write_section(ws, F, row, 0, 11, "Hour-of-day profile")
    hourly_table = write_table(ws, F, tables["hourly_performance"], row + 1, 0, header_labels={
        "hour": "Hour", "peak_period": "Period", "transactions": "Transactions", "failures": "Failures",
        "success_rate": "Success rate", "timeout_failures": "Timeout failures", "worst_hour_rank": "Worst-hour rank",
    })
    hourly = blocks["hourly_performance"]
    hour_chart = wb.add_chart({"type": "column"})
    hour_chart.add_series({
        "name": "Transactions", "categories": hourly.series("hour"), "values": hourly.series("transactions"),
        "fill": {"color": C["pale"]}, "border": {"color": C["teal"]}, "gap": 30,
    })
    hour_line = wb.add_chart({"type": "line"})
    hour_line.add_series({
        "name": "Success rate", "categories": hourly.series("hour"), "values": hourly.series("success_rate"),
        "y2_axis": True, "line": {"color": C["red"], "width": 2.25},
    })
    hour_chart.combine(hour_line)
    hour_chart.set_title({"name": "Volume and success rate by hour of day", "name_font": {"name": FONT, "size": 11, "color": C["navy"]}})
    hour_chart.set_y_axis({"num_format": NUM, "num_font": {"name": FONT, "size": 9},
                           "major_gridlines": {"visible": True, "line": {"color": C["line"]}}})
    hour_line.set_y2_axis({"num_format": PCT1, "min": 0.93, "max": 0.97, "num_font": {"name": FONT, "size": 9}})
    hour_chart.set_x_axis({"num_font": {"name": FONT, "size": 9}, "name": "Hour of day",
                           "name_font": {"name": FONT, "size": 9, "color": C["muted"]}})
    hour_chart.set_legend({"position": "bottom", "font": {"name": FONT, "size": 9}})
    hour_chart.set_chartarea({"border": {"color": C["line"]}})
    hour_chart.set_size({"width": 700, "height": 320})
    ws.insert_chart(row + 1, 8, hour_chart, {"x_offset": 12, "y_offset": 4})

    ws.set_landscape()
    ws.set_paper(9)
    ws.fit_to_pages(1, 0)


# ---------------------------------------------------------------------------
# Bank and Segment Analysis
# ---------------------------------------------------------------------------
def write_segment_sheet(wb, ws, F, blocks, tables, summary) -> None:
    ws.hide_gridlines(2)
    ws.set_tab_color(C["teal"])
    ws.set_column(0, 0, 22)
    ws.set_column(1, 9, 15)
    write_banner(ws, F, "BANK AND SEGMENT ANALYSIS",
                 "Rates are always shown next to absolute volumes and a 95% confidence interval, so a small or noisy "
                 "segment is never mistaken for an operational problem. " + DISCLAIMER, 11)
    ws.freeze_panes(3, 0)

    write_section(ws, F, 3, 0, 11, "Issuing-bank scorecard")
    bank_table = write_table(ws, F, tables["bank_performance"], 4, 0, header_labels={
        "segment": "Issuing bank", "transactions": "Transactions", "failures": "Failures",
        "success_rate": "Success rate", "timeout_failures": "Timeout failures",
        "failure_contribution": "Share of all failures", "success_rate_ci95_halfwidth": "95% CI (+/-)",
        "gap_to_portfolio_rate": "Gap to portfolio rate",
    })
    ws.set_column(bank_table.col("failure_contribution"), bank_table.col("gap_to_portfolio_rate"), 18)
    row = bank_table.last_row + 2
    ws.set_row(row, 32)
    ws.merge_range(row, 0, row, 7,
                   f"Read with care: the best and worst bank differ by {pp(summary['bank_spread_pp'])} pp over six months, "
                   f"but the 95% confidence half-width reaches ±{pp(summary['bank_max_ci95_pp'])} pp on the smallest bank. "
                   f"The intervals overlap, so the six-month ranking is not actionable. The defensible bank finding is the "
                   f"time-bounded {summary['incident_bank']} incident below.", F["note"])

    row += 2
    write_section(ws, F, row, 0, 11, f"Time-bounded incident: {summary['incident_bank']} versus all other banks, peak hours only")
    incident_table = write_table(ws, F, tables["incident_bank_comparison"], row + 1, 0, header_labels={
        "bank_group": "Bank group", "period": "Period", "transactions": "Transactions", "failures": "Failures",
        "success_rate": "Success rate", "timeout_failures": "Timeout failures", "timeout_rate": "Timeout rate",
    })
    ws.set_column(incident_table.col("period"), incident_table.col("period"), 24)

    incident = blocks["incident_bank_comparison"]
    incident_chart = wb.add_chart({"type": "column"})
    for i, group in enumerate(tables["incident_bank_comparison"].bank_group.unique()):
        subset = tables["incident_bank_comparison"].reset_index(drop=True)
        rows = subset.index[subset.bank_group.eq(group)].tolist()
        incident_chart.add_series({
            "name": group,
            "categories": [DATA_SHEET, incident.first_row + rows[0], incident.col("period"),
                           incident.first_row + rows[-1], incident.col("period")],
            "values": [DATA_SHEET, incident.first_row + rows[0], incident.col("success_rate"),
                       incident.first_row + rows[-1], incident.col("success_rate")],
            "fill": {"color": C["red"] if i == 0 else C["teal"]},
            "gap": 60,
            "data_labels": {"value": True, "num_format": PCT1, "font": {"name": FONT, "size": 9}},
        })
    incident_chart.set_title({"name": "Peak-hour success rate before, during and after the incident",
                              "name_font": {"name": FONT, "size": 11, "color": C["navy"]}})
    incident_chart.set_y_axis({"min": 0.87, "max": 0.97, "num_format": PCT1, "num_font": {"name": FONT, "size": 9},
                               "major_gridlines": {"visible": True, "line": {"color": C["line"]}}})
    incident_chart.set_x_axis({"num_font": {"name": FONT, "size": 9}})
    incident_chart.set_legend({"position": "bottom", "font": {"name": FONT, "size": 9}})
    incident_chart.set_chartarea({"border": {"color": C["line"]}})
    incident_chart.set_size({"width": 660, "height": 300})
    ws.insert_chart(row + 1, 8, incident_chart, {"x_offset": 12, "y_offset": 4})

    # The note stays clear of the floating chart by ending at column H.
    row = incident_table.last_row + 2
    ws.set_row(row, 30)
    ws.merge_range(row, 0, row, 7,
                   f"{summary['incident_bank']} peak-hour success fell to {summary['incident_bank_peak_during']:.2%} during "
                   f"10-20 March while every other bank held at {summary['other_banks_peak_during']:.2%} in the same hours, "
                   f"and its timeout rate rose from {summary['incident_bank_timeout_rate_before']:.2%} to "
                   f"{summary['incident_bank_timeout_rate_during']:.2%}. Both recovered afterwards, which is what "
                   f"distinguishes a contained incident from a structural problem.", F["note"])

    # Full-width blocks resume only below the incident chart.
    row = max(row + 2, 32)
    write_section(ws, F, row, 0, 11, "Segment scorecards across every dimension")
    segment_table = write_table(ws, F, tables["segment_performance"], row + 1, 0, header_labels={
        "dimension": "Dimension", "segment": "Segment", "transactions": "Transactions", "failures": "Failures",
        "success_rate": "Success rate", "payment_value_inr": "Payment value",
        "failure_contribution": "Share of dimension failures", "gap_to_portfolio_rate": "Gap to portfolio rate",
        "worst_rank_in_dimension": "Rank in dimension", "low_sample_flag": "Low sample (<1,000)",
    })
    ws.set_column(segment_table.col("dimension"), segment_table.col("segment"), 20)
    ws.set_column(segment_table.col("failure_contribution"), segment_table.col("low_sample_flag"), 18)
    ws.conditional_format(segment_table.first_row, segment_table.col("low_sample_flag"),
                          segment_table.last_row, segment_table.col("low_sample_flag"),
                          {"type": "cell", "criteria": "==", "value": 1,
                           "format": wb.add_format({"font_name": FONT, "font_size": 10, "bold": True,
                                                    "font_color": "#7A5C00", "bg_color": C["input"],
                                                    "num_format": NUM, "border": 1, "border_color": C["line"]})})

    row = segment_table.last_row + 2
    write_section(ws, F, row, 0, 11, "Highest-impact bank x hour combinations, ranked by absolute failures")
    write_table(ws, F, tables["bank_hour_impact"], row + 1, 0, header_labels={
        "issuing_bank": "Issuing bank", "hour": "Hour", "peak_period": "Period", "transactions": "Transactions",
        "failures": "Failures", "success_rate": "Success rate", "impact_rank": "Impact rank",
    })

    ws.set_landscape()
    ws.set_paper(9)
    ws.fit_to_pages(1, 0)


# ---------------------------------------------------------------------------
# Impact Simulator
# ---------------------------------------------------------------------------
def write_simulator(wb, ws, F, blocks, cells, summary) -> None:
    ws.hide_gridlines(2)
    ws.set_tab_color(C["inputborder"])
    ws.set_column(0, 0, 34)
    ws.set_column(1, 7, 17)
    ws.set_column(8, 8, 20)
    write_banner(ws, F, "IMPACT SIMULATOR",
                 "Edit the yellow cells to test remediation assumptions. Every output is a live formula over the "
                 "baselines below. Outputs are scenario estimates on synthetic data, not guaranteed benefits.", 8)
    ws.freeze_panes(3, 0)

    pools = blocks["simulator_baseline"]
    pool_rows = {name: pools.first_row + i for i, name in enumerate(pools.frame["pool"])}

    def pool_ref(pool: str, column: str) -> str:
        return sheet_cell(pool_rows[pool], pools.col(column))

    # --- Baselines ----------------------------------------------------------
    write_section(ws, F, 3, 0, 8, "Step 1 — Observed baselines (locked, sourced from the SQL analysis)")
    baseline_rows = [
        ("Total transactions in period", pool_ref("Total transactions", "transactions"), summary["transactions"], F["num"]),
        ("Successful transactions", pool_ref("Successful transactions", "transactions"), summary["successful_transactions"], F["num"]),
        ("Baseline success rate", None, None, F["pct2"]),
        ("Bank Timeout failures", pool_ref("Bank Timeout failures", "transactions"), summary["timeout_failures"], F["num"]),
        ("Average value of a timeout failure", pool_ref("Bank Timeout failures", "avg_value_inr"), summary["timeout_avg_value_inr"], F["money"]),
        ("Network Error failures", pool_ref("Network Error failures", "transactions"), summary["network_failures"], F["num"]),
        ("Average value of a network failure", pool_ref("Network Error failures", "avg_value_inr"), summary["network_avg_value_inr"], F["money"]),
        ("Retried, still failed (other causes)", pool_ref("Retried, still failed (other causes)", "transactions"), summary["retry_pool_transactions"], F["num"]),
        ("Average value of that retry pool", pool_ref("Retried, still failed (other causes)", "avg_value_inr"), summary["retry_pool_avg_value_inr"], F["money"]),
    ]
    first_baseline = 5
    ws.write(4, 0, "Baseline", F["header"])
    ws.write(4, 1, "Value", F["header"])
    for i, (label, ref, value, cell_format) in enumerate(baseline_rows):
        r = first_baseline + i
        ws.write(r, 0, label, F["text"])
        if label == "Baseline success rate":
            ws.write_formula(r, 1, f"=IFERROR({xl_rowcol_to_cell(first_baseline + 1, 1)}/{xl_rowcol_to_cell(first_baseline, 1)},0)",
                             F["pct2"], summary["success_rate"])
        else:
            ws.write_formula(r, 1, f"={ref}", cell_format, value)
    B = {label: xl_rowcol_to_cell(first_baseline + i, 1, row_abs=True, col_abs=True)
         for i, (label, *_rest) in enumerate(baseline_rows)}

    row = first_baseline + len(baseline_rows) + 1
    ws.set_row(row, 28)
    ws.merge_range(row, 0, row, 8,
                   "The three levers act on mutually exclusive pools: the retry lever deliberately excludes Bank Timeout "
                   "and Network Error failures, so no transaction can be counted twice.", F["note"])

    # --- Assumptions and outputs -------------------------------------------
    row += 2
    write_section(ws, F, row, 0, 8, "Step 2 — Editable assumptions (yellow) and modelled outcomes")
    header_row = row + 1
    headers = ["Scenario", "Timeout failures\nprevented", "Network failures\nprevented",
               "Retry recovery uplift\n(other causes)", "Recovered\ntransactions", "Recovered payment\nvalue",
               "Success-rate\nuplift", "Projected\nsuccess rate", "Status"]
    ws.set_row(header_row, 42)
    for j, text in enumerate(headers):
        ws.write(header_row, j, text, F["header"])

    scenarios = list(summary["scenarios"].items())
    first_scenario = header_row + 1
    for i, (name, values) in enumerate(scenarios):
        r = first_scenario + i
        ws.set_row(r, 22)
        ws.write(r, 0, name, F["text_bold"])
        ws.write_number(r, 1, values["timeout_reduction"], F["input"])
        ws.write_number(r, 2, values["network_reduction"], F["input"])
        ws.write_number(r, 3, values["retry_uplift"], F["input"])
        cell = lambda col, rr=r: xl_rowcol_to_cell(rr, col)  # noqa: E731 - local shorthand
        ws.write_formula(
            r, 4,
            f"=ROUND({B['Bank Timeout failures']}*{cell(1)}+{B['Network Error failures']}*{cell(2)}"
            f"+{B['Retried, still failed (other causes)']}*{cell(3)},0)",
            F["out_num"], values["recovered_transactions"])
        ws.write_formula(
            r, 5,
            f"={B['Bank Timeout failures']}*{cell(1)}*{B['Average value of a timeout failure']}"
            f"+{B['Network Error failures']}*{cell(2)}*{B['Average value of a network failure']}"
            f"+{B['Retried, still failed (other causes)']}*{cell(3)}*{B['Average value of that retry pool']}",
            F["out_money"], values["recovered_value_inr"])
        ws.write_formula(r, 6, f"=IFERROR({cell(4)}/{B['Total transactions in period']}*100,0)",
                         F["out_pp"], values["success_rate_uplift_pp"])
        ws.write_formula(r, 7,
                         f"=IFERROR(({B['Successful transactions']}+{cell(4)})/{B['Total transactions in period']},0)",
                         F["out_pct"], values["projected_success_rate"])
        ws.write(r, 8, "Scenario estimate", F["text"])

    last_scenario = first_scenario + len(scenarios) - 1
    ws.data_validation(first_scenario, 1, last_scenario, 3, {
        "validate": "decimal", "criteria": "between", "minimum": 0, "maximum": 1,
        "input_title": "Enter 0% to 100%", "input_message": "Share of that failure pool the change would remove.",
        "error_title": "Value out of range", "error_message": "Enter a percentage between 0% and 100%.",
    })

    row = last_scenario + 2
    ws.set_row(row, 26)
    ws.merge_range(row, 0, row, 8,
                   "Yellow cells are the only inputs. Recovered transactions = timeout failures x timeout reduction "
                   "+ network failures x network reduction + retry pool x recovery uplift. Recovered value applies each "
                   "pool's observed average failed value. Uplift = recovered transactions / total transactions.", F["note"])

    # --- Scenario chart -----------------------------------------------------
    row += 2
    write_section(ws, F, row, 0, 8, "Step 3 — Compare the scenarios")
    chart = wb.add_chart({"type": "column"})
    chart.add_series({
        "name": "Recovered transactions",
        "categories": ["Impact Simulator", first_scenario, 0, last_scenario, 0],
        "values": ["Impact Simulator", first_scenario, 4, last_scenario, 4],
        "fill": {"color": C["teal"]}, "gap": 70,
        "data_labels": {"value": True, "num_format": NUM, "font": {"name": FONT, "size": 10, "bold": True}},
    })
    uplift = wb.add_chart({"type": "line"})
    uplift.add_series({
        "name": "Success-rate uplift (pp)",
        "categories": ["Impact Simulator", first_scenario, 0, last_scenario, 0],
        "values": ["Impact Simulator", first_scenario, 6, last_scenario, 6],
        "y2_axis": True, "line": {"color": C["amber"], "width": 2.25},
        "marker": {"type": "circle", "size": 6, "fill": {"color": C["amber"]}},
    })
    chart.combine(uplift)
    chart.set_title({"name": "Recovered transactions and success-rate uplift by scenario",
                     "name_font": {"name": FONT, "size": 11, "color": C["navy"]}})
    chart.set_y_axis({"num_format": NUM, "num_font": {"name": FONT, "size": 9},
                      "major_gridlines": {"visible": True, "line": {"color": C["line"]}}})
    uplift.set_y2_axis({"num_format": DEC2, "num_font": {"name": FONT, "size": 9},
                        "name": "pp", "name_font": {"name": FONT, "size": 9, "color": C["muted"]}})
    chart.set_x_axis({"num_font": {"name": FONT, "size": 10}})
    chart.set_legend({"position": "bottom", "font": {"name": FONT, "size": 9}})
    chart.set_chartarea({"border": {"color": C["line"]}})
    chart.set_size({"width": 660, "height": 300})
    ws.insert_chart(row + 1, 0, chart, {"x_offset": 4, "y_offset": 6})

    row += 17
    write_section(ws, F, row, 0, 8, "What this model does not claim")
    ws.merge_range(row + 1, 0, row + 6, 8,
                   "This simulator converts a reduction in a measured failure pool into the transactions and payment "
                   "value that would have succeeded instead, at each pool's observed average value.\n"
                   "It does NOT model: customer re-attempt behaviour or substitution to another payment method; the "
                   "engineering cost, capacity or risk of each intervention; fraud or chargeback effects of retrying more; "
                   "interaction between the three levers; or any revenue attribution.\n"
                   "The underlying data is synthetic and generated for a portfolio project, so every figure is an "
                   "illustrative scenario estimate — not a forecast, a commitment, or evidence of realised benefit.",
                   F["panel_amber"])

    ws.set_landscape()
    ws.set_paper(9)
    ws.fit_to_pages(1, 0)
    ws.print_area(0, 0, row + 6, 8)


# ---------------------------------------------------------------------------
# Reference sheets
# ---------------------------------------------------------------------------
def write_reference_sheet(wb, ws, F, title: str, subtitle: str, columns: list[str],
                          rows: list[list[str]], widths: list[int], tab_color: str) -> None:
    ws.hide_gridlines(2)
    ws.set_tab_color(tab_color)
    write_banner(ws, F, title, subtitle, len(columns) - 1)
    for i, width in enumerate(widths):
        ws.set_column(i, i, width)
    ws.set_row(3, 28)
    for j, name in enumerate(columns):
        ws.write(3, j, name, F["header"])
    for i, record in enumerate(rows, start=4):
        ws.set_row(i, 30)
        for j, value in enumerate(record):
            ws.write(i, j, value, F["text_wrap"] if j else F["text_bold"])
    # These cells are deliberately text ("95.57%", "1.71 s"); suppress Excel's
    # green "number stored as text" markers so the sheet reads cleanly.
    ws.ignore_errors({"number_stored_as_text": xl_range_abs(4, 0, 3 + len(rows), len(columns) - 1)})
    ws.freeze_panes(4, 0)
    ws.set_landscape()
    ws.set_paper(9)
    ws.fit_to_pages(1, 0)
    ws.repeat_rows(3)
    ws.print_area(0, 0, 3 + len(rows), len(columns) - 1)


def kpi_definition_rows(summary: dict) -> list[list[str]]:
    return [
        ["Transaction success rate", "Share of payment attempts that succeeded on the initial attempt.",
         "COUNT(status = 'Success') / COUNT(*)", "All attempts in the period",
         f"{summary['success_rate']:.2%}", "Primary reliability KPI. Excludes later retry recoveries by design."],
        ["Value success rate", "Share of attempted payment value that succeeded.",
         "SUM(amount WHERE success) / SUM(amount)", "Total attempted value",
         f"{summary['value_success_rate']:.2%}", "Below the attempt rate means failures skew to larger payments."],
        ["Failure rate", "Share of attempts that failed.", "1 - success rate", "All attempts",
         f"{summary['failure_rate']:.2%}", "Always read next to the absolute failure count."],
        ["Average processing time", "Mean end-to-end attempt latency.", "AVG(processing_time_seconds)", "All attempts",
         f"{summary['avg_processing_seconds']:.2f} s", "Averages hide the tail; use P95 for user experience."],
        ["P95 processing time", "95th percentile latency (pandas linear interpolation).",
         "QUANTILE(processing_time_seconds, 0.95)", "All attempts",
         f"{summary['p95_processing_seconds']:.2f} s", "One in twenty attempts is slower than this."],
        ["Retry rate", "Share of all attempts that were retried.", "SUM(retry_flag) / COUNT(*)", "All attempts",
         f"{summary['retry_rate']:.2%}", "Only failed attempts are ever retried in this dataset."],
        ["Retry rate of failures", "Share of failed attempts that were retried.", "SUM(retry_flag) / COUNT(failures)",
         "Failed attempts", f"{summary['retry_rate_of_failures']:.1%}", "Shows retry coverage rather than retry volume."],
        ["Retry recovery rate", "Share of retried attempts that eventually succeeded.",
         "COUNT(outcome = 'Recovered') / SUM(retry_flag)", "Retried attempts",
         f"{summary['retry_recovery_rate']:.2%}", "Effectiveness of the retry policy, not of the platform."],
        ["Failure contribution", "Share of all failures caused by one root cause.",
         "cause failures / all failures", "Failed attempts",
         f"{summary['top_failure_reason']}: {summary['top_failure_contribution']:.1%}", "Drives Pareto prioritisation."],
        ["Peak-hour gap", "Non-peak success rate minus peak success rate.",
         "success_rate(non-peak) - success_rate(peak)", f"Peak = {summary['peak_definition']}",
         f"{pp(summary['peak_gap_pp'])} pp", "Positive means reliability degrades under evening load."],
        ["95% confidence half-width", "Normal-approximation uncertainty on a segment success rate.",
         "1.96 * SQRT(p * (1 - p) / n)", "One segment",
         f"up to ±{pp(summary['bank_max_ci95_pp'])} pp across banks", "Overlapping intervals mean a gap is not conclusive."],
        ["Success-rate uplift (pp)", "Modelled percentage-point gain from a scenario.",
         "recovered transactions / total transactions x 100", "All attempts",
         f"{HEADLINE_SCENARIO}: +{pp(summary['expected_success_rate_uplift_pp'])} pp",
         "Scenario estimate on synthetic data, not a guaranteed benefit."],
    ]


def data_dictionary_rows() -> list[list[str]]:
    return [
        ["transaction_id", "Text", "Unique synthetic identifier for one payment attempt.", "TXN000012345", "Always unique; 1 row = 1 attempt."],
        ["transaction_timestamp", "Datetime", "When the attempt was made (local synthetic context).", "2025-03-14 19:42:10", "Drives date, week, hour, peak and incident cuts."],
        ["customer_id", "Text", "Synthetic customer key (75,000 customers).", "CUS0012345", "No real person is represented."],
        ["merchant_id", "Text", "Synthetic merchant key (6,000 merchants).", "MER01234", "No real business is represented."],
        ["amount_inr", "Decimal", "Attempted payment value in INR.", "499.00", "Long-tailed: median far below the mean."],
        ["payment_method", "Category", "UPI, Debit Card, Credit Card, Wallet or Net Banking.", "UPI", "UPI is ~54% of volume by design."],
        ["issuing_bank", "Category", "Fictional issuing bank of the payment instrument.", "Crescent Bank", "Six fictional banks; none maps to a real bank."],
        ["merchant_category", "Category", "Merchant industry grouping.", "Grocery", "Eight categories."],
        ["city / state / region", "Category", "Synthetic geography of the attempt.", "Pune / Maharashtra / West", "Twelve cities across five regions."],
        ["device_type", "Category", "Android, iOS or Web.", "Android", "Android carries ~68% of volume."],
        ["network_type", "Category", "Connection quality at attempt time.", "4G", "2G/3G sessions fail measurably more often."],
        ["customer_tenure_group", "Category", "How long the customer has been on the platform.", "1-3 years", "Four tenure bands."],
        ["transaction_status", "Category", "Outcome of the INITIAL attempt: Success or Failed.", "Failed", "The success-rate KPI is built on this field."],
        ["failure_reason", "Category", "Root cause, populated only for failed attempts.", "Bank Timeout", "Blank on successful rows by design, not missing data."],
        ["processing_time_seconds", "Decimal", "End-to-end latency of the attempt.", "2.418", "Failures are slower; timeouts slowest."],
        ["retry_flag", "Integer", "1 if the failed attempt was retried, else 0.", "1", "Only failed attempts are ever retried."],
        ["retry_attempt", "Integer", "Retry ordinal, 0 when not retried.", "1", "Maximum of 3."],
        ["final_retry_outcome", "Category", "Not Retried, Recovered or Still Failed.", "Recovered", "Recovery is reported separately from the success rate."],
    ]


def read_me_rows(summary: dict) -> list[list[str]]:
    return [
        ["Purpose", "Diagnose payment reliability for the fictional company PayPulse: measure the success rate, find "
                    "where failures concentrate, and size the operational upside of fixing the largest causes."],
        ["SYNTHETIC DATA", DISCLAIMER + " Nothing here is derived from PhonePe or any other real organisation."],
        ["Reporting period", f"{summary['period_start']} to {summary['period_end']} ({summary['period_days']} days, "
                             f"{summary['transactions'] / 1000:.0f}K attempts, "
                             f"₹{summary['payment_value_inr'] / 1e6:.1f}M attempted value)."],
        ["How to read this workbook", "Executive Dashboard for the answer; Trend, Failure and Bank/Segment sheets for the "
                                      "evidence; Impact Simulator to test assumptions; KPI Definitions and Data Dictionary "
                                      "for the contract; Analysis Data for reconciliation."],
        ["Where the numbers come from", "Every presentation figure is a formula pointing at the Analysis Data sheet, which "
                                        "holds the exact CSV outputs of the SQL analysis. Nothing is typed in twice."],
        ["Status definition", "transaction_status is the outcome of the INITIAL attempt. Retry recovery is measured "
                              "separately through final_retry_outcome, so the headline success rate is never inflated by retries."],
        ["Peak hours", f"{summary['peak_definition']} local time, chosen because volume peaks there; the definition lives in "
                       "src/config.py and is injected into the SQL view so it cannot drift."],
        ["Incident window", f"{summary['incident_bank']}, {summary['incident_start']} to 20 March 2025, peak hours only. "
                            "This is engineered into the generator so the analysis has a known ground truth to recover."],
        ["Simulator inputs", "Yellow cells on the Impact Simulator sheet accept 0%-100%. Outputs are live formulas and are "
                             "scenario estimates, not guaranteed benefits."],
        ["Rebuilding", "From the repository root run: python src/run_pipeline.py — it regenerates the data, validates it, "
                       "rebuilds SQLite, re-runs the SQL, rebuilds this workbook and runs the tests."],
        ["Limitations", "Synthetic patterns are illustrative. No causal inference, revenue attribution, cost model, fraud "
                        "economics or production SLA is claimed anywhere in this workbook."],
    ]


# ---------------------------------------------------------------------------
def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    tables, summary = load_analysis()
    logger.info("Loaded %d analysis tables", len(tables))

    wb = xlsxwriter.Workbook(str(WORKBOOK_PATH), {"constant_memory": False})
    wb.set_properties({
        "title": "PayPulse Payment Reliability Analytics",
        "subject": "Synthetic transaction success-rate and failure root-cause analysis",
        "author": "Portfolio project",
        "company": "PayPulse (fictional)",
        "comments": DISCLAIMER,
        # Pinned so a rebuild is byte-identical; otherwise the current time is
        # stamped into the file and every pipeline run dirties the repository.
        "created": WORKBOOK_CREATED,
    })
    F = build_formats(wb)
    sheets = {name: wb.add_worksheet(name) for name in SHEET_ORDER}

    blocks, cells = write_analysis_data(wb, sheets[DATA_SHEET], F, tables, summary)
    logger.info("Wrote %s with %d blocks", DATA_SHEET, len(blocks))

    write_dashboard(wb, sheets["Executive Dashboard"], F, blocks, cells, summary)
    logger.info("Wrote Executive Dashboard")
    write_trend_sheet(wb, sheets["Trend Analysis"], F, blocks, tables, summary)
    logger.info("Wrote Trend Analysis")
    write_failure_sheet(wb, sheets["Failure Analysis"], F, blocks, tables, summary)
    logger.info("Wrote Failure Analysis")
    write_segment_sheet(wb, sheets["Bank and Segment Analysis"], F, blocks, tables, summary)
    logger.info("Wrote Bank and Segment Analysis")
    write_simulator(wb, sheets["Impact Simulator"], F, blocks, cells, summary)
    logger.info("Wrote Impact Simulator")

    write_reference_sheet(wb, sheets["KPI Definitions"], F, "KPI DEFINITIONS",
                          "One definition per KPI, used identically in SQL, Python, this workbook and the written report. "
                          + DISCLAIMER,
                          ["KPI", "Definition", "Formula", "Denominator", "This dataset", "How to interpret it"],
                          kpi_definition_rows(summary), [26, 46, 40, 26, 22, 52], C["muted"])
    logger.info("Wrote KPI Definitions")
    write_reference_sheet(wb, sheets["Data Dictionary"], F, "DATA DICTIONARY",
                          "All 20 fields of the synthetic transaction table. " + DISCLAIMER,
                          ["Field", "Type", "Business meaning", "Example", "Analysis note"],
                          data_dictionary_rows(), [26, 12, 58, 26, 52], C["muted"])
    logger.info("Wrote Data Dictionary")
    write_reference_sheet(wb, sheets["Read Me"], F, "READ ME",
                          "How this workbook is built, what it claims and what it does not.",
                          ["Topic", "Guidance"], read_me_rows(summary), [26, 118], C["muted"])
    logger.info("Wrote Read Me")

    sheets["Executive Dashboard"].activate()
    wb.close()
    logger.info("Workbook written to %s", WORKBOOK_PATH.relative_to(WORKBOOK_PATH.parents[1]))


if __name__ == "__main__":
    main()
