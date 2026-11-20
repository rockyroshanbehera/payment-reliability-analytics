"""Shared fixtures and a dependency-free XLSX reader for the test suite."""
from __future__ import annotations

import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

import config  # noqa: E402  (import after sys.path is extended)

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkg": "http://schemas.openxmlformats.org/package/2006/relationships",
}

ANALYSIS_TABLES = [
    "kpi_overview", "monthly_kpis", "daily_trend", "weekly_trend", "failure_pareto",
    "peak_comparison", "hourly_performance", "bank_performance", "segment_performance",
    "bank_hour_impact", "retry_analysis", "retry_by_attempt", "incident_analysis",
    "incident_bank_comparison", "impact_baseline", "simulator_baseline",
]


class Workbook:
    """Minimal read-only view over an .xlsx file, using only the standard library.

    Reading the raw XML is deliberate: it lets the tests assert on what is
    actually stored in the file - sheet order, formulas, cached results and
    error cells - without depending on a spreadsheet library.
    """

    def __init__(self, path: Path) -> None:
        self.path = path
        self.zip = zipfile.ZipFile(path)
        self.shared_strings = self._read_shared_strings()
        self.sheet_names, self._sheet_paths = self._read_sheet_index()
        self._cache: dict[str, dict[str, dict]] = {}

    def _read_shared_strings(self) -> list[str]:
        if "xl/sharedStrings.xml" not in self.zip.namelist():
            return []
        root = ElementTree.fromstring(self.zip.read("xl/sharedStrings.xml"))
        return ["".join(t.text or "" for t in si.iter(f"{{{NS['main']}}}t")) for si in root]

    def _read_sheet_index(self) -> tuple[list[str], dict[str, str]]:
        workbook = ElementTree.fromstring(self.zip.read("xl/workbook.xml"))
        rels = ElementTree.fromstring(self.zip.read("xl/_rels/workbook.xml.rels"))
        targets = {rel.get("Id"): rel.get("Target") for rel in rels}
        names, paths = [], {}
        for sheet in workbook.find(f"{{{NS['main']}}}sheets"):
            name = sheet.get("name")
            target = targets[sheet.get(f"{{{NS['rel']}}}id")].lstrip("/")
            names.append(name)
            paths[name] = target if target.startswith("xl/") else f"xl/{target}"
        return names, paths

    def cells(self, sheet: str) -> dict[str, dict]:
        """Return {cell_ref: {'value', 'formula', 'type'}} for one sheet."""
        if sheet not in self._cache:
            root = ElementTree.fromstring(self.zip.read(self._sheet_paths[sheet]))
            cells: dict[str, dict] = {}
            for cell in root.iter(f"{{{NS['main']}}}c"):
                formula = cell.find(f"{{{NS['main']}}}f")
                value = cell.find(f"{{{NS['main']}}}v")
                raw = value.text if value is not None else None
                cell_type = cell.get("t")
                if cell_type == "s" and raw is not None:
                    raw = self.shared_strings[int(raw)]
                cells[cell.get("r")] = {
                    "value": raw,
                    "formula": formula.text if formula is not None else None,
                    "type": cell_type,
                }
            self._cache[sheet] = cells
        return self._cache[sheet]

    def number(self, sheet: str, ref: str) -> float:
        return float(self.cells(sheet)[ref]["value"])

    def sheet_xml(self, sheet: str) -> str:
        return self.zip.read(self._sheet_paths[sheet]).decode("utf-8")

    def chart_count(self) -> int:
        return len([n for n in self.zip.namelist() if n.startswith("xl/charts/chart") and n.endswith(".xml")])


def require(path: Path, hint: str):
    if not path.exists():
        pytest.skip(f"{path.relative_to(ROOT).as_posix()} missing - run `python src/run_pipeline.py` ({hint})")
    return path


@pytest.fixture(scope="session")
def root() -> Path:
    return ROOT


@pytest.fixture(scope="session")
def summary() -> dict:
    require(config.SUMMARY_PATH, "analysis outputs")
    return json.loads(config.SUMMARY_PATH.read_text(encoding="utf-8"))


@pytest.fixture(scope="session")
def tables() -> dict[str, pd.DataFrame]:
    require(config.ANALYSIS_DIR / "kpi_overview.csv", "analysis outputs")
    return {name: pd.read_csv(config.ANALYSIS_DIR / f"{name}.csv") for name in ANALYSIS_TABLES}


@pytest.fixture(scope="session")
def transactions() -> pd.DataFrame:
    require(config.FULL_DATA_PATH, "generated dataset")
    return pd.read_csv(config.FULL_DATA_PATH)


@pytest.fixture(scope="session")
def workbook() -> Workbook:
    require(config.WORKBOOK_PATH, "Excel workbook")
    return Workbook(config.WORKBOOK_PATH)
