"""One-command reproduction of the whole project.

    python src/run_pipeline.py

Runs, in order: data generation -> data validation -> SQLite build -> SQL and
Python analysis -> Excel workbook -> dashboard preview -> test suite. Every step
is a separate process, so a failure stops the pipeline with the exit code and
output of the step that failed.

The preview step drives Excel through COM and therefore only runs on Windows
with Excel installed; everywhere else it is skipped with a clear message and the
committed PNG is left untouched. Nothing else in the pipeline needs Excel.
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import time

from config import DB_PATH, FULL_DATA_PATH, PREVIEW_PATH, ROOT, WORKBOOK_PATH, configure_logging

logger = configure_logging()
SRC = ROOT / "src"


def run(command: list[str], label: str) -> float:
    """Run one pipeline step, streaming its output; abort the run if it fails."""
    logger.info("-" * 72)
    logger.info("STEP: %s", label)
    started = time.perf_counter()
    result = subprocess.run(command, cwd=ROOT)
    elapsed = time.perf_counter() - started
    if result.returncode != 0:
        logger.error("STEP FAILED: %s (exit code %d)", label, result.returncode)
        raise SystemExit(result.returncode)
    logger.info("STEP OK: %s in %.1fs", label, elapsed)
    return elapsed


def preview_available() -> bool:
    """True when the Excel-based renderer can run on this machine."""
    return sys.platform == "win32" and shutil.which("powershell") is not None


def main() -> None:
    parser = argparse.ArgumentParser(description="Run the full PayPulse analytics pipeline.")
    parser.add_argument("--rows", type=int, default=None, help="override the number of generated transactions")
    parser.add_argument("--skip-generate", action="store_true", help="reuse the existing generated dataset")
    parser.add_argument("--skip-preview", action="store_true", help="do not rebuild the dashboard preview image")
    parser.add_argument("--skip-tests", action="store_true", help="do not run the test suite")
    args = parser.parse_args()

    python = sys.executable
    started = time.perf_counter()

    if args.skip_generate:
        if not FULL_DATA_PATH.exists():
            raise SystemExit(f"--skip-generate needs an existing dataset at {FULL_DATA_PATH}")
        logger.info("STEP SKIPPED: data generation (reusing %s)", FULL_DATA_PATH.name)
    else:
        command = [python, str(SRC / "generate_data.py")]
        if args.rows:
            command += ["--rows", str(args.rows)]
        run(command, "Generate synthetic transactions")

    run([python, str(SRC / "validate_data.py")], "Validate data quality")
    run([python, str(SRC / "build_database.py")], "Build SQLite database")
    run([python, str(SRC / "run_analysis.py")], "Run SQL analysis and export outputs")
    run([python, str(SRC / "build_excel_dashboard.py")], "Build Excel workbook")

    if args.skip_preview:
        logger.info("STEP SKIPPED: dashboard preview (--skip-preview)")
    elif preview_available():
        run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
             str(ROOT / "scripts" / "excel_qa.ps1")], "Verify workbook in Excel and render preview")
    else:
        logger.warning("STEP SKIPPED: dashboard preview needs Windows with Excel installed. "
                       "The committed %s is unchanged; every other output was rebuilt.", PREVIEW_PATH.name)

    if args.skip_tests:
        logger.info("STEP SKIPPED: test suite (--skip-tests)")
    else:
        run([python, "-m", "pytest", "-q", str(ROOT / "tests")], "Run test suite")

    logger.info("-" * 72)
    logger.info("PIPELINE COMPLETE in %.1fs", time.perf_counter() - started)
    for path in (FULL_DATA_PATH, DB_PATH, WORKBOOK_PATH, PREVIEW_PATH):
        status = "ok " if path.exists() else "MISSING"
        size = f"{path.stat().st_size / 1024 / 1024:.1f} MB" if path.exists() else "-"
        logger.info("  %s %-46s %s", status, path.relative_to(ROOT).as_posix(), size)


if __name__ == "__main__":
    main()
