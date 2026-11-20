"""Single source of truth for paths, dataset parameters and business definitions.

Every script in this project imports its constants from here so that the
generator, the SQLite view, the SQL analysis, the Excel workbook and the
documentation can never drift apart. All data produced by this project is
synthetic; no real company, customer, merchant or bank is represented.
"""
from __future__ import annotations

import logging
from pathlib import Path

# --- Paths (all relative to the repository root, never machine-specific) ----
ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
GENERATED_DIR = DATA_DIR / "generated"
SQL_DIR = ROOT / "sql"
OUTPUT_DIR = ROOT / "outputs"
ANALYSIS_DIR = OUTPUT_DIR / "analysis"
WORK_DIR = ROOT / "work"

DB_PATH = GENERATED_DIR / "paypulse.sqlite"
FULL_DATA_PATH = GENERATED_DIR / "transactions.csv"
SAMPLE_PATH = DATA_DIR / "sample_transactions.csv"
WORKBOOK_PATH = OUTPUT_DIR / "Payment_Reliability_Analytics_Dashboard.xlsx"
PREVIEW_PATH = OUTPUT_DIR / "dashboard_preview.png"
SUMMARY_PATH = ANALYSIS_DIR / "summary.json"

# --- Dataset parameters ----------------------------------------------------
SEED = 20250813
N_TRANSACTIONS = 180_000
SAMPLE_ROWS = 5_000
START_DATE = "2025-01-01"
END_DATE = "2025-06-30"
COMPANY = "PayPulse"

# --- Categorical domains ---------------------------------------------------
PAYMENT_METHODS = ["UPI", "Debit Card", "Credit Card", "Wallet", "Net Banking"]
BANKS = ["Aurora Bank", "Bharat Bank", "Crescent Bank", "Deccan Bank", "Eastern Bank", "Frontier Bank"]
MERCHANT_CATEGORIES = ["Grocery", "Food Delivery", "E-commerce", "Fuel", "Travel", "Utilities", "Entertainment", "Healthcare"]
DEVICES = ["Android", "iOS", "Web"]
NETWORKS = ["WiFi", "5G", "4G", "3G", "2G"]
TENURE_GROUPS = ["<3 months", "3-12 months", "1-3 years", "3+ years"]
FAILURE_REASONS = ["Bank Timeout", "Network Error", "Insufficient Funds", "Bank Declined", "Technical Error", "Risk Decline", "Merchant Error"]
STATUSES = ["Success", "Failed"]
RETRY_OUTCOMES = ["Not Retried", "Recovered", "Still Failed"]
MAX_RETRY_ATTEMPTS = 3

# --- Business definitions (mirrored verbatim in sql/ and in the workbook) ---
PEAK_START_HOUR = 18   # inclusive
PEAK_END_HOUR = 22     # inclusive, so peak covers 18:00:00-22:59:59
INCIDENT_BANK = "Crescent Bank"
INCIDENT_START = "2025-03-10"   # inclusive
INCIDENT_END = "2025-03-21"     # exclusive, so the window is 10-20 March 2025

# Failure causes the remediation scenarios target. Both are infrastructure
# failures that a payments platform can realistically influence; declines such
# as Insufficient Funds are customer-side and deliberately excluded.
TARGET_FAILURE_REASONS = ["Bank Timeout", "Network Error"]

# Scenario assumptions used by the Python summary and by the Excel simulator.
# Values are fractions of the relevant baseline pool.
SCENARIOS: dict[str, dict[str, float]] = {
    "Conservative": {"timeout_reduction": 0.10, "network_reduction": 0.10, "retry_uplift": 0.05},
    "Expected": {"timeout_reduction": 0.25, "network_reduction": 0.25, "retry_uplift": 0.10},
    "Ambitious": {"timeout_reduction": 0.40, "network_reduction": 0.40, "retry_uplift": 0.20},
}
HEADLINE_SCENARIO = "Expected"

# Plausibility guardrails asserted by validate_data.py and the test suite.
MIN_SUCCESS_RATE = 0.90
MAX_SUCCESS_RATE = 0.99

DISCLAIMER = (
    "Synthetic data. PayPulse is a fictional company; all transactions, banks, "
    "merchants, customers, incidents and impact estimates are simulated for "
    "portfolio purposes and do not represent any real organisation."
)


def configure_logging(level: int = logging.INFO) -> logging.Logger:
    """Configure root logging once and return the project logger."""
    logging.basicConfig(level=level, format="%(asctime)s | %(levelname)-7s | %(message)s", datefmt="%H:%M:%S")
    return logging.getLogger("paypulse")
