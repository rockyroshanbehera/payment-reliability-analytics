"""Load the generated CSV into a portable SQLite analytical database.

The physical table stays exactly as generated; every derived business concept
(reporting date, week, hour, peak window, amount band, incident period) lives in
the `transaction_enriched` view. The view is rendered from `config.py`, so the
peak-hour and incident definitions used by SQL can never drift from the ones
used by the generator, the Python summary and the documentation.
"""
from __future__ import annotations

import sqlite3
from contextlib import closing

import pandas as pd

from config import (
    DB_PATH, FULL_DATA_PATH, GENERATED_DIR, INCIDENT_BANK, INCIDENT_END, INCIDENT_START,
    PEAK_END_HOUR, PEAK_START_HOUR, configure_logging,
)

logger = configure_logging()

INDEX_DDL = """
CREATE UNIQUE INDEX IF NOT EXISTS idx_txn_id ON transactions(transaction_id);
CREATE INDEX IF NOT EXISTS idx_txn_ts ON transactions(transaction_timestamp);
CREATE INDEX IF NOT EXISTS idx_txn_status ON transactions(transaction_status);
CREATE INDEX IF NOT EXISTS idx_txn_bank ON transactions(issuing_bank);
"""


def view_ddl() -> str:
    """Return the enriched-view DDL with business definitions injected from config."""
    return f"""
DROP VIEW IF EXISTS transaction_enriched;
CREATE VIEW transaction_enriched AS
SELECT *,
       date(transaction_timestamp) AS transaction_date,
       -- Monday-anchored reporting week, so weekly charts plot as real dates.
       date(transaction_timestamp, '-6 days', 'weekday 1') AS week_start_date,
       substr(transaction_timestamp, 1, 7) AS transaction_month,
       CAST(strftime('%H', transaction_timestamp) AS INTEGER) AS transaction_hour,
       -- Peak = the {PEAK_START_HOUR}:00-{PEAK_END_HOUR}:59 evening load window.
       CASE WHEN CAST(strftime('%H', transaction_timestamp) AS INTEGER)
                 BETWEEN {PEAK_START_HOUR} AND {PEAK_END_HOUR} THEN 'Peak' ELSE 'Non-peak' END AS peak_period,
       CASE WHEN amount_inr < 200 THEN 'Under INR 200'
            WHEN amount_inr < 1000 THEN 'INR 200-999'
            WHEN amount_inr < 5000 THEN 'INR 1,000-4,999'
            WHEN amount_inr < 10000 THEN 'INR 5,000-9,999'
            ELSE 'INR 10,000+' END AS amount_band,
       CASE WHEN amount_inr < 200 THEN 1 WHEN amount_inr < 1000 THEN 2
            WHEN amount_inr < 5000 THEN 3 WHEN amount_inr < 10000 THEN 4 ELSE 5 END AS amount_band_order,
       -- Incident window for the engineered {INCIDENT_BANK} outage.
       CASE WHEN transaction_timestamp < '{INCIDENT_START}' THEN 'Before'
            WHEN transaction_timestamp < '{INCIDENT_END}' THEN 'During'
            ELSE 'After' END AS incident_period,
       CASE WHEN issuing_bank = '{INCIDENT_BANK}' THEN '{INCIDENT_BANK}' ELSE 'All other banks' END AS bank_group
FROM transactions;
"""


def main() -> None:
    GENERATED_DIR.mkdir(parents=True, exist_ok=True)
    df = pd.read_csv(FULL_DATA_PATH)
    logger.info("Loading %s rows into %s", f"{len(df):,}", DB_PATH.name)
    with closing(sqlite3.connect(DB_PATH)) as conn:
        df.to_sql("transactions", conn, if_exists="replace", index=False, chunksize=20_000)
        conn.executescript(INDEX_DDL)
        conn.executescript(view_ddl())
        conn.commit()
        rows = conn.execute("SELECT COUNT(*) FROM transaction_enriched").fetchone()[0]
    logger.info("Built SQLite database with %s rows and the transaction_enriched view", f"{rows:,}")


if __name__ == "__main__":
    main()
