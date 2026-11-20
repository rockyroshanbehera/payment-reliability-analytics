"""Generate the deterministic, entirely synthetic PayPulse transaction dataset.

The generator is seeded, so the same command always produces the same 180,000
rows and therefore the same findings, workbook and documentation. Failure
probability is built from small additive drivers (peak load, poor mobile
network, high value, device, payment method) plus one deliberately engineered,
time-bounded outage for a single fictional bank. Those drivers are the patterns
the SQL analysis is designed to recover.
"""
from __future__ import annotations

import argparse

import numpy as np
import pandas as pd

from config import (
    BANKS, DEVICES, END_DATE, FAILURE_REASONS, FULL_DATA_PATH, GENERATED_DIR,
    INCIDENT_BANK, INCIDENT_END, INCIDENT_START, MERCHANT_CATEGORIES,
    N_TRANSACTIONS, NETWORKS, PAYMENT_METHODS, PEAK_END_HOUR, PEAK_START_HOUR,
    SAMPLE_PATH, SAMPLE_ROWS, SEED, START_DATE, TENURE_GROUPS, configure_logging,
)

logger = configure_logging()

# Fictional but geographically plausible Indian metro/tier-2 mix.
CITY_DATA = [
    ("Mumbai", "Maharashtra", "West"), ("Pune", "Maharashtra", "West"),
    ("Delhi", "Delhi", "North"), ("Jaipur", "Rajasthan", "North"),
    ("Bengaluru", "Karnataka", "South"), ("Chennai", "Tamil Nadu", "South"),
    ("Hyderabad", "Telangana", "South"), ("Kolkata", "West Bengal", "East"),
    ("Bhubaneswar", "Odisha", "East"), ("Lucknow", "Uttar Pradesh", "North"),
    ("Ahmedabad", "Gujarat", "West"), ("Guwahati", "Assam", "Northeast"),
]

# Relative transaction volume by hour: a morning commute bump and a large
# 18:00-22:00 evening peak, which is where reliability is stressed.
HOUR_WEIGHTS = np.array([1, .7, .5, .4, .4, .6, 1, 1.8, 2.7, 2.5, 2.2, 2.2,
                         2.4, 2.3, 2.1, 2.0, 2.2, 3.0, 4.3, 4.8, 4.6, 3.7, 2.6, 1.7])

BASE_FAILURE_RATE = .034          # baseline probability an attempt fails
PEAK_FAILURE_UPLIFT = .010        # extra failure probability during 18:00-22:59
NETWORK_2G_UPLIFT = .040          # extra failure probability on 2G sessions
NETWORK_3G_UPLIFT = .018          # extra failure probability on 3G sessions
INCIDENT_UPLIFT = .075            # engineered outage uplift (one bank, peak, 11 days)
ANDROID_COHORT_UPLIFT = .010      # applies to a random 22% Android sub-cohort
HIGH_VALUE_UPLIFT = .012          # attempts of INR 10,000 or more
NET_BANKING_UPLIFT = .006         # redirect-based flow is less reliable
RETRY_PROPENSITY = .39            # share of failed attempts that are retried
BASE_RECOVERY_RATE = .48          # baseline probability a retry recovers


def generate(n: int = N_TRANSACTIONS, seed: int = SEED) -> pd.DataFrame:
    """Return a reproducible synthetic transaction dataset with engineered patterns."""
    rng = np.random.default_rng(seed)

    # --- When did the attempt happen? ---------------------------------------
    dates = pd.date_range(START_DATE, END_DATE, freq="D")
    chosen_dates = pd.to_datetime(rng.choice(dates.values, n))
    hours = rng.choice(24, n, p=HOUR_WEIGHTS / HOUR_WEIGHTS.sum())
    timestamps = chosen_dates + pd.to_timedelta(hours, unit="h") + pd.to_timedelta(rng.integers(0, 3600, n), unit="s")

    # --- Who, where, how ----------------------------------------------------
    city_idx = rng.choice(len(CITY_DATA), n, p=[.15, .07, .13, .05, .13, .09, .09, .08, .04, .05, .08, .04])
    cities = np.array([x[0] for x in CITY_DATA])[city_idx]
    states = np.array([x[1] for x in CITY_DATA])[city_idx]
    regions = np.array([x[2] for x in CITY_DATA])[city_idx]
    methods = rng.choice(PAYMENT_METHODS, n, p=[.54, .16, .12, .10, .08])
    banks = rng.choice(BANKS, n, p=[.21, .19, .18, .16, .14, .12])
    categories = rng.choice(MERCHANT_CATEGORIES, n, p=[.18, .16, .19, .10, .10, .11, .09, .07])
    devices = rng.choice(DEVICES, n, p=[.68, .22, .10])
    networks = rng.choice(NETWORKS, n, p=[.22, .09, .54, .12, .03])
    tenure = rng.choice(TENURE_GROUPS, n, p=[.12, .27, .38, .23])
    # Long-tailed retail payment values: many small payments, few large ones.
    amount = np.clip(rng.lognormal(mean=5.75, sigma=1.25, size=n), 10, 75000).round(2)

    # --- Additive failure drivers -------------------------------------------
    failure_p = np.full(n, BASE_FAILURE_RATE)
    peak = (hours >= PEAK_START_HOUR) & (hours <= PEAK_END_HOUR)
    failure_p += peak * PEAK_FAILURE_UPLIFT
    failure_p += np.select([networks == "2G", networks == "3G"], [NETWORK_2G_UPLIFT, NETWORK_3G_UPLIFT], default=0)
    incident = (banks == INCIDENT_BANK) & (timestamps >= INCIDENT_START) & (timestamps < INCIDENT_END) & peak
    failure_p += incident * INCIDENT_UPLIFT
    failure_p += ((devices == "Android") & (rng.random(n) < .22)) * ANDROID_COHORT_UPLIFT
    failure_p += (amount >= 10000) * HIGH_VALUE_UPLIFT
    failure_p += (methods == "Net Banking") * NET_BANKING_UPLIFT
    failed = rng.random(n) < failure_p

    # --- Root cause of each failed attempt ----------------------------------
    # Base mix is tilted by the same drivers, so causes stay diagnosable:
    # peak load produces timeouts, weak networks produce network errors.
    reasons = np.full(n, "", dtype=object)
    for i in np.where(failed)[0]:
        w = np.array([.25, .21, .20, .14, .10, .06, .04], dtype=float)
        if peak[i]: w[0] += .16
        if networks[i] in ("2G", "3G"): w[1] += .28
        if incident[i]: w[0] += .48
        if devices[i] == "Android": w[4] += .07
        if amount[i] >= 10000: w[3] += .07
        reasons[i] = rng.choice(FAILURE_REASONS, p=w / w.sum())

    # --- Retry behaviour ----------------------------------------------------
    # Only failed attempts are retried. Customer-side declines recover far less
    # often than infrastructure timeouts, which is what makes retry policy a
    # meaningful lever in the analysis.
    retry_flag = failed & (rng.random(n) < RETRY_PROPENSITY)
    retry_attempt = np.where(retry_flag, rng.choice([1, 2, 3], n, p=[.73, .21, .06]), 0)
    recovery_p = (BASE_RECOVERY_RATE
                  - (reasons == "Insufficient Funds") * .20
                  - (reasons == "Risk Decline") * .25
                  + (reasons == "Bank Timeout") * .08)
    recovered = retry_flag & (rng.random(n) < recovery_p)
    final_retry = np.where(~retry_flag, "Not Retried", np.where(recovered, "Recovered", "Still Failed"))

    # Failures are slower than successes; timeouts are slowest of all.
    processing = (rng.lognormal(.35, .45, n)
                  + failed * rng.lognormal(.55, .55, n)
                  + (reasons == "Bank Timeout") * rng.uniform(2, 7, n))

    return pd.DataFrame({
        "transaction_id": [f"TXN{i:09d}" for i in range(1, n + 1)],
        "transaction_timestamp": timestamps,
        "customer_id": [f"CUS{x:07d}" for x in rng.integers(1, 75001, n)],
        "merchant_id": [f"MER{x:05d}" for x in rng.integers(1, 6001, n)],
        "amount_inr": amount,
        "payment_method": methods,
        "issuing_bank": banks,
        "merchant_category": categories,
        "city": cities, "state": states, "region": regions,
        "device_type": devices, "network_type": networks,
        "customer_tenure_group": tenure,
        "transaction_status": np.where(failed, "Failed", "Success"),
        "failure_reason": reasons,
        "processing_time_seconds": processing.round(3),
        "retry_flag": retry_flag.astype(int),
        "retry_attempt": retry_attempt,
        "final_retry_outcome": final_retry,
    }).sort_values("transaction_timestamp").reset_index(drop=True)


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate synthetic PayPulse transactions.")
    parser.add_argument("--rows", type=int, default=N_TRANSACTIONS, help="number of transactions to generate")
    parser.add_argument("--seed", type=int, default=SEED, help="random seed controlling reproducibility")
    args = parser.parse_args()

    GENERATED_DIR.mkdir(parents=True, exist_ok=True)
    logger.info("Generating %s synthetic transactions (seed=%s)", f"{args.rows:,}", args.seed)
    df = generate(args.rows, args.seed)
    df.to_csv(FULL_DATA_PATH, index=False, date_format="%Y-%m-%d %H:%M:%S")
    sample = df.sample(min(SAMPLE_ROWS, len(df)), random_state=args.seed).sort_values("transaction_timestamp")
    sample.to_csv(SAMPLE_PATH, index=False, date_format="%Y-%m-%d %H:%M:%S")
    logger.info("Wrote %s rows to %s", f"{len(df):,}", FULL_DATA_PATH.relative_to(FULL_DATA_PATH.parents[2]))
    logger.info("Wrote %s-row public sample to %s", f"{len(sample):,}", SAMPLE_PATH.relative_to(SAMPLE_PATH.parents[1]))


if __name__ == "__main__":
    main()
