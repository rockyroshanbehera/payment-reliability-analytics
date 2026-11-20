# Data

**Everything in this folder is synthetic.** It is produced by [`../src/generate_data.py`](../src/generate_data.py), a seeded NumPy script. PayPulse is a fictional company. No transaction, customer, merchant, issuing bank or incident here corresponds to a real person, business or event, and **none of this data comes from PhonePe or any other real company**.

## What is committed, and what is not

| Path | Committed | Why |
|---|---|---|
| `sample_transactions.csv` | Yes | 5,000-row sample (~0.8 MB) so the schema and realism can be inspected without running anything. |
| `generated/transactions.csv` | No | ~30 MB and fully reproducible from the seed. Large generated files do not belong in Git. |
| `generated/paypulse.sqlite` | No | ~44 MB analytical database, rebuilt from the CSV in seconds. |

Both regenerated artefacts are listed in [`.gitignore`](../.gitignore). Recreate them with:

```bash
python src/run_pipeline.py
```

The aggregated results *are* committed, in [`../outputs/analysis/`](../outputs/analysis/) — 16 small CSV tables plus `summary.json`. That is the useful part to review, and it is what the workbook and the written report both read.

## Reproducibility

`SEED = 20250813` in [`../src/config.py`](../src/config.py) controls generation. The same seed always produces the same 180,000 rows, so every figure in the README, the executive summary and the Excel workbook is reproducible from a fresh clone. `tests/test_generator.py` asserts this determinism directly.

The committed sample is a deterministic 5,000-row draw from the same dataset (`random_state=SEED`), sorted by timestamp — it is a real slice of the full data, not separately generated, so aggregates computed from it are directionally comparable to the full run.

## Schema

20 columns, one row per payment attempt. Full definitions, types and analysis notes are in [`../docs/data_dictionary.md`](../docs/data_dictionary.md) and on the **Data Dictionary** sheet of the workbook.

Two conventions matter when reading the data:

- **`transaction_status` is the outcome of the initial attempt**, not the final outcome after retries. Retry results live in `final_retry_outcome`.
- **`failure_reason` is blank on successful rows by design.** That is not missing data, and the validation script excludes it from the completeness check for exactly this reason.

## Engineered patterns

The generator deliberately embeds signals so the analysis has a known ground truth to recover:

| Pattern | How it is built | Where it shows up |
|---|---|---|
| Evening peak stress | +1.0 pp failure probability during 18:00–22:59, and the peak tilts the failure mix towards Bank Timeout | Peak vs non-peak comparison, hourly profile |
| Weak mobile network | +4.0 pp on 2G, +1.8 pp on 3G, tilting the mix towards Network Error | Network segment scorecard |
| Bank incident | +7.5 pp for Crescent Bank, peak hours only, 10–20 March 2025, strongly tilted to Bank Timeout | Before/during/after incident comparison |
| High-value friction | +1.2 pp above ₹10,000 | Amount-band scorecard (flagged low-sample: 499 attempts) |
| Channel friction | +0.6 pp on Net Banking; +1.0 pp on a random 22% Android sub-cohort | Payment-method and device scorecards |
| Retry economics | Timeouts recover more often than declines; Insufficient Funds and Risk Decline recover far less | Retry analysis by cause |

Every other correlation in the data is coincidence — which is the point of reporting confidence intervals alongside the segment rates.
