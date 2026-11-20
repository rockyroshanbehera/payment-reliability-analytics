# Payment reliability: where we are losing transactions, and what it is worth to fix

**PayPulse · Reporting period 1 January – 30 June 2025 · Prepared for the payments leadership team**

> **This report is built on synthetic data.** PayPulse is a fictional company created for a portfolio project. Every transaction, bank, merchant, customer, incident and impact figure below is simulated. Nothing here describes PhonePe or any other real organisation, and no stated impact has actually occurred.

---

## The business problem

We process a large volume of small payments, and we only earn when a payment succeeds. Over the last six months **4.43% of payment attempts failed — 7,981 payments worth ₹5.5M of attempted value.**

The headline success rate has been stable at about 95.6%, which is why this has not previously been escalated. Stability, however, is not the same as health. A flat average can hide a serious problem in one part of the business, and this analysis was commissioned to answer three questions: where are we actually losing payments, how much of that loss can we influence, and what would it be worth to address it?

## Where the numbers stand

| Measure | Six-month result | What it means |
|---|---|---|
| Payments attempted | 180,000 (₹123.9M) | The base we are measured on |
| Payments succeeded | 172,019 (₹118.4M) | 95.57% of attempts, 95.53% of value |
| Payments failed | 7,981 (₹5.5M) | The addressable problem |
| Slowest 5% of payments | Over 3.44 seconds | Against a 1.71-second average |
| Failed payments we retry | 39% | Of which 44% eventually succeed |

**Movement over the period was minimal**: monthly success rate stayed inside a 0.36 percentage-point band (95.37% in March, the weakest month, to 95.72% in February). No trend needs explaining — but that flat line is exactly what concealed the findings below.

## What is actually causing the failures

**Two causes carry half of everything.** Bank Timeout (2,195 failures, 27.5%) and Network Error (1,771, 22.2%) together account for **49.7% of all failed payments**. Both are infrastructure problems we can influence. The next largest cause, Insufficient Funds (17.6%), is a customer's account balance and is not ours to fix — an important distinction when deciding where to spend engineering time.

**Failures concentrate in the evening.** The 18:00–22:59 window carries 38.6% of our volume but **44.8% of our failures**: 94.85% success against 96.02% for the rest of the day. The composition changes too — timeouts make up 33.5% of evening failures against 22.6% off-peak, and **55% of all our timeout failures happen in those five hours**. That is the signature of systems under load, not of customers behaving differently after work.

**Connection quality matters more than we assume.** Payments made on 2G succeed 92.44% of the time and on 3G 94.42%, against 95.86% on 4G. This is a smaller pool but a large per-customer effect, and it maps directly onto the Network Error failures.

**One issuing bank had a real, contained outage.** Between 10 and 20 March, Crescent Bank's evening success rate fell to **88.74%**, while every other bank in the same hours held at 95.08%. Its timeout rate went from 1.45% to 5.93% and back to 1.50% afterwards. Nothing changed on our side, and the effect ended when the window ended.

## How we chose what to recommend

Two disciplines shaped the recommendations, and both are worth stating because they changed the answer.

**We prioritised by volume, not by rate.** The worst success rate in the entire dataset belongs to 2G connections at 92.44% — but that is 5,383 attempts. Bank Timeout has a milder rate effect and 2,195 actual failed payments. Fixing the pool with the most failures in it beats fixing the pool with the worst percentage.

**We checked whether apparent differences are real.** Ranked over six months, our six issuing banks sit between 95.44% and 95.72%. That 0.28-point spread is smaller than the statistical uncertainty on the smallest bank (±0.27 points), so **the bank league table is not a finding and should not drive action** — the lowest-ranked bank on that table is not the bank that had the outage. The Crescent Bank incident only became visible when we cut the data by bank *and* hour *and* date at the same time. A monitoring approach based on daily averages would have missed it entirely; the portfolio average moved less than one point during the outage.

## Three recommended actions

**1. Cap and circuit-break slow issuing-bank calls during the evening peak.**
Targets the single largest addressable pool: 2,195 timeout failures, 55% of which land in the 18:00–22:59 window. When a bank stops responding within budget, we should fail fast and route around it rather than holding the customer in a spinner.

**2. Harden the payment path for weak connections.**
Lighter request payloads, longer client-side timeouts and offline confirmation for 2G/3G sessions. This addresses the 1,771 Network Error failures and improves the experience for the customers with the least margin for error.

**3. Make the retry policy cause-aware.**
We currently retry roughly 38–42% of failures regardless of cause, but recovery is wildly uneven: timeouts recover 53% of the time, while Insufficient Funds recovers 27% and Risk Decline 25%. We should retry infrastructure faults harder, stop retrying customer-side declines, and suppress retries to a bank that is in an active incident.

**Supporting change:** alert on issuing-bank success rate by hour, not on the daily platform average. That single change is what would have surfaced the March incident on day one instead of in a six-month review.

## What this could be worth

If the two infrastructure causes are reduced and retry recovery improves on the remaining causes:

| Scenario | Assumptions | Additional successful payments | Payment value recovered | Success rate |
|---|---|---:|---:|---|
| Conservative | −10% timeouts, −10% network errors, +5% retry recovery | 446 | ₹0.31M | 95.57% → 95.81% |
| **Expected** | **−25%, −25%, +10%** | **1,089** | **₹0.75M** | **95.57% → 96.17%** |
| Ambitious | −40%, −40%, +20% | 1,782 | ₹1.22M | 95.57% → 96.56% |

**In the expected case that is roughly 1,089 payments and ₹0.75M of value over six months, and a success rate improvement of 0.61 percentage points.**

## Risks and limitations

- **These are scenario estimates, not forecasts.** They assume a prevented failure becomes a completed payment at the average value we already observe for that failure type. They do not model the engineering cost or capacity required, the risk of retrying more aggressively (including fraud and chargeback exposure), customers who simply pay another way, or interaction between the three changes.
- **The bank incident analysis is a comparison, not an experiment.** It is persuasive because the affected bank moved sharply while a control group of every other bank did not, and because the effect ended when the window ended — but we did not run a controlled test and we should not claim causation.
- **The success rate measures the initial attempt only.** Retry recoveries are reported separately rather than blended in, which keeps the number honest but makes it look slightly worse than the customer's lived experience.
- **Some segments are too small to act on.** Payments above ₹10,000 represent only 499 attempts; the rate there is flagged as low-sample and no recommendation rests on it.
- **The underlying data is synthetic.** This is a portfolio exercise. Real payment data is messier and would not produce a Pareto this clean.

## Suggested next experiment

**Run a two-week controlled test of the timeout circuit-breaker on the evening peak.**

Split evening traffic by issuing bank into a treatment group with an aggressive per-bank timeout budget and automatic circuit-breaking, and a control group on current behaviour. Track transaction success rate, P95 processing time, Bank Timeout volume and retry recovery rate for both groups.

That design is small, reversible, and tests the largest single assumption in the model above — that a capped call turns a timeout into a success rather than into a different failure. A 25% reduction in timeout failures in the treatment group would validate the expected case; no difference would tell us the failures are happening at the bank rather than in our call pattern, which is equally worth knowing before investing further.

---

*Prepared from a fully reproducible analysis pipeline. Underlying figures: [`outputs/analysis/summary.json`](../outputs/analysis/summary.json). Interactive version with editable assumptions: [`outputs/Payment_Reliability_Analytics_Dashboard.xlsx`](../outputs/Payment_Reliability_Analytics_Dashboard.xlsx). Synthetic data throughout.*
