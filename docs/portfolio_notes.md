# Portfolio notes

How to talk about this project in a CV, a LinkedIn profile and an interview.

> **Always lead with the disclosure.** This project uses synthetic data for a fictional company. Say so before anyone has to ask — claiming or implying real company data would be the fastest way to lose an interview. The verbs below are chosen deliberately: **identified, estimated, simulated, demonstrated** — never "delivered", "saved" or "increased", because no impact has actually occurred.

---

## Resume bullets

**1.**
> Built an end-to-end payment-reliability analytics pipeline (Python → SQLite → SQL → Excel) on a **synthetic 180,000-transaction dataset**, and **identified** that 2 of 7 failure causes drove **49.7% of all failures**, with reliability degrading **1.17 pp** during the 18:00–22:59 peak window that carried 38.6% of volume but 44.8% of failures.

**2.**
> **Demonstrated** statistical discipline by pairing every segment rate with a 95% confidence interval, showing that a **0.28 pp** spread across six issuing banks fell inside the **±0.27 pp** margin — then **identified** the genuine signal through a bank × hour × date cut that isolated a contained 11-day outage (**88.7% vs 95.1%** peak-hour success against a control group of all other banks).

**3.**
> **Simulated** three remediation scenarios in a formula-driven Excel model with editable, validated assumptions and mutually exclusive impact pools, **estimating** a potential **+0.61 pp** success-rate improvement (**1,089 transactions, ₹0.75M** of payment value) in the expected case — with the model's exclusions stated explicitly alongside the result.

*Optional fourth bullet if the role is engineering-leaning:*
> Made the analysis reproducible and self-verifying: one command rebuilds data, database, SQL outputs, workbook and preview; **52 automated tests** reconcile SQL results against independent pandas calculations and assert the delivered workbook's formulas and cached values match the analysis outputs.

## 60-second interview explanation

> "It's a payment-reliability analysis for a fictional payments company — the data is fully synthetic, generated from a seeded script, so everything reproduces exactly.
>
> The business question was simple: 95.6% of payments succeed, so where are we losing the other 4.4%, and how much of it can we actually fix? I built the pipeline end to end — a seeded generator, a validation gate, SQLite, six SQL files, and a nine-sheet Excel workbook.
>
> Three things came out of it. First, two of seven failure causes carry half of all failures, and both are infrastructure rather than customer-side — so that's where engineering time is worth spending. Second, reliability degrades in the evening peak: 38% of volume but 45% of failures, and timeouts jump from 23% to 34% of the failure mix, which points at capacity rather than customer behaviour.
>
> The third one is the one I actually care about. The bank league table showed a worst bank, but I put a confidence interval on every rate and the whole spread across six banks was inside the margin of error — so that ranking wasn't a finding. When I cut by bank and hour and date together, a completely different bank had dropped to 88.7% in evening hours for eleven days in March while every other bank was flat, and then recovered. That's a real contained incident, and it was invisible in the daily average.
>
> Then I put an Excel simulator on top so a stakeholder can change the assumptions themselves — expected case is about 1,089 recovered transactions and 0.6 percentage points — with an explicit list of what the model doesn't account for."

## Five likely interview questions

**1. "Your success rate is 95.57%, but you also say retries recover 44%. Which is the real number?"**
Both, and they answer different questions. `transaction_status` is the outcome of the initial attempt, so 95.57% is how often we succeed on first contact. Retry recovery is measured separately in `final_retry_outcome` — blending them in would raise the headline by about 0.77 pp and hide the thing worth managing. I report both because the gap between them *is* the retry policy's performance. If a stakeholder needs the customer-experienced rate, it's a defined, documented derivation, not a redefinition of the KPI.

**2. "You named a worst bank and then said not to act on it. Why report it at all?"**
Because the reader will compute it anyway, and it's better to show it with its uncertainty than to leave it to be discovered without. The spread across six banks is 0.28 pp and the 95% half-width on the smallest bank is ±0.27 pp — the intervals overlap, so the ranking isn't distinguishable from sampling noise. I drew that chart with error bars specifically so the overlap is visible rather than something you have to read a footnote to learn. The defensible finding needed a time-bounded cut, and that's the one I put a recommendation behind.

**3. "How do you know the March incident is real and not just a bad week?"**
Three things have to line up, and they do. The affected bank drops 6.8 pp in peak hours during the window; a control group of every other bank in the same hours moves 0.2 pp; and the effect reverses when the window closes. The mechanism is consistent too — its timeout rate goes 1.45% → 5.93% → 1.50%, so the failure type matches the story. That's a difference-in-differences *style* comparison, not a controlled experiment, and I say so — I'd want an A/B test before claiming causation.

**4. "Walk me through how you prioritised the recommendations."**
By addressable volume, not by rate. The worst rate in the dataset is 2G at 92.44%, but that's 5,383 attempts and 407 failures. Bank Timeout has a milder rate effect and 2,195 failures, 55% of them concentrated in five hours a day — a much larger pool with a clear intervention. I also split causes by who owns them: Insufficient Funds is 17.6% of failures and isn't an engineering problem, so it's excluded from the scenario model. Prioritisation is volume × addressability, not the worst number on the page.

**5. "Your impact number — how much would you defend it?"**
As a sized hypothesis, fully. As a forecast, not at all. It's arithmetic on a measured pool: 2,195 timeout failures × an assumed reduction, at that pool's observed average failed value, with three levers on mutually exclusive pools so nothing is double-counted. What it doesn't model is on the sheet in writing — cost, capacity, customer substitution, fraud exposure from retrying more, and interaction between the levers. That's why the assumptions are editable: the number's job is to tell you whether this is a ₹0.3M problem or a ₹3M problem so you know how much engineering time it justifies. Then I'd run the two-week circuit-breaker test to find out if the central assumption actually holds.

## LinkedIn project description

> **Payment Reliability Analytics — Transaction Success Rate & Failure Root-Cause Analysis**
>
> End-to-end analytics project on a **synthetic** 180,000-transaction payments dataset (fictional company; no real company data). SQL → Python → Excel.
>
> Identified that two of seven failure causes drove 49.7% of all failures, and that reliability degraded 1.17 pp during the evening peak window carrying 38.6% of volume but 44.8% of failures. Showed the six-month issuing-bank ranking was inside its own confidence interval and therefore not actionable, then isolated a genuine contained 11-day bank outage (88.7% vs 95.1% peak-hour success versus a control group) using a bank × hour × date cut.
>
> Built a nine-sheet Excel dashboard with an editable, formula-driven scenario simulator estimating a potential +0.61 pp success-rate improvement in the expected case, with model exclusions stated explicitly. The whole pipeline reproduces from one command and is covered by 52 automated tests that reconcile SQL outputs, Python calculations and the delivered workbook.
>
> SQL (CTEs, window functions, LAG, RANK, Pareto, cohort/event analysis) · Python (pandas, NumPy, XlsxWriter, pytest) · Excel dashboarding · Root-cause analysis · Scenario modelling

## What I would do with real production data

The analysis pattern would survive; the certainty would not. Concretely:

**Data engineering.** Real payment events arrive as an append-only stream with late-arriving and duplicated events, so the first job is idempotent ingestion and a deduplication key — not `COUNT(*)`. Attempt-level and order-level identities need separating, because "one payment attempt" and "one customer trying to pay" are different grains and the success rate means different things on each.

**Definitions before dashboards.** I'd expect a week of arguing about what counts as a failure: user-cancelled, timed-out-then-succeeded, pending-then-settled, refunded, reversed. Every number in this project rests on `transaction_status`, and in production that field would be a derived state machine, not a column. I'd write the definition down, get it signed off, and version it.

**Confounding is the real work.** Here I engineered the drivers, so recovering them is fair. In production, bank, network, device, amount and time of day are all correlated with each other and with customer segment. I'd move from single-variable scorecards to a logistic model or a properly stratified comparison, so "2G is worse" doesn't turn out to be "2G users buy cheaper things on older Android devices in specific regions".

**Monitoring rather than a report.** The core finding here is that the platform average hid an eleven-day outage. In production I'd ship the bank × hour cut as an automated alert with a control-group comparison and a volume threshold, so the incident pages someone on day one. A six-month retrospective that finds an incident is a failure of monitoring, however good the analysis is.

**Validating the impact model.** The scenario numbers would become a hypothesis to test, not a deliverable. The two-week circuit-breaker experiment described in the executive summary is exactly the first thing I'd run — it tests the assumption the whole model rests on: that a prevented timeout becomes a completed payment rather than a different failure.

**Cost and risk on the other side of the ledger.** Recovered payment value is only half of a decision. Real prioritisation needs engineering cost, capacity headroom, and the fraud and chargeback exposure of retrying more aggressively. I'd want the risk team's numbers next to mine before recommending anyone change retry policy.
