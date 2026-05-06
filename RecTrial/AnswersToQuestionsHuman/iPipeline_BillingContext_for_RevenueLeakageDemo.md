# iPipeline Billing Context — Reference for Revenue Leakage Finder Demo

**Purpose:** Source material for building a synthetic data file (CSV) for the Revenue Leakage Finder Python demo. Pulled from the AlphaTrust billing template rebuild project files, prior session summaries, and direct inspection of `ATCustVolApril2026.xlsx`.

**Date compiled:** April 2026

---

## 1. Contract structure — what an iPipeline customer contract actually looks like

The product being billed is **AlphaTrust** (an iPipeline product family — e-signature / digital workflow for insurance and financial services). Contracts are tracked in a master file the team calls **CustVol** (`ATCustVolApril2026.xlsx`), with two sheets: `Customers` (contract metadata, 1 row per customer-deployment) and `Volumes` (transactional polling data, ~10,245 historical rows).

**Pricing model is transaction-based with banded overage tiers — not per-seat.** A contract has:

- **Base Fee** — flat annual (or per-period) fee. Sample distribution from the file: min $0, p25 $10K, median $15K, p75 $46.75K, max ~$333K. The very high values reflect a small number of large enterprise deals; the bulk of the customer base sits in $10K–$50K territory.
- **Base Quantity** — number of transactions included in the Base Fee. Sample: min 0, p25 7,000, median ~14,500, p75 85,000, max ~2,000,000. Three customers have Base Quantity = 0 (NY Life SaaS, Orthobanc OEM, ShareTec — flagged as data-quality issues, not real zero-quantity contracts).
- **Banded overage rates** — Band 1 through Band 10. Each band has a Lower Volume, Upper Volume, and Rate (USD per transaction). Band 1 Rate sample: min $0.12, median $1.41, max $2.00. Higher bands typically lower rate per transaction (volume discount). Most customers only use 1–3 bands; the 10-band schema accommodates the largest customers.
- **Transaction Multiplier** (CustVol col AZ) — a per-customer scalar applied to raw polled volume before billing. Purpose is undocumented (open question Q5/Q10 in the rebuild), values range roughly 0.2 to 5. Could plausibly model "this customer's transactions count as 0.2× because they're a downstream resold deal" or similar — but **don't claim a definitive interpretation in the demo**, because we don't have one.

**Billing cadence distribution from sampled rows:**

| Billing Basis | Approximate share |
|---|---|
| Annually | ~70% |
| Annually in arrears | ~7% |
| Monthly in arrears | ~8% |
| Monthly | ~5% |
| Quarterly in arrears | ~5% |

**The "annually in arrears" and "monthly in arrears" pattern matters for the demo** — overage is reconciled at the end of the period, not in real time. That's a structural reason a billed amount lags polled volume.

**Deployment types:** SaaS (~50%), On-Prem (~45%), Embedded (~5%). Each deployment is its own contract row — a single customer can have 2 or 3 deployments billed separately.

**Environments:** eSign, Pronto2, On-Prem, UK AWS. (Pronto2 and eSign are AlphaTrust product variants; UK AWS is the UK-region SaaS deployment.)

**Currency:** USD primarily, GBP for UK customers (Cirencester Friendly, The Cotswold Group, others — handful of accounts).

**Customer base profile:** insurance carriers and financial services firms dominate. Real names that have been worked through the rebuild include Prudential, Thrivent, Transamerica, Mutual of Omaha, Guardian, Lincoln, Principal, Pacific Life, Symetra, John Hancock, AIG, Allianz, ADP, plus mid-market brokers, healthcare orgs, and credit unions.

---

## 2. Legitimate reasons billed amount differs from contract amount

Drawn from what's actually structural in this billing model, not generic SaaS examples:

- **Overage on banded rates.** Customer exceeds Base Quantity → bills at Band 1 rate up to Band 1 upper, Band 2 rate above that, etc. Most common legitimate "billed > Base Fee" pattern. CustVol explicitly tracks `Overage (Transactions)` and `Overage (Revenue)` columns.
- **In-arrears reconciliation lag.** Annually-in-arrears and monthly-in-arrears contracts mean the volume billed in period N reflects period N-1 polling. A snapshot comparison at any single date will show "drift" that's actually just the lag.
- **Transaction Multiplier scaling.** Polled volume × multiplier = billed transactions. A customer whose polled volume is 100K but multiplier is 0.5 legitimately bills as 50K transactions.
- **Mid-term contract changes.** The CustVol structure has `Current Term Start`/`Current Term End` separate from `Effective Date`/`Expiry Date` — implies that a customer can renew or amend mid-fiscal-year, with the active term replacing the prior one. The rebuild explicitly does NOT model carryover-vs-reset of unused allotment at renewal (open question Q6/Q11), so legitimate variance arises here.
- **Annualization adjustments.** The Customers sheet has explicit `Volume (Annualized)` and `Expected Revenue (Annualized)` vs `(De-Annualized)` columns — meaning short partial periods are scaled to a full-year-equivalent for analysis, then de-annualized for actual billing. A snapshot mid-period legitimately differs from the annualized view.
- **Renewal ARR opportunity.** Contract has `Renewal ARR Opportunity Estimate (Lower Bound)` and `Upper Bound` columns — explicit acknowledgment that renewal pricing is negotiated within a range, not fixed. Lower bound sample max $500K, upper bound sample max $1M.

**What is NOT in our context:** explicit ramp pricing, trial periods, seat-count changes, or per-seat pricing. iPipeline's AlphaTrust pricing is transaction-volume based, not seat-based. The demo can include those concepts if you want general "revenue leakage" coverage, but **don't pretend they came from this dataset** — they didn't.

---

## 3. Real billing exceptions vs. theoretical ones

The rebuild surfaced these **real, documented** exception classes (these are what your demo should lean into for authenticity):

**Revenue leakage of the "polling without billing" type — the dominant finding:**

- **81 active customers polling into Volumes with no contract row in Customers.** Includes Prudential, Thrivent, Transamerica, Mutual of Omaha, Guardian, Lincoln, Principal, Pacific Life, Symetra, John Hancock, and ~70 others. These customers are using AlphaTrust today; whether they're being invoiced through some other system or genuinely going un-billed is the open question driving the rebuild. **This is the strongest "revenue leakage" signal in the actual dataset.**
- **38 active customers with expired Term End dates per CustVol** — most likely renewed but the contract record was never updated. Polling continues at old contract terms; new contract terms (probably negotiated at higher prices) aren't being applied. This is structurally **underbilling at the rate level** even when invoicing happens.
- **3 customers with Base Quantity = 0** (NY Life SaaS, Orthobanc OEM, ShareTec) — likely a missed contract field, not a real zero allotment. If the bill is calculated from Base Fee + (volume - Base Quantity) × rate, a zero Base Quantity bills the customer for every single transaction at overage rates. **Strong overbilling candidate** if the field is wrong, or strong underbilling if Base Fee is also missing.
- **ShareTec** — billing tab exists, both Term Start and Base Quantity blank.
- **The Standard** — billing tab has Base Quantity = 5000 but Term Start blank.

**Data quality drift exceptions:**

- 2,699 cells in Volumes column G (Polling Count) contain literal text `"NULL"` instead of empty/zero — 26% of all data rows. SUMIFS skips these naturally, but a less defensive analysis silently undercounts.
- 14 non-numeric values in Volumes column H (the canonical Volume column): 10 × `" -   "`, 3 × `"-"`, 1 × `"c"`.
- Row 79 in Customers contains literal `"x"` in customer/deployment/term fields — leftover footer marker.
- Row 37 of Customers is a byte-for-byte duplicate of row 36 (LandrumHR) — would cause double-billing if processed naively.
- 4 rows in Volumes for `&Partners` (likely internal); 37 rows where customer name is the literal string `"NULL"`.
- Internal/test entries mixed into customer data: `Amica QA`, `iPipeline Test Supplier`, `iPipeline Finance - Locked`, `iPipeline FP&A – Confidential` (13 known internal/test entries total).
- Name drift across systems — the same customer appears as `Sharetec` (Volumes) vs `ShareTec` (billing tab) vs `ShareTec (Bradford-Scott, NDS, GBS, DST)` (Customers sheet). Casing and parenthetical-tag variants. 14 confirmed alias mappings already documented; ~63 total reconciliation rows in the Name Reconciliation sheet.

**Theoretical for this dataset (don't claim these as real):** duplicate invoices, missing-invoice-entirely-with-active-service (the polling-without-contract finding above is the closest analog), wrong-product-billed (could happen with deployment-type confusion but no recorded instance), explicit double-billing.

---

## 4. Data structure specifics that make synthetic data look authentic

**The CustVol Customers sheet has 200 columns.** You don't need all of them. The shape that matters for a Revenue Leakage Finder demo:

Core contract fields (columns A–Q in CustVol):

```
Customer | Status | Deployment Type | Environment | Location |
Last Polling Date | Up to Date | Effective Date | Expiry Date |
Renewal | Current Term Start | Current Term End | Renewal |
Next True Up | Notice | Billing Basis
```

Pricing fields (columns AT–AW + DJ–EO):

```
Base Quantity | Base Fee | Base Fee Per Transaction | Band 1 Rate
... Band 1 Volume Lower | Band 1 Volume Upper | Band 1 Rate ...
... up through Band 10 ...
```

Analysis fields (revenue-related, columns AI–AS):

```
Volume | Expected Revenue (Trailing 12 Months; USD) |
Average Revenue Per Transaction | Marginal Revenue Per Transaction |
Overage (Transactions) | Overage (Revenue) |
Price Book Revenue Per Transaction | % of Price Book |
$ under Price Book |
Renewal ARR Opportunity Estimate (Price Book Δ) |
Renewal ARR Opportunity Estimate (over & above Price Book)
```

The `% of Price Book` and `$ under Price Book` columns are **strong candidates for your demo's leakage signal** — they explicitly track "what should this customer be paying vs. what they actually are paying" within the existing dataset.

**Volumes sheet structure (transactional polling table):**

```
Customer | Deployment Type | Environment | License Type |
Effective Date | Polling Date | Polling Count | Volume | Note
```

Per-poll Volume values sampled: min 1, p25 124, median ~1,400, p75 ~4,500, max ~307,000 (skewed — a few enterprise customers contribute the bulk of transaction volume).

**Status values:** `Active`, `Inactive`, `Internal`. (Plus the artifact `"x"` from row 79 — don't include in synthetic data.)

**No native "Customer ID" format exists in CustVol** — Customer Name is the join key. The billing template's K6 cell holds an internal sequential ID assigned during the rebuild, but that's local to the workbook, not a system-wide ID. **For the demo, invent a format.** Suggested: `ATC-NNNNN` (AlphaTrust Customer + 5-digit sequence), or use the customer's truncated name like the actual workbook does. Flag this as a synthetic-only construct.

**File format:** CSV is fine for the demo. The real data is XLSX, but billing exports for analysis would plausibly come out as CSV.

---

## 5. Volume and scale

| Metric | Real value |
|---|---|
| External customer-deployment combos | 175 (135 Active + 28 Stale + 12 Dormant) |
| Internal/test entries (filtered out) | 13 |
| Total customer-deployment rows in Customers sheet | ~75 (the gap vs 175 IS the 81-customer leakage finding) |
| Unique customers (one customer can have 2–3 deployments) | ~120 unique customers |
| Polling rows in Volumes (multi-year history) | 10,245 |
| Polls per customer-deployment per month | Roughly 1–4 (varies by environment) |

For your synthetic file, "a few hundred rows" is realistic for the **contract list** (target ~150–200 rows). For the **polling/usage history**, you want **thousands** if you're going to model multi-year multi-customer behavior — figure 5,000–15,000 rows depending on how much history you want to simulate. **For a focused leakage demo**, a 6-month or 12-month window is enough: roughly 175 customers × 4 polls/month × 12 months = ~8,400 polling rows, which is roughly the right shape.

For invoices: most contracts are annual, so most customers generate ~1 base invoice/year + 1 overage true-up invoice/year. Monthly-in-arrears customers generate 12 invoices/year. Reasonable target: ~250–400 invoices for a 12-month window across ~150 customers, with deliberate "missing invoice" gaps for your leakage scenarios.

---

## 6. Gaps — things you'll need to invent for the demo, flagged honestly

What our shared context does **not** cover, despite being on the original list:

- **MRR, specifically.** iPipeline's pricing is annual recurring or annual-with-overage, not monthly recurring. There's no "MRR" concept in CustVol — the analog is `Expected Revenue (Trailing 12 Months; USD)`, sample max ~$174K. If the demo needs MRR, divide ARR by 12 and call it that, but know it's a synthetic framing on top of an ARR-based reality.
- **Per-seat pricing.** Doesn't exist in this dataset. Don't claim it does.
- **Trial periods, ramp pricing, mid-year upgrade/downgrade pricing rules.** Not documented in our context. Plausible inventions for a demo, but flag them in code comments as "synthetic, not from real iPipeline data."
- **Discount mechanisms.** Not explicit in CustVol. The closest thing is the `% of Price Book` field (where a customer's actual rate sits below standard rate book) — that IS a real concept in this dataset and worth using in the demo.
- **Customer ID format.** No native one. Invent and flag.
- **Invoice-level data.** CustVol has expected revenue and overage estimates, but the actual **invoice ledger** (what was billed when, by line item) is downstream of CustVol and has not been seen in this project. If the demo compares invoices to contracts, the invoice file is fully synthetic.
- **Duplicate invoice / wrong product billed scenarios.** Theoretical for this dataset. Include them if the demo's scope is broader than the actual iPipeline context, but don't pretend they're documented findings.
- **The Transaction Multiplier semantics.** Confirmed real, purpose unconfirmed. Use the field in synthetic data, but don't overclaim its meaning.

---

## 7. One framing note on demo strategy

The single strongest "revenue leakage" finding in the actual iPipeline context is **upstream of invoicing** — it's the 81 customers polling without any contract record, and the 38 customers with expired terms still active. A typical Revenue Leakage Finder targets **invoice-vs-contract drift** (the customer was billed wrong). The iPipeline reality is more about **contract-vs-reality drift** (the customer doesn't have a contract record at all, or has a stale one).

If the demo is meant to feel authentic to this domain, lead with the contract-coverage class of leakage and treat invoice-line-item drift as a secondary class. If the demo is generic-SaaS-shaped, that's fine — but the iPipeline-flavored signal you'd be ignoring is the more interesting one.
