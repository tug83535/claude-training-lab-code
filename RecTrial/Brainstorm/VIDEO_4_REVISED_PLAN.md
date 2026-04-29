# Video 4 — Revised Production Plan

**Owner:** Connor Atlee — Finance & Accounting, iPipeline
**Date locked:** 2026-04-28
**Replaces:** Prior 4a+4b plan at `RecTrial\Brainstorm\VIDEO_4_CURRENT_PROPOSAL.md`
**Reference docs:** `VIDEO_4_REVIEW_DECISION_MEMO.md` (decisions), `SUPPORTED_WORKFLOWS_V1.md` (workflow detail)

---

## 1. Title and length target

**Official title:** Video 4 of 4 — Python Automation for Finance
**Target length:** 9–12 minutes
**Structure:** 8 chapters in a single video; no formal 4a/4b split
**Format:** Narrated screen recording (same approach as Videos 1–3 — ElevenLabs narration + screen capture)
**Audience:** 50–150 coworkers in Finance, Accounting, and adjacent operations at iPipeline. Non-developers, Excel-literate, zero Python exposure.
**Tone:** Plain English. Practical. "You can run this today" — not a research demo. Not a CFO pitch.

---

## 2. Eight-chapter outline with timing

| Chapter | Title | Target time |
|---|---|---|
| 1 | Why Python after Excel and VBA? | 45 sec |
| 2 | Safety first | 60 sec |
| 3 | Revenue Leakage Finder — hero demo | 2 min 30 sec – 3 min 30 sec |
| 4 | Data Contract Checker | 90 sec |
| 5 | Exception Triage Engine | 90 sec |
| 6 | Control Evidence Pack | 90 sec |
| 7 | Finance Automation Launcher | 60 sec |
| 8 | How to start | 30 sec |
| **Total** | | **9–12 min** |

---

## 3. Chapter-by-chapter demo beats

### Chapter 1 — Why Python after Excel and VBA? (45 sec)

**Message to land:**
- Excel and VBA are great for workbook-level automation — and you saw them in Videos 1–3.
- Python adds value for a different class of problems: multi-file workflows, data quality checks before analysis starts, repeatable structured reports from raw exports, and building evidence folders for audit/control work.
- This video shows 4 practical tools a Finance analyst can run without being a developer.

**On screen:** brief side-by-side text showing "Excel/VBA is for..." vs. "Python adds..." — can be a simple static slide or title card. No code on screen yet.

---

### Chapter 2 — Safety first (60 sec)

**Message to land:**
- Before any demo, state the ground rules clearly — not buried in a separate doc, but on screen while narration reads it.
- Scripts run on your local machine only. No internet, no AI, no external calls.
- Your input files are never changed. Everything goes to a separate output folder with a timestamp.
- A log is created for every run so you know exactly what happened.
- When something goes wrong, you get a plain error message — no cryptic Python stack trace on the main screen.

**On screen:** Show `PYTHON_SAFETY.md` open in Notepad or VS Code — a clean, readable file, not code. Scroll through the 14 rules while narration reads the key ones. Show the outputs folder path (`outputs/YYYYMMDD_HHMMSS_toolname/`).

**Why do this:** coworkers who have never run a Python script will be nervous. A 60-second "here is the safety contract" segment removes the hesitation. The fact that it's a visible, inspectable text file — not just a verbal promise — is what makes it credible.

---

### Chapter 3 — Revenue Leakage Finder — hero demo (2:30 – 3:30)

**The story:** "Python checks which customers are polling our system but don't have a matching contract on file — and which ones do have a contract but are being billed at old expired terms. That's the leakage."

**Demo sequence:**
1. Open Command Prompt. Navigate to `RecTrial\UniversalToolkit\python\ZeroInstall\`.
2. Run: `python revenue_leakage_finder.py --sample` — show it processing, show the output folder appear.
3. Open the HTML summary report in a browser. Show the headline numbers: total expected revenue vs. total billed revenue, net variance, exception count by type.
4. Highlight 2–3 specific exceptions on screen — a polling-without-contract customer, a stale-contract customer, a potential overbilling case.
5. Open the `top_10_action_list.csv` — show it ranked by severity. "This is where Finance needs to look first."
6. Closing line: "Run this against your own data with two CSV files — your contract list and your billing export."

**What's on screen during this chapter:** Command Prompt (brief) → browser with HTML report → two rows from exceptions CSV → ARR waterfall summary artifact (closing visual, ~20 sec).

**The ARR waterfall:** shown at the end of this chapter as a summary visual — "here's what the expected-vs-billed variance looks like aggregated by customer tier." It's the closing punctuation for the leakage story, not the main event. One screenshot or 20-second scroll — not a separate demo.

---

### Chapter 4 — Data Contract Checker (90 sec)

**The story:** "Before running any analysis, check that your input file has the right structure. A renamed column silently breaks every formula downstream. This tool catches it before the analysis starts."

**Demo sequence:**
1. Run `python data_contract_checker.py --sample --bad-file` — show red FAIL with specific error messages ("Missing required column: amount_billed", "invoice_date column has non-date values in 3 rows").
2. Show fixing one column in the CSV (rename it) — quick edit in Notepad.
3. Re-run → green PASS. Output folder with clean report.
4. "PASS means your file is safe to analyze. FAIL means fix the input first."

**What's on screen:** Command Prompt → red FAIL terminal output → Notepad (one edit) → green PASS → HTML report.

---

### Chapter 5 — Exception Triage Engine (90 sec)

**The story:** "Once you have a list of exceptions — billing mismatches, contract issues, data problems — Python can rank them so you know what to review first. Not all exceptions are equal."

**Demo sequence:**
1. Run `python exception_triage_engine.py --sample` — takes the Revenue Leakage output CSV as input.
2. Show ranked output: each exception gets a priority score based on dollar impact, confidence, and recency.
3. Open `top_10_action_list.csv` — "Row 1 is your highest-priority review. Each row tells you what it is, why it scored high, and what to check."

**What's on screen:** Command Prompt → terminal output → CSV ranked list open in Excel.

---

### Chapter 6 — Control Evidence Pack (90 sec)

**The story:** "After any significant analysis — especially one that may go to audit or leadership — create an evidence folder. Python logs exactly which files were analyzed, their hashes, and what the run produced. Repeatable and tamper-evident."

**Demo sequence:**
1. Run `python control_evidence_pack.py --sample --control-name "Revenue Leakage Review Q2 2026"` — show it scanning the outputs from the previous run.
2. Show the manifest: file names, sizes, timestamps, SHA-256 hashes.
3. Show the `evidence_summary.html` — one-page summary ready to attach to a ticket or email.
4. "If someone asks 'what files did you analyze and when?' — this folder answers that question."

**What's on screen:** Command Prompt → file list appearing → HTML evidence summary.

---

### Chapter 7 — Finance Automation Launcher (60 sec)

**The story:** "You don't have to remember command-line arguments. The launcher gives you a numbered menu."

**Demo sequence:**
1. Run `python finance_automation_launcher.py` — show the menu.
2. Pick option 1 (Revenue Leakage Finder, sample mode) — watch it run and show output folder path at the end.
3. Pick option 6 (Show safety rules) — the 14 rules print to screen.
4. Pick option 8 (Exit).

**What's on screen:** Command Prompt with menu → selection → output path printed → exit.

**Key line to narrate:** "Everything you just saw in the last four chapters is accessible from this one menu. Start here."

---

### Chapter 8 — How to start (30 sec)

**Message to land — four rules:**
1. Start with sample files. Don't run against real data until you understand the output.
2. Use the supported workflows first. They're the tested path.
3. Outputs go to the `/outputs/` folder. Never touches your input files.
4. Questions or something doesn't work — contact Connor.

**On screen:** Simple text card or title slide with the four rules. No demo.

---

## 4. Build effort estimate — Python scripts

| Script | Effort | Why |
|---|---|---|
| `common/safe_io.py` | S | Timestamped folder creation, read-only checks, CSV/JSON/HTML helpers — formulaic |
| `common/logging_utils.py` | S | Run start/end tracking, JSON + text log writers — straightforward |
| `common/report_utils.py` | S | Minimal HTML wrapper, table renderer, metric cards — small surface |
| `common/sample_data.py` | M | Must generate realistic AlphaTrust-shaped data (see Section 6). The realism bar is high; this takes care. |
| `data_contract_checker.py` | S | Schema validation: required columns, type checks, blank checks, business rules. Clean pattern. |
| `revenue_leakage_finder.py` | L | Core matching logic across 3 exception classes (polling-without-contract, stale contracts, invoice drift) + HTML report + ranked CSV + ARR waterfall artifact. Most complex script. |
| `exception_triage_engine.py` | M | Scoring formula, ranking, plain-English action lines. Logic is clear; output polish takes time. |
| `control_evidence_pack.py` | S | File hashing (SHA-256), manifest CSV, HTML summary. All stdlib. |
| `workbook_dependency_scanner.py` | M | zipfile + xml.etree parsing of .xlsx, regex formula extraction for cross-sheet refs. Edge cases take work. |
| `finance_automation_launcher.py` | S | Menu wrapper — built last after the 5 tools are stable. Minimal logic. |
| `smoke_test_video4_python.py` | S | Run each script in --sample mode, verify outputs exist. Straightforward. |
| `PYTHON_SAFETY.md` | Done | Already in this sprint. |
| `README_VIDEO4_PYTHON.md` | S | How to run, find outputs, troubleshoot. Write after scripts stabilize. |

**Total build estimate:** 5–8 days of focused work depending on who builds (Codex vs. Claude Code vs. both). The Codex build spec at `02_codex_build_spec.md` is detailed enough to hand over as-is. Build decision deferred until after all 5 sprint docs land.

**Recommended build order:** common utilities → sample_data.py → data_contract_checker.py → revenue_leakage_finder.py → exception_triage_engine.py → control_evidence_pack.py → workbook_dependency_scanner.py → finance_automation_launcher.py → docs → smoke test.

---

## 5. Recording effort estimate

| Task | Estimate |
|---|---|
| Write and record ElevenLabs narration clips (8 chapters) | 2–3 hours |
| Record screen capture (Command Prompt demos, browser HTML, CSV views) | 3–4 hours over 2–3 sessions |
| Edit and sync (clip playback + screen actions matched up) | 2–3 hours |
| Gemini review cycle (optional — V3 was reviewed, V4 may not need it) | 1–2 hours |
| **Total** | **~8–12 hours** |

Note: Video 4 is mostly Command Prompt + browser demos — simpler screen control than the Excel/VBA Videos 2–3 (no Director macro needed). Manual recording is appropriate.

---

## 6. Sample Data Design Lock — Revenue Leakage Finder

This section is the binding spec for the synthetic data files in `samples/revenue_leakage/`. **No V4 Python code is written until this section is read and the sample data generator (`sample_data.py`) is built to match it.**

### 6.1 Context — what the real data looks like

The product being modeled is **AlphaTrust** (iPipeline's e-signature / digital workflow product for insurance and financial services). The real contract file (`ATCustVolApril2026.xlsx`) has:
- ~175 customer-deployment rows (135 Active + 28 Stale + 12 Dormant), across ~120 unique customers
- Pricing is **transaction-based with banded overage** — not per-seat, not flat-rate SaaS
- Each contract has a Base Fee (flat), a Base Quantity (transactions included), and Bands 1–10 (overage rate tiers for transactions above Base Quantity)
- A **Transaction Multiplier** (range 0.2–5) scales polled volume before billing; exact purpose undocumented — include in synthetic data, keep semantics vague
- Billing cadence: ~70% annual, ~8% monthly in arrears, ~7% annually in arrears, ~5% quarterly
- Deployment types: SaaS (~50%), On-Prem (~45%), Embedded (~5%)
- Customer names come from insurance carriers and financial services firms (Prudential, Thrivent, Guardian, etc. in real data; synthetic data uses fictional equivalents)
- No native customer ID exists — invent `ATC-NNNNN` format for the demo and flag it as synthetic-only

### 6.2 Exception types — what to model (ranked by authenticity to the real dataset)

**Class 1 — Polling-without-contract (MOST AUTHENTIC — lead with this):**
The real dataset has 81 customers appearing in the polling/activity data with no matching contract row — the dominant leakage finding. In the demo, model this as: customers who show up in `billing.csv` (they're being invoiced or being polled) but have no row in `contracts.csv`. These are the most credible "you might be leaving money on the table" cases.

**Class 2 — Stale contract (VERY AUTHENTIC):**
38 real customers have expired Term End dates but are still polling. Model as: contracts with `term_end` in the past (expired 3–18 months ago) that still have active billing rows. This represents "we should have renegotiated this contract at a higher rate, but we haven't — so we're billing at a 2–3 year old price."

**Class 3 — Base Quantity = 0 anomaly (AUTHENTIC DATA QUALITY FINDING):**
3 real customers have Base Quantity = 0. If Base Quantity is 0, every transaction bills at overage rate — either a massive overbilling risk (if the zero was a data entry mistake) or a missing field (if the contract should have a free tier). Model 2–3 cases in the synthetic data.

**Class 4 — Invoice-to-expected-revenue drift (SECONDARY, but useful for demo variety):**
Differences between what the contract's `expected_annual_revenue` implies and what `amount_billed` shows in the billing export. Can be caused by: overage (legitimate), in-arrears reconciliation lag (legitimate), or billing mistake (leakage). Model a mix.

**Class 5 — Name drift / mapping gaps (DATA QUALITY):**
The real dataset has 14 confirmed customer alias mappings where the same customer appears under different name spellings across files (e.g., `Sharetec` in volumes vs `ShareTec` in billing). Model 3–5 cases where the customer name in `contracts.csv` doesn't exactly match `billing.csv` and the tool has to flag the ambiguous mapping.

**NOT modeled (don't claim these are real):** duplicate invoices, trial-period pricing, per-seat charges, explicit ramp pricing. Include at most 1–2 duplicate invoice cases as demo variety, but don't lead with them.

### 6.3 Synthetic customer names

Use fictional insurance/financial services firm names — do not use real iPipeline customer names. Suggested list:

```
Northstar Insurance Group         (large enterprise, SaaS, high volume)
Harbor Life Systems               (mid-market, On-Prem)
Pioneer Benefits Co.              (mid-market, SaaS)
Summit Policy Services            (small, Embedded)
Keystone Admin Solutions          (mid-market, On-Prem)
Atlas Brokerage Network           (mid-market, SaaS)
Meridian Life & Casualty          (large enterprise, SaaS + On-Prem deployments)
Lakewood Financial Group          (small, SaaS)
Crestview Insurance Partners      (mid-market, SaaS)
Ironbridge Benefits               (mid-market, On-Prem)
Stillwater Mutual                 (large enterprise, SaaS)
BluePeak Advisory                 (small, Embedded)
Ridgeline Health Plans            (mid-market, SaaS)
Coastal Re Solutions              (small, SaaS)
Pinnacle Trust Services           (large enterprise, On-Prem)
```

For the "polling without contract" leakage cases: add 10–15 customers who appear in `billing.csv` but NOT in `contracts.csv`. Name them similarly (e.g., `Westfield Annuity Group`, `Riverton Life Partners`, etc.).

### 6.4 Synthetic contracts.csv — column spec

```
customer_id         ATC-00001 through ATC-00NNN (synthetic, flag in code comments)
customer_name       fictional names from 6.3
deployment_type     SaaS | On-Prem | Embedded (50/45/5 distribution)
environment         eSign | Pronto2 | On-Prem | UK_AWS
billing_basis       Annually | Annually_in_arrears | Monthly_in_arrears | Monthly | Quarterly_in_arrears
status              Active | Stale | Dormant
base_fee            range $0 – $333K; typical $10K–$50K (p25 $10K, median $15K, p75 $46.75K)
base_quantity       range 0 – 2,000,000; typical 7,000–85,000 (p25 7K, median 14.5K, p75 85K)
band1_rate          range $0.12 – $2.00/transaction; median ~$1.41
transaction_multiplier  range 0.2 – 5.0 (include, flag as synthetic interpretation)
term_start          date; plausible 1–5 years ago
term_end            date; Active = future date; Stale = expired 3–18 months ago; Dormant = expired 18+ months ago
expected_annual_revenue  derived: base_fee + estimated overage (pre-calculated column, mimics CustVol)
currency            USD (except 2–3 UK accounts: GBP)
```

**Target rows:** 150 contracts (≈120 unique customers, some with 2–3 deployment rows)

### 6.5 Synthetic billing.csv — column spec

```
invoice_id          INV-2025-NNNN (sequential, year-stamped)
customer_id         matches contracts.csv ATC-NNNNN format (with deliberate nulls for leakage cases)
customer_name       matches contracts.csv (with deliberate name drift in 3–5 cases)
billing_period      YYYY-MM (year-month of the billing period)
amount_billed       USD; for Active/normal contracts ≈ (base_fee + overage) ± legitimate variance
invoice_date        date; typically 1–30 days into the following period (in-arrears pattern)
status              Paid | Outstanding | Disputed
notes               optional text field; can carry "Overage Reconciliation", "Annual True-Up", etc.
```

**Target rows:** 300–400 invoices (12-month window: ~150 contracts × 1 annual invoice OR 12 monthly invoices, with deliberate gaps for the leakage cases)

**Deliberate exception data to embed:**
- 12–15 customers in billing.csv with no matching contract_id in contracts.csv (Class 1 leakage)
- 10–12 customers with `term_end` expired in contracts.csv but active rows in billing.csv (Class 2)
- 2–3 contracts with `base_quantity = 0` (Class 3)
- 8–10 rows where `amount_billed` differs from `expected_annual_revenue` by more than 10% without overage explanation (Class 4)
- 3–5 name-drift cases where `customer_name` in billing.csv is spelled differently from contracts.csv (Class 5)
- 1–2 duplicate invoice rows (same invoice_id, different billing_period — late duplicate catch)

### 6.6 What NOT to model (synthetic-only flags)

The following concepts are **plausible for a generic SaaS demo but NOT from the real iPipeline dataset.** If they're included in the synthetic data for demo variety, the code must have a comment flagging them as synthetic constructs:
- Per-seat pricing or seat-count changes
- Trial period pricing or ramp pricing
- Discount mechanisms (except `% of Price Book` variance, which IS in the real data)
- Mid-year upgrade/downgrade pricing rules
- Explicit duplicate invoice detection (the LandrumHR duplicate is a data quality issue, not a billing process)

### 6.7 What makes this feel real, not toy

- Customer volume distribution is skewed: a few enterprise customers (Northstar, Stillwater, Meridian, Pinnacle) have base_fee $100K+ and base_quantity 500K+; the bulk of accounts are $10K–$50K / 10K–100K transactions
- Billing amounts should not be round numbers — they're calculated from rates × volume
- Include GBP amounts for 2–3 UK customers (different currency from USD default)
- The "polling without contract" leakage cases should look like real active customers (high volume, recent invoice dates) — not like test accounts
- Stale contracts should show term_end 3–18 months in the past with recent billing still happening
- Name drift cases should be realistic variations — `Atlas Brokerage` vs `Atlas Brokerage Network`, `Meridian L&C` vs `Meridian Life & Casualty`

---

## 7. Tradeoffs — what's good about this plan and what we're giving up

### What's strong
- One end-to-end story. Easier to watch, easier to record, easier to defend as a finished artifact.
- The Revenue Leakage Finder hero is grounded in a real iPipeline finding (81 polling-without-contract customers). This is not a made-up demo scenario — it reflects an actual analysis question Connor's team has worked through.
- The ARR waterfall closes the chapter as a clean executive visual rather than being the whole story.
- Chapters 4–6 follow a repeatable pattern (problem → script → output folder) that coworkers can learn and apply.
- All 4 demo tools are stdlib-only — no pip install barrier.
- The launcher menu at Chapter 7 gives non-developers a single entry point.

### What we're giving up
- The CFO-focused "executive showcase" framing from the original 4a plan. Accepted: the real V1 audience is 50–150 coworkers, not the C-suite.
- The 4a/4b split structure, which some stakeholders may have been expecting. Accepted: one strong 9–12 min video is better than two ~6 min videos with a fragmented narrative.
- xlwings Excel Button Edition — coworkers with locked-down laptops can't use it. Parked for v2.
- Depth on the banded overage pricing model — the demo simplifies to expected_annual_revenue vs. amount_billed rather than reconstructing the full Band 1–10 calculation. This is the right trade. The full calculation would add build complexity without making the demo clearer.

---

## 8. Optional recipe shorts roadmap (non-canonical, post-V4)

Once V4 ships, short standalone clips (2–5 min each) can be produced without being part of the official 4-video series. These don't require ElevenLabs narration or Director macro automation — a simple screen recording with voiceover is fine.

Candidates:
- "Run the Data Contract Checker on your own file" (practical how-to)
- "Use the Exception Triage Engine on any exception list" (generic version)
- "Build a revenue leakage check for a different billing model" (customization walkthrough)
- "What the Control Evidence Pack produces and why it matters for review" (deeper dive)

These are non-canonical backlog items — not blockers for V4, not part of the official video series.

---

**End of revised plan.**
