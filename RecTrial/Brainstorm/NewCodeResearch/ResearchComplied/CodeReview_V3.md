# iPipeline Video 4 — AI Handoff (2026-04-22)

## 1. What Is This File

This file contains one research-review session output: an inventory pass on 11 research files, a long-list candidate pass (16 accepted + 3 borderline + 21 rejected = 40 unique ideas evaluated), a scored finalists pass (9 candidates with on-camera demo narratives), and a combo-recommendation pass (best-3-per-combo plus opinionated personal pick with final review flags). All four passes are reproduced in full in Section 4 for Claude Code to act on.

## 2. Project Context

**Who.** Connor Atlee, Finance & Accounting analyst at iPipeline (~2,000-person SaaS for life insurance / financial services, owned by Roper Technologies). Non-developer; reads VBA/SQL/Python at working level; relies on Claude to author new code. All deliverables target a non-technical finance audience in plain-English voice.

**What's being built.** A 4-video internal demo series for 2,000+ coworkers plus CFO/CEO.

**Video status.**
- V1 — "What's Possible" (Excel + VBA highlight reel): **recorded, shipped**
- V2 — "Full Demo Walkthrough" (62 automated actions on a demo P&L workbook): **recorded, shipped**
- V3 — "Universal Tools" (VBA toolkit that plugs into any Excel file): **recorded, shipped**
- V4 — "Python Automation for Finance" (5–8 min target, open to going longer / splitting): **planning (this file)**

**Three Video 4 combos still weighed.**
1. **Finance Copilot menu** — single Python script, numbered menu, guided prompts per task; wraps existing scripts in one entry point.
2. **Excel Button Edition (xlwings)** — macro-enabled workbook with buttons; Python runs silently; results appear as new sheets. Zero Command Prompt exposure.
3. **Hero Demo + Cookbook** — one-command hero demo + 5 copy-paste recipes.

**Hard constraints.**
- No external AI API calls (OpenAI, Claude, Gemini, etc.)
- No Outlook / email automation
- No Windows Task Scheduler
- No internet scraping of company / paid data
- Approved Python packages only: `pandas`, `openpyxl`, `pdfplumber`, `python-docx`, `thefuzz`, `numpy`, `matplotlib`, `xlwings`, stdlib
- Non-developer audience voice
- iPipeline brand: Blue `#0B4779`, Navy `#112E51`, Innovation Blue `#4B9BCB`, Lime `#BFF18C`, Aqua `#2BCCD3`, Arctic White `#F9F9F9`, Charcoal `#161616`, Arial only

**Already-built Python scripts (do not re-suggest).** `aging_report`, `bank_reconciler`, `compare_files`, `forecast_rollforward`, `fuzzy_lookup`, `pdf_extractor`, `variance_analysis`, `variance_decomposition`, `clean_data`, `consolidate_files`, `multi_file_consolidator`, `date_format_unifier`, `two_file_reconciler`, `sql_query_tool`, `word_report`, `batch_process`, `regex_extractor`, `unpivot_data`, `pnl_forecast`, `pnl_dashboard`, `master_data_mapper`, `profile_workbook`, `sanitize_dataset`, `compare_workbooks`, `build_exec_summary`, `variance_classifier`, `scenario_runner`, `sheets_to_csv`.

## 3. Open Decisions

### D1 — Which combo ships?
- **Options:** Combo 1 (Finance Copilot menu) / Combo 2 (xlwings buttons) / Combo 3 (Hero + Cookbook) / Synthesis (Combo 3 video + Combo 1 deliverable).
- **Session finding:** Personal pick is the synthesis — ship Combo 3 as the video structure (F1 hero + F9/F4/F2/F6 recipes) with the download being a Combo 1 Finance Copilot menu wrapping those plus the 28 existing scripts. Combo 2 flagged as weakest for this specific video because it subsumes Python into the VBA narrative V1–V3 already covered.
- **Status:** Unconfirmed.

### D2 — Single video (5–8 min) or split (Part 1 / Part 2)?
- **Options:** Single / split.
- **Session finding:** If synthesis path (D1) is chosen, split is strongly recommended. Part 1 ~5–6 min CFO-leaning (F1 hero + F2 + F4). Part 2 ~5–6 min coworker-leaning (Finance Copilot menu tour + F9 + F6). If Combo 1 alone, single video is fine.
- **Status:** Unconfirmed.

### D3 — Does SOX apply to Connor's team?
- **Options:** Yes (F7 stays a finalist) / No (F7 drops; promote A14 Operating Metrics Monthly Tracker).
- **Assumption made:** iPipeline → Roper Technologies (public) means SOX applies somewhere, but whether Connor's team personally owns SOX evidence prep is unconfirmed.
- **Impact:** F7's CFO-wow score drops from 5 to ~2 if Connor's team doesn't own SOX work.

### D4 — F1 narrative risk (CSV reframing vs research files).
- **Context:** Research files describe "Revenue Leakage Finder" as live Salesforce + Zuora + AWS integration. Hard constraints rule that out. F1 is reframed as three-CSV-input tool (sold / billed / used).
- **Action needed:** Prepare one-line talking point acknowledging Phase 2 is live-wire integration. Not a blocker, but a narration risk.

### D5 — F5 vs F7 on-camera overlap.
- **Context:** Both produce audit-ready zipped deliverables with hash manifests and index workbooks. F5 = monthly close; F7 = quarterly SOX.
- **Action needed:** If both end up on the Combo 1 menu, demo only one on camera; mention the other in passing.

### D6 — Does Video 4 need a python-docx showcase?
- **Context:** None of the 9 finalists leans heavily on python-docx (F5 and F7 use it for 20-line cover memos).
- **Options:** (a) Accept current list, (b) scope F5/F7 bigger with auto-drafted memos, (c) swap in A14 which produces a full monthly Word narrative.
- **Status:** Unchosen.

### D7 — What data powers the F1 hero demo?
- **Context:** Combo 3 hinges on F1 producing a cinematic "we found $X in leakage" reveal. Requires realistic synthetic CRM / billing / usage CSVs with planted gaps.
- **Action needed:** Build F1 end-to-end first, screen-test the reveal before committing Combo 3. If the reveal collapses, fallback is Combo 1 with F5 + F6 + F1 as the menu trio.

## 4. Full Results

### Inventory Pass Results

**11 files analyzed. Content fully readable in all 11; three have file-extension drift.**

| # | Filename | Actual file type | One-sentence summary |
|---|---|---|---|
| 1 | `200_tools_catalog.md` | Markdown, 257 lines | Flat catalog of ~200 open-source Python and SQL libraries grouped by function, sourced from Best-of-Python and Awesome-SQL. |
| 2 | `CODE_CATALOG.md` | Markdown, 170 lines | Project-context doc for a sibling `MobileCLDCode/` folder containing 42 files (11 VBA, 13 Python, 6 SQL, 4 PowerShell, 3 OfficeScripts, 2 PowerAutomate, 2 PowerQuery) that showcase automations beyond native Excel/OneDrive, with full iPipeline brand specs. |
| 3 | `CodexCodeIdeas.md` | Markdown, 369 lines | Prioritized backlog: 10 SQL + 10 Python + 10 VBA + 5 other ideas, each with business outcome, use case, implementation approach, stack, and phase priority. |
| 4 | `deep-research-report.md` | Markdown, 299 lines | Synthesis report with a Python branch-synthesis harness, a 200+ external tool longlist, and a 5-idea future-state roadmap. |
| 5 | `deep-research-report_2.md` | Markdown, 231 lines | Alternate synthesis: audit-harness scaffolding, 5-tool cherry-pick (sqlglot, great_expectations, pandera, dedupe, unstructured), 5-idea roadmap. |
| 6 | `Executive_Automation_Catalog_2.docx` | **Plain markdown text mis-named `.docx`**, 391 lines | Enterprise catalog: SQL entitlement reconciler, Python Pydantic/Instructor validation, VBA-Web framework, 200+ tool list, 5 future-state ideas. |
| 7 | `Executive_Automation_Catalog.md` | Markdown, 266 lines | Template-style catalog spec with methodology, ID conventions, structure template, example SQL/Python/VBA entries, 5 roadmap ideas. |
| 8 | `Executive_Automation_Catalog__Master_Reference_for___.docx` | **Plain markdown text mis-named `.docx`**, 172 lines | Master synthesis: Compare-and-Classify SQL reconciliation, Pydantic+Instructor gatekeeper, VBA-Web legacy bridge, 5-tool cherry-pick (Splink, Instructor, dbt-audit-helper, VBA-Web, Pandera), 5 future-state ideas. |
| 9 | `report.md` | Markdown, 357 lines | SQL cross-database referential integrity triggers, Python ML/ETL/RPA, VBA UserForms and OLE automation, 7-row table of next-gen ideas. |
| 10 | `report_extended.md` | Markdown, 337 lines | Extended sibling of #9: Python/SQL reconciliation examples, massive open-source tool directory by domain, 5 future-state ideas. |
| 11 | `Exec_auto_master_fixed.pdf` | **ZIP archive mis-named `.pdf`** (6 JPEG pages + 6 `.txt` + `manifest.json`) | Byte-for-byte the same content as file #8, re-wrapped as page images with extracted text. |

**Flags raised during intake.**

1. **Three files have wrong extensions.** #6 and #8 are plain markdown text labeled `.docx`; #11 is a ZIP of JPEGs labeled `.pdf`. All three read cleanly once handled as their actual formats.
2. **Heavy content overlap.** Files #6, #8, #11 share large verbatim sections. Files #4 and #5 are two drafts of the same brief. Files #9 and #10 are a pair. Deduplication was required during the long-list pass.
3. **Formatting artifacts.** #4 has `citeturn##view#` citation residue throughout. #6 has `[span_XX](start_span)[span_XX](end_span)` artifacts embedded in SQL code blocks.
4. **`CODE_CATALOG.md` (#2) ends at an append-marker** — file tree complete, per-file deep detail never appended. Not a truncation; file ends intentionally.
5. **No critical content unreadable.** Every file rendered to text in full.

### Full Review Results (Long-List Candidate Pool)

**40 unique ideas evaluated. 16 ACCEPT + 3 BORDERLINE + 21 REJECT.**

#### ACCEPT — 16 candidates

| # | Candidate | One-line description | Key source(s) |
|---|---|---|---|
| A1 | Close Readiness Scorecard | Weighted 0–100 per entity per day (validations + feeds + postings on time), flags ready vs blocked with specific reasons. | CodexCodeIdeas SQL-02 |
| A2 | Exception Triage Engine | Config-driven `impact × confidence × recency` ranker so analysts work highest-value breaks first. | CodexCodeIdeas PY-01 |
| A3 | Control Evidence Pack Generator | Zips run logs + reconciliation outputs + hash manifest + index workbook into one audit-ready bundle. | CodexCodeIdeas PY-07 |
| A4 | SaaS ARR Waterfall Builder | Subscription snapshot CSV → new/expansion/contraction/churn/ending ARR with branded matplotlib waterfall. | CODE_CATALOG `saas_arr_waterfall.py` |
| A5 | Revenue Recognition Schedule Engine | Contract register → monthly rev rec schedule + deferred-rev roll-forward (ASC 606-style). | CODE_CATALOG `revenue_recognition_engine.py`; DRR1 |
| A6 | Cohort Retention Analyzer | Customer activity → retention heatmap grid. | CODE_CATALOG `cohort_retention_analyzer.py` |
| A7 | License Utilization Analyzer | Licenses sold vs used → over/under-provisioning with $ impact per account. | CODE_CATALOG `license_utilization_analyzer.py` |
| A8 | SOX Evidence Collector | Walks controls folder tree, indexes evidence, bundles quarterly deliverable. | CODE_CATALOG `sox_evidence_collector.py` |
| A9 | Vendor Payment Anomaly Scanner | Rolling median + MAD/z-score per vendor on amount and timing; flags outliers. | CodexCodeIdeas SQL-03 (Python reframe) |
| A10 | JE Duplicate-Ring Detector | Near-duplicate journal entries split across days/users/entities via amount + vendor + account similarity. | CodexCodeIdeas SQL-01 (Python reframe); R1E |
| A11 | Revenue Leakage Finder (CSV) | Reconciles sold/billed/used CSVs; classifies gaps as under-billed, ghost, orphan, expired. | DRR1, DRR2, EAC2, EACMR, R1E, Exec PDF |
| A12 | Root-Cause Reconciliation Assistant | On rec break, fuzzy-matches against historical break/resolution log to suggest cause categories. | CodexCodeIdeas PY-05 |
| A13 | Workbook Dependency Scanner | Parses formulas via openpyxl, outputs cross-sheet dependency graph as interactive HTML. | CodexCodeIdeas PY-08 |
| A14 | Operating Metrics Monthly Tracker | KPI config file → branded multi-metric trend pack for the controller. | CodexCodeIdeas "Success Metrics to Track Monthly" |
| A15 | Finance Data Contract Checker | Declarative YAML rules per feed (columns, types, ranges, uniqueness); fails gate on arrival. | CodexCodeIdeas PY-03 |
| A16 | Segregation-of-Duties Audit Pack | Cross-joins role-action matrix vs transaction log; flags who created + approved the same record. | CodexCodeIdeas SQL-09 (Python reframe) |

#### BORDERLINE — 3 candidates flagged for judgment call

| # | Candidate | Why borderline | Source(s) |
|---|---|---|---|
| B1 | Contract Clause Extractor (regex, no LLM) | pdfplumber + regex for renewal dates / notice windows / escalators. Overlaps existing `pdf_extractor` + `regex_extractor`. | R1E §3.2, EAC2 §3.2, EACMR §3.2, DRR2 |
| B2 | Close Bottleneck Heatmap | Close-task timestamp log → heatmap of step × entity × user delays. Conceptually close to A1. | CodexCodeIdeas SQL-10 (Python reframe) |
| B3 | Narrative Variance Writer (templates) | Deterministic threshold-driven commentary. Overlaps existing `build_exec_summary` + `variance_classifier`. | CodexCodeIdeas PY-02 |

#### REJECT — 21 candidates removed, reason-first

- **LLM contract extraction / multi-agent Instructor / Splink identity resolution** — hard constraint: no external AI APIs. (DRR1, DRR2, EAC2, EACMR, R1, R1E, Exec PDF.)
- **Live Salesforce + AWS CloudTrail entitlement audit** — requires live CRM/cloud API access. Idea kept alive as A11 (CSV reframe).
- **Forecast Ensemble Manager (PY-04)** — requires scikit-learn.
- **Close Calendar Risk Predictor (PY-06)** — requires scikit-learn; too ML-heavy for non-dev audience.
- **Data Drift Monitor (PY-10)** — requires scipy/statsmodels.
- **Isolation Forest / Prophet / SARIMA / LSTM anomaly detection** — ML libs not on approved list. Statistical portion captured by A9.
- **Legacy ERP bridge (pyautogui / selenium / pywinauto / WebView2)** — packages not approved; off-theme for Finance video.
- **email_to_structured_data** — email automation hard constraint.
- **jira_weekly_digest, support_ticket_triage, api_slo_tracker, aws_cost_optimizer** — all require external APIs.
- **Workbook-to-Source Reconciliation Mart (SQL-05)** — subsumed by existing `compare_workbooks` + `two_file_reconciler`.
- **Subledger Completeness Matrix (SQL-06)** — subsumed by A1.
- **Allocation Drift Tracker (SQL-04)** — subsumed by existing `variance_decomposition`.
- **Policy-as-Code Rule Engine (SQL-07)** — subsumed by A15.
- **Tenant Identity Resolution Fabric** — subsumed by existing `fuzzy_lookup` + `master_data_mapper`.
- **Shadow Revenue Journal** — too abstract to demo.
- **dbt-style model layer (OA-04)** — out of language scope.
- **.NET signed add-in (OA-02)** — out of language scope.
- **GitHub Actions validator (OA-05)** — infrastructure-tier, not Finance deliverable.
- **Internal Exception Status API (OA-03)** — too infra for a 5–8 min video.
- **customer_churn_risk_scorer** — needs CRM/support data + typically ML.
- **git_developer_metrics** — non-finance audience fit.

### Video 4 Finalists

**9 selected from the 16 ACCEPT pool.**

**Cuts made in narrowing 16 → 9:**
- A2 (Exception Triage) — strong tool but visually flat vs finalists.
- A7 (License Utilization) — narrow audience; overlaps A11's "over-provisioned" category.
- A10 (JE Duplicate Rings) — overlaps A9's forensic vibe; A9 is the simpler, stronger demo.
- A12, A14, A15, A16 — each loses head-to-head on "watchable in 5–8 minutes."
- All 3 borderlines dropped: B1 too close to `pdf_extractor`+`regex_extractor`, B2 too close to A1, B3 too close to `build_exec_summary`.

#### Scored Finalists Table

| # | Idea Name | What it does | Why perfect for Video 4 | Best Combo | CFO Wow | Coworker Use | Demo-ability | Effort | Packages | Source File(s) |
|---|---|---|---|---|---|---|---|---|---|---|
| F1 | Revenue Leakage Finder (3-way CSV) | Reconciles "sold" (CRM) vs "billed" (billing) vs "used" (usage log) CSVs; classifies every gap as under-billed / ghost-billed / orphan-usage / expired-still-consuming; quantifies $ per account. | Only idea that appears in every research file — crown-jewel "we found money" story. Three-way fuzzy reconciliation across messy CSVs is genuinely hard in Excel. | Combo 3 Hero | 5 | 3 | 5 | M | pandas, thefuzz, openpyxl, matplotlib | DRR1; DRR2 §1; EAC2 §3.1/§3.5; EACMR §3.1/§3.4; R1E §3.1; Exec PDF p.3–4 |
| F2 | SaaS ARR Waterfall Builder | Subscription snapshot CSV → new/expansion/contraction/churn/ending ARR + branded waterfall chart + board-ready roll-forward table. | Iconic SaaS metric in visual form. Every CFO has seen 100 of these and knows Excel pain. Directly mirrors iPipeline's business model. | Any (Hero-worthy) | 5 | 3 | 5 | M | pandas, matplotlib, openpyxl | CODE_CATALOG `saas_arr_waterfall.py` |
| F3 | Revenue Recognition Schedule Engine | Contract register (id/start/end/total/pattern) → monthly rev rec schedule per contract + deferred-rev roll-forward; handles partial months and mid-term mods. | ASC 606 is universal SaaS finance pain. "500 contracts, 2 seconds" productivity story controllers feel in their bones. | Combo 1 anchor / Combo 2 | 5 | 4 | 4 | M–L | pandas, openpyxl, numpy | CODE_CATALOG `revenue_recognition_engine.py`; DRR1 |
| F4 | Vendor Payment Anomaly Scanner | Rolling median + MAD/z-score per vendor on amount and timing; flags outliers with each vendor's baseline + flagged deviation. | Forensic demos play enormously well on camera. Pure numpy, no ML, but carries "suspicious pattern detected" weight. | Combo 3 Hero / any | 5 | 3 | 5 | M | pandas, numpy, matplotlib | CodexCodeIdeas SQL-03 (Python reframe) |
| F5 | Control Evidence Pack Generator | Walks close-run folder, indexes reconciliations/logs, computes SHA256 hashes, writes index workbook with hyperlinks, emits manifest.json, zips bundle + pre-filled cover memo. | "Monday morning I can actually use this." Multi-day audit prep scramble collapses to 5 seconds. Controller-grade output. | Combo 1 / Combo 2 | 5 | 4 | 4 | M | openpyxl, python-docx, stdlib (zipfile, hashlib, pathlib, json) | CodexCodeIdeas PY-07 |
| F6 | Close Readiness Scorecard | Close-checklist status + feed timestamps + validation pass/fail → weighted 0–100 per entity per day with traffic-light coding + specific blockers per red row. | Addresses controller's single most-asked question — "where are we?" — in one glance. Universal close pain across every finance org. | Any | 4 | 4 | 4 | M | pandas, openpyxl, matplotlib | CodexCodeIdeas SQL-02 (Python reframe) |
| F7 | SOX Evidence Collector | Walks controls folder tree (one per control), checks each has expected evidence, builds quarterly index workbook with pass/fail per control, bundles into auditor deliverable. | iPipeline → Roper (public) = SOX applies. Quarterly week-of-prep collapses to seconds. Massive CFO wow because SOX is existential at parent level. | Combo 1 / Combo 2 | 5 | 3 | 4 | M–L | pandas, openpyxl, python-docx, stdlib (pathlib, zipfile, hashlib) | CODE_CATALOG `sox_evidence_collector.py` |
| F8 | Workbook Dependency Scanner | Opens target xlsx, walks every formula via openpyxl, builds full cross-sheet dependency graph, emits interactive HTML + orphan/broken-ref/circular-chain report. | Visual is stunning on camera. Addresses universal "inherited ugly workbook" pain. Strongest possible case Python does things Excel structurally cannot. | Combo 3 Hero / Combo 2 | 3 | 5 | 5 | M | openpyxl, stdlib (html, json) | CodexCodeIdeas PY-08 |
| F9 | Cohort Retention Analyzer | Reads customer activity (id, signup month, active month), groups into cohorts, produces classic retention heatmap + underlying percentages table. | Cohort heatmaps are board-deck standard but tedious in Excel. Visually iconic on camera in iPipeline blue, 30 lines of Python. Perfect Combo 3 recipe. | Combo 3 recipe / Combo 2 | 4 | 3 | 5 | M | pandas, matplotlib, numpy | CODE_CATALOG `cohort_retention_analyzer.py` |

#### 30–60 Second Demo Narratives

**F1 — Revenue Leakage Finder.** Show three CSVs sitting in a folder: what Sales said we sold, what Billing actually billed, and what the product logs say customers consumed. Run one command: `python revenue_leakage.py ./inputs`. Two seconds. `leakage_findings.xlsx` opens on tab one: *"Under-billed: $42,300 across 7 accounts. Ghost billing: $1,100 across 2 accounts. Orphan usage: 31 tenants with no contract."* Jump to the detail tab. Point at one row: the CRM has "Acme Corporation," billing has "Acme Corp., Inc." — fuzzy-matched at 96%, under-billed by $4,500. Finish on the category summary chart, iPipeline blue bars. Closing line: *"Every one of those rows is a conversation with the billing team worth having Monday morning."*

**F2 — SaaS ARR Waterfall Builder.** Drag `subscriptions_march.csv` onto the script icon. One file, three columns: customer, month, MRR. Three seconds. A single branded PNG pops up: starting ARR on the left, green bars up for new logos and expansion, red bars down for contraction and churn, ending ARR on the right, net delta annotated in the middle. Behind it, an xlsx opens with the exact underlying roll-forward table — copy-paste ready for the board deck. Closing line: *"Every ARR deck I've ever seen took 2 hours in Excel. This is 10 seconds."*

**F3 — Revenue Recognition Schedule Engine.** Open the contract register: 500 rows, each with contract_id, start, end, total value, and revenue pattern. Close it. Run `python revrec.py contracts.xlsx`. Two seconds. `revrec_schedule.xlsx` opens to tab 1: every contract exploded into monthly recognition rows, roughly 6,000 rows of booked revenue, iPipeline-branded formatting. Flip to tab 2: the deferred revenue roll-forward — opening balance, contracts added, revenue released, closing balance, month by month. Tab 3: a deferred-balance trend chart over 24 months. Closing line: *"Doing this in Excel for 500 contracts is a full accountant's day. Python just gave you that day back."*

**F4 — Vendor Payment Anomaly Scanner.** Open the AP register CSV — 18 months, about 12,000 rows. Run `python vendor_anomaly.py ap_history.csv`. Terminal prints *"Top 10 anomalies detected."* Click the output xlsx. Row 1: *"XYZ Consulting. Typical payment: $1,200. This month: $15,400. Z-score 8.4."* Row 2: *"ABC Office Supplies. Typical timing: 28 days net. Last 3 invoices: paid in 4 days."* Scroll to the chart tab — red dots floating well above a blue baseline trend per vendor. Closing line: *"We didn't need machine learning. We needed a rolling median and someone to actually look."*

**F5 — Control Evidence Pack Generator.** Pan across the March close folder on screen — 40 scattered files: recon outputs, logs, screenshots, signoff PDFs. *"The auditor wants all of this packaged. Usually two days of copy-pasting."* Run `python evidence_pack.py --close 2026-03`. Four seconds. A new file appears: `close_evidence_2026-03.zip`. Open it. Inside: every original file plus `index.xlsx` (file name, SHA256 hash, size, category, hyperlink), `manifest.json` (machine-readable for the auditor's tool), and `cover_memo.docx` (pre-filled with close period and sign-offs). Closing line: *"Run this every month and audit evidence prep goes from a week to zero."*

**F6 — Close Readiness Scorecard.** It's Wednesday of close week. Run `python close_readiness.py`. The scorecard appears on screen: six entities listed — US, UK, Canada, Singapore, Australia, Roper-direct. Each row: overall score, validations-passed %, feeds received, postings-on-time %. US green at 94. UK green at 88. Canada yellow at 71. Singapore red at 42. Below each red row, the specific blocker: *"Singapore — payroll feed missing (expected 8 AM, not received), 3 unreconciled AR accounts, 2 failed FX validations."* Closing line: *"This is the view I want at 9 AM every close Monday. It tells me exactly where to send the fire brigade."*

**F7 — SOX Evidence Collector.** Open the controls folder: 80 folders labeled ITGC-01 through ITGC-80. *"Last quarter, cross-checking all of these was a week of my life."* Run `python sox_collect.py --q 2026Q1`. Seven seconds. Three outputs drop: `SOX_Q1_2026_Index.xlsx` (every control, evidence-complete or flagged-missing, pass rate 77/80), `SOX_Q1_2026_Bundle.zip` (all evidence in one archive), and `SOX_Q1_2026_Cover.docx` (pre-written cover memo). Open the index — three rows highlighted red at the top, with the exact expected file name they're missing. Closing line: *"Three specific files to chase. Everything else is auditor-ready right now."*

**F8 — Workbook Dependency Scanner.** Drop the nastiest inherited model on the desktop — 14 tabs, no documentation, the kind everyone is scared to edit. Run `python depgraph.py mystery_model.xlsx`. Three seconds. A browser tab opens: the workbook visualized as a network graph, boxes for sheets, lines for cross-sheet formula dependencies. Click "Summary!B12" — the graph lights up every cell feeding into it, 4 cells across 3 sheets. Sidebar counts: *"Orphan cells: 17. Broken refs: 2. Circular-adjacent chains: 1."* Click the circular chain — it lists the exact path. Closing line: *"Now I can change that one cell without blowing up the model. Or better yet, rebuild the whole thing knowing where the wiring actually goes."*

**F9 — Cohort Retention Analyzer.** Open `customer_activity.csv` — 18,000 rows, customer_id / signup_month / active_month. Close it. Run `python cohort.py customer_activity.csv`. Three seconds. A heatmap fills the screen: y-axis is signup cohort (24 monthly cohorts), x-axis is months since signup (0 to 24), each cell is the percentage of that cohort still active, dark iPipeline-blue for high retention, faded for low. The diagonal is solid. Point at one row: *"Our 2024-Q1 cohort drops sharply at month 4 — that's a product event finance should flag to CS."* Closing line: *"Thirty lines of Python, a permanent board-deck chart. Excel's pivot-table acrobatics couldn't do this in an hour."*

### Best Ideas Curation (Combo Recommendations + Personal Pick)

#### Combo 1 (Finance Copilot menu) — Best 3

**F5 + F6 + F1.** Menu format wins when each item needs short guided prompts (period? folder? threshold?) and when the trio covers distinct controller personas. F5 = packaging ("prep the audit bundle"), F6 = monitoring ("where are we right now?"), F1 = detective work ("where is money leaking?"). Most "Monday morning real" trio — every item is a recurring task done manually today.

#### Combo 2 (xlwings Excel buttons) — Best 3

**F8 + F2 + F6.** xlwings shines when the button is *on* the workbook and output is a new sheet/chart right there. F8 is the showstopper — a "Scan This Workbook" button mapping itself is the most memorable xlwings demo possible because the workbook inspects its own wiring. F2 is the CFO button (click → branded waterfall + roll-forward). F6 is the daily controller button (fresh scorecard sheet into close pack). All three feel native to Excel.

#### Combo 3 (Hero Demo + Cookbook) — Best 3

**F1 HERO + F9 + F4.** F1 is the only correct hero — no other finalist carries the "we found real dollars" punch. F9 is the most visually iconic recipe possible in 30 lines. F4 is the forensic recipe that thematically pairs with F1 (both tell "Python sees what humans miss" but with different mechanics: fuzzy reconciliation vs statistical outliers). 4th recipe: F2. 5th: F6.

#### Personal Pick

**Ship Combo 3 as the video and Combo 1 as the downloadable tool — and seriously consider splitting into Part 1 and Part 2.**

Mentor read: Combo 2 is the weakest choice for *this specific* video in *this specific* series. V1–V3 already taught the Excel+VBA world. Video 4's strategic job is to position Python as a *distinct* value proposition, not a hidden engine behind Excel buttons. Combo 2 subsumes Python into the VBA story already told — CFO walks away thinking "Python is what makes VBA fancier," which is the wrong lesson. Cross Combo 2 off.

Between Combo 1 and Combo 3: Combo 1 is safer and more complete-feeling; Combo 3 is more memorable, more CFO-friendly, more differentiated from VBA videos. The hero demo (F1) produces the single best 60-second clip CFO/CEO might share externally — huge asset for a 2,000-person internal demo. Cookbook format solves the audience-split problem better than a menu does: hero for leadership, recipes for coworkers.

Synthesis: ship Combo 3 as the video structure (hero → 4 recipes), but the *tool* people download is the Combo 1 Finance Copilot menu wrapping F1, F9, F4, F2, F6 plus the 28 existing scripts. Video teaches through spectacle; artifact in the downloads folder is a single `finance_copilot.py` they run Monday and keep running for years. Menu also solves the friction critique of Combo 3 (copy-pasting recipes is higher friction than clicking a menu).

Runtime: 5–8 min is tight for hero + 4 recipes + menu tour. **Split into two videos.** Part 1 (5–6 min, CFO-led): open with F1 hero, show F2 + F4 as supporting recipes, close with 20-second Finance Copilot preview. Part 2 (5–6 min, coworker-led): Finance Copilot menu walkthrough + F9 + F6 as recipes, plus "here's how to add your own" closing to teach ownership. Split is clean: leadership watches Part 1, analysts watch both. Doubles shelf life without doubling production (same code, same tool, two edits).

**Caveat:** All above assumes F1 demos cleanly in 60 seconds with a clear dollar-impact punchline. If the synthetic data doesn't produce a visually convincing reveal (one big red number, a few specific named-account rows), the hero collapses. Build F1 end-to-end first and screen-test the reveal before committing Combo 3. If F1 underwhelms on camera, fall back to Combo 1 with F5+F6+F1 as the menu trio.

#### Final Review Pass — Flags

1. **F7 placement is load-bearing on a SOX-scope assumption.** If Connor's team doesn't own SOX evidence work, F7's CFO-wow drops from 5 to 2 and A14 (Operating Metrics Tracker) should be promoted in.
2. **F1's CSV reframing softens the research-files narrative.** Have a one-liner ready: "Phase 1 is the engine; Phase 2 is live-wire integration to CRM/billing."
3. **F5 vs F7 overlap on camera.** Both produce audit-ready zipped deliverables. Demo only one; mention the other in passing.
4. **Personal pick assumes willingness to split the video.** If strictly one video 5–8 min, pick changes to **Combo 1 with F5 + F6 + F1** — safer, complete, lower production risk.
5. **python-docx usage is light across finalists.** If Video 4 needs to showcase docx specifically, either scope F5/F7 bigger (auto-drafted memos) or promote A14 (produces full monthly narrative Word doc).

## 5. Top Picks Consolidated

1. Revenue Leakage Finder (F1) | Python | M | Combo 3 Hero | Crown-jewel "we found money" story; in every research file
2. SaaS ARR Waterfall Builder (F2) | Python | M | Any Hero | Iconic SaaS chart; mirrors iPipeline's own business model
3. Revenue Recognition Schedule Engine (F3) | Python | M–L | Combo 1 anchor | ASC 606 pain; "500 contracts, 2 seconds" productivity story
4. Vendor Payment Anomaly Scanner (F4) | Python | M | Combo 3 Hero | Forensic flair, pure numpy — no ML needed
5. Control Evidence Pack Generator (F5) | Python | M | Combo 1 / Combo 2 | "Monday morning real" — week of audit prep collapses to 5 seconds
6. Close Readiness Scorecard (F6) | Python | M | Any | Controller's #1 question "where are we?" answered in one view
7. SOX Evidence Collector (F7) | Python | M–L | Combo 1 | Existential CFO concern at Roper parent level (if applies)
8. Workbook Dependency Scanner (F8) | Python | M | Combo 3 / Combo 2 | Visually stunning; workbook inspects its own wiring on camera
9. Cohort Retention Analyzer (F9) | Python | M | Combo 3 recipe | Iconic heatmap; 30 lines of Python, permanent board-deck chart
10. Exception Triage Engine (A2) | Python | M | Combo 1 recipe | Config-weighted scoring; analysts work highest-value breaks first
11. JE Duplicate-Ring Detector (A10) | Python | M | Combo 3 alt | Forensic JE pattern detection; pairs with F4 if F4 pulled
12. Operating Metrics Monthly Tracker (A14) | Python | M | Combo 1 recipe | F7 swap-in if SOX doesn't apply; full monthly Word narrative
13. Root-Cause Reconciliation Assistant (A12) | Python | M | Combo 1 recipe | Fuzzy-matches break history to suggest cause; extends `bank_reconciler`
14. Segregation-of-Duties Audit Pack (A16) | Python | M | Combo 1 recipe | Role-action matrix joins; flags creator = approver
15. License Utilization Analyzer (A7) | Python | S–M | Combo 1 recipe | Sold vs used per account, $ impact; narrow but fast to ship
