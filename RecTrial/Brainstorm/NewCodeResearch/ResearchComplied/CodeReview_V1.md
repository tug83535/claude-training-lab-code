# iPipeline Finance & Accounting Demo Series — AI Handoff

**Date:** 2026-04-22
**Owner:** Connor Atlee (Finance & Accounting, iPipeline)
**Audience for this file:** Claude Code (and any successor AI reviewer)

---

## 1. WHAT IS THIS FILE

This file contains the complete output of three review passes run against Connor Atlee's iPipeline Finance & Accounting demo-series research archive: a file-inventory pass across 11 source research files, a full curated review producing 60 deduplicated ideas across Sections A–D (Universal Toolkit Additions, Video 4 Candidates, Future, Skip), and a Top-10 best-ideas consolidation. Approximately 130+ raw ideas were inventoried before deduplication and constraint-filtering. All tables and narrative are preserved in full — no summarization.

---

## 2. PROJECT CONTEXT (compressed)

**Who / What.** Connor Atlee, Finance & Accounting analyst at iPipeline (SaaS in life insurance / financial services). Building a 4-video internal demo series for 2,000+ coworkers plus CFO and CEO demonstrating Excel + VBA + Python + SQL automation for non-technical Finance users. Every deliverable must be CFO-grade, plain English, runnable by someone with zero coding background.

**The 4 videos.**

| Video | Theme | Status |
|---|---|---|
| 1 | "What's Possible" — fast highlight reel (Excel + VBA focus) | **RECORDED** |
| 2 | "Full Demo Walkthrough" — end-to-end tour of a macro-enhanced P&L workbook with 62 automated actions | **RECORDED** |
| 3 | "Universal Tools" — VBA toolkit that works on any Excel file | **RECORDED** |
| 4 | "Python Automation for Finance" — must ship a downloadable tool coworkers can use the Monday after watching | **PLANNING NOW** |

**Three Video 4 combos still being weighed** *(inferred from session pushback — see Open Decision OD-1)*:
- **Combo 1 — Single Flagship (≈7 min):** One tool (B1 Exception Triage) shown in depth with before/after.
- **Combo 2 — Three-Act (≈15 min):** B1 + B2 + B9 as three distinct acts in one longer video.
- **Combo 3 — Split Series (2 videos):** Video 4a *Python Controls* (B1 + B2); Video 4b *Python Intelligence* (B8 + B9 + B11).

**Hard constraints — do not violate.**
- No external AI API calls (OpenAI, Claude, Gemini, etc.)
- No Outlook / email automation
- No Windows Task Scheduler dependencies
- No packages outside the approved list
- Non-developer audience — every feature must be explainable and runnable with zero coding background
- Plug-and-play preferred — tools that work on any coworker's file beat demo-specific ones
- iPipeline branding: Blue `#0B4779`, Navy `#112E51`, Arial fonts, plain-English output

**Approved Python packages.** `pandas`, `openpyxl`, `pdfplumber`, `python-docx`, `thefuzz`, `numpy`, `matplotlib`, `xlwings`, Python stdlib. Nothing else.

**Already-built VBA — 23 modules, ~140 universal tools.** data sanitizer, compare, consolidate, highlights, pivot tools, tab organizer, column ops, sheet tools, comments, validation builder, lookup builder, command center, exec brief, finance-specific tools, audit tools, branding.

**Already-built Python — 22+ scripts.** aging report, bank reconciler, compare files, forecast rollforward, fuzzy lookup, pdf extractor, variance analysis, variance decomposition, clean data, consolidate files, multi-file consolidator, date unifier, two-file reconciler, SQL query tool, word report, batch processor, regex extractor, unpivot, pnl forecast, pnl dashboard.

**Already-built stdlib-only Python.** 7 zero-install scripts (names not provided in research archive — see OD-4).

**Already-built SQL.** 4 scripts — staging, transformations, validations, enhancements.

---

## 3. OPEN DECISIONS

Decisions that are **not** locked in. Do not assume a default without confirming with Connor.

| ID | Decision | Options | Session findings that should inform the choice |
|---|---|---|---|
| OD-1 | Video 4 structure | (a) Single 7-min flagship; (b) Single 15-min three-act; (c) Split into Video 4a + 4b | 5–8 min is too tight for any single idea in Section B to breathe properly for a non-technical audience. Combo 1 risks feeling thin; Combo 2 is the lowest-risk narrative; Combo 3 maximizes content but doubles production effort. |
| OD-2 | Flagship Video 4 tool | B1 Exception Triage vs. B2 Journal Entry Duplicate-Ring Detector vs. B11 Rules-Based Contract Extractor | B1 has the cleanest chaos→clarity visual. B2 has the strongest CFO/SOX language hook. B11 has the most "wait, Python can do that?" wow-factor. Different winners for different audience emphases. |
| OD-3 | B5 Unified Reconciliation Engine — ship alongside or supersede? | (a) Ship B5 next to existing `two_file_reconciler.py` + `compare_files.py`; (b) Supersede: B5 becomes v2, old scripts deprecated; (c) Skip B5 entirely | Running two tools for the same job confuses coworkers. The 4-status classifier (Identical / Modified / Removed / Added) is a genuine rigor upgrade. Recommend (b) but Connor has final say on deprecation risk. |
| OD-4 | The 7 stdlib-only scripts — do any already cover B9 / B10 / B11? | Confirm each | Script names are unknown. If any of them already implements PSI drift, subledger completeness, or rules-based contract extraction, mark the corresponding B-entry as already-covered and do not rebuild. |
| OD-5 | A1 Formula Integrity Fingerprint — overlap with existing audit-tools VBA module? | (a) Net new; (b) Extends existing audit-tools; (c) Already covered | Existing audit-tools module content is not documented in the research files. Requires Connor to open and inspect before building. |
| OD-6 | A3 Controlled Snapshot Sign-off — overlap with existing audit-tools? | Same as OD-5 | Same reason. |
| OD-7 | A5 Auto-Repair Suggestions — overlap with data sanitizer VBA? | (a) Net new diagnose-and-choose layer; (b) Extends data sanitizer; (c) Already covered | Data sanitizer's scope is not documented. If sanitizer is the auto-apply version, A5 is the correct diagnose-and-confirm counterpart and both should coexist. |
| OD-8 | A15 Duplicate-Record Finder — overlap with `fuzzy_lookup.py`? | (a) Net new (record-to-record dedup); (b) Already covered (if fuzzy_lookup handles that mode) | fuzzy_lookup's mode coverage is not documented. If it is column-to-column only, A15 fills a real gap; if it handles record-to-record already, A15 is Section D. |
| OD-9 | Top-10 #10 slot — B11 (flashy) or A3 (audit posture)? | Swap based on CFO preference | B11 is the "wow factor" demo tool. A3 is the quieter audit-defensibility story. If CFO reads audit posture as higher priority than modern-and-flashy, swap. |
| OD-10 | C6 Tenant Identity Fabric — does a thefuzz-only "good-enough" version belong in B instead of C? | (a) Leave in C (needs Splink for precision); (b) Move to B with thefuzz + pandas | Listed in C because Splink isn't approved and tenant-identity data likely isn't in Connor's hands. If tenant data is accessible, a thefuzz-based v1 is buildable. |
| OD-11 | B12 CFO Pack Assembly Pipeline — build now or park? | (a) Keep in B; (b) Park to C until B1/B9/others ship | L-effort and depends on other B items as inputs. Cleanly the *last* thing to build, not the first. |

---

## 4. FULL RESULTS

### Inventory Pass Results

11 research files analyzed. 3 had misleading file extensions. No content truncation detected. 1 file is a visual duplicate of another.

| # | Filename | Actual format | One-sentence summary |
|---|---|---|---|
| 1 | `200_tools_catalog.md` | Markdown (257 lines) | Flat grouped list of ~200 open-source Python and SQL tools (serialization, dataframes, ORMs, ETL, utilities, etc.) pulled from "Best-of Python" and "Awesome SQL" — names only, no implementation guidance. |
| 2 | `CODE_CATALOG.md` | Markdown (169 lines) | Internal reference doc describing an **already-built** `MobileCLDCode/` folder of ~42 files (VBA, Python, SQL, PowerShell, Office Scripts, Power Automate, Power Query) inside iPipeline's training lab repo — this is a meta-doc about existing code, not a proposals file. |
| 3 | `CodexCodeIdeas.md` | Markdown (368 lines) | Forward-looking backlog of automation ideas (SQL, Python, VBA, integrations) for an existing F&A Excel-centric platform, organized by section with business outcomes and KPIs; explicitly lists what already exists to avoid duplication. |
| 4 | `deep-research-report.md` | Markdown (298 lines) | Three-part research deliverable — (1) Python deduplication/synthesis harness, (2) external open-source longlist, (3) five future-state ideas (entitlement-to-usage audit graph, contract extraction, legacy ERP bridge, revenue leakage resolver, SLA/renewal risk radar); upfront admits the three branches were not accessible. |
| 5 | `deep-research-report_2.md` | Markdown (230 lines) | Shorter three-part companion to #4 with different top-5 open-source stack (sqlglot, great_expectations, pandera, dedupe, unstructured) and five different future-state ideas (entitlement drift ledger, contract clause compiler, legacy ERP cockpit, tenant identity fabric, shadow revenue journal). |
| 6 | `Executive_Automation_Catalog.md` | Markdown (265 lines) | Framework/methodology document — defines the deduplication process, ID conventions (SQL-RO-001, PY-DI-014, etc.), template fields, and catalog structure; contains almost no concrete tool content, just the scaffolding. |
| 7 | `Executive_Automation_Catalog_2.docx` | **Actually Markdown** (390 lines, 230 pipe-table rows) | Most table-dense variant — Part 1 with specific reconciliation/UI examples, Part 2 with 200+ SaaS tool directory in table form, Part 3 with five roadmap ideas (cross-system audit, LLM contract extraction, last-mile UI wrappers, predictive capacity auditor, revenue leakage intelligence). |
| 8 | `Executive_Automation_Catalog__Master_Reference_for___.docx` | **Actually Markdown** (171 lines) | Dense narrative version — Part 1 (unified SQL reconciliation, Pydantic/Instructor schema enforcement, VBA legacy bridge), Part 2 profiling five libraries (Splink, Instructor, dbt-audit-helper, VBA-Web, Pandera), Part 3 with five roadmap ideas; ends with "Technical Metadata for AI Reviewers" footer and 29 citations. |
| 9 | `report.md` | Markdown (356 lines) | Repository synthesis with full **working code samples** — SQL cross-database `INSTEAD OF` triggers, Python ETL with pandas/rapidfuzz, VBA Office orchestration — capped by a 7-row future-state roadmap table. |
| 10 | `report_extended.md` | Markdown (336 lines) | Longer relative of #9 — Part 1 code samples (Python/SQL/VBA), Part 2 longlist grouped by domain (ETL, RPA, Browser Automation, APIs, Observability, Data Quality), Part 3 with five future-state ideas. |
| 11 | `Exec_auto_master_fixed.pdf` | **Actually a ZIP archive** (6 JPEG pages + 6 OCR .txt + manifest.json) | Visual/OCR rendering of the same content as file #8 — **duplicate in a different wrapper**, not new material. |

**Flags found during inventory.**
- Files #7 and #8 are not Word documents despite `.docx` extension — python-docx cannot open them. They are plain UTF-8 markdown.
- File #11 is not a PDF despite `.pdf` extension — it is a Zip of rendered page images plus OCR text. pdfplumber errors on open.
- No content truncation detected. Every file reaches a natural endpoint.

---

### Full Review Results

Full curated review of Sections A (Universal Toolkit Additions), C (Future Ideas), and D (Skip). Section B (Video 4 Candidates / Finalists) is presented separately below.

#### Section A — Universal Toolkit Additions

*Tools that live alongside existing `modUTL_*` VBA modules or `UniversalToolkit\python\` scripts. Plug-and-play on any workbook or file.*

| # | Idea name | What it does | Lang | Effort | Why worth including | Source(s) | Overlap flag |
|---|---|---|---|---|---|---|---|
| A1 | **Formula Integrity Fingerprint** | Stores a hash of every formula in a named range at baseline, then re-hashes on demand and shows a diff report of any formulas that silently changed. Catches the classic "somebody overwrote a formula with a typed-in number" bug. | VBA | M | CFO-grade control: turns silent formula corruption into a one-click audit. Pairs beautifully with the existing audit-tools module. | CodexCodeIdeas.md (VBA-02) | Possible adjacency to existing audit tools — verify (OD-5) |
| A2 | **Workbook Policy Validator** | Reads a small config section (required sheet names, font = Arial, header row must be blue #0B4779, named ranges present, tab order, no external links) and writes a pass/fail report sheet. Standardizes every team workbook automatically. | VBA | M | Enforces iPipeline brand and structural standards across 2,000 coworkers without a style guide document nobody reads. | CodexCodeIdeas.md (VBA-09) | None |
| A3 | **Controlled Snapshot Sign-off** | Locks the workbook state at close / sign-off: captures a checksum of key ranges, approver name, timestamp, and workbook hash into a hidden log sheet — so the submitted version is provably the version that was approved. | VBA | M | Gives the CFO a defensible, auditable "this is what we signed" artifact with zero reliance on file-system versioning. | CodexCodeIdeas.md (VBA-07) | Possible adjacency to existing audit tools — verify (OD-6) |
| A4 | **Dependency Impact Preview** | Before any destructive macro runs, shows the user every downstream cell, chart, pivot, and sheet that will be affected. User confirms or cancels. Uses Range.Precedents / Dependents tracing. | VBA | S–M | "Know before you hit run" — builds confidence and prevents the most common workbook accidents. Great Video-3-style universal feel. | CodexCodeIdeas.md (VBA-05) | None |
| A5 | **Auto-Repair Suggestions (not auto-apply)** | Detects common data issues (merged cells, stray spaces, inconsistent date formats, #REF!, duplicated headers) and presents a menu of fix options — user chooses what to apply. | VBA | M | Frames Excel errors as guided remediation instead of cryptic failures. Zero hidden changes, which coworkers and auditors love. | CodexCodeIdeas.md (VBA-06) | Partial overlap with data sanitizer — verify (OD-7); this is the "diagnose + choose fix" layer on top |
| A6 | **Data Entry Fraud Pattern Flags** | Event-based `Worksheet_Change` log that flags manual overrides meeting suspicious patterns (round-number entries above a threshold, end-of-period edits, same user touching the same cell repeatedly). | VBA | L | Detective control layer for shared workbooks. CFO-grade talking point: "every manual override is evidence-logged." | CodexCodeIdeas.md (VBA-10) | None |
| A7 | **Intelligent Rollforward Assistant** | Rolls month / quarter tabs forward with preflight checks (formula references still valid, named ranges present, mappings intact) and a staged apply that can be undone if any check fails. | VBA | L | Eliminates the #1 period-open error: half-rolled workbooks where some references moved and others didn't. | CodexCodeIdeas.md (VBA-03) | Adjacent to `forecast_rollforward.py` but different scope (tab structure, not forecast data) — different tool |
| A8 | **Exception Workbench Sheet** | A standardized Excel tab auto-built in any workbook: import button pulls exceptions from any CSV/xlsx into a ranked table with owner, due date, status, notes columns. One place, not twenty. | VBA | M | Every analyst everywhere has an "exceptions" tab they built from scratch. This ships the canonical version. | CodexCodeIdeas.md (VBA-04) | None |
| A9 | **Workbook Dependency Scanner** | Parses every formula with openpyxl and exports a JSON map of cell-to-cell dependencies across sheets. Produces an HTML or Excel table showing "if you change A, these 47 things move." | Python | M | Answers the scariest pre-close question: "What breaks if I change this?" Works on any `.xlsx`. | CodexCodeIdeas.md (PY-08) | None |
| A10 | **Missing-Value & Type Monitor** | Reads any Excel / CSV and writes a one-sheet quality report: null-rate per column, type-mismatch rows, out-of-range values. Uses only pandas — no Great Expectations dependency. | Python | S | Deterministic, deliverable in under 50 lines, and frames "data quality" in a way coworkers understand instantly. | report_extended.md (Missing-Value Monitor); CodexCodeIdeas.md (PY-03 lineage) | Adjacent to `clean_data.py` — differentiate as an *audit report* rather than a cleaner |
| A11 | **Evidence Pack Generator** | Takes a finished job's logs, outputs, and input files, computes SHA-256 hashes, zips everything with a `manifest.json`, and writes a one-page index. Pure stdlib, zero install. | Python | S | Audit prep goes from hours to seconds. CFO / SOX line: "every delivery ships with provable evidence." | CodexCodeIdeas.md (PY-07) | None |
| A12 | **Narrative Variance Writer (deterministic)** | Reads a variance table and writes plain-English commentary from templates plus threshold rules ("Revenue missed plan by $X driven primarily by [largest driver]"). No LLM. Fully auditable language. | Python | M | Solves the "CFO pack needs written commentary" problem without any AI risk. Template library grows over time. | CodexCodeIdeas.md (PY-02) | Adjacent to `variance_analysis.py` and `word_report.py` — this is the *prose layer* they don't currently produce |
| A13 | **Finance Data Contract Checker** | User writes a small YAML file listing expected columns, types, allowed values, and row-count bounds for any feed. Python validator runs it on every inbound file and blocks the job if the contract breaks. | Python | M | Stops silent upstream schema changes from poisoning downstream reports. Contract is a plain text file a non-coder can edit. | CodexCodeIdeas.md (PY-03); Master Ref (1.2 adapted); deep-research-report_2.md (pandera pattern, simplified) | None |
| A14 | **Policy-as-Code Rule Runner** | Finance rules (tolerance thresholds, sign conventions, approval limits) live in a CSV or YAML. A Python runner reads the rules and applies them to any file. Changing a rule is a one-cell edit, not a code change. | Python | M | Lets non-developers maintain finance policy without ever touching code. Huge CFO story. | CodexCodeIdeas.md (SQL-07 adapted to Python) | None |
| A15 | **Duplicate-Record Fuzzy Finder (generic)** | Drop in any two files (or one file against itself), pick a column, get a ranked list of near-duplicates with similarity score and a match confidence bucket (exact / high / borderline). thefuzz + pandas. | Python | S–M | Universal dedup tool beyond customer names — works on vendor lists, GL accounts, cost centers, SKUs. | report_extended.md (Duplicate-Customer Detector); Exec_Catalog_2 (1.2 Fuzzy Record Linkage); Master Ref (2.1 Splink pattern); report.md (ETL reconcile); deep-research-report.md (RapidFuzz profile); CodexCodeIdeas.md (PY-05 lineage) | **Possible overlap with `fuzzy_lookup.py` — verify (OD-8)** |

---

#### Section C — Future Ideas (parked for post-demo work)

| # | Idea | What it would do | Lang | Parking reason | Source(s) |
|---|---|---|---|---|---|
| C1 | **Cross-System Entitlement Audit Ledger** | Reconciles what customers are contractually entitled to (from the CRM) against what they actually consumed (from product telemetry and cloud usage logs). Output is an exception ledger categorized as *under-billed / over-provisioned / orphan usage / expired contract still consuming / SKU mismatch*. | Python + SQL | Requires access to live Salesforce/billing/AWS data; infrastructure work, not a solo build. Perfect *next* deliverable after the demo series lands. | Exec_Catalog_2.docx (3.1); Master Ref (3.1); deep-research-report.md; deep-research-report_2.md; report_extended.md |
| C2 | **AI-Powered Contract Clause Extractor** | Same use case as B11, but uses an LLM to extract complex, non-templated clauses (SLA credit language, rebate triggers, data-processing commitments). One row per obligation, written to Excel. | Python | Violates "no external AI API calls." Ship B11 (rules-based) first; this becomes v2 when AI APIs are on-table. | Exec_Catalog_2.docx (3.2); Master Ref (3.2); deep-research-report.md; deep-research-report_2.md; report_extended.md |
| C3 | **Legacy-ERP REST Facade** | VBA workbook acts as the operator console; a middleware layer handles auth, retries, logging, and writes to the ERP (via API or screen automation) when native connectors fail. "Last-mile" integration. | VBA + Python | Multi-system plumbing + requires VBA-Web library (non-approved) and an external middleware host. Large engineering project. | Exec_Catalog_2.docx (3.3); Master Ref (3.3); deep-research-report.md; deep-research-report_2.md; report_extended.md; report.md |
| C4 | **Weekly Revenue Leakage Diagnostic** | Every Monday, walks every unresolved mismatch between CRM quotes, invoices, billing usage, support credits, and contract amendments; produces a ranked queue with proposed match, confidence, impact $, and evidence. | Python + SQL | Cross-system data access; meaningful only with live billing/CRM feeds. Park as a Phase-2 deliverable. | Exec_Catalog_2.docx (3.5); Master Ref (3.4); deep-research-report.md; deep-research-report_2.md |
| C5 | **SLA Credit + Renewal Risk Radar** | Combines incident timelines, support severity history, account terms, and renewal windows into a forward-looking risk score per account. | Python + SQL | Requires support-system and CRM data access; cross-functional build with customer success and support. | deep-research-report.md |
| C6 | **Tenant Identity Resolution Fabric** | Canonical customer map resolving customer, workspace, subscription, reseller, product, and invoice identities across every source system — one "golden record" powering every downstream report. | Python + SQL | Needs Splink (not approved) for probabilistic matching at scale, plus data-warehouse access. thefuzz alone won't match at this precision. See OD-10 for caveat. | deep-research-report_2.md; Master Ref (3.5) |
| C7 | **ML Forecast Ensemble Manager** | Combines multiple forecast models (ARIMA, Prophet, gradient boosting) with backtest-based weighting and a champion/challenger registry. | Python | Requires scikit-learn / Prophet / statsmodels — none approved. Park until the approved-package list expands. | CodexCodeIdeas.md (PY-04); report_extended.md |
| C8 | **SARIMA / Prophet Time-Series Pipeline** | End-to-end forecasting pipeline with seasonality detection and auto-tuning. | Python | Requires statsmodels / Prophet. Strong candidate when ML packages are green-lit. | report.md; report_extended.md |
| C9 | **Close Calendar Risk Predictor** | Predicts SLA-miss probability per close task using historical cycle data. | Python | Needs scikit-learn. Valuable but blocked by package policy. | CodexCodeIdeas.md (PY-06) |
| C10 | **Anomaly Detection Pipeline (isolation forest)** | SQL extracts daily metrics into feature tables; Python applies isolation forest to flag outliers in revenue / usage; alerts flow back to Excel. | Python + SQL | scikit-learn not approved. B9 (PSI-based drift monitor) is the approved-package cousin and covers most of the value now. | report_extended.md |
| C11 | **Controlled Action Approvals (PIN-gated macros)** | High-impact macros require a manager PIN or approval record before execution. Approval log is signed and auditable. | VBA | Technically buildable in pure VBA, but touches shared-workbook governance and needs an approval-table rollout plan — too much change management to bundle into the demo. | CodexCodeIdeas.md (VBA-01) |
| C12 | **Macro Runtime Telemetry Dashboard** | Reads the existing VBA_AuditLog sheet and renders runtime, error rate, and usage frequency by Command Center action as a KPI dashboard. | VBA | Depends on audit-log conventions being consistent across workbooks — which varies. Worth doing as a v2 of Command Center once logging format is canonicalized. | CodexCodeIdeas.md (VBA-08) |
| C13 | **Internal Exception-Status API** | Flask/FastAPI service Excel, VBA, and Python all read/write against — single source of truth for exception status instead of fragmented tabs. | Python | Flask/FastAPI not approved; also requires a hosted endpoint. Big architectural shift — right thing long-term, wrong thing for this quarter. | CodexCodeIdeas.md (OA-03) |
| C14 | **Pipeline Orchestrator (Airflow / Prefect / Dagster)** | Centrally schedules all Python + SQL + VBA jobs with dependencies, retries, alerting, and a UI. | Python | Requires Airflow or Prefect (not approved) and a host; also conflicts with the "no Task Scheduler" intent by introducing a heavier scheduler. Revisit when infrastructure is on the table. | deep-research-report.md; report_extended.md |
| C15 | **SOX Segregation-of-Duties Audit Query Pack** | Role-action matrix joins to flag conflicting permissions (e.g., same user posting and approving the same JE). | SQL | Needs direct access to the ERP/identity database and SOX scope alignment. Worth doing but is a controls-team project. | CodexCodeIdeas.md (SQL-09) |
| C16 | **Close Bottleneck Heatmap** | Decomposes lag per close step, entity, and user using event timestamps; produces a process-improvement roadmap dataset. | SQL | Needs close-process event telemetry that doesn't exist in Excel today; prerequisite is instrumenting the close. | CodexCodeIdeas.md (SQL-10) |
| C17 | **Cross-Database Referential Integrity Triggers** | `INSTEAD OF` triggers enforce FK-like rules across databases; audit tables record every change. | SQL | Requires SQL Server admin rights and prod DDL changes. This is a DBA deliverable, not a Finance-analyst tool. | report.md |
| C18 | **Office Scripts + Power Automate Close Trigger** | File lands in a "Ready for Close" SharePoint folder → Power Automate flow kicks off the Python/SQL close job. | Office Scripts + Power Automate | Power Automate is cloud-scheduled — same reliability-and-replicability problems as Task Scheduler for a non-developer audience. Revisit only if IT standardizes on it. | CodexCodeIdeas.md (OA-01) |

---

#### Section D — Skip (and why)

| # | Idea | Skip reason | Source(s) |
|---|---|---|---|
| D1 | **Monthly Billing Reconciler (RapidFuzz + pandas, basic)** | **Already covered.** Directly duplicates existing `two_file_reconciler.py` + `bank_reconciler.py`. B5 is the upgrade — ship that instead if you want the pattern. | report_extended.md; report.md |
| D2 | **Basic Two-List Fuzzy Match** | **Already covered** by `fuzzy_lookup.py`. A15 is the different-use-case variant (record-to-record dedup, not column lookup) — use A15 only if `fuzzy_lookup` doesn't already handle that mode. | report.md; Exec_Catalog_2.docx (1.2); Master Ref (2.1); deep-research-report.md; report_extended.md |
| D3 | **Custom Data Entry UserForm** | **Already covered** by the validation-builder VBA module. A UserForm with drop-downs and input validation is exactly what validation builder produces. | report.md |
| D4 | **Excel Dashboard Refresher (RefreshAll button)** | **Trivial + likely already in Command Center.** `ThisWorkbook.RefreshAll` is a 3-line macro. | report_extended.md |
| D5 | **Sales Stage Categorization (Select Case mapping)** | **Trivial.** A 10-line Select Case block. Better solved with a lookup table and XLOOKUP natively. | Exec_Catalog_2.docx (1.3) |
| D6 | **Automated Slide Generator (python-pptx)** | **Package not approved** (python-pptx not on the list). Existing `word_report.py` + `exec_brief` VBA cover the executive-deliverable use case using approved tooling. | report_extended.md |
| D7 | **Access → Excel → Word OLE Orchestration** | **Off-mission.** Brittle multi-app automation; Access isn't part of the iPipeline F&A stack. | report.md |
| D8 | **Interactive UserForm Dashboards (multi-page)** | **Already covered** by Command Center + exec-brief. A second dashboarding paradigm inside the same workbook fragments the UX. | report.md |
| D9 | **Cross-Platform RPA (pyautogui / selenium / pywinauto)** | **Packages not approved** and RPA targeting desktop UIs creates a maintenance nightmare for non-developers. | report.md |
| D10 | **SQLAlchemy ETL Job with warehouse loads** | **Package not approved** and the ETL pattern is already handled by existing `SQL query tool` + `consolidate_files.py` + `pnl_dashboard.py`. | report.md |
| D11 | **.NET / VSTO Signed Add-In** | **Off-mission for this series.** Requires Visual Studio, code signing, IT deployment — nothing a Finance analyst can replicate. | CodexCodeIdeas.md (OA-02) |
| D12 | **GitHub Actions Validation Bundle** | **Off-mission.** Developer CI/CD is not a coworker-facing tool. | CodexCodeIdeas.md (OA-05) |
| D13 | **dbt-style SQL Model Layer** | **Requires dbt** — warehouse-engineering tooling the audience can't run. | CodexCodeIdeas.md (OA-04); deep-research-report.md; Master Ref (2.3) |
| D14 | **Predictive Capacity Auditor (Spanner tenant migration)** | **Not relevant** — iPipeline F&A isn't operating Google Spanner. | Exec_Catalog_2.docx (3.4) |
| D15 | **The 200+ SaaS Product Directory (Salesforce, HubSpot, Stripe, QuickBooks, etc.)** | **Not automation ideas.** These are commercial products, not things you build. Same for workflow-engine, BI-tool, and RPA vendor lists scattered across files. | Exec_Catalog_2.docx (Part 2); 200_tools_catalog.md; report_extended.md (§2.2); deep-research-report.md; deep-research-report_2.md |

---

### Video 4 Finalists

*Section B — Python-focused ideas specifically suited to the "Python Automation for Finance" theme. Ordered by Video 4 narrative strength. All use approved packages only.*

| # | Idea name | What it does | Lang | Effort | Why it's a Video 4 candidate | Source(s) | Overlap flag |
|---|---|---|---|---|---|---|---|
| B1 | **Exception Triage Engine** | Reads exception rows from any CSV or xlsx (validation failures, reconciliation breaks, AR aging items), scores each by `impact × confidence × recency` using weights in a config file, and writes a ranked priority workbench to Excel with iPipeline styling. | Python | M | Perfect "chaos → clarity" visual for Video 4. Every F&A coworker has a messy exception list. Scoring weights are editable by non-coders. | CodexCodeIdeas.md (PY-01) | None — genuinely new |
| B2 | **Journal Entry Duplicate-Ring Detector** | Reads JE data, groups near-duplicate entries split across days, users, or entities within a tolerance window (amount ± $, date ± N days, vendor fuzzy-match). Flags suspicious rings for review. | Python | M | Fraud/SOX story lands hard with the CFO. Shows Python doing what no pivot table can. thefuzz + pandas only. | CodexCodeIdeas.md (SQL-01, adapted) | None |
| B3 | **Vendor Payment Velocity Monitor** | Reads AP history, computes rolling median + MAD per vendor, flags payments whose timing or amount deviates beyond a z-score threshold. Output is a ranked exceptions tab. | Python | M | AP fraud and duplicate-pay detection is a universal finance pain point. numpy + pandas, no scipy needed. | CodexCodeIdeas.md (SQL-03, adapted) | None |
| B4 | **Close Readiness Score (0–100)** | Points Python at a folder of period snapshot files, evaluates whether each required feed arrived on time with expected row counts and no validation failures, then writes a single per-entity readiness score plus a red/yellow/green summary. | Python | M | The close-pain video moment. "Is entity X ready to close?" becomes one click instead of a 30-minute meeting. | CodexCodeIdeas.md (SQL-02, adapted); SQL-06 complement | **Requires realistic sample data for the demo** |
| B5 | **Unified Reconciliation Engine (Compare-and-Classify)** | Full-outer-join pattern: every row gets labeled *Identical / Modified / Removed / Added*, and modifications include a field-level diff. Reads any two tabular files with a shared key column; writes a diff workbook. | Python | M | More rigorous successor to classic two-file compare. The four-status framework is CFO-briefable on one slide. | Master Ref 1.1; deep-research-report_2.md (dbt-audit-helper pattern) | **Adjacent to `two_file_reconciler.py` and `compare_files.py` — see OD-3** |
| B6 | **Root-Cause Reconciliation Assistant** | Pairs with B5. When a diff row is Modified or Removed, searches a historical break-log (Excel) for similar past breaks using thefuzz on the description column, and suggests the top-3 most likely cause categories with confidence scores. | Python | M | "Our reconciliation tool now learns from history without any AI." Pure lookup + fuzzy. | CodexCodeIdeas.md (PY-05) | Adjacent to `fuzzy_lookup.py` — reuses same primitive, different application |
| B7 | **Allocation Drift Tracker** | Reads monthly allocation percentage files, computes per-cost-center delta vs. prior month, flags cost centers whose allocation drifted beyond tolerance and requires a reason-code entry for each silent drift. | Python | M | Margin-governance story for CFO. Answers "why did margin move?" without a human hunt. | CodexCodeIdeas.md (SQL-04, adapted) | None |
| B8 | **Forecast Backtest Tracker** | Every time a forecast runs, Python writes the run's assumptions, outputs, and timestamp to a persistent log workbook. When actuals land, it computes error per run and produces a forecast-accuracy history (MAPE by entity/model). | Python | M | Objective model accountability. "Which forecast is actually right?" becomes a chart instead of an argument. | CodexCodeIdeas.md (SQL-08, adapted) | Adjacent to `pnl_forecast.py` — this is the accuracy-over-time layer, not a new forecaster |
| B9 | **Data Drift Monitor (PSI)** | Computes Population Stability Index on any numeric column vs. a baseline snapshot. Threshold bands (<0.1 stable / 0.1–0.25 shifting / >0.25 alert) plus a one-sheet drift report. Pure numpy + pandas. | Python | S–M | Proactive alert that answers: "Are any of our KPIs quietly drifting?" Strong CFO hook. | CodexCodeIdeas.md (PY-10, simplified to approved packages) | None |
| B10 | **Subledger Completeness Gate** | A tiny "guardrail" script: reads a control table of expected feeds (name, arrival window, row-count bounds) and returns PASS/FAIL before downstream steps run. Can be chained with B4. | Python | S | Prevents half-close accidents. 50-line tool, massive narrative weight ("we don't close on incomplete data anymore"). | CodexCodeIdeas.md (SQL-06, adapted) | None |
| B11 | **Vendor Contract Field Extractor (rules-based, no LLM)** | Point pdfplumber at a folder of vendor contracts, use regex templates (renewal date, notice period, auto-renew Y/N, term length) to extract structured fields into an Excel contract register. One row per contract. | Python | M | Contract-intake story without AI risk. Rules engine is editable by non-coders via a small pattern file. Very polished demo. | Exec_Catalog_2.docx (3.2, adapted); Master Ref (3.2, adapted); deep-research-report.md; report_extended.md | Adjacent to `pdf_extractor.py` — this is the *field-register* variant, probably new |
| B12 | **CFO Pack Assembly Pipeline** | Reads a config file listing which charts, tables, and commentary files go into this month's pack. python-docx + matplotlib build the final Word/PDF deliverable. Locks inputs via hash once approved so the released pack is reproducible. | Python | L | Ties every Video 4 tool together: B2–B11 produce artifacts, B12 assembles the monthly deliverable. Closes the loop. | CodexCodeIdeas.md (PY-09) | **Adjacent to `word_report.py` and `exec_brief` VBA — see OD-11** |

**Video 4 planning pushback (from session, preserved verbatim).**
1. **5–8 minutes is tight.** B1, B2, B5, or B11 each need 3–4 minutes to land for a non-technical audience (setup, problem, demo, result). Pick one flagship and montage the rest (30 seconds) OR split into two videos. A single 15-minute "Python Automation for Finance" covering B1 + B2 + B9 as three acts is probably stronger than a rushed 7-minute single-tool show.
2. **"Monday-morning usable" is a real constraint.** B4, B5, B11 require coworkers to have their own real data in a consistent shape. For each to actually get used on Monday, ship with: (a) a sample input file, (b) a 1-page README with exact file-placement instructions, and (c) a dry-run mode that generates a demo output using the sample. Otherwise downloads sit in the Downloads folder untouched.
3. **B5 almost certainly overlaps `two_file_reconciler.py` and `compare_files.py`.** Don't ship B5 alongside the old ones. Either supersede (B5 = v2, deprecate originals) or skip B5 entirely. Lean supersede: the 4-status classifier is a genuine upgrade. See OD-3.

---

### Best Ideas Curation

Top-10 picks across all sections, ranked by bang-for-buck (value ÷ effort) with an eye on what gives Video 4 a spine.

| Rank | ID | Idea | Lang | Effort | One-sentence rationale |
|---|---|---|---|---|---|
| 1 | B1 | **Exception Triage Engine** | Python | M | Flagship Video 4 candidate; cleanest before/after story; config-driven weights editable by non-coders. |
| 2 | A11 | **Evidence Pack Generator** | Python (stdlib) | S | Tiny script, enormous CFO/SOX talking weight: every deliverable ships with hashes, logs, and a manifest. |
| 3 | B9 | **Data Drift Monitor (PSI)** | Python | S–M | Proactive alerting using only numpy + pandas; answers "are any KPIs quietly drifting?" |
| 4 | B2 | **Journal Entry Duplicate-Ring Detector** | Python | M | Fraud/SOX language gets CFO attention faster than anything else in the catalog. |
| 5 | A1 | **Formula Integrity Fingerprint** | VBA | M | Solves the universal silent-formula-corruption problem with a hash baseline. CFO-grade control. |
| 6 | B10 | **Subledger Completeness Gate** | Python | S | 50 lines of code, massive narrative: "we don't run close on incomplete data anymore." |
| 7 | A4 | **Dependency Impact Preview** | VBA | S–M | Universal safety feature preventing the most common workbook accident. |
| 8 | A12 | **Narrative Variance Writer (deterministic)** | Python | M | Actually solves "CFO pack needs written commentary" with zero AI risk. |
| 9 | B8 | **Forecast Backtest Tracker** | Python | M | Rare accountability tool — finally answers "which forecast was right?" objectively. |
| 10 | B11 | **Contract Extractor (rules-based)** | Python | M | Flashy demo moment without AI; makes coworkers say "wait, Python can do that?" — swap with A3 if CFO weights audit posture over modernity (see OD-9). |

---

## 5. TOP PICKS CONSOLIDATED

Single flat list — highest-confidence recommendations. 15 items max.

1. **B1** Exception Triage Engine | Python | M | Video 4 flagship | Cleanest chaos→clarity visual; config-driven weights
2. **A11** Evidence Pack Generator | Python stdlib | S | Universal toolkit | Tiny script, massive audit/SOX talking weight
3. **B9** Data Drift Monitor (PSI) | Python | S–M | Video 4 | Proactive KPI drift alerting with numpy + pandas only
4. **B2** Journal Entry Duplicate-Ring Detector | Python | M | Video 4 | Fraud/SOX hook with CFO language; pandas + thefuzz
5. **A1** Formula Integrity Fingerprint | VBA | M | Universal toolkit | CFO-grade control against silent formula corruption
6. **B10** Subledger Completeness Gate | Python | S | Video 4 | 50-line guardrail preventing half-close accidents
7. **A4** Dependency Impact Preview | VBA | S–M | Universal toolkit | Pre-run preview of downstream impact on any workbook
8. **A12** Narrative Variance Writer (deterministic) | Python | M | Universal toolkit | Deterministic CFO-pack commentary, zero AI risk
9. **B8** Forecast Backtest Tracker | Python | M | Video 4 | Objective accuracy-over-time layer for existing forecast tools
10. **B11** Contract Extractor (rules-based) | Python | M | Video 4 | Polished demo moment; pdfplumber + regex templates
11. **A3** Controlled Snapshot Sign-off | VBA | M | Universal toolkit | Defensible close-snapshot artifact with checksum + approver log
12. **A9** Workbook Dependency Scanner | Python (openpyxl) | M | Universal toolkit | "What breaks if I change this?" on any .xlsx
13. **A14** Policy-as-Code Rule Runner | Python | M | Universal toolkit | Finance rules editable by non-coders via CSV/YAML
14. **B3** Vendor Payment Velocity Monitor | Python | M | Video 4 | Rolling median + z-score AP anomaly detection
15. **A6** Data Entry Fraud Pattern Flags | VBA | L | Universal toolkit | Worksheet-level override logging with suspicious-pattern detection

---

*End of handoff file. If any section is ambiguous or superseded, start with Section 3 (Open Decisions) before modifying Section 4 tables.*
