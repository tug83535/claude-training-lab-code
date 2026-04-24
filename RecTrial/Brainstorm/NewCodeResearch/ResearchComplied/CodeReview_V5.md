# iPipeline Finance Automation Demo — Master Brief for Claude Code

**Audience for this file:** You are Claude Code. You will read this file before touching code and return to it whenever a decision is ambiguous. Every section below is a rule set, a spec, or a reference table. None of it is filler.

**Orientation:**
- `HARD CONSTRAINTS` is law. Violations are blockers.
- `ALREADY BUILT` is the do-not-duplicate list. Check it before writing new code.
- `BUILD BACKLOG` is the authoritative list of what to build next, curated from a de-duplicated 156-idea inventory (see `APPENDIX A`).
- Every backlog item has a stable ID (`T-##`, `V4-##`, `F-##`). Use those IDs in commits, branch names, and handoffs.
- When a request is vague, resolve it against this file first. Only then ask the user.

---

## 1. PROJECT CONTEXT

| Field | Value |
|---|---|
| Company | iPipeline — ~2,000-person life-insurance SaaS |
| User | Connor Atlee, Finance & Accounting analyst, non-developer, reads code at working level |
| Deliverable | 4-video internal demo series for 2,000+ coworkers + CFO/CEO |
| Current state | Videos 1–2 recorded. Video 3 mid-debug after silent-wrapper refactor. Video 4 ready to record. |
| Stack | VBA · Python · SQL (see `HARD CONSTRAINTS` for exact package list) |
| Repo | `tug83535/claude-training-lab-code` |
| Active branch (parent project) | `claude/resume-ipipeline-demo-qKRHn` |
| Delivery folder | `FinalExport/` (single source of truth for all shipped assets) |

**Audience assumption for all generated artifacts:** zero coding background. If a feature cannot be explained in one plain-English sentence, simplify or split it.

---

## 2. HARD CONSTRAINTS

**Every rule below is non-negotiable.** If a user request would violate one, stop and flag it.

### 2.1 Forbidden
- No external AI API calls (OpenAI, Claude API, Gemini, etc.) — deterministic logic only
- No Outlook or email automation
- No Windows Task Scheduler dependencies
- No `scikit-learn`, no `statsmodels`, no `prophet`, no `rapidfuzz`, no `pyyaml`
- No hardcoded sheet names, file paths, or row numbers — everything plug-and-play
- No Command Prompt exposure in any demo-facing script

### 2.2 Approved Python packages (only these)
`pandas`, `openpyxl`, `pdfplumber`, `python-docx`, `thefuzz`, `numpy`, `matplotlib`, `xlwings`, Python standard library.

If you need something outside this list, use stdlib instead. Example: YAML contracts → use JSON (`json` module); SARIMA forecasts → use `numpy` rolling averages; fuzzy matching → use `thefuzz` (already approved), not `rapidfuzz`.

### 2.3 Branding — apply to every styled output
| Element | Value |
|---|---|
| Primary Blue | `#0B4779` / `RGB(11, 71, 121)` |
| Navy | `#112E51` |
| Innovation Blue | `#4B9BCB` |
| Lime accent | `#BFF18C` |
| Aqua accent | `#2BCCD3` |
| Arctic White | `#F9F9F9` |
| Charcoal | `#161616` |
| Font family | Arial only |

**Styled header rule:** iPipeline Blue band, white bold Arial text.
**Excel output default:** `.xlsx` via `openpyxl`, board-ready formatting.

### 2.4 Quality bar — every artifact
- Header comment block with `PURPOSE`, `WHY THIS IS NOT NATIVE`, `USE CASE`
- No hardcoded secrets — VBA reads named ranges; Python reads env vars
- Plug-and-play: works on any coworker's file, no setup
- Comments explain *why*, not *what* (reader is Finance, not a dev)

---

## 3. ALREADY BUILT — DO NOT RE-IMPLEMENT

Before writing anything, check these lists. If a requested feature matches, ask the user whether to extend or replace instead of duplicating.

### 3.1 VBA — 23 modules, ~140 tools
Sanitizer, compare, consolidate, highlights, pivot tools, tab organizer, column ops, sheet tools, comments, validation builder, lookup builder, command center, exec brief, finance tools, audit tools, branding.

### 3.2 Python — 28 scripts
`aging_report`, `bank_reconciler`, `compare_files`, `forecast_rollforward`, `fuzzy_lookup`, `pdf_extractor`, `variance_analysis`, `variance_decomposition`, `clean_data`, `consolidate_files`, `multi_file_consolidator`, `date_format_unifier`, `two_file_reconciler`, `sql_query_tool`, `word_report`, `batch_process`, `regex_extractor`, `unpivot_data`, `pnl_forecast`, `pnl_dashboard`, `master_data_mapper`, `profile_workbook`, `sanitize_dataset`, `compare_workbooks`, `build_exec_summary`, `variance_classifier`, `scenario_runner`, `sheets_to_csv`.

### 3.3 SQL — 4 scripts
Staging, transformations, validations, enhancements.

---

## 4. VIDEO 4 — PRIMARY DECISION

### 4.1 The three combos under evaluation

| Combo | What it is | Strength | Weakness |
|---|---|---|---|
| 1. Finance Copilot Menu | Numbered CLI menu wrapping existing Python scripts in one entry point | Fastest to build; reuses what exists | Shows Command Prompt — intimidating for non-dev audience |
| 2. Excel Button Edition (xlwings) | Excel buttons trigger Python silently; results appear as new sheets | Excel-native; zero Command Prompt; CFO-friendly | More engineering upfront |
| 3. Hero Demo + Cookbook | One dramatic end-to-end hero demo + 5 copy-pasteable recipe scripts | Strongest story; cookbook extends reach | Recipes re-introduce Command Prompt for coworkers running them |

### 4.2 Recommended approach — **Combo 2 as Video 4, with Combo 3 as a companion SharePoint drop**

Combo 2 is the correct primary for this audience because the audience lives in Excel, xlwings is on the approved list, and "click button → new sheet appears" is the most CFO-legible experience any Python demo can deliver. Combo 1's Command Prompt surface is a demo killer for a 2,000-person audience. Combo 3's hero moment is valuable, but its recipes hit the same Command Prompt wall the moment coworkers try them locally.

The bridge: record Video 4 as the Button Edition (6–8 minutes, six feature buttons, every result lands as a new sheet in the workbook), and ship a companion `FinalExport/Video4_Cookbook/` folder containing the 5 recipe scripts from Combo 3 as plain Python files with one-page walkthroughs — for the small subset of coworkers who want to go deeper. This preserves the hero-demo value without polluting the main video with Command Prompt.

**If runtime flexibility is exercised:** split into Video 4a (Button Edition, 8 min) + Video 4b (Hero Demo narrative deep-dive on one flagship feature, 6–8 min). Given the existing scripts in `ALREADY BUILT` already cover the hero material, the split is cheap.

### 4.3 Ordered build list for Video 4 Button Edition

Build in this order. Each row corresponds to a backlog ID in Section 6.

| Order | Button Label (Excel UI) | Backlog ID | Script it wraps | Output sheet |
|---|---|---|---|---|
| 1 | Score This Sheet | V4-01 | Data Quality Scorecard | `UTL_QualityScore` |
| 2 | Tag Material Rows | V4-02 | Materiality Classifier | `UTL_Materiality` |
| 3 | Explain Each Row | V4-03 | Exception Narrative Generator | `UTL_Narratives` |
| 4 | Compare To Prior File | V4-04 | Zero-Install Workbook Compare | `UTL_Compare` |
| 5 | Run Scenarios | V4-05 | Zero-Install Scenario Runner | `UTL_Scenarios` |
| 6 | Build Exec Summary | V4-06 | Executive Summary Builder | `UTL_ExecSummary` |

All six are in `ALREADY BUILT` or the Toolkit backlog — zero greenfield. The xlwings wrapper is the only new code surface for Video 4.

---

## 5. WORKFLOW — WHEN USER SAYS "BUILD X"

1. **Resolve X against `BUILD BACKLOG` (Section 6).** If a backlog ID matches, use that spec as authoritative.
2. **Cross-check against `ALREADY BUILT` (Section 3).** If a match exists, ask user: extend, replace, or skip.
3. **Verify against `HARD CONSTRAINTS` (Section 2).** If any constraint is violated, stop and flag — do not silently substitute.
4. **Choose the path:**
   - VBA → `FinalExport/UniversalToolkit/vba/modUTL_*.bas`
   - Python → `FinalExport/UniversalToolkit/python/*.py`
   - SQL → `FinalExport/UniversalToolkit/sql/*.sql`
5. **Write the file with the Section 2.4 header block.**
6. **Apply branding from Section 2.3 to any styled output.**
7. **Add a one-page plain-English guide** in `FinalExport/Guides_v2/` using the same voice as existing guides.
8. **Commit with the backlog ID in the message:** `T-03: add Data Quality Scorecard`.

---

## 6. BUILD BACKLOG — CURATED 60 PICKS (from 156-idea inventory)

Drawn from `APPENDIX A`. Every ID here is actionable under current constraints. Items that required banned packages in the raw inventory have been reframed inline — notes flag those.

### 6.1 Section A — Toolkit additions (universal, plug-and-play)

Ideas that belong in the Universal Toolkit and work on **any** coworker file. Priority order within the table.

| ID | Idea | Language | Pass 1 # | Effort | Notes for Claude Code |
|---|---|---|---|---|---|
| T-01 | Header Row Auto-Detect | VBA | #4 | S | Foundational — builds on this for T-02, T-03, T-06. Implement first. |
| T-02 | Materiality Classifier | VBA | #1 | S | Depends on T-01. Configurable $ and % thresholds, auto-detects Current/Prior columns. |
| T-03 | Data Quality Scorecard | VBA | #3 | S | 0–100 score from blanks/errors. Writes `UTL_QualityScore` tab. Branded header. |
| T-04 | Exception Narrative Generator | VBA | #2 | S | Depends on T-02. Plain-English row commentary from Materiality Status column. |
| T-05 | Quick Row Compare Count | VBA | #5 | S | Hash-based pre-check before full compare. Returns mismatch count. |
| T-06 | Run Receipt Sheet | VBA | #6 | S | Appends timestamped execution receipt to `UTL_RunReceipt` tab on every macro run. |
| T-07 | Cover Show-Tools Button Installer | VBA | #7 | S | One-time macro. Adds branded blue launcher button to Cover sheet → Command Center. |
| T-08 | Intelligence Category in Command Center | VBA | #8 | S | Pins T-02, T-04, T-03 near top of tool list. |
| T-09 | Zero-Install Workbook Profiler | Python | #9 | S | stdlib only — no pip install. Inventories sheets, ranges, VBA flag. First Python demo asset. |
| T-10 | Word Report Talking Points | Python | #10 | S | Adds `--talking-points` flag to existing `word_report.py`. 3–5 auto-built CFO bullets from variance data. Template-based, no LLM. |
| T-11 | Formula Integrity Fingerprinting | VBA | #29 | M | Hash critical formula ranges at baseline; compare on demand. Surfaces silent formula drift. |
| T-12 | Macro Runtime Telemetry Dashboard | VBA | #31 | M | Reads existing `VBA_AuditLog` sheet. Surfaces runtime, error rate, usage frequency per action. |
| T-13 | Controlled Snapshot Sign-off | VBA | #32 | M | Captures checksum + approver metadata + timestamp at monthly sign-off. Writes to a locked log sheet. |
| T-14 | Intelligent Rollforward Assistant | VBA | #76 | M | Preflight checks before month rollforward. Staged apply with undo. Prevents setup errors. |
| T-15 | Workbook Policy Validator | VBA | #79 | M | Enforces naming standards, required sheets, tab order, Arial font, brand colors. Emits compliance report sheet. |
| T-16 | Dependency Impact Preview | VBA | #77 | M | Before a destructive action runs, show which cells/charts will change. Popup summary. |
| T-17 | Auto-Repair Suggestions | VBA | #78 | M | Menu of fix options for detected data issues — user picks. Never auto-applies. |
| T-18 | FISCAL_YEAR Startup Check | VBA | #86 | S | On workbook open, compare `modConfig.FISCAL_YEAR_4` to current year. Show one-time warning if mismatched. |
| T-19 | Quick Demo Mode Macro | VBA | #87 | S | One button auto-runs 5 marquee features back-to-back. Use for the "can you show me in 2 minutes?" ask. |
| T-20 | What's New Sheet | VBA | #88 | S | In-workbook change log tab. Every version bump appends a row. Returning coworkers read this first. |

### 6.2 Section B — Video 4 candidates

Python-first, approved packages, designed to wrap in xlwings buttons for Combo 2. Items V4-01 through V4-06 are the six buttons in Video 4's ordered build list (Section 4.3).

| ID | Idea | Language | Pass 1 # | Role in demo | Notes |
|---|---|---|---|---|---|
| V4-01 | Data Quality Scorecard | Python | #3, #65 | Button 1 | Python variant of T-03 for xlwings wrapping. Emits new `UTL_QualityScore` sheet. |
| V4-02 | Variance Classifier (Zero-Install) | Python | #12 | Button 5 companion | stdlib. Labels rows Over / Under / On-target vs baseline. Already built — wrap in xlwings. |
| V4-03 | Exception Narrative Generator | Python | #2, #69 | Button 3 | Template-based (no LLM). Writes `UTL_Narratives` sheet with one narrative per material row. |
| V4-04 | Zero-Install Workbook Compare | Python | #11 | Button 4 | stdlib. Two-file row-by-row diff → `UTL_Compare` sheet. Existing script — wrap only. |
| V4-05 | Scenario Runner (Zero-Install) | Python | #13 | Button 5 | stdlib. Percentage shocks to a metric column → scenario sheets. |
| V4-06 | Executive Summary Builder | Python | #15 | Button 6 | stdlib. Markdown exec summary from workbook tabs. Lands as `UTL_ExecSummary` sheet + optional .md export. |
| V4-07 | Sheets-to-CSV Batch Export | Python | #14 | Cookbook | Recipe script. One sheet per CSV. Good "Python bridge" demo. |
| V4-08 | Exception Triage Engine | Python | #17 | Cookbook / Video 4b | Ranks exceptions by impact × confidence × recency. Config-driven weights in JSON (not YAML). |
| V4-09 | Control Evidence Pack Generator | Python | #18 | Cookbook / Video 4b | Zips macro logs + validation outputs + manifest into an audit bundle. No external deps. |
| V4-10 | Finance Data Contract Checker | Python | #19 | Cookbook | **Reframed:** use JSON contracts (`json` stdlib), not YAML — `pyyaml` is not approved. |
| V4-11 | Workbook Dependency Scanner | Python | #20 | Video 4b hero candidate | `openpyxl` formula parser → impact graph JSON/HTML. Visually striking end-to-end. |
| V4-12 | Close Readiness Score View | SQL | #16 | Cookbook / companion | Per-entity 0–100 score. Best CFO-language output in the whole catalog. |
| V4-13 | Exception Workbench Sheet | VBA | #30 | Bridge artifact | Imports the outputs of V4-07/V4-08 into a tracked workbench tab. Closes the Python → Excel loop. |

### 6.3 Section C — Future (parked, post-demo)

Real value, but post-Video 4. Group by type for scanability.

#### SQL controls and marts
| ID | Idea | Pass 1 # | Notes |
|---|---|---|---|
| F-01 | Allocation Drift Tracker | #21 | Monthly delta view with tolerances and required reason codes. |
| F-02 | Forecast Backtest Warehouse | #22 | Three tables: run, assumption, actual. Enables MAPE tracking. |
| F-03 | Subledger Completeness Control Matrix | #23 | Gate close steps on expected feed times and row-count bounds. |
| F-04 | Workbook-to-Source Reconciliation Mart | #24 | Standardized recon tables + variance reason taxonomy. |
| F-05 | Vendor Payment Velocity Baselines | #25 | Rolling medians per vendor. MAD/z-score thresholds. |
| F-06 | Journal Entry Duplicate Ring Detection | #26 | Similarity windows on amount/date/vendor/account. Use `thefuzz`, not `rapidfuzz`. |
| F-07 | Close Bottleneck Heatmap | #27 | Lag decomposition by step/entity/user from event timestamps. |
| F-08 | Segregation-of-Duties Audit Pack | #28 | Role-action matrix joins + exception materialized views. |
| F-09 | Policy-as-Code Rule Engine Tables | #68 | Metadata-driven rule catalog + dynamic execution proc. |
| F-10 | Cohort Retention Matrix | #60 | SQL-side only. Good for finance + CS crossover. |

#### Python engines
| ID | Idea | Pass 1 # | Notes |
|---|---|---|---|
| F-11 | Forecast Ensemble Manager | #70 | Backtest-weighted combination of forecast models. No ML libs — use `numpy`-based baselines. |
| F-12 | Root Cause Reconciliation Assistant | #71 | Deterministic rules + `thefuzz` similarity against historical break/resolution log. |
| F-13 | CFO Pack Assembly Pipeline | #73 | Release-tagged assembly of approved charts, tables, commentary. `python-docx` + `matplotlib`. |
| F-14 | Data Drift Monitor Service | #74 | **Reframed:** use `numpy` rolling-window distribution compare, not `scipy` PSI/KS. |
| F-15 | Lightweight Internal Exception Status API | #83 | Flask or stdlib `http.server`. Single source for exception state across VBA/Python/Excel. |

#### VBA workflow
| ID | Idea | Pass 1 # | Notes |
|---|---|---|---|
| F-16 | Controlled Action Approvals | #75 | Manager PIN / approval record before high-impact macros run. |
| F-17 | Data Entry Fraud Pattern Flags | #80 | Event log of manual cell edits + rule windows scoring suspicious patterns. |
| F-18 | Approval Stamp and Audit Trail Writer | #109 | Records who approved each adjustment. Complements T-13. |
| F-19 | Multi-Workbook Diff | #55 | Extends existing `compare_files` to N files. |
| F-20 | Renewal Alert Engine | #54 | Scans contract dates, raises window-based alerts in-workbook. |
| F-21 | Invoice PDF Generator | #51 | Branded invoice PDFs from workbook rows. Uses `pdfplumber` for templates (stdlib for output). |
| F-22 | One-Click Board Pack Builder | #112 | Button-driven refresh + assembly. |
| F-23 | Git-Friendly VBA Module Exporter | #113 | On-save export of `.bas` modules to source folder. Enables Git diffs. |
| F-24 | Guided Adjustment Wizard | #107 | UserForm walk-through for approving billing variances. |
| F-25 | Legacy ERP Export Cleaner | #110 | Reshape CSV exports into standardized tables. |

#### Platform and quality
| ID | Idea | Pass 1 # | Notes |
|---|---|---|---|
| F-26 | dbt-Style Model Layer for Finance SQL | #84 | Adopt dbt or dbt-inspired SQL transformations. |
| F-27 | GitHub Actions Validation Bundle | #85 | Lint + tests + data contract checks on every push. |

---

## 7. TOP 10 PICKS + COMBO RECOMMENDATION

Highest bang-for-buck across all 156 inventory items, weighted for demo impact and CFO-level message:

| Rank | ID | Idea | Why it wins |
|---|---|---|---|
| 1 | V4-12 | Close Readiness Score View | One number per entity. Highest-leverage CFO-language artifact in the whole catalog. |
| 2 | V4-08 | Exception Triage Engine | Every analyst uses it every month-end. Direct workflow improvement. |
| 3 | T-03 / V4-01 | Data Quality Scorecard | Visual, zero-setup, perfect live demo moment. |
| 4 | V4-09 | Control Evidence Pack Generator | Cuts audit-prep hours. Leadership cares deeply about audit speed. |
| 5 | T-02 | Materiality Classifier | Transforms any flat sheet into a risk-tiered view in seconds. |
| 6 | T-10 | Word Report Talking Points | AI-style output with zero AI calls — on-brand for the constraint story. |
| 7 | V4-04 | Zero-Install Workbook Compare | Plug-and-play Python on a locked laptop = convincing for coworkers. |
| 8 | T-01 | Header Row Auto-Detect | Foundational helper — makes every other tool reliable. |
| 9 | T-12 | Macro Runtime Telemetry Dashboard | Shows toolkit is production-grade, not just a demo. |
| 10 | V4-13 | Exception Workbench Sheet | Creates one action hub everyone can use after the demo — turns a video into a habit. |

**Video 4 combo recommendation — Combo 2 (Excel Button Edition) as the primary video, with Combo 3's cookbook scripts (V4-07, V4-08, V4-09, V4-10, V4-11) shipped as a companion SharePoint folder.** Combo 2 is the only option that survives the "2,000-person non-developer audience + C-suite on the call" test: xlwings is on the approved package list, Excel is where the audience already lives, and buttons-to-new-sheets is CFO-legible in a way that Command Prompt output will never be. Combo 1 fails the audience test the moment the terminal appears. Combo 3's hero arc is a strength, so absorb it by splitting into Video 4a (Button Edition) + optional Video 4b (Hero Demo, running V4-11 end-to-end on a real sample workbook) if runtime flexibility is exercised — Video 4b is cheap because every underlying script already exists.

---

## APPENDIX A — Full Pass 1 Inventory (156 unique ideas)

Reference only. Raw, deduplicated list from all 13 source files. Use when a user asks about an idea not in the curated backlog.

| # | Idea Name | One-Line Description | Language | Source File(s) |
|---|-----------|----------------------|----------|----------------|
| 1 | Materiality Classifier | Tags each row as Material increase/decrease, Watch, or Normal using configurable $ and % thresholds with auto-detected Current/Prior columns. | VBA | BranchIdeasReview_April2026.md |
| 2 | Exception Narrative Generator | Writes plain-English row narratives based on a Materiality Status column to produce CFO-ready wording automatically. | VBA | BranchIdeasReview_April2026.md |
| 3 | Data Quality Scorecard | Scores a sheet 0–100 from blanks and errors and writes a formatted quality report tab. | VBA | BranchIdeasReview_April2026.md |
| 4 | Header Row Auto-Detect | Scans top rows and picks the most-likely header row, removing the need for hardcoded row numbers. | VBA | BranchIdeasReview_April2026.md |
| 5 | Quick Row Compare Count | Fast pre-check that hashes rows and returns mismatch count before running a full compare. | VBA | BranchIdeasReview_April2026.md |
| 6 | Run Receipt Sheet | Writes a timestamped execution receipt to a UTL_RunReceipt tab on every macro run for audit evidence. | VBA | BranchIdeasReview_April2026.md |
| 7 | Cover Show-Tools Button Installer | One-time macro that adds a branded blue launcher button to the Cover sheet pointing to Command Center. | VBA | BranchIdeasReview_April2026.md |
| 8 | Intelligence Category in Command Center | Pins Materiality, Narratives, and Scorecard tools to the top of the Command Center tool list. | VBA | BranchIdeasReview_April2026.md |
| 9 | Zero-Install Workbook Profiler | Inventories workbook sheets, ranges, and VBA flag with stdlib only, no pip install required. | Python | BranchIdeasReview_April2026.md |
| 10 | Word Report Talking Points Flag | Adds a talking-points flag to word_report.py that auto-builds three-to-five CFO narrative bullets. | Python | BranchIdeasReview_April2026.md |
| 11 | Zero-Install Workbook Compare | Compares two workbooks row by row and exports diffs to CSV using stdlib. | Python | BranchIdeasReview_April2026.md |
| 12 | Zero-Install Variance Classifier | Labels rows as Over, Under, or On-target vs a baseline using rules only. | Python | BranchIdeasReview_April2026.md |
| 13 | Zero-Install Scenario Runner | Applies percentage shocks to a metric column and exports all scenarios. | Python | BranchIdeasReview_April2026.md |
| 14 | Sheets-to-CSV Batch Export | Exports every sheet in a workbook to its own CSV file. | Python | BranchIdeasReview_April2026.md |
| 15 | Executive Summary Builder | Builds a Markdown executive summary from CSV outputs with stats and highlights. | Python | BranchIdeasReview_April2026.md |
| 16 | Close Readiness Score View | SQL view returning a per-entity 0–100 close readiness score from failed checks, missing feeds, and late postings. | SQL | BranchIdeasReview_April2026.md, CodexCodeIdeas.md |
| 17 | Exception Triage Engine | Ranks exceptions by impact × confidence × recency using config-driven weights. | Python | BranchIdeasReview_April2026.md, CodexCodeIdeas.md |
| 18 | Control Evidence Pack Generator | Packages macro logs and validation results into a zipped audit evidence bundle with manifest; overlaps SOX evidence collector. | Python | BranchIdeasReview_April2026.md, CodexCodeIdeas.md, CODE_CATALOG.md |
| 19 | Finance Data Contract Checker | Validates incoming data against YAML schema/quality contracts before downstream use. | Python | BranchIdeasReview_April2026.md, CodexCodeIdeas.md |
| 20 | Workbook Dependency Scanner | Parses formulas and named ranges to map a change-impact graph as JSON or HTML. | Python | BranchIdeasReview_April2026.md, CodexCodeIdeas.md |
| 21 | Allocation Drift Tracker | Detects silent drift in cost allocation percentages month-over-month with threshold flags and reason codes. | SQL | BranchIdeasReview_April2026.md, CodexCodeIdeas.md |
| 22 | Forecast Backtest Warehouse | Stores every forecast run, its assumptions, and realized actuals for accuracy comparison. | SQL | BranchIdeasReview_April2026.md, CodexCodeIdeas.md |
| 23 | Subledger Completeness Control Matrix | Checks that all required upstream feeds are present before close steps execute. | SQL | BranchIdeasReview_April2026.md, CodexCodeIdeas.md |
| 24 | Workbook-to-Source Reconciliation Mart | Reconciles workbook aggregates against warehouse source-of-truth tables with variance reason taxonomy. | SQL | BranchIdeasReview_April2026.md, CodexCodeIdeas.md |
| 25 | Vendor Payment Velocity Baselines | Flags abnormal timing or amount shifts by vendor using rolling medians and MAD/z-score thresholds. | SQL | BranchIdeasReview_April2026.md, CodexCodeIdeas.md |
| 26 | Journal Entry Duplicate Ring Detection | Finds near-duplicate journal entry patterns split across users, days, vendors, or entities. | SQL | BranchIdeasReview_April2026.md, CodexCodeIdeas.md |
| 27 | Close Bottleneck Heatmap | Decomposes where close-cycle delays occur by step, entity, and user using event timestamp lag. | SQL | BranchIdeasReview_April2026.md, CodexCodeIdeas.md |
| 28 | Segregation-of-Duties Audit Pack | Flags conflicting role and action combinations in the transaction lifecycle via role-action matrix joins. | SQL | BranchIdeasReview_April2026.md, CodexCodeIdeas.md |
| 29 | Formula Integrity Fingerprinting | Hash-checks critical formula zones to catch silent changes against a stored baseline. | VBA | BranchIdeasReview_April2026.md, CodexCodeIdeas.md |
| 30 | Exception Workbench Sheet | Central Excel tab for assigning, tracking, and closing exceptions with owner/due-date workflow. | VBA | BranchIdeasReview_April2026.md, CodexCodeIdeas.md |
| 31 | Macro Runtime Telemetry Dashboard | Summarizes run times, error rates, and usage frequency by Command Center action from audit log. | VBA | BranchIdeasReview_April2026.md, CodexCodeIdeas.md |
| 32 | Controlled Snapshot Sign-off | Captures approved monthly workbook state with checksum, approver metadata, and read-only snapshots. | VBA | BranchIdeasReview_April2026.md, CodexCodeIdeas.md, Executive_Automation_Catalog.md |
| 33 | Outlook Mail Merge with Attachments | VBA module that runs a mail merge with file attachments through Outlook. | VBA | BranchIdeasReview_April2026.md, CODE_CATALOG.md |
| 34 | Calendar Appointment Builder | VBA module that creates Outlook calendar appointments from workbook data. | VBA | BranchIdeasReview_April2026.md, CODE_CATALOG.md |
| 35 | JIRA Bridge and Weekly Digest | Pulls JIRA tickets into Excel and emits a weekly digest summary. | combo | BranchIdeasReview_April2026.md, CODE_CATALOG.md |
| 36 | Slack Notifier | VBA module that posts messages to Slack via webhook. | VBA | BranchIdeasReview_April2026.md, CODE_CATALOG.md |
| 37 | Teams Notifier and Webhook on Threshold | VBA module plus Office Script that posts Teams messages when a metric crosses a threshold. | combo | BranchIdeasReview_April2026.md, CODE_CATALOG.md |
| 38 | AWS Cost Optimizer | Python script that analyzes AWS usage and flags cost-saving opportunities. | Python | BranchIdeasReview_April2026.md, CODE_CATALOG.md |
| 39 | Customer Churn Risk Scorer | ML-based Python script that scores customer churn likelihood. | Python | BranchIdeasReview_April2026.md, CODE_CATALOG.md |
| 40 | Support Ticket Triage | ML-based Python script that classifies and prioritizes inbound support tickets. | Python | BranchIdeasReview_April2026.md, CODE_CATALOG.md |
| 41 | AD Inactive User Audit | PowerShell script that finds stale Active Directory accounts. | PowerShell | CODE_CATALOG.md, BranchIdeasReview_April2026.md |
| 42 | SharePoint Site Storage Audit | PowerShell script that inventories SharePoint site storage usage. | PowerShell | CODE_CATALOG.md, BranchIdeasReview_April2026.md |
| 43 | SSL Cert Expiry Monitor | PowerShell script that checks SSL certificate expiration dates. | PowerShell | CODE_CATALOG.md, BranchIdeasReview_April2026.md |
| 44 | New Hire Account Provisioner | PowerShell script that provisions new user accounts across systems. | PowerShell | CODE_CATALOG.md, BranchIdeasReview_April2026.md |
| 45 | Purchase Approval Flow | Power Automate flow template for multi-step purchase approval routing. | Power Automate | CODE_CATALOG.md, BranchIdeasReview_April2026.md |
| 46 | New Hire Onboarding Flow | Power Automate flow template for new-hire onboarding steps. | Power Automate | CODE_CATALOG.md, BranchIdeasReview_April2026.md |
| 47 | REST API Pagination (Power Query) | Power Query M script that handles paginated REST API responses. | Power Query | CODE_CATALOG.md |
| 48 | Multi-Folder Merge (Power Query) | Power Query M script that merges files across multiple folders into one table. | Power Query | CODE_CATALOG.md |
| 49 | Daily Metrics Export (Office Script) | TypeScript Office Script that exports daily metrics on a schedule. | Office Scripts | CODE_CATALOG.md |
| 50 | Bulk Format Data Import (Office Script) | TypeScript Office Script that bulk-formats imported data. | Office Scripts | CODE_CATALOG.md |
| 51 | Invoice PDF Generator | VBA module that generates branded invoice PDFs from workbook rows. | VBA | CODE_CATALOG.md |
| 52 | SharePoint Sync | VBA module that uploads and downloads files between Excel and SharePoint. | VBA | CODE_CATALOG.md |
| 53 | SQL Server Runner from Excel | VBA module that executes SQL Server queries directly from an Excel workbook. | VBA | CODE_CATALOG.md |
| 54 | Renewal Alert Engine | VBA module that scans contract data and raises renewal alerts based on date windows. | VBA | CODE_CATALOG.md |
| 55 | Multi-Workbook Diff | VBA module that diffs multiple workbooks and reports differences. | VBA | CODE_CATALOG.md |
| 56 | Folder Organizer | VBA module that auto-organizes files into folder structures based on rules. | VBA | CODE_CATALOG.md |
| 57 | SaaS ARR Waterfall | Python script that builds an ARR waterfall from contract and churn data. | Python | CODE_CATALOG.md |
| 58 | License Utilization Analyzer | Python script that compares purchased seats to actual active users. | Python | CODE_CATALOG.md |
| 59 | API SLO Tracker | Python script that measures API uptime and latency against SLO thresholds. | Python | CODE_CATALOG.md |
| 60 | Cohort Retention Analyzer / Matrix | Produces a cohort retention table by signup month and months-since-start. | combo | CODE_CATALOG.md, Executive_Automation_Catalog.md |
| 61 | Email to Structured Data | Script that turns unstructured support emails into structured records via LLM or regex. | Python | CODE_CATALOG.md, report.md |
| 62 | Git Developer Metrics | Python script that computes developer activity metrics from Git history. | Python | CODE_CATALOG.md |
| 63 | SaaS Metrics Suite | SQL script bundle that computes core SaaS metrics (MRR, ARR, churn, etc.). | SQL | CODE_CATALOG.md |
| 64 | Sales Pipeline Velocity | SQL script that calculates pipeline velocity by stage and rep. | SQL | CODE_CATALOG.md |
| 65 | Data Quality Audit SQL | SQL script that audits data quality across warehouse tables. | SQL | CODE_CATALOG.md |
| 66 | Revenue Recognition Schedule / Waterfall | SQL or Python routine that generates a monthly revenue recognition schedule from multi-year contracts. | combo | CODE_CATALOG.md, Executive_Automation_Catalog.md |
| 67 | Slow Query Tuner | SQL script that identifies slow queries and suggests optimizations. | SQL | CODE_CATALOG.md |
| 68 | Policy-as-Code Rule Engine Tables | Metadata-driven rule catalog table plus dynamic execution procedure that reads finance policy rules at runtime. | SQL | CodexCodeIdeas.md |
| 69 | Narrative Variance Writer | Python template library that generates draft commentary using deterministic templates (no LLM). | Python | CodexCodeIdeas.md |
| 70 | Forecast Ensemble Manager | Combines multiple forecast models with backtest-based weighting and a champion/challenger registry. | Python | CodexCodeIdeas.md |
| 71 | Root Cause Reconciliation Assistant | Proposes likely cause categories for reconciliation breaks using deterministic rules plus similarity against past resolved issues. | Python | CodexCodeIdeas.md |
| 72 | Close Calendar Risk Predictor | Predicts SLA miss probability for each close task using a lightweight ML baseline. | Python | CodexCodeIdeas.md |
| 73 | CFO Pack Assembly Pipeline | Compiles approved charts, tables, and commentary into one monthly release artifact with release tagging. | Python | CodexCodeIdeas.md |
| 74 | Data Drift Monitor Service | Monitors distribution drift in critical metrics via PSI and KS tests and alerts when thresholds trip. | Python | CodexCodeIdeas.md |
| 75 | Controlled Action Approvals | Requires manager PIN or approval record before high-impact macros execute. | VBA | CodexCodeIdeas.md |
| 76 | Intelligent Rollforward Assistant | Rolls month tabs with formula and mapping preflight checks plus undo capability. | VBA | CodexCodeIdeas.md |
| 77 | Dependency Impact Preview | Traces precedents and dependents and surfaces a summary popup before an action executes. | VBA | CodexCodeIdeas.md |
| 78 | Auto-Repair Suggestions | Recommends fix options for detected data issues as a menu the user chooses from (no auto-apply). | VBA | CodexCodeIdeas.md |
| 79 | Workbook Policy Validator / Template Enforcer | Enforces naming standards, required sheets, tab order, and font/color standards and emits a compliance report. | VBA | CodexCodeIdeas.md, Executive_Automation_Catalog.md |
| 80 | Data Entry Fraud Pattern Flags | Logs manual cell edits and scores suspicious timing/threshold patterns as a detective control. | VBA | CodexCodeIdeas.md |
| 81 | Office Scripts + Power Automate Close Trigger | Uses an Office Script to trigger Python or SQL runs when files reach controlled states. | combo | CodexCodeIdeas.md |
| 82 | .NET Add-In for Signed Enterprise Deployment | Moves critical controls into a signed managed .NET add-in for orgs that block unsigned macros. | .NET/VSTO | CodexCodeIdeas.md |
| 83 | Lightweight Internal Exception Status API | Flask or FastAPI service that Excel, VBA, and Python all read and write for a single exception state. | Python | CodexCodeIdeas.md |
| 84 | dbt-Style Model Layer for Finance SQL | Adopts a dbt or dbt-inspired pattern for versioned, tested, documented SQL transformations. | SQL | CodexCodeIdeas.md, deep-research-report.md |
| 85 | GitHub Actions Validation Bundle | Runs lint, tests, and data contract checks on every push via CI. | combo | CodexCodeIdeas.md |
| 86 | FISCAL_YEAR Startup Check | Startup routine that compares modConfig FISCAL_YEAR against the current year and warns on mismatch. | VBA | GitAgentIdeas2.md |
| 87 | Quick Demo Mode Macro | One-button macro that auto-runs five marquee features back-to-back for two-minute ad hoc demos. | VBA | GitAgentIdeas2.md |
| 88 | What's New Sheet | In-workbook change log tab that records what changed and when for returning coworkers. | VBA | GitAgentIdeas2.md |
| 89 | Monthly Billing Reconciler | Reconciles invoiced amounts against usage and entitlement tables for the billing period with flagged variances. | combo | Executive_Automation_Catalog.md, report_extended.md |
| 90 | Entitlement vs Usage / Revenue Leakage Engine | Cross-system control that reconciles CRM entitlements, contract limits, usage logs, and billing into an exception ledger with categorized leakage types. | combo | Executive_Automation_Catalog.md, Executive_Automation_Catalog_2.docx, Executive_Automation_Catalog__Master_Reference_for___.docx, Exec_auto_master_fixed.pdf, deep-research-report.md, deep-research-report_2.md, report.md, report_extended.md |
| 91 | Churn and Downgrade Revenue Tracker | SQL script that calculates ARR impact of churns, downgrades, and expansions at contract line level. | SQL | Executive_Automation_Catalog.md |
| 92 | Cross-System Key Consistency Checker | Detects orphaned or mismatched keys between CRM, billing, and ERP. | SQL | Executive_Automation_Catalog.md |
| 93 | Generic Table Audit Trail Trigger | Generic trigger-based audit table that captures before/after DML images and adapts to schema changes. | SQL | Executive_Automation_Catalog.md, report.md |
| 94 | Idempotent Replay Ledger | Tracks processed upstream events to prevent double-processing. | SQL | Executive_Automation_Catalog.md |
| 95 | ARR/NRR Fact Table Builder | Transforms raw contract and billing data into a star-schema ARR/NRR fact table. | SQL | Executive_Automation_Catalog.md |
| 96 | Finance Close Data Snapshotter | Takes end-of-month immutable snapshots of key finance tables with timestamps. | SQL | Executive_Automation_Catalog.md |
| 97 | Excel Billing Pack Generator | Produces standardized Excel billing packs from SQL result sets. | Python | Executive_Automation_Catalog.md |
| 98 | Revenue Recognition Simulator | Runs what-if simulations on revenue schedules under different recognition policies. | Python | Executive_Automation_Catalog.md, CODE_CATALOG.md |
| 99 | Usage Aggregation Orchestrator | Ingests raw usage logs and aggregates into billing-ready tables. | Python | Executive_Automation_Catalog.md |
| 100 | Cross-System Reconciliation Runner | Python orchestrator that runs SQL checks and emits consolidated Excel or CSV reports. | Python | Executive_Automation_Catalog.md |
| 101 | Contract Metadata Validator | Applies rule-based and regex validations to exported contract CSVs. | Python | Executive_Automation_Catalog.md |
| 102 | Schema Drift Monitor | Compares current database schemas against a stored baseline and flags drift. | Python | Executive_Automation_Catalog.md |
| 103 | Fleet/Network Automation Patterns | Reusable patterns for managing fleets of systems and executing standardized command sets. | Python | Executive_Automation_Catalog.md |
| 104 | Multi-Tenant Reporting Engine | Parameterized engine that runs a bundle of SQL queries per tenant and renders Excel or PDF outputs. | Python | Executive_Automation_Catalog.md |
| 105 | Slack/Email Distribution Bot | Pushes finalized Excel or PDF dashboards to Slack channels or email lists. | Python | Executive_Automation_Catalog.md |
| 106 | Git-Driven Report Definition Loader | Loads report definitions from a Git repo for version-controlled reporting. | Python | Executive_Automation_Catalog.md |
| 107 | Guided Adjustment Wizard | Excel UserForm that walks finance users through reviewing and approving billing variances. | VBA | Executive_Automation_Catalog.md |
| 108 | Multi-Workbook Consolidator | Macro that ingests multiple regional billing workbooks into a central master model. | VBA | Executive_Automation_Catalog.md |
| 109 | Approval Stamp and Audit Trail Writer | Captures who approved each adjustment and writes a row to an audit log. | VBA | Executive_Automation_Catalog.md |
| 110 | Legacy ERP Export Cleaner | Cleans and reshapes CSV exports from legacy ERPs into standardized tables. | VBA | Executive_Automation_Catalog.md |
| 111 | Multi-App Office Automation Patterns | Macro patterns that orchestrate Excel, Word, Outlook, Access, and PowerPoint in one flow. | VBA | Executive_Automation_Catalog.md, report.md |
| 112 | One-Click Board Pack Builder | Button-driven macro that refreshes queries and assembles a full board pack. | VBA | Executive_Automation_Catalog.md |
| 113 | Git-Friendly VBA Module Exporter | Exports VBA modules to text files on save events for Git version control. | VBA | Executive_Automation_Catalog.md |
| 114 | LLM Contract/Vendor PDF Extractor to Excel | Python pipeline that OCRs PDFs, extracts clauses via an LLM into a strict schema, and writes obligations as structured Excel rows. | Python | Executive_Automation_Catalog.md, Executive_Automation_Catalog_2.docx, Executive_Automation_Catalog__Master_Reference_for___.docx, Exec_auto_master_fixed.pdf, deep-research-report.md, deep-research-report_2.md, report.md, report_extended.md, CODE_CATALOG.md |
| 115 | Legacy ERP Bridge / REST Facade | VBA-plus-API last-mile bridge that calls REST endpoints (or an AI-driven legacy UI wrapper) so Excel can read and write legacy ERP fields. | combo | Executive_Automation_Catalog.md, Executive_Automation_Catalog_2.docx, Executive_Automation_Catalog__Master_Reference_for___.docx, Exec_auto_master_fixed.pdf, deep-research-report.md, deep-research-report_2.md, report.md, report_extended.md |
| 116 | Revenue Risk / SLA Credit + Renewal Risk Radar | Hybrid SQL plus Python engine that fuses incident timelines, entitlement drift, and renewal windows into a forward-looking risk score. | combo | Executive_Automation_Catalog.md, deep-research-report.md |
| 117 | Unified Close Orchestrator | SQL materializes close-ready tables, Python orchestrates stored procedures and quality checks, and VBA provides the workbook trigger UI. | combo | Executive_Automation_Catalog.md |
| 118 | Tenant Identity Resolution / Customer 360 | Canonical tenant map resolving customer, subscription, product, and invoice identity across systems using deterministic plus active-learning matching. | combo | Executive_Automation_Catalog__Master_Reference_for___.docx, Exec_auto_master_fixed.pdf, deep-research-report_2.md, CODE_CATALOG.md |
| 119 | Shadow Revenue Journal | Append-only operational ledger recording quotes, orders, amendments, credits, provisioning, usage, and cash events in one causal stream. | combo | deep-research-report_2.md |
| 120 | Predictive Capacity Auditor | Analyzes tenant primary-key usage to predict Spanner/DB scaling bottlenecks and suggest tenant migration. | Python | Executive_Automation_Catalog_2.docx, Exec_auto_master_fixed.pdf |
| 121 | Anomaly Detection Pipeline | Combines SQL feature extraction with Python models (isolation forest, Prophet) to flag outliers in revenue/usage. | combo | report_extended.md |
| 122 | Auto-Tuned Forecasting Service | Microservice that trains and selects the best forecasting model (SARIMA, Prophet, LSTM) automatically. | Python | report.md |
| 123 | Headless RPA Orchestration | Python script that drives legacy desktop apps and web portals via pyautogui, selenium, and pywinauto for UI-only systems. | Python | report.md |
| 124 | Distributed Cross-DB Referential Integrity Service | INSTEAD OF triggers plus stored procs that enforce referential integrity across multiple databases and log violations. | SQL | report.md |
| 125 | SARIMA Time-Series Forecasting Pipeline | End-to-end statsmodels SARIMA forecasting pipeline with train/test split and plotting. | Python | report.md |
| 126 | Custom Data Entry UserForm | Excel UserForm with validation, drop-down population from lookup sheets, and write-back to a table. | VBA | report.md |
| 127 | Interactive Dashboard UserForms | Rich VBA UserForms with multi-page navigation, dynamic charts, and tree views that mimic modern web UIs. | VBA | report.md |
| 128 | Multi-Tenant Data Segmentation | SQL views and stored procs that partition shared tables into per-tenant schemas with row-level security and on-demand rebuilds. | SQL | report_extended.md |
| 129 | Subscription MRR Forecast with Churn Decay | SQL aggregation of MRR with exponential decay based on churn_rate to forecast next month. | SQL | report_extended.md |
| 130 | Duplicate-Customer Levenshtein Detector | SQL CTE that computes pairwise Levenshtein distance on normalized customer names and returns pairs below a threshold. | SQL | report_extended.md |
| 131 | Missing-Value Monitor | Python Great Expectations script that validates critical columns for nulls and writes an integrity report to Excel. | Python | report_extended.md |
| 132 | Automated Slide Generator | Python script using python-pptx that builds a KPI PowerPoint deck from a CSV. | Python | report_extended.md |
| 133 | Excel Dashboard Refresher | VBA macro that refreshes every pivot table and data connection in the workbook. | VBA | report_extended.md |
| 134 | xlwings Python-Excel UDF | xlwings-decorated Python function exposed as an Excel user-defined function (e.g., amortization schedule). | combo | report_extended.md |
| 135 | Sales Funnel Orchestration | VBA macro that categorizes sales stages into standardized labels and refreshes the workbook. | VBA | Executive_Automation_Catalog_2.docx, Exec_auto_master_fixed.pdf |
| 136 | Compare-and-Classify SQL Reconciliation Pattern | CTE pattern using full outer joins with presence indicators to tag rows Identical, Modified, Added, or Removed. | SQL | Executive_Automation_Catalog__Master_Reference_for___.docx, Exec_auto_master_fixed.pdf |
| 137 | Pydantic Gatekeeper Pattern | Pydantic BaseModel plus Instructor-style schema enforcement with validation retry loops for multi-shape payloads. | Python | Executive_Automation_Catalog__Master_Reference_for___.docx, Exec_auto_master_fixed.pdf, deep-research-report.md |
| 138 | Apache Airflow DAG Orchestration | Airflow DAGs with dynamic task mapping (expand/partial) wrapping SQL and Python steps for monthly pipelines. | Python | deep-research-report.md, report_extended.md |
| 139 | Great Expectations Checkpoint + Data Docs | Expectation Suite → Validation Definition → Checkpoint → action_list pattern producing human-readable Data Docs. | Python | deep-research-report.md, deep-research-report_2.md, report_extended.md |
| 140 | RapidFuzz Two-Stage Fuzzy Reconciliation | Exact-key match first, then RapidFuzz process.cdist/extractOne scoring with confidence buckets for the residue. | Python | deep-research-report.md, Executive_Automation_Catalog_2.docx, Exec_auto_master_fixed.pdf, report.md, report_extended.md |
| 141 | dbt-utils union_relations / deduplicate / unique_combination_of_columns | Warehouse-side dbt-utils macros for ragged-union stitching, explicit deduplication, and composite-key uniqueness tests. | SQL | deep-research-report.md |
| 142 | sqlglot Cross-Dialect SQL Normalizer | Python wrapper using sqlglot parse/transpile/diff to standardize SQL, detect drift, and extract lineage. | Python | deep-research-report_2.md |
| 143 | Pandera DataFrame Contracts | Class-based Pandera DataFrameModel with column, groupby, and DataFrame-wide Check rules around imports and exports. | Python | deep-research-report_2.md, Executive_Automation_Catalog__Master_Reference_for___.docx, Exec_auto_master_fixed.pdf |
| 144 | Active-Learning Entity Resolution (dedupe / Splink) | ML-based entity resolution using dedupe or Splink with human review loop in Excel for uncertain pairs. | Python | deep-research-report_2.md, Executive_Automation_Catalog__Master_Reference_for___.docx, Exec_auto_master_fixed.pdf |
| 145 | unstructured partition_pdf Section Extractor | Document ETL using unstructured with hi_res strategy, table inference, and title-aware chunking before schema validation. | Python | deep-research-report_2.md |
| 146 | dbt-audit-helper Compare-and-Classify Macro | Native dbt package macros for reconciling legacy queries against new dbt models during migrations. | SQL | Executive_Automation_Catalog__Master_Reference_for___.docx, Exec_auto_master_fixed.pdf |
| 147 | Instructor Structured LLM Extraction | Instructor library with Pydantic response_model pattern for validated, type-safe LLM outputs. | Python | Executive_Automation_Catalog__Master_Reference_for___.docx, Exec_auto_master_fixed.pdf |
| 148 | VBA-Web REST Framework | VBA-Web WebClient/WebRequest/WebResponse modular framework adding REST and OAuth 2.0 capability to Excel and Access. | VBA | Executive_Automation_Catalog__Master_Reference_for___.docx, Exec_auto_master_fixed.pdf |
| 149 | Branch Synthesis / Dedup Harness | Python script that recursively inventories, fingerprints, scores, and collapses duplicate SQL/Python/VBA across branches into a single catalog. | Python | deep-research-report.md, deep-research-report_2.md |
| 150 | LLM-Driven Unstructured Data ETL | LLM-based pipeline that extracts structured records from emails, PDFs, or chat transcripts with deterministic validation. | Python | report.md, report_extended.md |
| 151 | Fuzzywuzzy ETL Customer Reconciler | Python ETL skeleton using fuzzywuzzy to reconcile customer records across CRM and ERP into a warehouse table. | Python | report.md |
| 152 | Access-to-Excel-to-Word Report Macro | VBA macro that runs an Access query, exports to Excel, and generates a Word report from a template in one pass. | VBA | report.md |
| 153 | Office-to-Mainframe Integration | VBA macro that uses Windows API and COM objects to push and pull data between Excel and SAP GUI or AS/400 emulators. | VBA | report.md |
| 154 | VBA Version Control Helpers | Workbook save-event export/import of VBA components to a source folder for Git-based version control. | VBA | Executive_Automation_Catalog.md |
| 155 | External Python/SQL Library Reference Catalog | Curated inventory of 200+ open-source Python, SQL, and infrastructure libraries grouped by function for developer reference. | reference | 200_tools_catalog.md, deep-research-report.md, report_extended.md |
| 156 | VBA Ecosystem Library Catalog | Longlist of VBA libraries (stdVBA, VBA-Web, Rubberduck, VBA-FastJSON, WebView2 for Excel, etc.) as a discovery layer for VBA projects. | VBA | deep-research-report.md |

---

## APPENDIX B — Decision rules for ambiguous requests

If a user request maps to something in this list, follow the adjacent rule without re-deriving.

| Request pattern | Rule |
|---|---|
| "Add LLM / AI to X" | Refuse under HARD CONSTRAINT 2.1. Offer the deterministic template equivalent (e.g., T-10 for narrative, V4-03 for explanations). |
| "Use package Y" where Y ∉ approved list | Refuse. Substitute from approved list. Note the substitution in code comments. |
| "Email this out" / "send to Outlook" | Refuse under 2.1. Offer file-based output (Word via `python-docx`, Excel via `openpyxl`, Markdown via stdlib). |
| "Schedule this to run nightly" | Refuse Task Scheduler. Offer a manual trigger (button in Excel via xlwings, or a one-line CLI recipe in the cookbook). |
| "Build <thing already in Section 3>" | Ask user: extend, replace, or skip. Do not silently duplicate. |
| "Make it work on this specific sheet" | Refuse hardcoding. Use Header Row Auto-Detect (T-01) or ask for a column contract. |
| "Port this to scikit-learn / statsmodels / prophet" | Refuse. Use `numpy` rolling statistics or a `thefuzz`-based heuristic instead. |

---

## APPENDIX C — Voice and quality reference

Every artifact shipped from this repo is written in a voice that a non-developer Finance coworker can read. Concrete rules:

- **Header comments:** three labeled blocks — `PURPOSE`, `WHY THIS IS NOT NATIVE`, `USE CASE`. Each one to three sentences.
- **Inline comments:** explain *why*, not *what*. Reader is Finance, not an engineer.
- **Variable names:** business-first. `current_month_revenue` beats `cmr` every time.
- **Error messages:** full sentences. No stack-trace-only failures in demo-facing code.
- **Plain-English guides:** every new feature gets a one-page `.md` in `FinalExport/Guides_v2/` in the same voice as the existing 00–10 numbered guides. Start with a real problem, end with "you just did X."
- **Branding:** any styled output applies Section 2.3. Default Excel header: Primary Blue band, white bold Arial. No exceptions.

---

*End of brief. If a request cannot be resolved against any section of this file, ask the user — do not infer.*
