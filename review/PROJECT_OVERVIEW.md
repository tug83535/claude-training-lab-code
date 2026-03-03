# iPipeline P&L Demo — Full Project Overview
**For External Claude Review**
**Date:** 2026-03-02
**Project:** APCLDmerge — iPipeline Finance & Accounting Demo

---

## ⚠ READ FIRST — Two Review Prompts

This document is written for a separate Claude account to review. There are two distinct prompts below. You should answer **both independently**. Read the full document before answering either.

---

### PROMPT 1 — Demo Focus Cut

> You are reviewing a Finance & Accounting demo project built in Excel with VBA macros, Python scripts, and SQL. The demo will be presented live to the CFO, CEO, and 2,000+ employees at iPipeline. It will also be recorded as a video walkthrough.
>
> The project currently has **62 Command Center actions** across 32 VBA modules, plus 14 Python scripts, 4 SQL scripts, and a second demo phase planned around 76 Universal Tool candidates.
>
> **Your job for Prompt 1:** Identify the single best set of actions and features to actually show in the live demo. Assume the demo is 10–15 minutes maximum. The audience is non-technical Finance & Accounting staff, plus executive leadership. They do not care how the code works — they care about what it does and how impressive it looks.
>
> Recommend:
> - Which specific Command Center actions (out of 62) are the most visually impressive and relevant to show
> - Which Python and SQL capabilities are worth mentioning or demoing (even briefly)
> - Whether the Universal Tools second demo is worth including or should be cut for time
> - A suggested demo flow/order that builds toward a strong finish
>
> Do NOT suggest cutting the codebase itself — the full project stays. This prompt is only about what to show in the demo.

---

### PROMPT 2 — Overall Project Cut

> You are reviewing a Finance & Accounting demo project that has grown very large over many development sessions. The full inventory is listed below: 32 VBA modules, 14 Python scripts, 4 SQL scripts, and 76 Universal Tool candidates.
>
> **Your job for Prompt 2:** Evaluate the entire project honestly and recommend what should be cut, deferred, or simplified permanently — not just for the demo, but for the overall health and achievability of the project.
>
> Consider:
> - Which modules or scripts appear redundant or overlapping in purpose
> - Which features are too technical or complex for a Finance & Accounting audience who will actually use this
> - Which Universal Tool candidates are low-value or duplicate tools already in the main project
> - Whether 76 Universal Tool candidates is a realistic build target, and what a smarter focused list would look like
> - What the project should look like if you had to ship it cleanly in the next 30 days
>
> Be direct and honest. It is better to have 20 great things than 80 half-finished things.
>
> **Do NOT suggest cutting these** — the user has already permanently declined them and does not want them brought up again:
> - Backup Workbook with Timestamp macro
> - Cell Change / Timestamp Audit Trail on edits
> - Export Charts to PowerPoint
> - Email/Outlook integration

---
---

## SECTION 1 — What This Project Is

**Project Name:** APCLDmerge — iPipeline P&L Demo
**File:** KeystoneBenefitTech_PL_Model.xlsm (Excel macro-enabled workbook)

**What it is:**
A fully automated Profit & Loss financial model built in Excel. The workbook has a **Command Center** — a professional UserForm with buttons for 62 actions. Finance staff open the file, click a button, and the macro runs instantly. No coding knowledge required.

**The goal:**
Demonstrate to the CFO, CEO, and 2,000+ iPipeline employees that Finance & Accounting has built world-class automation tools. Show that Excel can do things most people don't know are possible. Inspire other Finance teams to adopt similar approaches.

**Who it's for:**
- **Primary audience:** Non-technical Finance & Accounting staff at iPipeline
- **Demo audience:** CFO, CEO, and 2,000+ employees
- **Sharing plan:** Share the final `.xlsm` file directly. Coworkers open it and use the Command Center. All 32 modules are already embedded in the file.

**Demo format:**
- Live walkthrough presentation
- Recorded video for coworkers who can't attend live

**Second demo phase (planned — not built yet):**
A separate set of "Universal Tools" — VBA and Python tools that work on ANY Excel file, not just this one. These would be packaged as an Add-In. Currently in candidate list phase — 76 tools identified, no code written yet.

---

## SECTION 2 — Current Status

| Item | Status |
|------|--------|
| VBA Modules | 32 modules imported into workbook — Debug > Compile passes clean |
| Python Scripts | 14 scripts complete and functional |
| SQL Scripts | 4 scripts complete |
| Command Center Actions | 62 actions wired and routing correctly |
| T1 Testing | T1.01–T1.07 complete (all PASS). T1.08 (pip install) pending. |
| T2–T4 Testing | Not yet started (live action testing) |
| Universal Tools | Candidate list only — 76 tools identified, no code written |
| Demo video | Not yet started |
| Coworker training guide | Not yet started |

---

## SECTION 3 — Main Demo: VBA Modules (32 Modules)

These are all imported into the Excel workbook. Coworkers never see this code — they only see the Command Center buttons.

### Core Infrastructure (always run in background — not shown in demo)

| Module | Purpose |
|--------|---------|
| `modConfig_v2.1.bas` | Single source of truth for all constants — sheet names, product names, fiscal year, colors, thresholds. Every other module reads from here. |
| `modLogger_v2.1.bas` | Runtime audit log — writes timestamped entries to a hidden sheet (VBA_AuditLog) every time an action runs. |
| `modPerformance_v2.1.bas` | TurboMode — disables screen updates and auto-calc during macros for fast execution. Also has a precision timer for benchmarking. |
| `modFormBuilder_v2.1.bas` | Builds the Command Center UserForm at runtime and routes all 62 actions to the correct module. This is the central switchboard. |
| `modMasterMenu_v2.1.bas` | Fallback InputBox menu (4 pages, 62 items) if the UserForm is not installed. |
| `frmCommandCenter_code.txt` | Code-behind for the Command Center UserForm — 62 action buttons, search filter, category navigation, Run/Run+Close buttons. |

### Navigation & Presentation (Command Center actions)

| Module | Purpose |
|--------|---------|
| `modNavigation_v2.1.bas` | Jump to any sheet instantly via table of contents. GoHome button. Keyboard shortcuts (Ctrl+Shift+M for menu). Toggle Executive Mode (hides internal sheets). |
| `modDemoTools_v2.1.bas` | Add control buttons directly to the Report tab. Set a parameterized print area. Generate and print a clean executive summary. |

### Data Quality & Cleaning (Command Center actions)

| Module | Purpose |
|--------|---------|
| `modDataQuality_v2.1.bas` | Six-scan data quality check — finds text-stored numbers, blanks, duplicates, formatting errors. Optionally auto-fixes them. |
| `modDataSanitizer_v2.1.bas` | Numeric-only sanitizer — converts text-stored numbers and fixes floating-point precision errors. Never touches dates, names, or ID columns. |
| `modDataGuards_v2.1.bas` | Pre-run safety guards — validates Assumptions tab is populated, confirms allocation shares sum to 100%, scans for negatives/zeros/suspicious round numbers. |
| `modSearch_v2.1.bas` | Search any keyword across all sheets simultaneously. Returns up to 200 results with clickable navigation. Results highlighted yellow. |

### Financial Analysis (Command Center actions)

| Module | Purpose |
|--------|---------|
| `modVarianceAnalysis_v2.1.bas` | Month-over-month variance detection — flags lines >15% change, auto-generates plain-English commentary for top variances. |
| `modSensitivity_v2.1.bas` | What-if sensitivity analysis — varies Assumptions drivers ±10% and ±20%, shows revenue and margin impact, outputs a tornado table. |
| `modAWSRecompute_v2.1.bas` | AWS cost allocation recalculation — reads revenue shares, verifies they sum to 100%, recalculates per-product AWS cost allocations. |
| `modForecast_v2.1.bas` | Rolling forecast — builds a 3-month rolling average forecast, appends completed month actuals to the trend sheet. |
| `modTrendReports_v2.1.bas` | Trend reporting and archiving — creates a rolling 12-month P&L view, builds a reconciliation trend chart, saves dated snapshots. |

### Reconciliation & Audit (Command Center actions)

| Module | Purpose |
|--------|---------|
| `modReconciliation_v2.1.bas` | Automated reconciliation runner — runs all PASS/FAIL checks on the Checks sheet, cross-sheet validations, outputs summary report. |
| `modDrillDown_v2.1.bas` | Reconciliation drill tools — adds hyperlinks from Checks rows to GL source data, applies heatmap coloring to reconciliation items, runs a golden file compare baseline. |
| `modAuditTools_v2.1.bas` | Audit and compliance tools — appends change log entries, finds and fixes external links, audits hidden sheets, creates a masked/redacted copy, exports error summary to clipboard. |
| `modIntegrationTest_v2.1.bas` | Integration test suite — 18 automated tests covering sheets, formulas, macros, and data integrity. Also has a quick health check mode. |

### Reporting & Export (Command Center actions)

| Module | Purpose |
|--------|---------|
| `modDashboard_v2.1.bas` | Dynamic chart and dashboard generation — builds/refreshes 3 charts on Report tab, executive dashboard, waterfall chart, product comparison, small multiples grid. |
| `modPDFExport_v2.1.bas` | Batch PDF export — exports selected sheets to a single professional PDF with print settings, company headers, and date stamps. |

### Scenario & Version Management (Command Center actions)

| Module | Purpose |
|--------|---------|
| `modScenario_v2.1.bas` | Scenario management — save Assumptions snapshots as named scenarios, load any scenario instantly, compare two scenarios side-by-side, delete. |
| `modVersionControl_v2.1.bas` | Version control — save timestamped workbook snapshots with metadata, compare versions, restore a prior version, list version history. |

### Operations & Admin (Command Center actions)

| Module | Purpose |
|--------|---------|
| `modImport_v2.1.bas` | Data import pipeline — import GL data from CSV or Excel, validate required columns, check for duplicates, append or replace existing data. |
| `modAllocation_v2.1.bas` | Cost allocation engine — reads GL transactions, applies allocation shares from Assumptions tab, outputs cost breakdown by product and department. |
| `modConsolidation_v2.1.bas` | Multi-entity consolidation — loads P&L data from external entity files, consolidates to master, manages intercompany eliminations. |
| `modAdmin_v2.1.bas` | Auto-documentation and change management — generates documentation of workbook structure, tracks modification requests, produces status reports. |
| `modMonthlyTabGenerator_v2.1.bas` | Monthly tab generator — clones the March template to create April–December tabs. Calendar-aware next-month prep marks the upcoming month column yellow. |
| `modUtilities_v2.1.bas` | 12 quick utility macros — delete blank rows, unhide all, freeze panes, autofit, protect/unprotect, and other workbook maintenance actions (Command Center actions 51–62). |
| `modETLBridge_v2.1.bas` | Python ETL bridge — triggers the Python data pipeline from a button inside Excel, then imports the cleaned ETL output back into the workbook. |

---

## SECTION 4 — Main Demo: Python Scripts (14 Scripts)

These run from the command line or from a button in Excel (via modETLBridge). Finance staff do not need to write Python — the scripts are pre-built.

| Script | Purpose |
|--------|---------|
| `pnl_config.py` | Shared configuration — single source of truth for file paths, constants, and utility functions. Imported by all other scripts. |
| `pnl_cli.py` | Master CLI — single command-line interface for the entire toolkit. One command replaces 10+ separate script calls. |
| `pnl_runner.py` | Unified entry point — simplified command syntax for running any P&L operation (even easier than pnl_cli for daily use). |
| `pnl_month_end.py` | Month-end close automation — validates GL, verifies allocations, runs all checks, generates plain-English commentary, produces the full close package. |
| `pnl_forecast.py` | Rolling forecast model — multiple methods: moving average, exponential smoothing, trend/seasonality decomposition. Outputs scenarios. |
| `pnl_monte_carlo.py` | Monte Carlo risk simulation — runs 10,000 P&L scenarios with randomized inputs, outputs probability distributions and stress test results. |
| `pnl_snapshot.py` | Point-in-time snapshots — saves timestamped P&L snapshots to a SQLite database. Enables tracking through close cycles and comparing prior periods. |
| `pnl_dashboard.py` | Streamlit web dashboard — interactive web-based P&L dashboard. No Python knowledge required to use — just open a browser. |
| `pnl_allocation_simulator.py` | Allocation what-if simulator — change revenue, AWS, or headcount shares and instantly see the P&L impact per product. |
| `pnl_ap_matcher.py` | AP invoice matching — fuzzy vendor name matching + amount/date proximity to identify likely duplicates and flag unmatched invoices. |
| `pnl_tests.py` | Automated test suite — unit and integration tests: 100% coverage of pnl_config, 80% coverage of pnl_month_end and pnl_allocation_simulator. |
| `build_charts.py` | Fortune 100 chart builder — builds the executive dashboard with 8 charts in a two-column grid layout. Applies iPipeline brand standards. |
| `redesign_pl_model.py` | Model redesign script — fixed all reconciliation failures, applied iPipeline brand guidelines, created the executive dashboard. (Run once — historical.) |

---

## SECTION 5 — Main Demo: SQL Scripts (4 Scripts)

These run in SQLite (zero-install, portable) and simulate what a real corporate database would provide.

| Script | Purpose |
|--------|---------|
| `staging.sql` | GL staging and normalization — imports raw GL data, normalizes dimensions (product, department, account), deduplicates entries. |
| `transformations.sql` | Allocation pivot and summary views — creates allocation shares table, product/department summary views, month-over-month variance calculations. |
| `validations.sql` | Data integrity checks — referential integrity, orphan detection, completeness checks, debit/credit balance, allocation reconciliation. |
| `pnl_enhancements.sql` | Advanced SQL features — budget vs actual view, allocation audit trail with triggers, rolling 12-month view, vendor contracts, full reconciliation view. |

---

## SECTION 6 — Second Demo Phase: Universal Tools (Candidate List — Not Built Yet)

These are tools for a **second demo** showing automation that works on **any Excel file** — not just the P&L model. This would be packaged as a Personal Macro Workbook Add-In (VBA) and standalone Python scripts.

**Status:** Candidate list only. 76 tools identified. No code written yet. Awaiting approval before building.

**Concept:** Coworkers install the Add-In once and have these tools available in every Excel session — whether they're working on a budget, an invoice tracker, an expense report, or any other file.

### Tier 1 VBA — Build First (24 tools, highest Finance value)

| Tool | Purpose |
|------|---------|
| Unmerge Cells & Fill Down | Unmerges all merged cells in selection, fills top value downward. Fixes every Finance export instantly. |
| Fill Blanks Down | Fills blank cells with the value from above. Fixes stacked layouts from GL exports and pivot tables. |
| Convert Text to Numbers | Fixes text-stored numbers so they actually sum. Most common Finance data export pain point. |
| Remove Leading/Trailing Spaces | Trims all text. Kills lookup failures caused by invisible spaces. |
| Delete Blank Rows | Removes empty rows from selection. Instant cleanup on any raw data export. |
| Replace Error Values | Replaces all #N/A, #REF!, #VALUE!, #DIV/0! errors with blank or a value. Cleans files before sharing. |
| Highlight Duplicate Rows | Colors duplicates for review — does NOT delete. Safe for invoice and vendor checks. |
| Remove Duplicate Rows | Deletes confirmed duplicate rows. Use after Highlight version confirms what to remove. |
| AutoFit All Columns & Rows | Auto-fits every column and row in the workbook. One click to make any file readable. |
| Freeze Top Row on All Sheets | Applies Freeze Panes to row 1 on every sheet. Finance standard for every file. |
| Number Format Standardizer | Applies #,##0.00 to all numeric cells. Forces consistency across any file. |
| Currency Format Standardizer | Applies $#,##0.00 to all currency values across all sheets. |
| Date Format Standardizer | Normalizes all dates to MM/DD/YYYY across all sheets. Kills the text-date import problem. |
| Conditional Formatting — Negatives in Red | Highlights all negatives red across the workbook. Finance standard for instant loss visibility. |
| Find & Replace Across All Sheets | Global find-and-replace across every sheet simultaneously. Massive time saver. |
| Unhide All Sheets, Rows & Columns | Makes every hidden worksheet, row, and column visible instantly. |
| Export All Sheets as One Combined PDF | Combines all visible sheets into one multi-page PDF. Send one file instead of many. |
| External Link Finder | Lists all cells referencing external workbooks with file paths. Critical audit tool. |
| Circular Reference Detector | Finds and reports circular references by exact location. |
| Workbook Error Scanner | Lists every cell with an error value — sheet name and cell address. Full report in one scan. |
| Search Across All Sheets | Finds any value across every sheet, returns location of every match. |
| Duplicate Invoice Detector | Matches on vendor + amount + date + invoice number. Flags potential duplicates for review. |
| Auto-Balancing GL Validator | Sums debit and credit columns, flags any imbalance, optionally adds a balancing plug line. |
| Multi-Replace Data Cleaner | Batch find-and-replace based on a two-column mapping table. Standardizes account names, cost center codes, category labels. |

### Tier 1 Python — Build First (5 tools)

| Tool | Purpose |
|------|---------|
| Universal Data Cleaner | One command: remove duplicates, drop empty rows/cols, trim spaces, fix text-numbers, standardize dates on any file. |
| Compare Two Excel Files | Cell-level diff report between any two workbooks. Perfect for version comparison. |
| Budget vs. Actual Consolidator | Merges 50+ department budget files into one master with $ and % variance columns. |
| AR/AP Aging Report | Auto-buckets invoices by days overdue (0–30, 31–60, 61–90, 90+) from any date column. |
| Multi-File Data Consolidator | Combines hundreds of Excel files from a folder into one master sheet with Source_File column. |

### Tier 2 VBA — Build Later (34 tools)

| Tool | Purpose |
|------|---------|
| Create Table of Contents Sheet | Generates a clickable index with hyperlinks to every worksheet. |
| Sort Worksheets Alphabetically | Reorders all sheet tabs A–Z. |
| Data Quality Scorecard | Summary: total rows, blank cells, errors, duplicates, data types per column. |
| Formula Auditor — Inconsistent Formulas | Flags cells in a column where the formula differs from the rest. |
| Protect All Sheets | Applies password protection to every sheet at once. |
| Unprotect All Sheets | Removes worksheet protection from every sheet. |
| Lock All Formula Cells | Locks formula cells, leaves input cells editable. |
| Export Active Sheet as PDF | Saves the active sheet as a PDF. |
| Export All Sheets as Individual PDFs | Saves each sheet as its own PDF. |
| Aging Bucket Calculator | Assigns records to 0–30, 31–60, 61–90, 90+ buckets from any date column. |
| Variance Analysis Template | Adds $ and % variance columns next to Actual and Budget columns. |
| Trial Balance Checker | Verifies debits equal credits and highlights any imbalance. |
| Journal Entry Validator | Checks every journal entry for balanced debits and credits. |
| Flux Analysis / Period Comparison | Flags lines with large period-over-period changes above a set threshold. |
| AP Aging Summary Generator | Generates AP aging report with standard buckets. |
| AR Aging Summary Generator | Generates AR aging report with standard buckets. |
| Ratio Analysis Dashboard Builder | Calculates Current Ratio, ROE, Gross Margin from statement data. |
| Build Distribution-Ready Copy | Creates a clean copy: formulas as values, metadata stripped, formatting standardized. |
| Reset All Filters | Clears all AutoFilter and Advanced Filter criteria on every sheet. |
| Workbook Health Check | Full diagnostic: file size, formula count, error count, external links, pivot tables. |
| Phantom Hyperlink Purger | Removes all embedded hyperlinks from the selected range or entire sheet. |
| Formula-to-Value Hardcoder | Converts all formulas in the selected range to static values. |
| Multi-Sheet Batch Renamer | Replaces a text string across all sheet tab names at once (e.g., "2024" → "2025"). |
| Convert Numbers to Words | Translates numeric values to written text (e.g., 1,250.00 → "One Thousand Two Hundred Fifty Dollars"). For invoices and checks. |
| Quick Corkscrew Builder | Builds a standard financial corkscrew (roll-forward) schedule for any balance sheet line item. |
| Financial Number Formatting Suite | Applies standard financial formats — Normal Accounting, Factor (000s), Percentage — via hotkeys. |
| General Ledger Journal Mapper | Transforms a raw trial balance into a journal entry upload template for accounting systems. |
| Financial Period Roll-Forward | Updates all header dates and clears prior-period input cells for a new month-end. |
| Multi-Currency Consolidation Aggregator | Consolidates subsidiary sheets with different currencies using a budget FX rate table. |
| Named Range Auditor | Reports all Named Ranges and flags any with hidden or broken references. |
| Conditional Format Purger | Lists all conditional formatting rules, allows batch deletion of redundant rules causing file bloat. |
| Data Validation Checker | Scans for Data Validation dropdowns with broken source ranges. |
| Print Header/Footer Standardizer | Updates print headers and footers across all selected sheets to match organizational standards. |
| External Link Severance Protocol | Replaces external file references with static values, logs original formula in a comment. |

### Tier 2 Python — Build Later (13 tools)

| Tool | Purpose |
|------|---------|
| Variance Analysis Generator | Compares Actual vs Budget across multiple files, creates summary and waterfall chart. |
| GL Reconciliation Engine | Matches two large transaction lists and flags unmatched items. |
| Fuzzy Match / Fuzzy Lookup | Matches records between datasets using fuzzy string matching — catches name typos. |
| Batch Process Folder of Files | Runs the same cleaning or transformation script on every Excel file in a folder. |
| Forecast Roll-Forward | Takes last month actuals + assumptions and builds the next 12-month forecast. |
| Unstructured Data Unpivoter | Converts wide pivot-style data (one column per month) to tall database format (one row per record). |
| PDF Tabular Extractor | Extracts tables from uneditable PDFs directly into Excel. Requires pdfplumber library. |
| Regex Text Extractor | Extracts structured patterns (invoice numbers, account codes, tax IDs) from free-text columns. |
| Word Report Generator | Reads Excel outputs and generates a formatted Word document automatically. |
| Fuzzy-Match Bank Reconciler | Probabilistically matches ledger descriptions against bank statement text using fuzzy logic. |
| Dynamic Master Data Mapper | SQL-style joins between datasets on common keys, replacing complex VLOOKUP arrays. |
| Reconciliation Exception Generator | Filters out matched transactions, outputs only unmatched exceptions into a new workbook. |
| Variance Decomposition Analyzer | Breaks financial variances into price, volume, and mix components automatically. |

---

## SECTION 7 — Constraints for the Review Claude

**These were permanently declined by the user. Do NOT recommend, suggest, or reference them:**
- Backup Workbook with Timestamp macro
- Cell-change / Timestamp Audit Trail (recording every edit)
- Export Charts to PowerPoint
- Email / Outlook automation integration

**The user is not a developer.** All guides, instructions, and recommendations must be in plain English. The audience for the tools is non-technical Finance & Accounting staff.

**The Excel file is binary.** Claude cannot read it directly. All code reference must go through the `.bas` source files in the `vba/` folder.

**The primary deliverable is the demo.** Everything else is secondary until the demo is done.
