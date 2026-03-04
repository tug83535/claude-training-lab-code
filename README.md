# APCLDmerge — P&L Demo Project
### iPipeline Finance & Accounting | Built for 2,000+ Employees

---

## What This Is

A world-class Profit & Loss (P&L) demonstration system built for iPipeline's
Finance & Accounting department. The project combines an Excel workbook with
VBA macros, a full SQL data pipeline, and a Python automation toolkit to deliver
a polished, fully interactive financial model designed to be presented to the
entire company — including the CFO and CEO.

A professional video walkthrough is also planned to help Finance & Accounting
coworkers understand and use all the tools included in this project.

---

## What's Inside

### Excel + VBA (The Centerpiece)
The main file is `excel/ExcelDemoFile_adv.xlsm` — a 13-sheet P&L
model powered by VBA macros organized into 32 modules:

| Module | What It Does |
|--------|-------------|
| `modConfig` | All constants — sheet names, products, departments, fiscal year, colors. The foundation every other module depends on. |
| `modFormBuilder` | Builds the Command Center UserForm. 62 actions across 15 categories. Single routing table (ExecuteAction). |
| `modMasterMenu` | InputBox fallback menu (4 pages, 62 items) for when the UserForm isn't installed. |
| `modNavigation` | Table of contents, GoHome, QuickJump, keyboard shortcuts (Ctrl+Shift+M/H/J/R). |
| `modDashboard` | 6 chart and dashboard types: revenue trend, margin %, product mix, Executive KPI cards, Waterfall, Product Comparison. |
| `modDataQuality` | 6 data scans + fix routines: duplicates, mixed dates, text-stored numbers, assumption errors, misspellings, blank cells. |
| `modReconciliation` | Checks sheet PASS/FAIL validation + 4 cross-sheet reconciliation checks with configurable tolerance. |
| `modVarianceAnalysis` | Month-over-month variance engine with Favorable/Unfavorable logic and auto-written English commentary for top variances. |
| `modPDFExport` | Batch 7-sheet PDF export with professional headers, footers, landscape layout, and Save As dialog. |
| `modPerformance` | TurboMode on/off, elapsed timer, ForceRecalc, status bar progress updates. |
| `modMonthlyTabGenerator` | Auto-generates Apr–Dec monthly tabs by cloning the Mar template. One-tab-at-a-time option included. |
| `modSearch` | Cross-sheet keyword search, 200-result cap, yellow highlights, hyperlinked results sheet. |
| `modUtilities` | 12 utility macros (actions 51–62): delete blank rows, unhide all sheets, sort tabs, freeze panes, convert formulas to values, AutoFit columns, protect/unprotect sheets, find & replace all sheets, highlight hardcoded numbers, presentation mode, unmerge and fill down. |
| `modLogger` | Runtime audit log. Every VBA macro run is timestamped and logged to a hidden sheet (VBA_AuditLog) with username, module, procedure, and status. |
| `modSensitivity` | Sensitivity analysis on Assumptions sheet drivers. |
| `modAWSRecompute` | AWS allocation validation and recalculation. |
| `modImport` | CSV/Excel data import pipeline (Command 17). |
| `modForecast` | Rolling forecast + append to trend sheets. |
| `modScenario` | Scenario save, load, compare, and delete. |
| `modAllocation` | Cost allocation engine + preview (Commands 24-25). |
| `modConsolidation` | Multi-entity consolidation + IC eliminations (Commands 26-30). |
| `modVersionControl` | Version save, compare, restore, and list (Commands 31-35). |
| `modAdmin` | Auto-documentation and change management (Commands 36-40). |
| `modIntegrationTest` | 18-test suite + quick health check (Commands 44-45). |
| `modDemoTools` | Control sheet buttons, parameterized print areas, executive summary. |
| `modDataGuards` | Validates assumptions presence, checks driver sums, finds negative/zero/round numbers. |
| `modDrillDown` | Reconciliation drill links, heatmap, golden file compare. |
| `modAuditTools` | Change log, external links finder/fixer, hidden sheet audit, masked copy. |
| `modETLBridge` | Triggers ETL locally and imports ETL output. |
| `modTrendReports` | Rolling 12-month view, reconciliation trend chart, archive results. |
| `modDataSanitizer` | Numeric-only sanitizer — never touches dates, names, or customer IDs. |

---

### SQL Pipeline (SQLite 3)

| File | What It Does |
|------|-------------|
| `staging.sql` | Full ETL pipeline. Creates 5 dimension tables + 1 normalized fact table. Duplicate detection, indexing, dimensional lookups. |
| `transformations.sql` | Allocation framework. Revenue, AWS compute, and headcount share tables. 8 analytical views including MoM variance with FLAG/NEW/OK status. |
| `pnl_enhancements.sql` | 5 advanced additions: Budget vs Actual tracking, allocation audit trail with SQL triggers, rolling 12-month views, vendor contract calendar, allocation reconciliation checks. |
| `validations.sql` | 20+ validation views: referential integrity, ETL completeness, data quality, balance checks, and a consolidated pass/fail summary view. |

---

### Python Toolkit (13 Scripts)

| Script | What It Does |
|--------|-------------|
| `pnl_config.py` | Central configuration: all file paths, constants, products, departments, revenue shares, sheet names, thresholds. Single source of truth for every other script. |
| `pnl_runner.py` | Master orchestrator. Chains all pipeline steps in order (staging → transformations → validations → close → snapshot). |
| `pnl_month_end.py` | Month-end close automation. 6-check QA pipeline with PASS/FAIL/WARN for each check. Excel export. |
| `pnl_forecast.py` | Forecasting engine with 4 methods: Simple Moving Average, Exponential Smoothing, Linear Trend, and Scenario-based. Confidence intervals included. |
| `pnl_allocation_simulator.py` | What-if scenario engine. Recalculates product P&L under different revenue share assumptions. 3 presets + custom input. |
| `pnl_monte_carlo.py` | **Monte Carlo risk simulation.** Runs 10,000+ iterations with randomized allocation shares (Dirichlet) and expense amounts (Normal distribution) to produce full probability distributions: P5/P25/P50/P75/P95, Value at Risk, per-product breakdown, 4-panel chart, and Excel export. |
| `pnl_snapshot.py` | Point-in-time P&L snapshots stored in SQLite. Enables period-over-period comparison of historical close states. |
| `pnl_dashboard.py` | Interactive Streamlit web dashboard. Filter by product, department, month. Revenue trends, margin charts, heatmaps, top vendors. |
| `pnl_ap_matcher.py` | AP invoice matching engine. Fuzzy vendor name matching, duplicate detection, unmatched items flagged for review. |
| `pnl_tests.py` | Full pytest test suite. 100% coverage on config and allocation logic. 80%+ on close and forecasting. |
| `pnl_cli.py` | Master command-line interface. Single entry point for every script. Run any module with one command. |
| `build_charts.py` | Chart generation utility for the demo P&L file. |
| `redesign_pl_model.py` | One-time script used to redesign the Excel workbook to Fortune 100 FP&A standard. |

---

## How to Run the Python Toolkit

```bash
# Show toolkit status and check all dependencies
python pnl_cli.py status

# Run the full pipeline (staging → close → forecast → report)
python pnl_cli.py run-all

# Month-end close for March
python pnl_cli.py close --month 3 --export

# 6-month forecast using exponential smoothing
python pnl_cli.py forecast --months 6 --method ets

# Monte Carlo simulation — 10,000 iterations
python pnl_cli.py monte-carlo --export results.xlsx

# What-if allocation scenarios
python pnl_cli.py simulate --presets

# Launch the interactive web dashboard
python pnl_cli.py dashboard
```

---

## Repository Structure

```
APCLDmerge/
├── excel/                  Main demo P&L Excel file (uploaded each session)
├── vba/                    VBA modules (.bas files) — 32 modules, 62 actions
├── sql/                    SQL pipeline — 4 scripts, SQLite 3
├── python/                 Python toolkit — 14 scripts
├── docs/
│   ├── overview/           Architecture diagram, code comparison report
│   ├── day-to-day/         Quick-reference guides for daily operations
│   └── setup/              Step-by-step setup and implementation guides
├── training/               Final training materials for coworkers
├── qa/                     QA tracking, test plans, bug logs, validation reports
├── tasks/                  Session management — todo.md and lessons.md
├── DemofileChartBuild/     Chart sheet files, work in progress
├── NewTesting/             Research and ideas not yet ready for main project
└── CompletePackageStorage/
    ├── production/         Live, ready-to-go final files
    └── backups/            Versioned backups of completed work
```

---

## Project Status — v2.1.0 (2026-02-28)

| Layer | Status |
|-------|--------|
| VBA Macros | 32 modules, 62 actions — fully operational |
| SQL Pipeline | 4 scripts — production-quality ETL + analytics |
| Python Toolkit | 13 scripts including Monte Carlo simulation |
| Documentation | Architecture, runbook, training guides, QA reports |
| Test Coverage | 100% on config/allocation, 80%+ on close/forecast |

**Current branch:** `claude/review-project-status-ntucB`

---

## Quality Standard

Every file in this repository is held to a world-class standard.
The goal is for this project to represent the best work of the best employee
at the best company — something the CFO and CEO would be proud to present
to 2,000+ employees.

---

*Finance & Accounting — iPipeline | Project: APCLDmerge*
