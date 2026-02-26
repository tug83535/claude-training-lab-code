# KBT P&L Automation Toolkit — Start-to-Finish Guide

**Everything you need to go from "I have these files" to running a live demo.**

---

## Table of Contents

1. [What You Received](#1-what-you-received)
2. [What Every File Does](#2-what-every-file-does)
3. [Prerequisites — What You Need Installed](#3-prerequisites)
4. [Step 1: Set Up the Excel Workbook](#4-step-1-set-up-the-excel-workbook)
5. [Step 2: Import VBA Modules](#5-step-2-import-vba-modules)
6. [Step 3: Build the Command Center](#6-step-3-build-the-command-center)
7. [Step 4: Test That Everything Works](#7-step-4-test-that-everything-works)
8. [Step 5: Set Up Python (Optional)](#8-step-5-set-up-python-optional)
9. [Step 6: Set Up SQL (Optional)](#9-step-6-set-up-sql-optional)
10. [How to Use the System Day-to-Day](#10-how-to-use-the-system-day-to-day)
11. [How to Prepare and Run a Demo](#11-how-to-prepare-and-run-a-demo)
12. [Troubleshooting](#12-troubleshooting)
13. [Quick Reference Cheat Sheet](#13-quick-reference-cheat-sheet)

---

## 1. What You Received

You have a zip file called `Excel-PnL-Automation-Package.zip`. When you unzip it, you get this folder structure:

```
Excel-PnL-Automation-Package/
│
├── CHANGELOG.md                          ← Version history
│
├── 01_Solution_Overview/                 ← High-level docs for leadership/IT
│   ├── EXECUTIVE_SUMMARY.md
│   └── ARCHITECTURE_DIAGRAM.md
│
├── 02_Workbook/                          ← Instructions for the Excel file
│   └── WORKBOOK_SETUP_NOTES.md
│
├── 03_Code/                              ← All the code
│   ├── VBA/                              ← 13 files (Excel macros)
│   ├── Python/                           ← 12 files (analytics scripts)
│   └── SQL/                              ← 4 files (database queries)
│
├── 04_Docs/                              ← User documentation
│   ├── QUICK_START.md
│   ├── IMPLEMENTATION_GUIDE.md
│   ├── USER_TRAINING_GUIDE.md
│   ├── OPERATIONS_RUNBOOK.md
│   └── SANITIZATION_PLAYBOOK.md
│
├── 05_Templates/                         ← Support files
│   └── logging_template.csv
│
└── 06_QA/                                ← Quality assurance
    ├── TEST_PLAN.md
    ├── VALIDATION_REPORT.md
    ├── INTEGRATION_TEST_GUIDE.md
    └── ISSUE_CLOSURE.md
```

You also need the **source workbook**: `KeystoneBenefitTech_PL_Model.xlsx`. This is the Excel file all the code operates on. The code does NOT work without this workbook.

**The 3 layers of this system:**

| Layer | Required? | What It Does |
|-------|-----------|-------------|
| **Excel + VBA** | Yes — this is the core | 50 automated commands inside Excel. Everything from data quality checks to PDF report export. |
| **Python** | Optional | Analytics that run OUTSIDE Excel: interactive dashboard, statistical forecasting, fuzzy matching, automated testing. |
| **SQL** | Optional | Database queries for when you outgrow Excel — staging tables, allocation pivots, validation views. |

---

## 2. What Every File Does

### 03_Code/VBA/ — The Excel Macros (13 files)

These are VBA modules that get imported into your Excel workbook. They ARE the automation. Each `.bas` file is a separate module with a specific job.

**Foundation modules (the engine room — everything else depends on these):**

| File | What It Does | Why It Matters |
|------|-------------|----------------|
| `modConfig_v2.1.bas` | Central configuration. Contains every constant: sheet names, product names, department names, colors, thresholds. Also has helper functions like `SafeDeleteSheet`, `StyleHeader`, `SheetExists`. | If another module needs to know "what's the GL sheet called?" it asks modConfig. Change a sheet name here and it changes everywhere. |
| `modPerformance_v2.1.bas` | Speed optimization. `TurboOn()` turns off screen updating, auto-calculation, and event handling before a macro runs. `TurboOff()` turns them back on. Also has a timer for measuring how long things take. | Without this, every macro would be visibly slow. Every public sub calls TurboOn at the start and TurboOff at the end. |
| `modNavigation_v2.1.bas` | Keyboard shortcuts and navigation. Assigns Ctrl+Shift+M (menu), Ctrl+Shift+H (home), Ctrl+Shift+J (jump to sheet), Ctrl+Shift+R (reconciliation). | This is what makes Ctrl+Shift+M open the Command Center. It runs automatically when you open the workbook. |

**Menu system (how you interact with all 50 commands):**

| File | What It Does | Why It Matters |
|------|-------------|----------------|
| `modMasterMenu_v2.1.bas` | The 50-item menu router. When you pick a command number, this module calls the right function in the right module. | This is the traffic cop. "User picked #6" → calls `modVarianceAnalysis.RunVarianceAnalysis`. |
| `modFormBuilder_v2.1.bas` | Builds the Command Center UserForm (the nice GUI). Contains `BuildCommandCenter` (creates the form automatically) and `ExecuteAction` (routes all 50 commands). Also contains the `GetFormInstallGuide` for manual setup instructions. | The largest file (699 lines). This is what makes the polished pop-up interface with categories, search, and double-click. |
| `frmCommandCenter_code.txt` | The UserForm's event code (what happens when you click buttons, type in search, select a category). This is NOT a .bas file — it gets pasted into the UserForm code window. | You only need this if you're building the form manually (Mode B). If you use `BuildCommandCenter` (Mode A), this code is generated automatically. |

**Feature modules (the actual automation):**

| File | What It Does | Commands It Powers |
|------|-------------|-------------------|
| `modMonthlyTabGenerator_v2.1.bas` | Creates and deletes monthly P&L tabs (Apr–Dec) by cloning the March template. Also has `GenerateNextMonthOnly` for single-month creation. | #1 Generate Monthly Tabs, #2 Delete Generated Tabs, plus #42 (single month) |
| `modReconciliation_v2.1.bas` | Runs PASS/FAIL checks comparing totals across sheets. Also has `ValidateCrossSheet` for 4 deep computed validations (GL vs Trend, GL vs Functional, Product check, Mirror check). | #3 Run Reconciliation, #4 Export Recon Report, #47 Cross-Sheet Validation |
| `modDataQuality_v2.1.bas` | Scans the entire workbook for data issues: duplicates, text-stored numbers, blanks, outliers. Fix commands only operate on pre-flagged cells (safety feature). | #7 Scan Data Quality, #8 Fix Text Numbers, #9 Fix Duplicates |
| `modVarianceAnalysis_v2.1.bas` | Compares current month to prior month, flags variances >15%. `GenerateCommentary` auto-writes an executive narrative for the top 5 variances. | #6 Variance Analysis, #46 Variance Commentary |
| `modDashboard_v2.1.bas` | Creates charts: revenue trend bar chart, contribution margin line chart, revenue mix pie chart. Also has `CreateExecutiveDashboard` (KPI tiles), `WaterfallChart`, and `ProductComparison`. | #12 Build Dashboard, plus executive dashboard features |
| `modPDFExport_v2.1.bas` | Exports sheets to PDF with professional formatting — headers, footers, page numbers, timestamps. Can export a full report package or single sheet. | #10 Export Report Package, #11 Export Active Sheet |
| `modSearch_v2.1.bas` | Full-text search across all sheets. Shows results on a "Search Results" sheet with hyperlinks back to the source cells. Caps at 200 results with a warning showing total matches. | Various search functions, #47 (partial) |

**What about the other 15 VBA modules?** The package only includes the 12 `.bas` files + 1 `.txt` that were updated to v2.1. The remaining 15 modules (modAdmin, modAllocation, modForecast, modFormatting, modImport, modIntegrationTest, modRefresh, modScenario, modSetup, modSnapshot, modConditionalFormat, modEmailSummary, modValidation, modSensitivity, modLogger) ship with the original v2.0 workbook and are already present if you have the full system. They don't need to be re-imported unless you're starting from scratch.

---

### 03_Code/Python/ — Analytics Scripts (12 files)

These run OUTSIDE of Excel in a terminal/command prompt. They read the `.xlsx` file, do analysis, and produce output files or web dashboards.

| File | What It Does | How to Run It |
|------|-------------|---------------|
| `pnl_config.py` | Shared configuration for all Python scripts. Product names, revenue shares, department lists, fiscal year, formatting functions. The Python equivalent of modConfig. | `python pnl_config.py` — prints a self-test showing all config values |
| `pnl_runner.py` | **The unified entry point.** Routes to all other scripts via sub-commands. This is the ONE file you actually run. | `python pnl_runner.py --help` — shows all 9 commands |
| `pnl_dashboard.py` | Interactive web dashboard using Streamlit. Revenue trends, product comparisons, department breakdowns, allocation what-if sliders — all in a browser. | `python pnl_runner.py dashboard` — opens in your browser |
| `pnl_month_end.py` | Automated month-end close checklist. Runs 6 categories of checks (data completeness, reconciliation, allocation, trend, quality, executive readiness) and produces a close report. | `python pnl_runner.py month-end --month 1` |
| `pnl_forecast.py` | Statistical forecasting. Simple moving average, exponential smoothing, and trend analysis for remaining months. | `python pnl_runner.py forecast --months 3` |
| `pnl_allocation_simulator.py` | What-if allocation scenarios. "What happens if we shift 10% of iGO's revenue share to DocFast?" Shows before/after comparison with Greek delta (Δ) change indicators. | `python pnl_runner.py allocate` |
| `pnl_snapshot.py` | Save, load, compare, and restore P&L snapshots. Like version control for your financial data. | `python pnl_runner.py snapshot --save "Q1 Close"` |
| `pnl_ap_matcher.py` | Fuzzy vendor name matching using Levenshtein distance. Finds "Amzon Web Srvcs" matches "Amazon Web Services" with a confidence score. | `python pnl_runner.py match` |
| `pnl_email_report.py` | Generates a formatted HTML executive email summary with KPIs, traffic-light status indicators, and trend sparklines. | `python pnl_runner.py email --month 1` |
| `pnl_cli.py` | Alternative CLI using Click framework (legacy — `pnl_runner.py` is preferred). | Not typically used directly |
| `pnl_tests.py` | Automated test suite. 116 test methods across 17 classes. Tests config validation, currency formatting, month-end checks, allocation math, and smoke-tests all modules. | `python pnl_runner.py test` or `pytest pnl_tests.py -v` |
| `requirements.txt` | List of Python packages needed. | `pip install -r requirements.txt` |

---

### 03_Code/SQL/ — Database Queries (4 files)

These are SQLite scripts for when you want to move data into a proper database. Run order matters.

| File | What It Does | Run Order |
|------|-------------|-----------|
| `staging.sql` | Creates dimension tables (product, department, expense category, date), a GL staging table, and a normalized fact table. Includes the ETL to move data from staging → fact. | Run 1st |
| `transformations.sql` | Creates allocation share tables, department×product pivot views, product/department summaries, month-over-month variance views, and expense category mix analysis. | Run 2nd |
| `validations.sql` | Creates validation views: referential integrity checks, orphan detection, completeness analysis, balance validation, data quality checks (blanks, outliers, zero amounts), and a consolidated summary. | Run 3rd |
| `pnl_enhancements.sql` | 5 advanced additions: budget vs. actual tables, allocation audit trail with triggers, rolling 12-month P&L view, vendor contract renewal calendar, allocation reconciliation queries. | Run 4th (optional) |

Each SQL file also contains the equivalent Power Query M code as comments at the bottom, so you can implement the same logic in Excel's Power Query if you prefer.

---

### Documentation Files

| File | Who It's For | What's In It |
|------|-------------|-------------|
| `EXECUTIVE_SUMMARY.md` | CFO, FP&A manager | 2-page business case: problem, solution, outcomes, architecture overview |
| `ARCHITECTURE_DIAGRAM.md` | IT, developers | Full ASCII system diagrams: layer architecture, data flow, module dependencies |
| `WORKBOOK_SETUP_NOTES.md` | Anyone doing setup | How to convert .xlsx → .xlsm, sheet inventory, layout contract, backup tips |
| `QUICK_START.md` | First-time users | 10-minute path from zero to running your first 5 commands |
| `IMPLEMENTATION_GUIDE.md` | IT admin, power users | Trust Center settings, VBA import procedure, UserForm build (Mode A & B), Python env setup, named ranges, fiscal year rollover |
| `USER_TRAINING_GUIDE.md` | Finance team | All 50 commands documented: what it does, when to use it, what could go wrong |
| `OPERATIONS_RUNBOOK.md` | Month-end operator | Monthly cadence (open/mid/close), step-by-step procedures, failure scenarios with fixes |
| `SANITIZATION_PLAYBOOK.md` | Demo preparer | How to mask company name, vendor names, dollar amounts for external presentation |
| `CHANGELOG.md` | Everyone | Version history from v1.0 → v2.0 → v2.1 with breaking changes flagged |
| `TEST_PLAN.md` | QA, auditor | 40 test cases across 7 categories with pass/fail criteria |
| `VALIDATION_REPORT.md` | QA, auditor | Actual workbook data validated: 510 GL rows, $3.7M total, formula scan results |
| `INTEGRATION_TEST_GUIDE.md` | IT admin | How to run and interpret the built-in integration test suite |
| `ISSUE_CLOSURE.md` | Project stakeholders | All 15 pre-audit issues: what they were, how they were fixed, how to verify |
| `logging_template.csv` | Admin | Empty CSV template for the audit log (Timestamp, Module, Procedure, Detail, User, Workbook) |

---

## 3. Prerequisites

### Must Have (for the core Excel system)

| Software | Version | Where to Get It |
|----------|---------|-----------------|
| **Microsoft Excel** | 2019 or Microsoft 365 (Windows) | Already installed on most corporate PCs |
| **Windows** | 10 or 11 | Mac Excel has VBA limitations — Windows is required |

That's it. The core VBA system has zero external dependencies.

### Optional (for Python analytics)

| Software | Version | Where to Get It |
|----------|---------|-----------------|
| **Python** | 3.11 or newer | python.org/downloads |
| **pip** | (comes with Python) | Included with Python installer |

### Optional (for SQL database)

| Software | Version | Where to Get It |
|----------|---------|-----------------|
| **SQLite** | 3.x | sqlite.org/download (or use DB Browser for SQLite for a GUI) |

---

## 4. Step 1: Set Up the Excel Workbook

### 1A — Unzip the Package

1. Right-click `Excel-PnL-Automation-Package.zip` → **Extract All**
2. Choose a permanent location (e.g., `C:\KBT_PnL_Toolkit\`)
3. You should see the folder structure from Section 1 above

### 1B — Copy the Workbook

1. Put `KeystoneBenefitTech_PL_Model.xlsx` into `C:\KBT_PnL_Toolkit\` (or wherever you extracted)
2. Make a backup copy somewhere safe — you'll always want an untouched original

### 1C — Convert to Macro-Enabled Format

The workbook ships as `.xlsx` (no macros allowed). You need to save it as `.xlsm`:

1. Open `KeystoneBenefitTech_PL_Model.xlsx` in Excel
2. Click **File** → **Save As**
3. In the "Save as type" dropdown, select **Excel Macro-Enabled Workbook (*.xlsm)**
4. Click **Save**
5. If Excel asks "Do you want to keep the workbook in this format?", click **Yes**
6. Close the file. From now on, only use the `.xlsm` version.

### 1D — Enable Macros in Trust Center

This tells Excel "I trust the code in this file."

1. Open your new `.xlsm` file
2. Click **File** → **Options** → **Trust Center** → **Trust Center Settings...**
3. In the left panel, click **Macro Settings**
4. Select **Enable all macros** (or "Disable all macros with notification" if your IT policy requires — you'll just click "Enable" each time you open the file)
5. **CRITICAL:** Check the box that says **"Trust access to the VBA project object model"** — this is required for the automatic Command Center builder
6. Click **OK** → **OK**

### 1E — Verify the Workbook

Before importing any code, confirm the workbook is intact:

1. You should see 13 sheet tabs at the bottom:
   - CrossfireHiddenWorksheet
   - Assumptions
   - Data Dictionary
   - AWS Allocation
   - Report-->
   - P&L - Monthly Trend
   - Product Line Summary
   - Functional P&L - Monthly Trend
   - Functional P&L Summary - Jan 25
   - Functional P&L Summary - Feb 25
   - Functional P&L Summary - Mar 25
   - US January 2025 Natural P&L
   - Checks

2. Click on the **CrossfireHiddenWorksheet** tab — you should see 510 rows of GL data with columns: ID, Date, Department, Product, Expense Category, Vendor, Amount

3. Click on the **Checks** tab — you should see 9 reconciliation checks (3 PASS, 6 FAIL — the FAILs are pre-existing data discrepancies, not bugs)

If anything is missing, your source workbook may be damaged. Go back to the original.

---

## 5. Step 2: Import VBA Modules

This is where you put the automation code into the workbook.

### 2A — Open the VBA Editor

1. With the `.xlsm` file open, press **Alt+F11**
2. The Visual Basic Editor opens. On the left side you'll see the "Project Explorer" panel showing your workbook's project.

### 2B — Import the 12 .bas Files

1. In the VBA Editor, click on your project name (e.g., "VBAProject (KeystoneBenefitTech_PL_Model.xlsm)")
2. Go to **File** → **Import File...** (or right-click the project → **Import File...**)
3. Navigate to `C:\KBT_PnL_Toolkit\Excel-PnL-Automation-Package\03_Code\VBA\`
4. Select **modConfig_v2.1.bas** and click **Open**
5. It appears in the Project Explorer under "Modules"
6. **Repeat** for all 12 `.bas` files:

```
Import in this order (order matters for dependencies):

1.  modConfig_v2.1.bas           ← Must be first (everything depends on it)
2.  modPerformance_v2.1.bas      ← Must be second (used by all modules)
3.  modNavigation_v2.1.bas
4.  modMasterMenu_v2.1.bas
5.  modFormBuilder_v2.1.bas
6.  modMonthlyTabGenerator_v2.1.bas
7.  modReconciliation_v2.1.bas
8.  modDataQuality_v2.1.bas
9.  modVarianceAnalysis_v2.1.bas
10. modDashboard_v2.1.bas
11. modPDFExport_v2.1.bas
12. modSearch_v2.1.bas
```

**Important:** If the project already has modules with these names (without "_v2.1"), the old and new will coexist. You may want to delete the old ones first to avoid confusion. Right-click the old module → **Remove modXxx...** → **No** (don't export).

### 2C — Handle the "_v2.1" Suffix

The imported modules will appear as "modConfig_v2_1" (VBA replaces the dot). This is fine — the code references module names internally, not file names. If you want cleaner names, you can rename them:

1. Click on the module in Project Explorer
2. In the Properties window (bottom-left), change the **(Name)** property
3. Remove the "_v2_1" suffix so it reads just "modConfig", "modPerformance", etc.

**Note:** If you already have v2.0 modules, renaming causes a conflict. In that case, delete the v2.0 modules first, then import and rename the v2.1 ones.

### 2D — Verify Compilation

1. In the VBA Editor, click **Debug** → **Compile VBAProject**
2. If there are no errors, nothing happens (that's good!)
3. If you get an error, it will highlight the problem line. The most common cause is a missing module — make sure all 12 are imported.

### 2E — Save

Press **Ctrl+S** to save the workbook with the new modules.

---

## 6. Step 3: Build the Command Center

The Command Center is the pop-up interface with categories, search, and all 50 commands. There are two ways to build it.

### Mode A — Automatic (Recommended)

This creates the UserForm programmatically in about 2 seconds.

1. In the VBA Editor, press **Ctrl+G** to open the **Immediate Window** (bottom panel)
2. Type the following and press Enter:

```
modFormBuilder.BuildCommandCenter
```

3. You should see a message: "Command Center built successfully!"
4. In the Project Explorer, you'll now see "frmCommandCenter" under "Forms"
5. Close the VBA Editor (Alt+F11 again, or click X)

**If this fails** with "Programmatic access to Visual Basic Project is not trusted," go back to Step 1D and make sure you checked "Trust access to the VBA project object model."

### Mode B — Manual (If Mode A Doesn't Work)

1. In the VBA Editor, go to **Insert** → **UserForm**
2. A blank form appears. In the Properties window, change the **(Name)** to `frmCommandCenter`
3. Don't design anything on the form — the code creates all controls at runtime
4. Double-click the form to open its code window
5. Delete any default code (like `Private Sub UserForm_Click()`)
6. Open `03_Code/VBA/frmCommandCenter_code.txt` in a text editor (Notepad is fine)
7. Select all the text (Ctrl+A), copy it (Ctrl+C)
8. Go back to the VBA code window and paste (Ctrl+V)
9. Save (Ctrl+S)

### Test the Command Center

1. Go back to Excel (close or minimize the VBA Editor)
2. Press **Ctrl+Shift+M**
3. The Command Center should appear!

**What you should see:**
- A pop-up window with a categories panel on the left (All Actions, Monthly Operations, Analysis, etc.)
- A list of actions on the right (numbered 1–50)
- A search box at the top
- Run and Close buttons at the bottom

**If nothing happens when you press Ctrl+Shift+M:**
- The shortcut might not be assigned yet. Close and reopen the workbook (the Workbook_Open event assigns shortcuts).
- Or, in the Immediate Window, type: `modNavigation.AssignShortcuts` and press Enter. Then try Ctrl+Shift+M again.

**If you get an InputBox instead of the pretty form:**
- That's the fallback mode — it still works. Type a command number (1-50) and press OK. The form may not have been built correctly. Try Mode A again.

---

## 7. Step 4: Test That Everything Works

Run these 5 commands in order. Each one exercises different parts of the system.

### Test 1: Data Quality Scan (Command #7)

1. Press **Ctrl+Shift+M** to open Command Center
2. Click **"Data Quality"** in the left categories panel
3. Double-click **"#7 Scan for Data Quality Issues"** (or single-click + click Run)
4. Wait 5-10 seconds
5. **Expected result:** A new sheet called "Data Quality Report" appears with scan results

### Test 2: Reconciliation (Command #3)

1. Ctrl+Shift+M → Monthly Operations → **#3 Run Reconciliation Checks**
2. **Expected result:** The Checks sheet updates. You should see 3 PASS and 6 FAIL (the FAILs are pre-existing data discrepancies — the system is working correctly by flagging them)

### Test 3: Variance Analysis (Command #6)

1. Ctrl+Shift+M → Analysis → **#6 Run Variance Analysis**
2. **Expected result:** A "Variance Analysis" sheet appears comparing months with red/green formatting on items that moved >15%

### Test 4: Build Dashboard (Command #12)

1. Ctrl+Shift+M → Reporting → **#12 Build Dashboard Charts**
2. **Expected result:** A "Dashboard" sheet appears with 3 charts (revenue bar chart, CM% line chart, revenue mix pie chart)

### Test 5: About (Command #50)

1. Ctrl+Shift+M → Advanced → **#50 About This Toolkit**
2. **Expected result:** A message box showing version 2.1.0, build date, and toolkit summary

**If all 5 tests work: your VBA system is fully operational.** You can stop here if you don't need Python or SQL.

---

## 8. Step 5: Set Up Python (Optional)

Skip this section entirely if you only want the Excel/VBA automation.

### 5A — Install Python

1. Go to [python.org/downloads](https://python.org/downloads)
2. Download Python 3.11 or newer
3. **CRITICAL during install:** Check the box that says **"Add Python to PATH"**
4. Click "Install Now"

### 5B — Verify Python Installation

Open a command prompt (Windows Key → type "cmd" → Enter):

```
python --version
```

You should see `Python 3.11.x` or newer. If you get "not recognized," Python isn't in your PATH — reinstall with the PATH checkbox checked.

### 5C — Install Dependencies

In the same command prompt, navigate to the Python folder and install:

```
cd C:\KBT_PnL_Toolkit\Excel-PnL-Automation-Package\03_Code\Python
pip install -r requirements.txt
```

This installs: pandas, numpy, openpyxl, matplotlib, streamlit, plotly, statsmodels, scikit-learn, thefuzz, click, pytest. It may take 2-5 minutes.

### 5D — Configure the File Path

The Python scripts need to know where your Excel workbook is. Two options:

**Option A — Environment variable (set once):**
```
set KBT_SOURCE_FILE=C:\KBT_PnL_Toolkit\KeystoneBenefitTech_PL_Model.xlsm
```

**Option B — Pass it each time:**
```
python pnl_runner.py config --file "C:\KBT_PnL_Toolkit\KeystoneBenefitTech_PL_Model.xlsm"
```

### 5E — Test Python

```
python pnl_runner.py --help
```

You should see a banner and 9 available commands:

```
Available commands:

  dashboard       Launch interactive Streamlit dashboard
  month-end       Run month-end close checklist
  forecast        Run forecast engine
  allocate        What-if allocation simulator
  snapshot        Manage P&L snapshots
  match           AP fuzzy matching
  email           Generate executive email report
  test            Run automated test suite
  config          Show current configuration
```

### 5F — Run the Test Suite

```
python pnl_runner.py test
```

This runs 116 automated tests. You should see mostly passes (some may skip if the workbook isn't in the expected location — that's OK).

### 5G — Try the Dashboard

```
python pnl_runner.py dashboard
```

This launches a Streamlit web server. Your browser opens automatically to `http://localhost:8501` showing an interactive dashboard with:
- Revenue trend charts
- Product comparison
- Department breakdowns
- Allocation what-if sliders

Press **Ctrl+C** in the command prompt to stop the dashboard.

---

## 9. Step 6: Set Up SQL (Optional)

Skip this section entirely if you don't need a database.

### 6A — Install SQLite

Download DB Browser for SQLite from [sqlitebrowser.org](https://sqlitebrowser.org) (GUI tool, easier than command line). Or use the command-line `sqlite3` tool.

### 6B — Create the Database

**Using command line:**
```
cd C:\KBT_PnL_Toolkit\Excel-PnL-Automation-Package\03_Code\SQL
sqlite3 keystone_pnl.db < staging.sql
sqlite3 keystone_pnl.db < transformations.sql
sqlite3 keystone_pnl.db < validations.sql
sqlite3 keystone_pnl.db < pnl_enhancements.sql
```

**Using DB Browser:**
1. Open DB Browser → **New Database** → save as `keystone_pnl.db`
2. Go to **Execute SQL** tab
3. Open `staging.sql`, paste it in, click **Execute All**
4. Repeat for `transformations.sql`, `validations.sql`, `pnl_enhancements.sql` (in that order)

### 6C — Load GL Data

You need to export the GL data from the CrossfireHiddenWorksheet to a CSV file first, then import it into the staging table.

### 6D — Run Validations

```sql
SELECT * FROM v_validation_summary;
```

This shows you a full PASS/FAIL report for all data quality checks.

---

## 10. How to Use the System Day-to-Day

### The Monthly Workflow

Every month, you follow this cycle. The `OPERATIONS_RUNBOOK.md` has the full detail, but here's the summary:

**Week 1 (Month-Open):**
1. Open the workbook, Ctrl+Shift+M
2. **#20 Save Current Scenario** — name it "FY25_MonXX_Opening"
3. **#17 Import GL Data Pipeline** — load the new month's GL extract
4. **#7 Scan Data Quality** — fix any issues found
5. **#3 Run Reconciliation** — baseline check

**Week 2 (Mid-Month):**
6. **#17 Import GL Data** again if updated extracts arrive
7. **#18 Rolling Forecast** — update projections
8. **#6 Variance Analysis** — spot-check emerging trends

**Week 3-4 (Month-Close):**
9. **#17 Import GL Data** — final extract
10. **#7 Scan Data Quality** — must be zero issues
11. **#3 Reconciliation** — must be all PASS
12. **#6 Variance Analysis** — full analysis
13. **#46 Variance Commentary** — auto-generated narrative
14. **#47 Cross-Sheet Validation** — deep validation
15. **#24 Run Allocation Engine** — allocate costs
16. **#12 Build Dashboard** — charts for leadership
17. **#10 Export Report Package** — PDF for distribution
18. **#19 Append Month to Trend** — finalize
19. **#20 Save Current Scenario** — name it "FY25_MonXX_Close"

### Keyboard Shortcuts You'll Use Every Day

| Shortcut | What It Does |
|----------|-------------|
| **Ctrl+Shift+M** | Open Command Center (the main one) |
| **Ctrl+Shift+H** | Jump to Report (home) sheet |
| **Ctrl+Shift+J** | Quick jump to any sheet |
| **Ctrl+Shift+R** | Run reconciliation checks |

---

## 11. How to Prepare and Run a Demo

### Phase 1: Sanitize the Data (if showing externally)

If the demo audience is outside your company, you need to mask real financial data. Follow `SANITIZATION_PLAYBOOK.md`. Quick version:

1. **Save the original** — make a copy of the `.xlsm` called `_ORIGINAL`
2. **Scale dollar amounts** — pick a random factor (e.g., 1.47), multiply all GL amounts by it. This preserves ratios and trends but makes absolute numbers unrecognizable.
3. **Mask vendor names** — replace real vendor names with `Vendor_001`, `Vendor_002`, etc.
4. **Replace company name** — Ctrl+H, find "Keystone BenefitTech, Inc.", replace with "Acme Corp" across entire workbook.
5. **Check file properties** — File → Info → Check for Issues → Inspect Document → remove personal information and author names.

If the demo is internal, skip this phase.

### Phase 2: Reset to a Clean State

You want the demo to start from a "before" state so you can show the automation doing its thing live.

1. Open the workbook
2. Delete any previously generated sheets (Dashboard, Variance Analysis, Data Quality Report, Cross-Sheet Validation, etc.) — the demo will regenerate them
3. Save

### Phase 3: Demo Script (15-20 Minutes)

Here's a suggested demo flow that shows the most impressive features:

**Opening (2 min):**
- Show the workbook: 13 sheets, 510 GL transactions, $3.7M in data
- "Previously, processing this took 15+ hours per month. Now watch."
- Press **Ctrl+Shift+M** — the Command Center appears
- Point out: 50 commands, 14 categories, search box, keyboard shortcuts

**Data Quality (3 min):**
- Run **#7 Scan Data Quality**
- Show the Data Quality Report sheet — "It found X issues in seconds"
- Run **#8 Fix Text-Stored Numbers** — "One click to fix"
- Re-run **#7** — "Clean now"

**Reconciliation (2 min):**
- Run **#3 Run Reconciliation Checks** (or press Ctrl+Shift+R)
- Show the Checks sheet: PASS/FAIL results with color coding
- "This replaces manual cross-checking between 13 sheets"

**Variance Analysis (3 min):**
- Run **#6 Variance Analysis**
- Show the variance sheet: dollar changes, percent changes, red flags
- Run **#46 Variance Commentary**
- Show the auto-generated executive narrative — "AI-style narrative in seconds"

**Dashboard (2 min):**
- Run **#12 Build Dashboard Charts**
- Show the 3 charts: revenue trend, CM%, revenue mix
- "Ready for the board meeting"

**Report Package (2 min):**
- Run **#10 Export Report Package (PDF)**
- Open the PDF — professional headers, footers, page numbers
- "This used to take an hour of formatting"

**Cross-Sheet Validation (1 min):**
- Run **#47 Cross-Sheet Validation**
- Show the 4 computed validations: GL vs Trend, GL vs Functional, Product check, Mirror check
- "Deep validation that catches what manual checks miss"

**Executive Mode (1 min):**
- Run **#48 Executive Mode Toggle**
- Show how technical sheets disappear, leaving only report-ready tabs
- "One click to make it boardroom-ready"

**Python Dashboard (2 min, optional):**
- Switch to command prompt
- Run `python pnl_runner.py dashboard`
- Show the interactive Streamlit dashboard in the browser
- "Same data, interactive web interface — drag sliders to do what-if analysis"

**Closing (1 min):**
- Run **#50 About This Toolkit** — shows version info
- "50 commands, 29 modules, full audit trail, zero external dependencies"

### Phase 4: Prep for Questions

Common questions you'll get and where to find answers:

| Question | Answer / Reference |
|----------|-------------------|
| "Can we customize the products/departments?" | Yes — change constants in modConfig. See `IMPLEMENTATION_GUIDE.md` Section 5. |
| "Does it work on Mac?" | VBA has limitations on Mac. Windows is required for full functionality. |
| "How does it handle security?" | Audit trail logs every action. Executive Mode hides technical sheets. See `OPERATIONS_RUNBOOK.md`. |
| "What if something breaks?" | Built-in integration test (#44) checks all 29 modules. See `INTEGRATION_TEST_GUIDE.md`. |
| "Can it handle more data?" | Current workbook has 510 rows. VBA handles 100K+ rows. For bigger datasets, use the SQL layer. |
| "What's the learning curve?" | `QUICK_START.md` gets you running in 10 minutes. `USER_TRAINING_GUIDE.md` covers all 50 commands. |

---

## 12. Troubleshooting

### Excel / VBA Issues

| Problem | Cause | Fix |
|---------|-------|-----|
| Ctrl+Shift+M does nothing | Shortcuts not assigned | Close and reopen the workbook, or run `modNavigation.AssignShortcuts` in Immediate Window |
| "Macros have been disabled" yellow bar | Trust Center settings | Go to File → Options → Trust Center → Enable macros (see Step 1D) |
| "Compile error" when opening VBA Editor | Missing module dependency | Make sure modConfig is imported first. Then Debug → Compile. |
| "Subscript out of range" | A sheet was renamed or deleted | Check the error message for which sheet. Verify tab names match modConfig constants. |
| Everything runs slowly | TurboOn not called, or manual calc mode | In Immediate Window: `Application.ScreenUpdating = True` then `Application.Calculation = xlAutomatic` |
| Command Center shows InputBox instead of form | UserForm not built | Run `modFormBuilder.BuildCommandCenter` in Immediate Window (see Step 3) |
| "Programmatic access not trusted" | Trust Center missing checkbox | File → Options → Trust Center → "Trust access to the VBA project object model" |

### Python Issues

| Problem | Cause | Fix |
|---------|-------|-----|
| `python` not recognized | Python not in PATH | Reinstall Python with "Add to PATH" checked |
| `ModuleNotFoundError` | Missing package | Run `pip install -r requirements.txt` |
| `FileNotFoundError` | Can't find the workbook | Set `KBT_SOURCE_FILE` env var or use `--file` flag |
| Dashboard won't open in browser | Firewall blocking port 8501 | Try `streamlit run pnl_dashboard.py --server.port 8502` |
| Tests fail with "file not found" | Workbook path not configured | Set the env var or put the .xlsm in the same folder as the Python scripts |

---

## 13. Quick Reference Cheat Sheet

### The 10 Commands You'll Use Most Often

| # | Command | When | Shortcut |
|---|---------|------|----------|
| 3 | Run Reconciliation Checks | After every data change | Ctrl+Shift+R |
| 6 | Run Variance Analysis | Monthly analysis | — |
| 7 | Scan Data Quality | After every import | — |
| 10 | Export Report Package (PDF) | Month-end | — |
| 12 | Build Dashboard Charts | Before meetings | — |
| 15 | Quick Jump to Sheet | Navigation | Ctrl+Shift+J |
| 16 | Go Home | Navigation | Ctrl+Shift+H |
| 17 | Import GL Data | When new data arrives | — |
| 20 | Save Scenario | Before major changes | — |
| 46 | Variance Commentary | For exec summaries | — |

### The 9 Python Commands

```
python pnl_runner.py dashboard     ← Interactive web dashboard
python pnl_runner.py month-end     ← Automated close checklist
python pnl_runner.py forecast      ← Statistical forecasting
python pnl_runner.py allocate      ← What-if allocation scenarios
python pnl_runner.py snapshot      ← Save/load/compare snapshots
python pnl_runner.py match         ← Fuzzy vendor matching
python pnl_runner.py email         ← Executive email report
python pnl_runner.py test          ← Run 116 automated tests
python pnl_runner.py config        ← Show configuration
```

### File Quick Reference

```
NEED TO...                          → READ THIS FILE
─────────────────────────────────────────────────────
Get started in 10 minutes           → QUICK_START.md
Set up everything from scratch      → IMPLEMENTATION_GUIDE.md
Learn what a specific command does   → USER_TRAINING_GUIDE.md
Run the monthly close               → OPERATIONS_RUNBOOK.md
Prepare for an external demo        → SANITIZATION_PLAYBOOK.md
See what changed in this version    → CHANGELOG.md
Run quality assurance tests         → TEST_PLAN.md
Verify workbook data integrity      → VALIDATION_REPORT.md
Understand the system architecture  → ARCHITECTURE_DIAGRAM.md
Check if all bugs were fixed        → ISSUE_CLOSURE.md
```
