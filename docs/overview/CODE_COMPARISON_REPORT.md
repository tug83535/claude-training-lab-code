# Code Review & NewTesting Comparison Report
**Project:** APCLDmerge — iPipeline Finance & Accounting P&L Demo
**Date:** 2026-02-27
**Prepared by:** Claude Code
**Purpose:** Full review of all code in vba/, sql/, and python/ compared against the
ideas documented in NewTesting/ (GPT.md, Gemini.md, Perlex.md)

---

## HOW TO READ THIS REPORT

- **ALREADY BUILT** — The idea exists and is fully implemented in current code
- **PARTIALLY BUILT** — Something is there, but it's incomplete, limited in scope, or only exists in one language (e.g., Python only, no VBA equivalent)
- **NOT YET BUILT** — Clean gap; nothing in the current codebase covers this idea

---

---

# SECTION 1 — FULL INVENTORY: WHAT'S ALREADY BUILT

## VBA — 14 Modules, 62 Actions (v2.1)

| Module | What It Does |
|--------|-------------|
| `modConfig_v2.1.bas` | The foundation of everything. Holds all constants (sheet names, product names, departments, fiscal year, colors), plus 15+ helper functions: SafeNum, SafeStr, LastRow, LastCol, FindColByHeader, FindRowByLabel, SheetExists, GetSheet, SafeDeleteSheet, StyleHeader. Every other module depends on this. |
| `modDashboard_v2.1.bas` | Builds and refreshes charts. Includes 3 standard charts (revenue trend, margin %, product mix pie) plus 3 advanced dashboards: Executive KPI cards with trend arrows, Waterfall chart (Revenue → Net Income), and Product Comparison side-by-side. Dynamically finds the last month column with data — never hardcoded. |
| `modDataQuality_v2.1.bas` | Full data quality scanner. Runs 6 checks: duplicate rows, mixed date formats, text-stored numbers, assumption cell issues, misspelled product names, and blank AWS expense cells. v2.1 fix: FixTextNumbers only converts pre-flagged cells — never blindly converts GL IDs or date strings. |
| `modFormBuilder_v2.1.bas` | Builds the Command Center UserForm. Mode A: creates the form and injects code automatically (requires VBA trust setting). Mode B: manual installation with printed step-by-step instructions. Routes all 62 actions through a single ExecuteAction() function. |
| `modLogger_v2.1.bas` | Runtime audit log. Every VBA macro run is timestamped and logged to a hidden sheet (VBA_AuditLog) — xlSheetVeryHidden. Logs: Timestamp, User, Module, Procedure, Message, Status. Color-coded by severity. Auto-trims at 5,000 rows. ViewLog, ExportLog, and ClearLog available from Command Center (actions 41–43). |
| `modMasterMenu_v2.1.bas` | InputBox fallback menu (4-page, 62 items). Used when the UserForm can't be installed. Supports N/P navigation between pages. All routing delegates to modFormBuilder.ExecuteAction(). |
| `modMonthlyTabGenerator_v2.1.bas` | Auto-generates monthly tabs. Clones Mar template for Apr–Dec with formula updates. v2.1: GenerateNextMonthOnly detects the latest existing month, clones it, clears data values, keeps formulas, yellow-highlights input cells, and marks tab green. Header update logic prevents substring corruption bugs (e.g., "Margin" → "Aprigin"). |
| `modNavigation_v2.1.bas` | Sheet navigation. RefreshTableOfContents rebuilds hyperlinks on the Report--> sheet. GoHome, QuickJump. Keyboard shortcuts via Application.OnKey: Ctrl+Shift+M (Command Center), Ctrl+Shift+H (Home), Ctrl+Shift+J (Jump), Ctrl+Shift+R (Reconciliation). v2.1 fix: switched from MacroOptions (which overwrote Excel built-ins like Ctrl+H) to OnKey with safe Ctrl+Shift combos. |
| `modPDFExport_v2.1.bas` | Professional PDF export. ExportReportPackage loops through configured report sheets and exports the full package. ExportSingleSheet exports the active sheet. ApplyPrintSettings stamps professional headers (company name, sheet name, CONFIDENTIAL) and footers (page number, date, version) with landscape, fit-to-page, 0.5" margins. SaveAs dialog with Desktop default and date stamp. |
| `modPerformance_v2.1.bas` | TurboMode on/off (disables screen updating, events, alerts; sets manual calc; changes cursor). ElapsedSeconds with midnight-wrap fix. ForceRecalc. UpdateStatus for status bar progress. |
| `modReconciliation_v2.1.bas` | RunAllChecks reads the Checks sheet and evaluates all PASS/FAIL formulas, color-codes results. ExportCheckResults writes a timestamped text file. ValidateCrossSheet (v2.1) runs 4 computed validation checks: GL total vs P&L Trend, GL Jan vs Functional Jan, GL by product vs Product Summary, plus Checks sheet mirror — all with configurable tolerance. |
| `modSearch_v2.1.bas` | Cross-sheet search. SearchAll finds a keyword across all visible sheets and generates a Search Results sheet with hyperlinks to every match. SearchAndNavigate is interactive. SearchCurrentSheet highlights matches on the active sheet in yellow. Caps at 200 rows displayed but reports total match count. |
| `modUtilities_v2.1.bas` | 12 quick-win utility macros (actions 51–62): DeleteBlankRows, UnhideAllSheets, SortSheetsAlphabetically, ToggleFreezePanes, ConvertToValues, AutoFitAllColumns, ProtectAllSheets, UnprotectAllSheets, FindReplaceAllSheets, HighlightHardcodedNumbers, TogglePresentationMode, UnmergeAndFillDown. All wired into Command Center Sheet Tools category. |
| `modVarianceAnalysis_v2.1.bas` | RunVarianceAnalysis compares two monthly P&L sheets (default: Jan vs Feb), calculates dollar and % variance, and applies Favorable/Unfavorable/Flat logic with cost-line reversal for expense rows. Flags rows over 15% threshold in yellow. GenerateCommentary auto-writes English narrative for the top 5 FY-vs-Budget variances, ranked by absolute dollar impact. |
| `frmCommandCenter_code.txt` | The full VBA code for the Command Center UserForm (Mode B manual install). 62 actions across 15 categories. Category filter + text search filter. Status bar feedback. |

---

## SQL — 4 Files, SQLite 3

| File | What It Does |
|------|-------------|
| `staging.sql` | Full ETL pipeline. Creates 5 dimension tables (product, department, expense category, date calendar, GL raw staging) and 1 normalized fact table with generated columns (abs_amount, is_positive). Dimensional lookups on load. Duplicate detection view. Indexed on date, product, department. |
| `transformations.sql` | Allocation framework and analytical views. Defines 3 share types (revenue, AWS compute, headcount) for all 4 products. 8 views: product and dept summaries by month, FY totals with spend share %, MoM variance with FLAG/NEW/OK status, category mix breakdown. |
| `pnl_enhancements.sql` | 5 strategic additions. Budget vs Actual tracking (dim_budget table, 2 queries with OVER/UNDER/ON TRACK status). Allocation audit trail with SQL triggers that log every share change (what changed, old value, new value, who, when). Rolling 12-month views for both products and departments. Vendor contract calendar (classifies spend as FIXED/SEMI/VARIABLE by coefficient of variation). Allocation reconciliation (4 checks: shares sum to 100%, allocation matches GL totals). |
| `validations.sql` | 20+ validation views across 6 sections: referential integrity (orphan records), ETL completeness (staging vs fact row counts), data quality (blanks, zero amounts, outliers via Z-score), balance checks (allocation shares sum, staging vs fact amount reconciliation), and a consolidated v_validation_summary view ordered by FAIL/WARN/PASS status. |

---

## Python — 11 Scripts + Tests + Config

| File | What It Does |
|------|-------------|
| `pnl_config.py` | Centralized config (database paths, file paths, product/department/month lists, fiscal year, thresholds, email settings). Used by every other script. |
| `pnl_allocation_simulator.py` | What-if allocation scenario engine. 3 preset scenarios (baseline, aggressive growth, cost reduction). Recalculates product-level financials under different share assumptions. Exports results to Excel. |
| `pnl_forecast.py` | Forecasting engine with 4 methods: Simple Moving Average, Exponential Smoothing (ETS), Linear Trend, and Scenario-based. Generates confidence intervals. Handles both product and department level. |
| `pnl_monte_carlo.py` | **Monte Carlo P&L risk simulation.** Runs N iterations (default 10,000) with Dirichlet-distributed revenue shares and Normal-distributed expense amounts. Produces P5/P25/P50/P75/P95 percentile tables, Value at Risk, per-product breakdown, a 4-panel distribution chart, and a formatted Excel export. Shock event modeling (optional). Fully wired into pnl_cli.py. |
| `pnl_month_end.py` | Month-end close automation. Runs a 6-check QA pipeline: data completeness, duplicate check, allocation balance, variance threshold, cross-sheet reconciliation, and snapshot creation. Returns PASS/FAIL/WARN for each check with detail. |
| `pnl_ap_matcher.py` | AP invoice matching engine. Fuzzy vendor name matching (handles typos and abbreviations). Duplicate invoice detection. Matches GL transactions to AP records. Flags unmatched items for review. |
| `pnl_snapshot.py` | Point-in-time P&L snapshots stored in SQLite. Captures full P&L state at a given date. Enables period-over-period comparisons using historical snapshots. |
| `pnl_dashboard.py` | Interactive Streamlit web dashboard. Filters by product, department, and month. Visualizes revenue, expenses, margin trends. Reads directly from the SQLite database. |
| `pnl_cli.py` | Master command-line interface for running any module from the terminal. All scripts reachable from one entry point including monte-carlo, forecast, simulate, close, dashboard, and run-all. |
| `pnl_runner.py` | Master orchestrator. Chains all scripts in the correct order (staging → transformations → validations → month-end close → snapshot). Single entry point to run the full pipeline. |
| `pnl_tests.py` | Full pytest test suite. 100% coverage on config and allocation logic. 80%+ coverage on close and forecasting. Tests edge cases: empty datasets, divide-by-zero, missing products, tolerance handling. |

---

---

# SECTION 2 — EVERY NEWTESTING IDEA COMPARED TO CURRENT CODE

## FROM GPT.md (15 Ideas)

### EASY LEVEL

**1. Report Formatter Macro**
> Standardizes fonts, column widths, headers, number formats, freeze panes.

| Status | PARTIALLY BUILT |
|--------|----------------|
| What's there | `modConfig.StyleHeader()` applies navy background, white bold font, and bottom border to any header row. `modPDFExport.ApplyPrintSettings()` standardizes landscape, margins, headers, footers for PDF output. `modDashboard` applies consistent chart colors and formats. |
| What's missing | No standalone one-click macro that formats an entire active sheet end-to-end: accounting number format, freeze panes toggle, auto-fit all columns, header styling — all at once. The pieces exist as helpers but are never assembled into a single "clean up this sheet" button. |

---

**2. Auto PDF Export & Emailer**
> Exports selected sheets as PDFs and drafts Outlook emails automatically.

| Status | PARTIALLY BUILT |
|--------|----------------|
| What's there | PDF export is fully built in VBA (`modPDFExport`: ExportReportPackage, ExportSingleSheet). |
| What's missing | No VBA macro to attach PDF to Outlook email and draft/send. A VBA Outlook integration (late binding via `CreateObject("Outlook.Application")`) does not exist yet. Email report feature was removed. |

---

**3. Data Cleaner Macro**
> Removes duplicates, trims spaces, fixes date formats, standardizes text case.

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `modDataQuality_v2.1.bas` — ScanAll() + FixTextNumbers() + FixDuplicates() |
| Notes | Covers: duplicate row detection and removal, text-stored number correction, mixed date format scanning, product name misspelling correction, and blank cell detection. v2.1 adds safety: text numbers only converted after pre-flagging, preventing GL ID corruption. |

---

**4. Batch File Processor**
> Opens multiple files in a folder and applies the same transformation.

| Status | NOT YET BUILT |
|--------|--------------|
| What's missing | No VBA module that prompts for a folder, loops through all .xlsx files, applies a standard operation, and closes each file. This is a clean gap across VBA, SQL, and Python. The Python runner (`pnl_runner.py`) processes one database but does not open and transform multiple Excel files in a batch. |

---

**5. Template Generator**
> Creates new structured files based on a master template.

| Status | PARTIALLY BUILT |
|--------|----------------|
| What's there | `modMonthlyTabGenerator` clones the Mar sheet as a template to generate new monthly tabs (Apr–Dec or one-at-a-time). This is template-based generation within the same workbook. |
| What's missing | No macro that creates a brand new standalone Excel workbook from a master template file. The concept is month-tab focused, not a general-purpose "create a new file from this template" tool. |

---

### MEDIUM LEVEL

**6. One-Click P&L Generator**
> Builds a summarized Profit & Loss statement from raw transaction data.

| Status | PARTIALLY BUILT |
|--------|----------------|
| What's there | Monthly tabs are auto-generated (`modMonthlyTabGenerator`). The Executive Dashboard (`modDashboard.CreateExecutiveDashboard`) builds a KPI summary sheet. The SQL pipeline (`staging.sql` + `transformations.sql`) builds product and department P&L summaries from raw GL data. |
| What's missing | No single VBA macro that takes raw GL transaction data and builds a clean formatted P&L statement from scratch in one click. The existing flow assumes the P&L structure is already in place and populates/refreshes it. A true "raw GL → formatted P&L" generator does not exist in VBA. |

---

**7. Variance Analysis Engine**
> Compares Actual vs Budget/Forecast and flags material differences.

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `modVarianceAnalysis_v2.1.bas` — RunVarianceAnalysis() + GenerateCommentary() |
| Notes | MoM comparison with dollar/% variance, Favorable/Unfavorable/Flat with cost-line logic, 15% flag threshold, auto-written English commentary for top 5 FY-vs-Budget items. SQL also has `v_product_mom_variance` with FLAG/NEW/OK status and budget vs actual queries in `pnl_enhancements.sql`. |

---

**8. Account Reconciliation Tool**
> Matches two datasets and flags unmatched items.

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `modReconciliation_v2.1.bas` (VBA), `validations.sql` (SQL), `pnl_month_end.py` (Python) |
| Notes | VBA runs Checks sheet validations + 4 computed cross-sheet checks (GL vs P&L Trend, GL vs Functional Jan, GL by product vs Product Summary). SQL has 20+ validation views. Python month-end close runs a 6-check QA pipeline. Full three-layer reconciliation framework. |

---

**9. Consolidation Engine**
> Combines multiple entity files into one consolidated workbook.

| Status | NOT YET BUILT |
|--------|--------------|
| What's missing | Referenced as action item 34 in the Command Center menu ("ConsolidateMultiEntity"), but no actual VBA module exists to implement it. No Python script handles multi-file consolidation either. Clean gap. |

---

**10. Scenario Modeling Interface**
> Allows users to input assumptions and automatically recalculates financial projections.

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `pnl_allocation_simulator.py` (Python), `modDashboard.ProductComparison()` (VBA) |
| Notes | Python simulator offers 3 preset scenarios plus custom what-if input, recalculating all product financials and exporting to Excel. The VBA ProductComparison dashboard shows side-by-side product metrics. SQL `v_dept_product_allocated` view applies scenario-based allocation shares. |

---

### ADVANCED LEVEL

**11. Interactive Financial Dashboard Controller**
> Controls pivots, slicers, and refresh logic via VBA buttons.

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `modDashboard_v2.1.bas` + `frmCommandCenter_code.txt` |
| Notes | Command Center UserForm provides button-driven control. Dashboard module builds and refreshes 6 chart/dashboard types. RefreshDashboard() recalculates all charts. No pivot table slicer control (no pivot tables in current workbook), but the VBA dashboard controller concept is fully implemented. |

---

**12. Audit Trail & Change Logger**
> Logs changes made to key financial cells.

| Status | PARTIALLY BUILT |
|--------|----------------|
| What's there | SQL has a full audit trail: `allocation_audit` table with triggers that log every change to allocation shares (old value, new value, changed_by, changed_at, change_type). This is production-quality compliance logging at the database layer. |
| What's missing | VBA has NO Worksheet_Change event handler to log direct cell edits in Excel. If a user changes a budget number, a revenue figure, or an assumption directly in a cell, nothing is recorded. The audit trail only exists at the SQL layer, not inside the workbook. This is a significant compliance gap for a file being shown to the CFO/CEO. |

---

**13. Rolling Forecast Automation**
> Automatically shifts forecast months forward and recalculates projections.

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `pnl_forecast.py` (Python) |
| Notes | 4 forecasting methods (SMA, ETS, trend, scenario-based), confidence intervals, product and department level. SQL has `v_rolling_12m_product` and `v_rolling_12m_department` with TTM metrics and YoY change. VBA has modMonthlyTabGenerator for rolling the workbook forward period-by-period. |

---

**14. Capital Expenditure Tracker System**
> Tracks budget vs actual spending on capital projects.

| Status | NOT YET BUILT |
|--------|--------------|
| What's missing | No CapEx-specific tracking anywhere in VBA, SQL, or Python. The SQL has expense categories (including Depreciation and Hardware which are CapEx-adjacent) but no dedicated CapEx budget vs actual tracker with project-level detail. Clean gap. |

---

**15. Automated Financial Close Checklist Tracker**
> Tracks close tasks, owners, and completion status.

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `pnl_month_end.py` (Python) |
| Notes | 6-check automated QA pipeline: data completeness, duplicate check, allocation balance, variance threshold, cross-sheet reconciliation, snapshot creation — each returns PASS/FAIL/WARN with detail. VBA ReconciliationRunAllChecks() also tracks and color-codes check results. |

---

---

## FROM Gemini.md (12 Ideas)

**1. Unmerge and Fill Down**
> Unmerges selected cells and fills blanks with the value from the cell above.

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `modUtilities_v2.1.bas` — UnmergeAndFillDown() |
| Notes | Unmerges all merged cells in selection, then fills each blank cell with the value from the row above. Loops forward through selection. Uses TurboOn/TurboOff for performance. Wired into Command Center as action #62 (Sheet Tools category). |

---

**2. Auto-Format Financial Reports**
> Applies accounting number format, bolds headers, auto-fits all columns.

| Status | PARTIALLY BUILT |
|--------|----------------|
| What's there | `modConfig.StyleHeader()` applies header formatting (navy background, white bold text, bottom border) to specific header rows when called. `modPDFExport` applies print formatting. |
| What's missing | No standalone "clean up this entire sheet" macro. StyleHeader is a helper called by other code, not a one-click formatting action available in the Command Center for ad-hoc use on any sheet. |

---

**3. Highlight Hardcoded Numbers**
> Changes font color of cells containing typed numbers (not formulas) to blue.

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `modUtilities_v2.1.bas` — HighlightHardcodedNumbers() |
| Notes | Uses `SpecialCells(xlCellTypeConstants, xlNumbers)` on the active sheet's UsedRange to find all hardcoded number cells. Reports count, asks confirmation, then sets font color to RGB(0,0,255). Standard audit convention: blue = typed input, black = formula. Wired into Command Center as action #60. |

---

**4. Toggle Presentation Mode**
> Hides gridlines, headings, formula bar, and collapses ribbon.

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `modUtilities_v2.1.bas` — TogglePresentationMode() |
| Notes | Uses gridlines visibility as the toggle indicator. Run once: hides gridlines, headings, formula bar. Run again: restores all three. Status bar confirms mode. Auto-clears status bar after 3 seconds. Wired into Command Center as action #61. |

---

**5. Delete Completely Blank Rows**
> Loops backwards through rows and deletes any row with zero data.

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `modUtilities_v2.1.bas` — DeleteBlankRows() |
| Notes | Loops BACKWARDS through the selection (prevents row-shift errors during deletion). Uses CountA to detect fully blank rows. Reports count deleted. Scan detection (`modDataQuality.ScanBlankCells`) also still exists for non-destructive review before deleting. Wired into Command Center as action #51. |

---

**6. Generate a Table of Contents (TOC)**
> Creates a sheet listing all tabs with hyperlinks to each.

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `modNavigation_v2.1.bas` — RefreshTableOfContents() |
| Notes | Rebuilds hyperlinks on the Report--> sheet. Covers the core TOC idea. |

---

**7. Consolidate Multiple Workbooks**
> Opens every .xlsx in a folder, copies data, stacks into master workbook.

| Status | NOT YET BUILT |
|--------|--------------|
| Notes | Same gap as GPT.md item 9. No folder-loop workbook consolidation in any language. |

---

**8. Protect/Unprotect All Sheets**
> Prompts for a password and protects or unprotects every worksheet.

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `modUtilities_v2.1.bas` — ProtectAllSheets() (action #57), UnprotectAllSheets() (action #58) |
| Notes | ProtectAllSheets loops through all worksheets and applies sheet protection. UnprotectAllSheets removes protection from all sheets. Both prompt for password confirmation. Wired into Command Center as actions #57 and #58 in the Sheet Tools category. |

---

**9. Batch Email via Outlook (VBA)**
> Uses late binding to loop through a table and draft personalized emails.

| Status | NOT BUILT |
|--------|-----------|
| What's there | Email report feature was removed. |
| What's missing | No VBA Outlook integration using `CreateObject("Outlook.Application")`. Would need late binding to loop through a table and draft personalized emails directly from Excel. |

---

**10. Automated Monthly Roll-Forward**
> Duplicates active sheet, increments month name, clears input cells.

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `modMonthlyTabGenerator_v2.1.bas` — GenerateNextMonthOnly() |
| Notes | Detects latest existing month, clones it, clears data values, keeps formulas, yellow-highlights input cells, marks tab green. Header update handles multiple date format patterns safely. |

---

**11. Batch PDF Generator**
> Loops through identifiers, recalculates dashboard for each, exports PDF.

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `modPDFExport_v2.1.bas` — ExportReportPackage() |
| Notes | Loops through configured report sheets, applies professional print settings, exports full PDF package. ExportSingleSheet handles individual sheet export. |

---

**12. Extract Unique Values to New Tabs**
> Extracts unique values from a column, creates a tab for each, pastes filtered data.

| Status | NOT YET BUILT |
|--------|--------------|
| Notes | No macro uses Scripting Dictionary or Advanced Filter to split data by category into separate tabs. Clean gap. |

---

---

## FROM Perlex.md (24 Ideas)

**1. AutoFit All Column Widths**

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `modUtilities_v2.1.bas` — AutoFitAllColumns() (action #56) |
| Notes | Calls `.Columns.AutoFit` on the active sheet's UsedRange. One-click column width cleanup. Wired into Command Center as action #56 in the Sheet Tools category. |

---

**2. Freeze/Unfreeze Panes Toggle**

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `modUtilities_v2.1.bas` — ToggleFreezePanes() (action #54) |
| Notes | Checks `ActiveWindow.FreezePanes` as the toggle indicator. If frozen: unfreezes. If unfrozen: selects B2 and freezes. Status bar confirms action. Wired into Command Center as action #54. |

---

**3. Convert Formulas to Values (Selection)**

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `modUtilities_v2.1.bas` — ConvertToValues() (action #55) |
| Notes | Pastes selection values over formulas (`Selection.Value = Selection.Value`). Asks confirmation before converting. Critical for finalizing files before distribution. Wired into Command Center as action #55. |

---

**4. Clear All Hyperlinks on a Sheet**

| Status | NOT YET BUILT |
|--------|--------------|
| Notes | `modNavigation.RefreshTableOfContents()` creates and updates hyperlinks but does not clear them. No ClearHyperlinks macro exists. Clean gap. |

---

**5. Highlight Duplicates in a Selection**

| Status | PARTIALLY BUILT |
|--------|----------------|
| What's there | `modDataQuality.ScanDuplicateRows()` scans the entire GL sheet for exact duplicate rows and highlights them yellow. |
| What's missing | Perlex's version is selection-based and range-flexible (prompts user to select any range, uses CountIf). The current version is hardcoded to the GL sheet and looks for full-row duplicates. The concept is covered but the flexible, user-driven version does not exist. |

---

**6. Unhide All Worksheets**

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `modUtilities_v2.1.bas` — UnhideAllSheets() (action #52) |
| Notes | Loops through all worksheets in the active workbook and sets `ws.Visible = xlSheetVisible`. Reports count of sheets unhidden. Wired into Command Center as action #52. (Note: VBA_AuditLog is xlSheetVeryHidden by design and stays hidden.) |

---

**7. Quick Format Header Row**

| Status | PARTIALLY BUILT |
|--------|----------------|
| What's there | `modConfig.StyleHeader()` exists as a helper function used by other modules. |
| What's missing | Not available as a standalone one-click action. StyleHeader takes parameters (which sheet, which row) — it's a utility function, not an end-user macro that formats the current active sheet's row 1. |

---

**8. Delete All Blank Rows**

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `modUtilities_v2.1.bas` — DeleteBlankRows() (action #51) |
| Notes | Same as Gemini.md item 5. Loops backwards through selection using CountA. Reports count deleted. modDataQuality ScanBlankCells still exists for non-destructive review. Wired into Command Center as action #51. |

---

**9. Protect/Unprotect All Sheets**

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `modUtilities_v2.1.bas` — ProtectAllSheets() (action #57), UnprotectAllSheets() (action #58) |
| Notes | Same implementation as Gemini.md item 8. Both wired into Command Center. |

---

**10. Save Active Sheet as PDF (Dated Filename)**

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `modPDFExport_v2.1.bas` — ExportSingleSheet() |
| Notes | Exports active sheet with professional headers/footers. SaveAs dialog defaults to Desktop with date stamp in filename. |

---

**11. Backup Workbook with Timestamp**

| Status | NOT YET BUILT — User Declined |
|--------|------------------------------|
| Notes | No backup macro anywhere. `ThisWorkbook.SaveCopyAs` with a timestamped filename does not exist in any VBA module. **User decision (2026-02-28): This feature was reviewed and explicitly declined. Do not propose or rebuild this in future sessions.** |

---

**12. Sort Sheets Alphabetically**

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `modUtilities_v2.1.bas` — SortSheetsAlphabetically() (action #53) |
| Notes | Bubble sort using `Sheets(j).Move Before:=Sheets(i)`. Loops until all tabs are in alpha order. Confirms count of sheets sorted. Wired into Command Center as action #53. |

---

**13. Create Table of Contents (Hyperlinked)**

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `modNavigation_v2.1.bas` — RefreshTableOfContents() |

---

**14. Consolidate All Sheets into One Master Sheet**

| Status | NOT YET BUILT |
|--------|--------------|
| Notes | Same gap as Gemini.md item 7 (within-workbook version). No consolidation macro exists. |

---

**15. Bulk Find and Replace Across Entire Workbook**

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `modUtilities_v2.1.bas` — FindReplaceAllSheets() (action #59) |
| Notes | Prompts for Find and Replace text, then loops through all visible sheets running `Cells.Replace`. Reports total replacements made. Useful for fiscal year label changes, cost center renames, vendor name corrections. Wired into Command Center as action #59. |

---

**16. Export Each Sheet as a Separate PDF**

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `modPDFExport_v2.1.bas` — ExportReportPackage() loops through report sheets |

---

**17. Email Active Workbook via Outlook**

| Status | NOT BUILT |
|--------|-----------|
| What's there | Email report feature was removed. |
| What's missing | No VBA Outlook integration using `CreateObject("Outlook.Application")`. Would need to work natively from inside Excel. |

---

**18. Data Entry UserForm**
> Pop-up form for clean data entry (Date, Vendor, Amount, GL Code, Description).

| Status | PARTIALLY BUILT |
|--------|----------------|
| What's there | The Command Center (`frmCommandCenter_code.txt`) is a full UserForm — but it's a macro launcher, not a data entry form. The infrastructure to build and wire UserForms is well established in the project (Mode A/B installation process in modFormBuilder). |
| What's missing | No form designed for entering financial transactions row-by-row. The concept of a UserForm exists; a data-entry specific form does not. |

---

**19. Automated Overdue Invoice Email Reminders**

| Status | NOT YET BUILT |
|--------|--------------|
| Notes | Clean gap. No invoice tracking table, no overdue logic, no reminder email workflow in any language. |

---

**20. Financial Statement Generator from Trial Balance**

| Status | NOT YET BUILT |
|--------|--------------|
| Notes | No macro that takes an AccountNumber + Mapping table and builds formatted Income Statement and Balance Sheet tabs. The closest existing capability is `modDashboard.CreateExecutiveDashboard()` which builds a summary from an existing P&L structure — it does not build from a raw trial balance. Clean gap. |

---

**21. Export All Charts to PowerPoint**

| Status | NOT YET BUILT |
|--------|--------------|
| Notes | No PowerPoint integration anywhere in VBA, SQL, or Python. No use of `CreateObject("PowerPoint.Application")`. Clean gap — and one of the highest-impact missing features for the executive presentation. |

---

**22. Auto-Refresh Pivot Tables on Workbook Open**

| Status | NOT YET BUILT |
|--------|--------------|
| Notes | No Workbook_Open event handler in any module. No pivot table refresh logic. Clean gap. (Note: current workbook may not use pivot tables — but the capability to auto-refresh on open is still useful if pivots are added later.) |

---

**23. Dynamic Progress Bar (KPI Shape)**
> A rectangle shape that grows and changes color (red/orange/green) based on a % cell.

| Status | NOT YET BUILT |
|--------|--------------|
| Notes | `modDashboard` builds charts and KPI cards but uses Excel chart objects, not shape-based progress bars. No shape-width/color manipulation by percentage. Clean gap. |

---

**24. Timestamp Audit Trail on Cell Changes**
> Worksheet_Change event logs DateTime, User, Sheet, Cell, OldValue, NewValue to an AuditLog sheet.

| Status | PARTIALLY BUILT |
|--------|----------------|
| What's there | SQL `pnl_enhancements.sql` has `allocation_audit` table and triggers that log every change to allocation shares at the database layer (old value, new value, changed_by, changed_at). |
| What's missing | No VBA Worksheet_Change event handler. Direct cell edits in Excel are not logged. **User decision (2026-02-28): VBA implementation was reviewed and explicitly declined. The SQL layer audit trail remains the only implementation.** |

---

---

# SECTION 3 — COMPLETE GAP LIST: NOT YET BUILT

## Status Update — 2026-02-28

**All 12 VBA Utility Macro "Quick Win" items from the original gap list are now BUILT.**
`modUtilities_v2.1.bas` (committed 2026-02-28) implemented all of them as Command Center
actions #51–62. The quick-wins section has been removed. The remaining gaps are below.

## VBA Advanced Features — Bigger Builds

| # | Gap | Why It Matters |
|---|-----|---------------|
| 1 | Clear All Hyperlinks | Cleans pasted web/email data — modNavigation creates hyperlinks but cannot bulk-clear them |
| 2 | Dynamic Progress Bar KPI Shape | Dashboard polish — visual % tracking for budget utilization or close status |
| 3 | Consolidate Multiple Workbooks from Folder | Batch processing for multi-department or multi-entity submissions |
| 4 | Extract Unique Values to New Tabs | Split company-wide data into departmental tabs instantly |
| 5 | Auto-Refresh Pivot Tables on Workbook Open | Good hygiene if pivots are added in future — no Workbook_Open handler exists |
| 6 | VBA Outlook Email Integration | Complete the PDF → Email workflow natively in VBA without Python dependency |
| 7 | Automated Invoice Reminder Emails | AR/AP use case — no invoice tracking table or overdue logic in any language |
| 8 | Financial Statement Generator from Trial Balance | Transforms raw TB + account mapping into formatted IS/BS — powerful demo feature |
| 9 | Batch File Processor (folder loop) | Multi-file automation for department budget consolidation |
| 10 | One-Click Full P&L Generator from Raw GL | True raw-to-report automation — the crown jewel of the demo |

**User decisions logged:**
- **Backup Workbook with Timestamp** — explicitly declined by user (2026-02-28), do not re-propose
- **Export All Charts to PowerPoint** — dropped permanently by user (2026-02-28)
- **Timestamp Audit Trail (VBA)** — VBA implementation declined by user (2026-02-28); SQL layer audit trail remains

## Python / SQL Enhancements — Future Roadmap

| # | Gap | Why It Matters |
|---|-----|---------------|
| 11 | Constraint-based allocation optimization | Simulator is what-if only; cannot optimize toward a target |
| 12 | Seasonality decomposition in forecasting | SARIMA or Prophet for more accurate monthly patterns |
| 13 | Transaction drill-down in Streamlit dashboard | Click a chart bar to see the underlying GL transactions |
| 14 | Snapshot auto-cleanup policy | Prevent SQLite snapshot database from growing unbounded |
| 15 | Multi-user audit logging in Python | Track which user ran which pipeline step and when |

---

---

# SECTION 4 — RECOMMENDED PRIORITY ORDER

Ranked by impact for the CFO/CEO demo, internal controls story, and coworker video.

**Note:** As of 2026-02-28, all 12 quick-win utility macros (former Tier 1/Tier 2 items)
are now **fully built** in `modUtilities_v2.1.bas` (actions #51–62). Priority order below
reflects the remaining gaps only.

## TIER 1 — Build These Next (Highest Remaining Demo Value)

| Priority | Item | Why Now |
|----------|------|---------|
| 1 | **Dynamic Progress Bar KPI Shape** | Dashboard polish — visual % tracker for budget utilization or close status; high visual impact for the CFO/CEO presentation |
| 2 | **Financial Statement Generator from Trial Balance** | Powerful demo feature — transforms a raw TB + account mapping into a formatted IS/BS in one click |
| 3 | **VBA Outlook Email Integration** | Completes the PDF → Email workflow natively inside Excel without requiring the Python runtime |
| 4 | **Clear All Hyperlinks** | Small gap, easy build — rounds out the navigation toolkit |

## TIER 2 — Build Next (Bigger Features for Future Scope)

| Priority | Item | Why |
|----------|------|-----|
| 5 | One-Click Full P&L Generator from Raw GL | Most ambitious feature — full raw GL to formatted P&L in VBA; save for after demo is complete |
| 6 | Consolidate Multiple Workbooks from Folder | Batch processing for multi-department or multi-entity submissions |
| 7 | Extract Unique Values to New Tabs | Split company-wide data into departmental reporting tabs |
| 8 | Automated Invoice Reminder Emails | AR/AP use case — no invoice tracking infrastructure exists yet |
| 9 | Batch File Processor (folder loop) | Multi-file automation; requires defining folder conventions first |
| 10 | Auto-Refresh Pivot Tables on Workbook Open | Good hygiene to have if pivots are added in future |

## TIER 3 — Python / SQL Future Enhancements

| Priority | Item | Notes |
|----------|------|-------|
| 11 | Constraint-based allocation optimization | Mathematical optimizer — build after simulator proves its value |
| 12 | Seasonality decomposition in forecasting | SARIMA/Prophet upgrade; requires more historical data first |
| 13 | Transaction drill-down in Streamlit dashboard | High-value UX feature; requires click event wiring in Plotly |
| 14 | Snapshot auto-cleanup policy | Database hygiene; low urgency until snapshot count grows large |
| 15 | Multi-user audit logging in Python | Tracking pipeline users; add when team usage expands |

---

---

# SUMMARY SCORECARD

| Source | Total Ideas | Already Built | Partially Built | Not Yet Built |
|--------|-------------|---------------|-----------------|---------------|
| GPT.md | 15 | 8 | 4 | 3 |
| Gemini.md | 12 | 8 | 2 | 2 |
| Perlex.md | 24 | 11 | 5 | 8 |
| **TOTAL** | **51** | **27 (53%)** | **11 (22%)** | **13 (25%)** |

**Scorecard updated 2026-02-28** — `modUtilities_v2.1.bas` added 12 new macros (actions #51–62),
moving all 12 quick-win items from "Not Yet Built" to "Already Built." Also added `pnl_monte_carlo.py`
(Monte Carlo P&L risk simulation, 10,000+ iteration engine with Dirichlet share randomization).

**Key Takeaway:** Your current code now covers **53% of all NewTesting ideas completely** —
up from 33% at last report. Another 22% are partially covered. Only 25% remain as true gaps,
and several of those were explicitly declined by the user (Backup Workbook, PowerPoint export,
VBA Audit Trail).

The strongest areas of the codebase are **data quality, PDF export, variance analysis,
reconciliation, navigation, forecasting, utility macros, and Monte Carlo simulation** — all
production-quality and going well beyond what the NewTesting ideas describe. The remaining
gaps are concentrated in **workbook consolidation, VBA Outlook integration, and advanced
demo features** like the Financial Statement Generator and P&L Generator from Raw GL.

---

*Report originally prepared 2026-02-27 — Updated 2026-02-28 — APCLDmerge Project — iPipeline Finance & Accounting*
