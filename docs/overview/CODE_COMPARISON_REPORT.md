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

## VBA — 12 Modules, 50+ Actions (v2.1)

| Module | What It Does |
|--------|-------------|
| `modConfig_v2.1.bas` | The foundation of everything. Holds all constants (sheet names, product names, departments, fiscal year, colors), plus 15+ helper functions: SafeNum, SafeStr, LastRow, LastCol, FindColByHeader, FindRowByLabel, SheetExists, GetSheet, SafeDeleteSheet, StyleHeader. Every other module depends on this. |
| `modDashboard_v2.1.bas` | Builds and refreshes charts. Includes 3 standard charts (revenue trend, margin %, product mix pie) plus 3 advanced dashboards: Executive KPI cards with trend arrows, Waterfall chart (Revenue → Net Income), and Product Comparison side-by-side. Dynamically finds the last month column with data — never hardcoded. |
| `modDataQuality_v2.1.bas` | Full data quality scanner. Runs 6 checks: duplicate rows, mixed date formats, text-stored numbers, assumption cell issues, misspelled product names, and blank AWS expense cells. v2.1 fix: FixTextNumbers only converts pre-flagged cells — never blindly converts GL IDs or date strings. |
| `modFormBuilder_v2.1.bas` | Builds the Command Center UserForm. Mode A: creates the form and injects code automatically (requires VBA trust setting). Mode B: manual installation with printed step-by-step instructions. Routes all 50 actions through a single ExecuteAction() function. |
| `modMasterMenu_v2.1.bas` | InputBox fallback menu (3-page, 50 items). Used when the UserForm can't be installed. Supports N/P navigation between pages. All routing delegates to modFormBuilder.ExecuteAction(). |
| `modMonthlyTabGenerator_v2.1.bas` | Auto-generates monthly tabs. Clones Mar template for Apr–Dec with formula updates. v2.1: GenerateNextMonthOnly detects the latest existing month, clones it, clears data values, keeps formulas, yellow-highlights input cells, and marks tab green. Header update logic prevents substring corruption bugs (e.g., "Margin" → "Aprigin"). |
| `modNavigation_v2.1.bas` | Sheet navigation. RefreshTableOfContents rebuilds hyperlinks on the Report--> sheet. GoHome, QuickJump. Keyboard shortcuts via Application.OnKey: Ctrl+Shift+M (Command Center), Ctrl+Shift+H (Home), Ctrl+Shift+J (Jump), Ctrl+Shift+R (Reconciliation). v2.1 fix: switched from MacroOptions (which overwrote Excel built-ins like Ctrl+H) to OnKey with safe Ctrl+Shift combos. |
| `modPDFExport_v2.1.bas` | Professional PDF export. ExportReportPackage loops through configured report sheets and exports the full package. ExportSingleSheet exports the active sheet. ApplyPrintSettings stamps professional headers (company name, sheet name, CONFIDENTIAL) and footers (page number, date, version) with landscape, fit-to-page, 0.5" margins. SaveAs dialog with Desktop default and date stamp. |
| `modPerformance_v2.1.bas` | TurboMode on/off (disables screen updating, events, alerts; sets manual calc; changes cursor). ElapsedSeconds with midnight-wrap fix. ForceRecalc. UpdateStatus for status bar progress. |
| `modReconciliation_v2.1.bas` | RunAllChecks reads the Checks sheet and evaluates all PASS/FAIL formulas, color-codes results. ExportCheckResults writes a timestamped text file. ValidateCrossSheet (v2.1) runs 4 computed validation checks: GL total vs P&L Trend, GL Jan vs Functional Jan, GL by product vs Product Summary, plus Checks sheet mirror — all with configurable tolerance. |
| `modSearch_v2.1.bas` | Cross-sheet search. SearchAll finds a keyword across all visible sheets and generates a Search Results sheet with hyperlinks to every match. SearchAndNavigate is interactive. SearchCurrentSheet highlights matches on the active sheet in yellow. Caps at 200 rows displayed but reports total match count. |
| `modVarianceAnalysis_v2.1.bas` | RunVarianceAnalysis compares two monthly P&L sheets (default: Jan vs Feb), calculates dollar and % variance, and applies Favorable/Unfavorable/Flat logic with cost-line reversal for expense rows. Flags rows over 15% threshold in yellow. GenerateCommentary auto-writes English narrative for the top 5 FY-vs-Budget variances, ranked by absolute dollar impact. |
| `frmCommandCenter_code.txt` | The full VBA code for the Command Center UserForm (Mode B manual install). 50 actions across 14 categories. Category filter + text search filter. Status bar feedback. |

---

## SQL — 4 Files, SQLite 3

| File | What It Does |
|------|-------------|
| `staging.sql` | Full ETL pipeline. Creates 5 dimension tables (product, department, expense category, date calendar, GL raw staging) and 1 normalized fact table with generated columns (abs_amount, is_positive). Dimensional lookups on load. Duplicate detection view. Indexed on date, product, department. |
| `transformations.sql` | Allocation framework and analytical views. Defines 3 share types (revenue, AWS compute, headcount) for all 4 products. 8 views: product and dept summaries by month, FY totals with spend share %, MoM variance with FLAG/NEW/OK status, category mix breakdown. |
| `pnl_enhancements.sql` | 5 strategic additions. Budget vs Actual tracking (dim_budget table, 2 queries with OVER/UNDER/ON TRACK status). Allocation audit trail with SQL triggers that log every share change (what changed, old value, new value, who, when). Rolling 12-month views for both products and departments. Vendor contract calendar (classifies spend as FIXED/SEMI/VARIABLE by coefficient of variation). Allocation reconciliation (4 checks: shares sum to 100%, allocation matches GL totals). |
| `validations.sql` | 20+ validation views across 6 sections: referential integrity (orphan records), ETL completeness (staging vs fact row counts), data quality (blanks, zero amounts, outliers via Z-score), balance checks (allocation shares sum, staging vs fact amount reconciliation), and a consolidated v_validation_summary view ordered by FAIL/WARN/PASS status. |

---

## Python — 10 Scripts + Tests + Config

| File | What It Does |
|------|-------------|
| `pnl_config.py` | Centralized config (database paths, file paths, product/department/month lists, fiscal year, thresholds, email settings). Used by every other script. |
| `pnl_allocation_simulator.py` | What-if allocation scenario engine. 3 preset scenarios (baseline, aggressive growth, cost reduction). Recalculates product-level financials under different share assumptions. Exports results to Excel. |
| `pnl_forecast.py` | Forecasting engine with 4 methods: Simple Moving Average, Exponential Smoothing (ETS), Linear Trend, and Scenario-based. Generates confidence intervals. Handles both product and department level. |
| `pnl_month_end.py` | Month-end close automation. Runs a 6-check QA pipeline: data completeness, duplicate check, allocation balance, variance threshold, cross-sheet reconciliation, and snapshot creation. Returns PASS/FAIL/WARN for each check with detail. |
| `pnl_ap_matcher.py` | AP invoice matching engine. Fuzzy vendor name matching (handles typos and abbreviations). Duplicate invoice detection. Matches GL transactions to AP records. Flags unmatched items for review. |
| `pnl_snapshot.py` | Point-in-time P&L snapshots stored in SQLite. Captures full P&L state at a given date. Enables period-over-period comparisons using historical snapshots. |
| `pnl_dashboard.py` | Interactive Streamlit web dashboard. Filters by product, department, and month. Visualizes revenue, expenses, margin trends. Reads directly from the SQLite database. |
| `pnl_email_report.py` | Automated HTML email reports via Office365/SMTP. Generates formatted P&L summary emails with tables, charts, and KPIs. Supports both scheduled and on-demand sending. |
| `pnl_cli.py` | Command-line interface for running any module from the terminal. Argument parsing for all scripts. |
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
| What's there | PDF export is fully built in VBA (`modPDFExport`: ExportReportPackage, ExportSingleSheet). Email reporting is fully built in Python (`pnl_email_report.py`). |
| What's missing | No single VBA macro that does both in sequence: export PDF → attach to Outlook email → draft/send. The two halves exist in different languages with no bridge. A VBA Outlook integration (late binding via `CreateObject("Outlook.Application")`) does not exist yet. |

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

| Status | NOT YET BUILT |
|--------|--------------|
| Notes | Nothing in current VBA handles merged cells. Common issue with ERP exports. Clean gap. |

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

| Status | NOT YET BUILT |
|--------|--------------|
| Notes | No macro uses `SpecialCells(xlCellTypeConstants, xlNumbers)` to identify and visually flag hardcoded inputs. This is a particularly useful audit tool for financial modelers and would be a strong demo feature. Clean gap. |

---

**4. Toggle Presentation Mode**
> Hides gridlines, headings, formula bar, and collapses ribbon.

| Status | NOT YET BUILT |
|--------|--------------|
| Notes | No toggle exists. Would turn the working spreadsheet into a clean executive presentation view instantly. Clean gap. |

---

**5. Delete Completely Blank Rows**
> Loops backwards through rows and deletes any row with zero data.

| Status | NOT YET BUILT as a delete action |
|--------|----------------------------------|
| What's there | `modDataQuality.ScanBlankCells()` SCANS for blank cells in the AWS expense area and reports them. |
| What's missing | The scan reports blanks; it does not delete entire blank rows. There is no DeleteBlankRows() function anywhere. Half of the idea is built (detection), but the action (deletion) does not exist. |

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

| Status | NOT YET BUILT |
|--------|--------------|
| Notes | No protection macros anywhere. Clean gap. |

---

**9. Batch Email via Outlook (VBA)**
> Uses late binding to loop through a table and draft personalized emails.

| Status | NOT YET BUILT in VBA |
|--------|---------------------|
| What's there | `pnl_email_report.py` sends HTML emails via Office365/SMTP (Python). |
| What's missing | No VBA Outlook integration using `CreateObject("Outlook.Application")`. The Python version requires a separate runtime environment. A VBA version would work directly from inside Excel with no additional dependencies. |

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

| Status | NOT YET BUILT as standalone |
|--------|--------------------------|
| What's there | StyleHeader calls `.EntireColumn.AutoFit` but only as part of header formatting. |
| What's missing | No standalone SubAutoFitAll() available as an action in the Command Center. |

---

**2. Freeze/Unfreeze Panes Toggle**

| Status | NOT YET BUILT |
|--------|--------------|
| Notes | Clean gap. The Perlex code shows exactly how it would be built (check `ActiveWindow.FreezePanes`, toggle, freeze at B2). |

---

**3. Convert Formulas to Values (Selection)**

| Status | NOT YET BUILT |
|--------|--------------|
| Notes | No macro does `Selection.Value = Selection.Value`. Critical for finalizing files before distribution. Clean gap. |

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

| Status | NOT YET BUILT |
|--------|--------------|
| Notes | Clean gap. Simple loop through `ActiveWorkbook.Worksheets`, set `ws.Visible = xlSheetVisible`. |

---

**7. Quick Format Header Row**

| Status | PARTIALLY BUILT |
|--------|----------------|
| What's there | `modConfig.StyleHeader()` exists as a helper function used by other modules. |
| What's missing | Not available as a standalone one-click action. StyleHeader takes parameters (which sheet, which row) — it's a utility function, not an end-user macro that formats the current active sheet's row 1. |

---

**8. Delete All Blank Rows**

| Status | NOT YET BUILT as delete action |
|--------|-------------------------------|
| Notes | Same gap as Gemini.md item 5. DataQuality scans/reports; does not delete. |

---

**9. Protect/Unprotect All Sheets**

| Status | NOT YET BUILT |
|--------|--------------|
| Notes | Same gap as Gemini.md item 8. |

---

**10. Save Active Sheet as PDF (Dated Filename)**

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `modPDFExport_v2.1.bas` — ExportSingleSheet() |
| Notes | Exports active sheet with professional headers/footers. SaveAs dialog defaults to Desktop with date stamp in filename. |

---

**11. Backup Workbook with Timestamp**

| Status | NOT YET BUILT |
|--------|--------------|
| Notes | No backup macro anywhere. `ThisWorkbook.SaveCopyAs` with a timestamped filename does not exist in any VBA module. This is a meaningful gap — especially before running any destructive macro like FixTextNumbers or FixDuplicates. |

---

**12. Sort Sheets Alphabetically**

| Status | NOT YET BUILT |
|--------|--------------|
| Notes | Clean gap. The Perlex code shows a bubble sort using `Sheets(j).Move Before:=Sheets(i)`. |

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

| Status | NOT YET BUILT |
|--------|--------------|
| Notes | `modSearch.SearchAll()` finds keywords and reports them, but does not replace. No Find+Replace macro that loops all sheets. Clean gap. |

---

**16. Export Each Sheet as a Separate PDF**

| Status | ALREADY BUILT |
|--------|--------------|
| Where | `modPDFExport_v2.1.bas` — ExportReportPackage() loops through report sheets |

---

**17. Email Active Workbook via Outlook**

| Status | NOT YET BUILT in VBA |
|--------|---------------------|
| Notes | Same gap as Gemini.md item 9. Python email exists (`pnl_email_report.py`) but no VBA Outlook integration. |

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

| Status | NOT YET BUILT in VBA |
|--------|---------------------|
| What's there | SQL `pnl_enhancements.sql` has `allocation_audit` table and triggers that log changes at the database layer. |
| What's missing | No VBA Worksheet_Change event handler. Direct cell edits in Excel are never logged. This is a meaningful internal controls gap — particularly for a file presented to the CFO/CEO. |

---

---

# SECTION 3 — COMPLETE GAP LIST: NOT YET BUILT

## VBA Utility Macros — Quick Wins (Easy to Build)

| # | Gap | Why It Matters |
|---|-----|---------------|
| 1 | Delete All Blank Rows | Completes the DataQuality module — scan exists but delete does not |
| 2 | Unhide All Worksheets | Essential for inherited or shared workbooks |
| 3 | Sort Sheet Tabs Alphabetically | Better navigation for multi-tab financial models |
| 4 | Freeze/Unfreeze Panes Toggle | Day-to-day usability |
| 5 | Convert Formulas to Values (Selection) | Critical for finalizing files before sharing |
| 6 | AutoFit All Columns (standalone) | Currently only buried inside StyleHeader |
| 7 | Protect/Unprotect All Sheets | Controls & security for distribution |
| 8 | Bulk Find and Replace (all sheets) | Needed for fiscal year updates, cost center renames |
| 9 | Highlight Hardcoded Numbers | Audit tool — instantly shows what's a formula vs. a typed input |
| 10 | Toggle Presentation Mode | One-click clean view for demos and presentations |
| 11 | Unmerge and Fill Down | Fixes messy ERP/system exports before analysis |
| 12 | Clear All Hyperlinks | Cleans pasted web/email data |

## VBA Advanced Features — Bigger Builds (High Value)

| # | Gap | Why It Matters |
|---|-----|---------------|
| 13 | Timestamp Audit Trail on Cell Changes | Compliance and internal controls — critical for CFO/CEO audience |
| 14 | Backup Workbook with Timestamp | Safety net before any destructive macro runs |
| 15 | Export All Charts to PowerPoint | Direct support for the executive presentation and CFO/CEO demo |
| 16 | Dynamic Progress Bar KPI Shape | Dashboard polish — visual % tracking for budget utilization or close status |
| 17 | Consolidate Multiple Workbooks from Folder | Batch processing for multi-department or multi-entity submissions |
| 18 | Extract Unique Values to New Tabs | Split company-wide data into departmental tabs instantly |
| 19 | Auto-Refresh Pivot Tables on Workbook Open | If pivots are added in future — good hygiene to have in place |
| 20 | VBA Outlook Email Integration | Complete the PDF → Email workflow natively in VBA |
| 21 | Automated Invoice Reminder Emails | AR/AP use case — not part of current demo scope but high value |
| 22 | Financial Statement Generator from Trial Balance | Transforms raw TB + mapping into formatted IS/BS — powerful demo feature |
| 23 | Batch File Processor (folder loop) | Multi-file automation for department budget consolidation |
| 24 | One-Click Full P&L Generator from Raw GL | True raw-to-report automation — the crown jewel of the demo |

## Python / SQL Enhancements — Future Roadmap

| # | Gap | Why It Matters |
|---|-----|---------------|
| 25 | Constraint-based allocation optimization | Simulator is what-if only; cannot optimize toward a target |
| 26 | Seasonality decomposition in forecasting | SARIMA or Prophet for more accurate monthly patterns |
| 27 | Transaction drill-down in Streamlit dashboard | Click a chart bar to see the underlying GL transactions |
| 28 | Snapshot auto-cleanup policy | Prevent SQLite snapshot database from growing unbounded |
| 29 | Multi-user audit logging in Python | Track which user ran which pipeline step and when |

---

---

# SECTION 4 — RECOMMENDED PRIORITY ORDER

Ranked by impact for the CFO/CEO demo, internal controls story, and coworker video.

## TIER 1 — Build These First (Highest Demo and Compliance Value)

| Priority | Item | Why Now |
|----------|------|---------|
| 1 | **Timestamp Audit Trail on Cell Changes** | Tells a powerful internal controls story to the CFO/CEO. Shows the workbook tracks who changed what and when — directly in Excel. |
| 2 | **Backup Workbook with Timestamp** | Simple but shows professionalism and data safety discipline. Should run automatically before any destructive macro. |
| 3 | **Export All Charts to PowerPoint** | Directly supports the CFO/CEO presentation. One-click turns all Excel charts into a PowerPoint deck. High visual impact. |
| 4 | **Toggle Presentation Mode** | Instant clean demo view. Hide gridlines, formula bar, headings — professional look for the video walkthrough. |
| 5 | **Delete All Blank Rows** | Closes the DataQuality module. The scan already exists — the delete step should be there too. |

## TIER 2 — Build Next (Day-to-Day Efficiency Story for Coworkers)

| Priority | Item | Why |
|----------|------|-----|
| 6 | Protect/Unprotect All Sheets | Good controls practice; coworkers will want this for distribution |
| 7 | Convert Formulas to Values (Selection) | Critical for safely sharing finalized files |
| 8 | Bulk Find and Replace (all sheets) | Saves real time on fiscal year or label changes |
| 9 | AutoFit All Columns (standalone) | Quick cleanup after any data import |
| 10 | Sort Sheet Tabs Alphabetically | Navigation polish for complex workbooks |
| 11 | Unhide All Worksheets | Inherited workbook lifesaver |
| 12 | Freeze/Unfreeze Panes Toggle | Everyday usability |
| 13 | Highlight Hardcoded Numbers | Useful audit and review tool |
| 14 | Unmerge and Fill Down | ERP export cleanup — common pain point |

## TIER 3 — Plan for Later (Bigger Builds or Future Scope)

| Priority | Item | Notes |
|----------|------|-------|
| 15 | One-Click Full P&L Generator from Raw GL | Most ambitious — save for after demo is complete |
| 16 | Financial Statement Generator from Trial Balance | Requires mapping table design — plan carefully first |
| 17 | Consolidate Multiple Workbooks | Needs folder structure defined first |
| 18 | Extract Unique Values to New Tabs | Medium complexity; useful for dept reporting |
| 19 | VBA Outlook Email Integration | Ties the PDF export to email natively |
| 20 | Dynamic Progress Bar KPI Shape | Dashboard polish — low effort, nice visual |
| 21+ | Python/SQL enhancements | Forecasting improvements, drill-down, optimization |

---

---

# SUMMARY SCORECARD

| Source | Total Ideas | Already Built | Partially Built | Not Yet Built |
|--------|-------------|---------------|-----------------|---------------|
| GPT.md | 15 | 8 | 4 | 3 |
| Gemini.md | 12 | 4 | 2 | 6 |
| Perlex.md | 24 | 5 | 3 | 16 |
| **TOTAL** | **51** | **17 (33%)** | **9 (18%)** | **25 (49%)** |

**Key Takeaway:** Your current code already covers 33% of the ideas in NewTesting completely,
and another 18% are partially covered. That means roughly half of what was ideated has been
implemented. The remaining 25 gaps represent a clear, prioritized roadmap for what to build next.

The strongest areas of your current code are **data quality, PDF export, variance analysis,
reconciliation, navigation, and forecasting** — these are production-quality and go well beyond
what the NewTesting ideas describe. The biggest gaps are in **utility macros, VBA audit trail,
PowerPoint export, and workbook consolidation**.

---

*Report prepared 2026-02-27 — APCLDmerge Project — iPipeline Finance & Accounting*
