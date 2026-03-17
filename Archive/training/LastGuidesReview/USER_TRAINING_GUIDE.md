# KBT P&L Toolkit — User Training Guide

> **Audience:** Finance team members who work in Excel daily but have never run a macro.
> All 65 commands are documented with their number, what they do, when to use them, and what to do if something goes wrong.
> **Last Updated:** 2026-03-12 | **Version:** 2.1.1 (65 actions across 16 categories)

---

## How to Run Any Command

**Option A — Command Center (recommended):**
Press **Ctrl+Shift+M**, find the command by name or number, double-click or click **Run**.

**Option B — InputBox fallback:**
If the UserForm is not installed, press Ctrl+Shift+M, type the command number, press OK.

**Option C — Direct keyboard shortcut** (for the 4 most common actions):

| Shortcut | Action |
|----------|--------|
| Ctrl+Shift+M | Open Command Center |
| Ctrl+Shift+H | Go Home (Report sheet) |
| Ctrl+Shift+J | Quick Jump to any sheet |
| Ctrl+Shift+R | Run All Reconciliation Checks |

---

## Monthly Operations (Commands 1–4)

These are the commands you run every month as part of your standard close cycle.

---

### Command 1 — Generate Monthly Tabs

**What it does:** Creates monthly P&L summary tabs (Apr through Dec) by cloning the March template. Each new tab gets the correct month name, updated formulas pointing to the right columns, and proper headers.

**When to use:** Once at the beginning of the fiscal year, or when you need a new month's tab. After the initial batch, use Command 42 to generate one month at a time.

**What could go wrong:**
- "Tabs already exist" — The tabs have already been created. To recreate, delete them first with Command 2.
- Formula errors (#REF!) on new tabs — The March template may have changed. Verify the March tab formulas are correct before generating.

---

### Command 2 — Delete Generated Tabs

**What it does:** Removes all auto-generated monthly tabs (Apr through Dec), leaving the original Jan/Feb/Mar tabs intact.

**When to use:** If you need to regenerate monthly tabs from scratch, or during fiscal year rollover cleanup.

**What could go wrong:**
- Confirmation prompt — You must click "Yes" to confirm deletion. This is irreversible.
- Data loss — Any manually entered data on generated tabs will be lost. Save a snapshot first (Command 20).

---

### Command 3 — Run Reconciliation Checks

**What it does:** Runs a series of PASS/FAIL validation checks across all sheets. Checks include GL totals vs P&L Trend totals, allocation share sums, formula integrity, and cross-sheet consistency. Results appear on the **Checks** sheet with green (PASS) or red (FAIL) formatting.

**When to use:** After importing new data, after any manual edits, and as part of month-end close. This is your primary "is everything still correct?" command.

**What could go wrong:**
- FAIL results — This is the system working correctly. Read the detail column to understand what's mismatched, then fix the source data.
- Checks sheet is blank — Ensure the GL tab has data and sheet names match modConfig constants.

**Keyboard shortcut:** Ctrl+Shift+R

---

### Command 4 — Export Reconciliation Report

**What it does:** Exports the Checks sheet results as a formatted text file for audit trail or email attachment.

**When to use:** After running Command 3, when you need to share reconciliation results with auditors, managers, or external parties.

**What could go wrong:**
- Empty export — Run Command 3 first to populate check results.

---

## Analysis (Commands 5–6)

Analytical tools for understanding P&L performance.

---

### Command 5 — Run Sensitivity Analysis

**What it does:** Creates a "Sensitivity Analysis" sheet showing how 4 key P&L metrics change when you vary input drivers at 4 different levels (+/- 5%, 10%, 15%, 20%). Useful for understanding which assumptions have the biggest impact.

**When to use:** During budgeting season, when presenting to leadership, or when evaluating what-if scenarios.

**What could go wrong:**
- Results seem unrealistic — Check that the Assumptions sheet has reasonable baseline values.
- Sheet already exists — It will be replaced with fresh analysis each time.

---

### Command 6 — Run Variance Analysis

**What it does:** Compares the current month to the prior month for every P&L line item. Shows dollar change, percent change, and flags any items that moved more than 15% (the threshold from modConfig). Creates a "Variance Analysis" sheet with red/green formatting.

**When to use:** Monthly, after all data has been imported and reconciled. This is your main tool for explaining "what changed this month."

**What could go wrong:**
- Only 1 month of data — Variance requires at least 2 months. The command will warn you.
- Threshold too sensitive/insensitive — The 15% threshold is in modConfig (`VARIANCE_PCT`). Ask your admin to adjust if needed.

---

## Data Quality (Commands 7–9)

Data hygiene tools that find and fix common data problems.

---

### Command 7 — Scan for Data Quality Issues

**What it does:** Scans the entire workbook for: duplicate GL rows, text-stored numbers, blank required fields, date formatting issues, misspelled product/department names, and outlier amounts. Results appear on a "Data Quality Report" sheet.

**When to use:** Every time new data is imported or manually entered. Run this BEFORE reconciliation.

**What could go wrong:**
- Many issues flagged — This is the scanner working correctly. Review each issue type and decide which to fix.
- Scan takes a long time — Large GL datasets (10,000+ rows) may take 10-30 seconds. The status bar shows progress.

---

### Command 8 — Fix Text-Stored Numbers

**What it does:** Converts text-stored numbers (cells that look like numbers but are stored as text) to actual numeric values. **Only converts cells that were flagged by Command 7** — it will not touch GL IDs, dates, or headers.

**When to use:** After running Command 7, if text-stored numbers were found in the Amount column.

**What could go wrong:**
- "Run Scan first" message — You must run Command 7 before this command will work. This is a safety feature.
- Confirmation prompt — Shows how many cells will be converted. Review the count before clicking Yes.

---

### Command 9 — Fix Duplicate Rows

**What it does:** Identifies and optionally removes exact duplicate rows from the GL data. Shows duplicates for review before deletion.

**When to use:** After Command 7 flags duplicates. Common after accidental double-imports.

**What could go wrong:**
- False positives — Two transactions can look identical but be legitimate (e.g., two $500 office supply orders on the same day). Review carefully before confirming deletion.

---

## Reporting (Commands 10–12)

Output generation for stakeholders.

---

### Command 10 — Export Report Package (PDF)

**What it does:** Exports all report sheets (P&L Trend, Product Summary, Functional Trend, monthly summaries, Checks) as a single professional PDF with company headers, footers, page numbers, and timestamps.

**When to use:** Month-end, for distribution to leadership, board, or audit.

**What could go wrong:**
- Blank pages in PDF — Ensure each report sheet has data populated.
- File save dialog — You'll be prompted to choose a save location and filename.

---

### Command 11 — Export Active Sheet (PDF)

**What it does:** Exports only the currently active sheet as a PDF.

**When to use:** When you need a quick PDF of a single sheet (e.g., just the Dashboard or just the Variance Analysis).

**What could go wrong:**
- Wrong sheet exported — Make sure you're on the correct sheet before running.

---

### Command 12 — Build Dashboard Charts

**What it does:** Creates a "Dashboard" sheet with 3 charts: (1) Monthly revenue trend bar chart, (2) Contribution margin % trend line chart, (3) Revenue mix pie chart by product. Charts auto-detect which months have data.

**When to use:** Monthly, after data is finalized. Great for presentations and quick visual health checks.

**What could go wrong:**
- Charts appear blank — Need at least 1 month of data in P&L Trend.
- Old charts remain — The dashboard sheet is rebuilt from scratch each time.

---

## Utilities (Commands 13–16)

Navigation and maintenance tools.

---

### Command 13 — Refresh Table of Contents

**What it does:** Updates the Report sheet's table of contents with hyperlinks to all worksheets in the workbook.

**When to use:** After adding or removing sheets (e.g., after generating monthly tabs).

---

### Command 14 — Recalculate AWS Allocations

**What it does:** Validates that AWS allocation shares sum to 100% and forces a recalculation of all allocation formulas on the AWS Allocation sheet.

**When to use:** After changing allocation percentages on the Assumptions sheet.

**What could go wrong:**
- "Shares don't sum to 100%" — Fix the percentages on the Assumptions sheet first.

---

### Command 15 — Quick Jump to Sheet

**What it does:** Shows a list of all sheets and lets you jump to any one instantly.

**When to use:** Anytime you need to navigate a large workbook quickly.

**Keyboard shortcut:** Ctrl+Shift+J

---

### Command 16 — Go Home (Report Sheet)

**What it does:** Jumps to the Report sheet, cell A1.

**When to use:** When you want to get back to the starting point.

**Keyboard shortcut:** Ctrl+Shift+H

---

## Data & Import (Command 17)

---

### Command 17 — Import GL Data Pipeline

**What it does:** Imports GL transaction data from an external CSV or Excel file into the CrossfireHiddenWorksheet (GL) tab. Validates column headers, checks for required fields, and appends new rows.

**When to use:** When you receive a new data extract from the accounting system.

**What could go wrong:**
- Column mismatch — The source file must have the 7 standard GL columns (ID, Date, Department, Product, Expense Category, Vendor, Amount).
- File format error — Ensure the file is `.csv` or `.xlsx`.

---

## Forecasting (Commands 18–19)

---

### Command 18 — Rolling Forecast

**What it does:** Generates a statistical forecast for the remaining months of the fiscal year based on historical trends. Uses simple moving average methodology.

**When to use:** Mid-month or during planning cycles to project full-year results.

**What could go wrong:**
- Insufficient data — Need at least 3 months of actuals for meaningful forecasts.

---

### Command 19 — Append Month to Trend

**What it does:** Takes the current month's finalized data and appends it as a new column on the P&L Trend sheet.

**When to use:** As the final step of month-end close, after all data is reconciled and approved.

**What could go wrong:**
- Column already exists — The month may already be on the trend sheet. Check before running.

---

## Scenarios (Commands 20–23)

Save, load, and compare different versions of the P&L.

---

### Command 20 — Save Current Scenario

**What it does:** Saves a named snapshot of all current P&L values (not formulas — just the computed results). Useful as a restore point before making changes.

**When to use:** Before any major edits, before month-end adjustments, at the start of each month.

---

### Command 21 — Load Scenario

**What it does:** Restores a previously saved scenario, overwriting current values with the saved snapshot.

**When to use:** When you need to revert to a prior state or compare a saved version.

**What could go wrong:**
- Data overwrite — Loading a scenario replaces current values. Save your current state first (Command 20).

---

### Command 22 — Compare Scenarios

**What it does:** Creates a side-by-side comparison of two saved scenarios showing differences by line item.

**When to use:** During reviews to show what changed between versions, or to compare budget vs. actual scenarios.

---

### Command 23 — Delete Scenario

**What it does:** Permanently removes a saved scenario from the workbook.

**When to use:** To clean up old or no-longer-needed snapshots.

---

## Allocation (Commands 24–25)

Cost allocation engine.

---

### Command 24 — Run Allocation Engine

**What it does:** Allocates costs across the 4 product lines using the allocation methods defined for each department (revenue share, headcount-based, or blended). Creates an "Allocation Output" sheet with a Department x Product pivot.

**When to use:** Monthly, after GL data is imported and validated.

**What could go wrong:**
- Allocation shares don't sum to 100% — Run Command 14 first to validate.

---

### Command 25 — Allocation Scenario Preview

**What it does:** Shows a preview of what allocations would look like under different share assumptions, without committing the changes.

**When to use:** During planning to evaluate "what if we shifted more spend to iGO?"

---

## Consolidation (Commands 26–30)

Multi-entity consolidation tools.

---

### Command 26 — Consolidation Menu

**What it does:** Opens a sub-menu for entity consolidation operations.

---

### Command 27 — Add Entity File

**What it does:** Loads a subsidiary or business unit P&L workbook for consolidation.

**When to use:** When preparing a consolidated P&L across multiple entities.

---

### Command 28 — Generate Consolidated P&L

**What it does:** Combines all loaded entity P&Ls into a single consolidated view with elimination entries.

---

### Command 29 — View Loaded Entities

**What it does:** Lists all currently loaded entity files and their status.

---

### Command 30 — Add Elimination Entry

**What it does:** Adds an intercompany elimination entry to the consolidation.

**When to use:** To remove intercompany transactions from the consolidated P&L.

---

## Version Control (Commands 31–35)

Track and manage workbook versions.

---

### Command 31 — Version Control Menu

**What it does:** Opens the version control sub-menu.

---

### Command 32 — Save Version

**What it does:** Saves a timestamped version record with a description of changes.

**When to use:** Before and after significant changes. Creates an audit trail.

---

### Command 33 — Compare Versions

**What it does:** Shows differences between two saved versions.

---

### Command 34 — Restore Version

**What it does:** Reverts the workbook to a previously saved version state.

**What could go wrong:**
- Irreversible — Save the current state first (Command 32 or 20).

---

### Command 35 — List Versions

**What it does:** Displays all saved versions with timestamps and descriptions.

---

## Governance (Commands 36–40)

Documentation and change management.

---

### Command 36 — Auto-Documentation

**What it does:** Generates a "Tech Documentation" sheet listing all VBA modules, their public subs/functions, and descriptions. Useful for auditors and new team members.

**When to use:** When onboarding new team members or preparing for an audit.

---

### Command 37 — Change Management Menu

**What it does:** Opens the change request management sub-menu.

---

### Command 38 — Add Change Request

**What it does:** Creates a new change request (CR) record on the Change Management Log sheet. Prompts for description, priority, and requester.

**When to use:** When someone requests a modification to the toolkit or workbook structure.

---

### Command 39 — Update CR Status

**What it does:** Updates the status of an existing change request (e.g., from "Open" to "In Progress" to "Closed").

---

### Command 40 — CR Summary Report

**What it does:** Generates a summary of all change requests grouped by status.

**When to use:** During team meetings or governance reviews.

---

## Admin & Testing (Commands 41–45)

Audit log management and system testing.

---

### Command 41 — View Audit Log

**What it does:** Shows the VBA_AuditLog sheet which tracks every command execution with timestamps, module names, and details.

**When to use:** When troubleshooting ("what was the last thing that ran?") or for audit compliance.

---

### Command 42 — Export Audit Log

**What it does:** Exports the audit log to a CSV file for external analysis or archival.

---

### Command 43 — Clear Audit Log

**What it does:** Clears all entries from the audit log. Asks for confirmation first.

**When to use:** Periodically, after exporting, to prevent the log from growing too large.

---

### Command 44 — Full Integration Test

**What it does:** Runs a comprehensive test suite that exercises all 23+ modules, verifies cross-module dependencies, and reports PASS/FAIL for each test case. Results go to an "Integration Test Report" sheet.

**When to use:** After importing new module versions, after major changes, or as a quarterly system health check.

**What could go wrong:**
- Some tests fail — Read the failure details. Common causes: missing sheets, renamed tabs, or broken formulas.

---

### Command 45 — Quick Health Check

**What it does:** A lightweight version of Command 44 that tests only the most critical module connections (Config loaded, Logger working, Performance timer OK, key sheets exist).

**When to use:** Quick daily check or when something "feels off."

---

## Advanced (Commands 46–50)

Power-user and administrative tools.

---

### Command 46 — Variance Commentary

**What it does:** Scans the variance analysis results (from Command 6) and auto-generates a plain-English executive narrative covering the top 5 most impactful variances. Creates a "Variance Commentary" sheet with explanatory text.

**When to use:** After running Command 6, when you need to explain variances to leadership in words rather than numbers.

---

### Command 47 — Cross-Sheet Validation

**What it does:** Performs 4 computed cross-sheet validations: (1) GL total vs P&L Trend FY total, (2) GL January vs Functional P&L Jan, (3) GL by Product vs Product Line Summary, (4) Mirror Checks. Results written to a "Cross-Sheet Validation" sheet with PASS/FAIL/REVIEW status.

**When to use:** As an additional layer of validation beyond Command 3, especially before publishing reports.

---

### Command 48 — Executive Mode Toggle

**What it does:** Toggles between "Executive Mode" (hides technical sheets, shows only report-ready tabs) and "Full Mode" (shows all sheets).

**When to use:** Before presenting to leadership or sharing the workbook externally.

---

### Command 49 — Force Recalculate All

**What it does:** Forces Excel to recalculate every formula in the workbook. Useful when automatic calculation is turned off or when formulas seem stale.

**When to use:** If numbers look wrong or haven't updated after a data change.

---

### Command 50 — About This Toolkit

**What it does:** Displays version information, build date, and a summary of the toolkit's capabilities.

---

## Sheet Tools (Commands 51–62)

Everyday worksheet utilities for cleanup, formatting, and maintenance.

---

### Command 51 — Delete All Blank Rows

**What it does:** Deletes every completely blank row within your current selection. Only removes rows where every cell is empty — rows with even one value are kept.

**When to use:** After imports that leave blank rows, or to clean up messy data before analysis.

**What could go wrong:**
- "Please select a range first" — Select the area you want to clean before running.
- Rows with hidden data — If a row appears blank but has a space character, it won't be deleted. Use Data Quality Scan (Command 7) to find hidden spaces.

---

### Command 52 — Unhide All Worksheets

**What it does:** Makes every hidden and very-hidden worksheet visible.

**When to use:** When sheets have been hidden for presentation mode (Command 48) and you need to see everything, or when troubleshooting missing data.

---

### Command 53 — Sort Sheets Alphabetically

**What it does:** Reorders all worksheet tabs A-Z by name.

**When to use:** When the workbook has many tabs and you want them organized for easy navigation.

---

### Command 54 — Toggle Freeze Panes

**What it does:** Toggles freeze panes on or off. When turning on, it freezes at cell B2 so row 1 (headers) and column A (labels) stay visible while scrolling.

**When to use:** When working with large data sheets where you need headers to stay visible.

---

### Command 55 — Convert Formulas to Values

**What it does:** Replaces all formulas in the current selection with their computed values. This is a one-way operation — the formulas are gone permanently.

**When to use:** When you need to "lock in" calculated results, or before sharing a sheet externally without exposing formulas.

**What could go wrong:**
- Data loss — This cannot be undone. Save a version (Command 32) or snapshot (Command 20) first.

---

### Command 56 — AutoFit All Columns

**What it does:** Auto-sizes every column on the active sheet to fit its widest content.

**When to use:** After importing data or pasting values when columns are too narrow or too wide.

---

### Command 57 — Protect All Sheets

**What it does:** Applies password protection to every worksheet in the workbook. Prompts you for a password.

**When to use:** Before sharing the workbook externally to prevent accidental edits.

---

### Command 58 — Unprotect All Sheets

**What it does:** Removes password protection from every worksheet. Prompts for the password.

**When to use:** When you need to edit protected sheets.

---

### Command 59 — Find & Replace (All Sheets)

**What it does:** Performs a find-and-replace operation across every worksheet in the workbook at once. Prompts for the search text and replacement text.

**When to use:** When you need to change a value, label, or department name everywhere in the workbook (e.g., renaming "Dept A" to "Engineering" on all sheets).

---

### Command 60 — Highlight Hardcoded Numbers

**What it does:** Scans the active sheet and changes the font color to blue for any cell that contains a hardcoded number (not a formula). Formula cells are left unchanged.

**When to use:** During audits or model reviews to identify which numbers are manually entered vs. calculated.

---

### Command 61 — Toggle Presentation Mode

**What it does:** Toggles gridlines, row/column headings, and the formula bar on or off. When off, the workbook looks clean and polished for screen sharing or presentations.

**When to use:** Before presenting to leadership or sharing your screen in a meeting.

---

### Command 62 — Unmerge and Fill Down

**What it does:** Unmerges all merged cells in the current selection and fills the blank cells below with the value from the cell above. This is essential for making merged-cell data usable for sorting, filtering, and PivotTables.

**When to use:** When working with reports that use merged cells (common in downloaded reports from ERP systems).

---

## What-If Demo (Commands 63–65)

Live scenario modeling tools for presentations and planning.

---

### Command 63 — Run What-If Scenario Demo

**What it does:** Shows a menu of 7 preset what-if scenarios plus a custom option and restore option. When you pick a scenario, it saves a backup of your current Assumptions values, applies the percentage change to the matching drivers, recalculates the entire P&L model, and creates a styled "What-If Impact" report sheet showing before/after values.

**Preset scenarios:**
1. Revenue drops 15%
2. Revenue increases 10%
3. AWS costs increase 25%
4. Headcount grows 20%
5. All expenses cut 10%
6. Best case: Revenue +15%, Expenses -5%
7. Worst case: Revenue -20%, Expenses +15%
8. Custom (pick your own driver and %)
9. Restore original values

**When to use:** During presentations to leadership, budget planning sessions, or when someone asks "what if revenue drops?" and you want to show the answer in 3 seconds.

**What could go wrong:**
- "Assumptions sheet not found" — The Assumptions sheet must exist with driver names in column A and values in column B.
- Impact report shows 0 changes — No drivers matched the category keywords. Check that Assumptions has drivers with names containing "revenue", "aws", "headcount", etc.

---

### Command 64 — Custom What-If Analysis

**What it does:** Shows all drivers from the Assumptions sheet with their current values. You pick a specific driver and enter any percentage change. Same backup/apply/report workflow as Command 63.

**When to use:** When someone asks about a specific driver that isn't in the 7 presets (e.g., "What if licensing revenue goes up 20%?").

---

### Command 65 — Restore Baseline (Undo What-If)

**What it does:** Restores all Assumptions values back to their originals from before the last What-If scenario. Deletes the backup and impact report sheets and recalculates the model.

**When to use:** After demonstrating a What-If scenario, to reset the workbook to its original state before running another scenario or continuing with normal work.

**What could go wrong:**
- "No baseline saved" — You need to run a What-If scenario (Command 63 or 64) first.

---

## Quick Reference — All 65 Commands

| # | Category | Command | Key Use Case |
|---|----------|---------|-------------|
| 1 | Monthly Ops | Generate Monthly Tabs | Start of FY |
| 2 | Monthly Ops | Delete Generated Tabs | FY cleanup |
| 3 | Monthly Ops | Run Reconciliation Checks | After every data change |
| 4 | Monthly Ops | Export Reconciliation Report | Audit trail |
| 5 | Analysis | Sensitivity Analysis | Budget planning |
| 6 | Analysis | Variance Analysis | Monthly close |
| 7 | Data Quality | Scan Data Quality | After data import |
| 8 | Data Quality | Fix Text-Stored Numbers | After scan flags issues |
| 9 | Data Quality | Fix Duplicate Rows | After scan flags dupes |
| 10 | Reporting | Export Report Package (PDF) | Month-end distribution |
| 11 | Reporting | Export Active Sheet (PDF) | Ad hoc export |
| 12 | Reporting | Build Dashboard Charts | Monthly visuals |
| 13 | Utilities | Refresh Table of Contents | After sheet changes |
| 14 | Utilities | Recalculate AWS Allocations | After assumption edits |
| 15 | Utilities | Quick Jump to Sheet | Navigation |
| 16 | Utilities | Go Home (Report) | Navigation |
| 17 | Data & Import | Import GL Data Pipeline | Monthly data load |
| 18 | Forecasting | Rolling Forecast | Mid-month planning |
| 19 | Forecasting | Append Month to Trend | Month-end final step |
| 20 | Scenarios | Save Current Scenario | Before major edits |
| 21 | Scenarios | Load Scenario | Revert to prior state |
| 22 | Scenarios | Compare Scenarios | Version comparison |
| 23 | Scenarios | Delete Scenario | Cleanup |
| 24 | Allocation | Run Allocation Engine | Monthly allocation |
| 25 | Allocation | Allocation Preview | What-if planning |
| 26 | Consolidation | Consolidation Menu | Multi-entity |
| 27 | Consolidation | Add Entity File | Load subsidiary |
| 28 | Consolidation | Generate Consolidated P&L | Combine entities |
| 29 | Consolidation | View Loaded Entities | Status check |
| 30 | Consolidation | Add Elimination Entry | Intercompany adj |
| 31 | Version Control | Version Control Menu | Track changes |
| 32 | Version Control | Save Version | Before changes |
| 33 | Version Control | Compare Versions | Diff versions |
| 34 | Version Control | Restore Version | Rollback |
| 35 | Version Control | List Versions | History |
| 36 | Governance | Auto-Documentation | Audit prep |
| 37 | Governance | Change Management Menu | CR workflow |
| 38 | Governance | Add Change Request | New CR |
| 39 | Governance | Update CR Status | CR lifecycle |
| 40 | Governance | CR Summary Report | Governance review |
| 41 | Admin & Testing | View Audit Log | Troubleshooting |
| 42 | Admin & Testing | Export Audit Log | Archival |
| 43 | Admin & Testing | Clear Audit Log | Maintenance |
| 44 | Admin & Testing | Full Integration Test | System health |
| 45 | Admin & Testing | Quick Health Check | Daily check |
| 46 | Advanced | Variance Commentary | Exec narratives |
| 47 | Advanced | Cross-Sheet Validation | Deep validation |
| 48 | Advanced | Executive Mode Toggle | Presentation prep |
| 49 | Advanced | Force Recalculate All | Formula refresh |
| 50 | Advanced | About This Toolkit | Version info |
| 51 | Sheet Tools | Delete All Blank Rows | Data cleanup |
| 52 | Sheet Tools | Unhide All Worksheets | Show hidden sheets |
| 53 | Sheet Tools | Sort Sheets Alphabetically | Tab organization |
| 54 | Sheet Tools | Toggle Freeze Panes | Header visibility |
| 55 | Sheet Tools | Convert Formulas to Values | Lock in results |
| 56 | Sheet Tools | AutoFit All Columns | Column sizing |
| 57 | Sheet Tools | Protect All Sheets | Security |
| 58 | Sheet Tools | Unprotect All Sheets | Remove locks |
| 59 | Sheet Tools | Find & Replace (All Sheets) | Bulk rename |
| 60 | Sheet Tools | Highlight Hardcoded Numbers | Audit review |
| 61 | Sheet Tools | Toggle Presentation Mode | Clean display |
| 62 | Sheet Tools | Unmerge and Fill Down | Fix merged data |
| 63 | What-If Demo | Run What-If Scenario Demo | Live demo scenarios |
| 64 | What-If Demo | Custom What-If Analysis | Custom driver changes |
| 65 | What-If Demo | Restore Baseline (Undo) | Reset after what-if |
