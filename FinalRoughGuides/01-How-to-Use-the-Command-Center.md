# How to Use the Command Center

## iPipeline P&L Automation Toolkit — Complete User Guide

**Document Version:** 1.0
**Last Updated:** March 5, 2026
**Audience:** All iPipeline Finance & Accounting Employees
**Prepared by:** Finance Automation Team, iPipeline

---

## Table of Contents

1. [What Is the Command Center?](#1-what-is-the-command-center)
2. [How to Open the Command Center](#2-how-to-open-the-command-center)
3. [Understanding the Command Center Layout](#3-understanding-the-command-center-layout)
4. [How to Find and Run an Action](#4-how-to-find-and-run-an-action)
5. [All 62 Actions — Complete Reference](#5-all-62-actions--complete-reference)
6. [Category-by-Category Walkthrough](#6-category-by-category-walkthrough)
7. [Keyboard Shortcuts](#7-keyboard-shortcuts)
8. [Tips for Daily Use](#8-tips-for-daily-use)
9. [What to Do If Something Goes Wrong](#9-what-to-do-if-something-goes-wrong)
10. [Frequently Asked Questions](#10-frequently-asked-questions)

---

## 1. What Is the Command Center?

The Command Center is a single control panel built into the iPipeline P&L Excel workbook. It gives you one-click access to **62 automated actions** — everything from generating monthly reports to running data quality scans to exporting PDF packages.

**Think of it like a remote control for the entire workbook.** Instead of navigating menus, writing formulas, or remembering where things are, you open the Command Center, pick the action you want, and click Run.

### What It Replaces

| Before (Manual) | After (Command Center) |
|---|---|
| Manually copying sheets for each month | Action 1: Generate Monthly Tabs — done in seconds |
| Eyeballing numbers for errors | Action 7: Scan for Data Quality Issues — catches everything |
| Building charts by hand | Action 12: Build Dashboard Charts — instant professional charts |
| Ctrl+F on one sheet at a time | Action 15: Quick Jump to Sheet — searches everything at once |
| Emailing spreadsheets back and forth | Action 10: Export Report Package (PDF) — polished 7-sheet PDF |

### Who Should Use It

Everyone. If you open the P&L file, you should use the Command Center. It is designed for Finance and Accounting professionals who work in Excel every day but do not need to know anything about VBA, macros, or code.

---

## 2. How to Open the Command Center

There are **three ways** to open the Command Center. Use whichever is easiest for you.

### Method 1: Keyboard Shortcut (Fastest)

1. Make sure the P&L workbook is open and is the active window
2. Press **Ctrl + Shift + M** on your keyboard (hold all three keys at the same time)
3. The Command Center window will appear in the center of your screen

> **What you should see:** A pop-up window titled "AUTOMATION COMMAND CENTER" with a list of categories on the left and actions on the right.

### Method 2: Click the Button on the Home Sheet

1. Open the P&L workbook
2. Click on the **"Home"** sheet tab at the bottom of Excel (it should be the first tab)
3. You will see a large button labeled **"Open Command Center"**
4. Click that button once
5. The Command Center window will appear

> **What you should see:** The Home sheet has two buttons — "Open Command Center" and "View Sheet Index". Click the first one.

### Method 3: Run from the VBA Editor (Advanced — Not Recommended for Most Users)

1. Press **Alt + F8** to open the Macro dialog box
2. In the list, find **LaunchCommandCenter**
3. Click **Run**
4. The Command Center window will appear

> **Note:** Method 3 is only if Methods 1 and 2 are not working. Most users should never need this.

### Troubleshooting: Command Center Won't Open

| Problem | Solution |
|---|---|
| Nothing happens when I press Ctrl+Shift+M | Make sure the P&L workbook is the active window (click on it first). Make sure macros are enabled (see the First Time Setup guide). |
| I get a "Macros have been disabled" yellow bar | Click "Enable Content" in the yellow bar at the top of Excel. |
| I get an error message | Close the Command Center if it's partially open, then try again. If the error persists, close and reopen the workbook. |
| I don't see a Home sheet | The Home sheet may not have been created yet. Use Method 1 (Ctrl+Shift+M) instead. |

---

## 3. Understanding the Command Center Layout

When the Command Center opens, you will see a professional window with several sections. Here is exactly what each part does:

```
+------------------------------------------------------------------+
|              AUTOMATION COMMAND CENTER                             |
|              Version 2.1.0                                        |
+------------------------------------------------------------------+
|  [_______ Search actions... _______]                              |
+------------------------------------------------------------------+
|                          |                                        |
|  CATEGORIES              |  ACTIONS                               |
|                          |                                        |
|  > All Actions           |  #1   Generate Monthly Tabs            |
|  > Monthly Operations    |  #2   Delete Generated Tabs            |
|  > Analysis              |  #3   Run Reconciliation Checks        |
|  > Data Quality          |  #4   Export Reconciliation Report      |
|  > Reporting             |  #5   Run Sensitivity Analysis         |
|  > Utilities             |  #6   Run Variance Analysis            |
|  > Data & Import         |  #7   Scan for Data Quality Issues     |
|  > Forecasting           |  ...                                   |
|  > Scenarios             |                                        |
|  > Allocation            |                                        |
|  > Consolidation         |                                        |
|  > Version Control       |                                        |
|  > Governance            |                                        |
|  > Admin & Testing       |                                        |
|  > Advanced              |                                        |
|  > Sheet Tools           |                                        |
|                          |                                        |
+------------------------------------------------------------------+
|  [ Run Selected ]    [ Run & Close ]    [ Close ]                 |
+------------------------------------------------------------------+
|  Status: Ready                                                    |
+------------------------------------------------------------------+
```

### Section-by-Section Breakdown

**Title Bar (Top)**
- Shows "AUTOMATION COMMAND CENTER" and the version number (2.1.0)
- Also displays how many sheets are in the workbook

**Search Box**
- Type any word to instantly filter the action list
- Example: Type "PDF" and only PDF-related actions will show
- Example: Type "variance" and only variance-related actions will show
- The search checks both the action name AND the category name
- To clear your search, delete all text from the box

**Categories Panel (Left Side)**
- Lists 16 categories including "All Actions" at the top
- Click a category to filter the action list to only that category
- Click "All Actions" to see everything again
- Categories are: Monthly Operations, Analysis, Data Quality, Reporting, Utilities, Data & Import, Forecasting, Scenarios, Allocation, Consolidation, Version Control, Governance, Admin & Testing, Advanced, Sheet Tools

**Actions Panel (Right Side)**
- Shows the list of available actions with their number and name
- Each action has a number (#1 through #62) and a description
- Click once on an action to select it (it will highlight)
- Double-click an action to run it immediately

**Buttons (Bottom)**
- **Run Selected** — Runs the selected action and keeps the Command Center open (useful when you want to run multiple actions in a row)
- **Run & Close** — Runs the selected action and closes the Command Center (useful when you only need to do one thing)
- **Close** — Closes the Command Center without running anything

**Status Bar (Very Bottom)**
- Shows "Ready" when the Command Center is waiting for you
- Shows the name of the last action you ran and whether it succeeded

---

## 4. How to Find and Run an Action

### Option A: Browse by Category

1. Open the Command Center (Ctrl + Shift + M)
2. In the **Categories** panel on the left, click a category name
   - Example: Click **"Data Quality"** to see only data quality actions
3. The **Actions** panel on the right will update to show only actions in that category
4. Click once on the action you want to select it
5. Click **"Run Selected"** (to keep the Command Center open) or **"Run & Close"** (to close it after running)

### Option B: Search by Keyword

1. Open the Command Center (Ctrl + Shift + M)
2. Click in the **Search box** at the top
3. Type a keyword related to what you want to do
   - Example: Type **"reconciliation"** to find all reconciliation-related actions
   - Example: Type **"export"** to find all export actions
   - Example: Type **"sheet"** to find sheet management tools
4. The action list will filter in real time as you type
5. Click on the action you want
6. Click **"Run Selected"** or **"Run & Close"**

### Option C: Double-Click to Run Instantly

1. Open the Command Center (Ctrl + Shift + M)
2. Find the action you want (browse or search)
3. **Double-click** on the action name
4. The action will run immediately — no need to click any buttons

### Option D: Run by Action Number (If You Know It)

If you already know the action number from the Quick Reference Card:

1. Open the Command Center (Ctrl + Shift + M)
2. Type the number in the search box (e.g., **"7"** for Scan for Data Quality Issues)
3. Click the matching action
4. Click **"Run Selected"** or **"Run & Close"**

### What Happens After You Run an Action

- The Command Center may briefly hide while the action runs
- You may see a progress bar or status messages in the bottom-left corner of Excel
- When the action finishes, you will typically see one of these:
  - A **message box** telling you the results (e.g., "Data quality scan complete: 3 issues found")
  - A **new sheet** created with the results (e.g., a "Variance Analysis" sheet)
  - A **file saved** to your computer (e.g., a PDF export)
  - The Command Center's **status bar** will update with the result

---

## 5. All 62 Actions — Complete Reference

Below is every single action available in the Command Center, organized by category. For each action, you will find:

- **What it does** — Plain English explanation
- **When to use it** — The situation where this action helps you
- **What to expect** — What you will see after running it

---

### Monthly Operations (Actions 1–4)

#### Action 1: Generate Monthly Tabs

- **What it does:** Creates individual monthly P&L summary sheets for each month of the fiscal year. Each tab is a copy of the template with the correct month name and dates.
- **When to use it:** At the beginning of the fiscal year, or whenever you need to set up the monthly tabs for the first time.
- **What to expect:** New sheet tabs will appear at the bottom of your workbook (e.g., "Functional P&L Summary - Apr 25", "Functional P&L Summary - May 25", etc.). A message box will confirm how many tabs were created.
- **Important:** If the tabs already exist, the action will tell you. It will not create duplicates.

#### Action 2: Delete Generated Tabs

- **What it does:** Removes all the monthly tabs that were created by Action 1. Does NOT delete the original template tabs (Jan, Feb, Mar).
- **When to use it:** If you need to start over, or if the monthly tabs were created with incorrect data and you want to regenerate them fresh.
- **What to expect:** The generated tabs will disappear from the bottom of your workbook. A message box will confirm how many tabs were deleted. Your original tabs (Jan 25, Feb 25, Mar 25) will NOT be touched.
- **Warning:** This cannot be undone. Make sure you want to delete before confirming.

#### Action 3: Run Reconciliation Checks

- **What it does:** Runs a comprehensive set of reconciliation checks across the entire workbook. It compares numbers between sheets to make sure everything ties out — revenue totals match, allocations balance, and cross-sheet references are consistent.
- **When to use it:** During month-end close, after importing new data, or any time you want to verify the workbook is accurate.
- **What to expect:** The "Checks" sheet will be updated with PASS or FAIL for each check. Green = PASS, Red = FAIL. A summary message box will tell you how many checks passed and how many failed.
- **Tip:** Run this after every major change to catch errors early.

#### Action 4: Export Reconciliation Report

- **What it does:** Exports the results from the Checks sheet into a clean, formatted report. This gives you a printable record of which reconciliation checks passed and which failed.
- **When to use it:** After running Action 3 (Reconciliation Checks), when you need a formal record for audit files or management review.
- **What to expect:** A formatted report showing all check results with timestamps.

---

### Analysis (Actions 5–6)

#### Action 5: Run Sensitivity Analysis

- **What it does:** Tests how changes in your key assumptions (revenue growth rate, cost percentages, allocation shares) would affect the bottom line. It runs multiple scenarios automatically — "What if revenue drops 10%? What if costs increase 5%?"
- **When to use it:** During budgeting, forecasting, or when leadership asks "What would happen if...?" questions.
- **What to expect:** A new "Sensitivity Analysis" sheet will be created showing a table of results. Each row represents a different assumption change, and the columns show the impact on key P&L lines.

#### Action 6: Run Variance Analysis

- **What it does:** Compares the current month to the prior month (Month-over-Month) and calculates the dollar and percentage variance for every line item. Any variance greater than 15% is automatically flagged.
- **When to use it:** During month-end review, when preparing commentary for management, or any time you need to understand what changed and by how much.
- **What to expect:** A new "Variance Analysis" sheet will be created with a formatted table. Significant variances (over 15%) are highlighted. The report shows both dollar amounts and percentages.

---

### Data Quality (Actions 7–9)

#### Action 7: Scan for Data Quality Issues

- **What it does:** Performs a comprehensive scan of the entire workbook looking for six types of data problems:
  1. Text-stored numbers (numbers that Excel thinks are text)
  2. Blank cells in critical ranges
  3. Duplicate rows
  4. Formula errors (#REF!, #VALUE!, #DIV/0!, etc.)
  5. Inconsistent formatting
  6. Suspicious values (outliers)
- **When to use it:** Before month-end close, after importing new data, or any time the numbers look "off." This is your first line of defense against data errors.
- **What to expect:** A "Data Quality Report" sheet will be created with a summary count of each issue type, followed by a detailed list of every issue found (sheet name, cell address, what the problem is). The report now also includes a **Letter Grade (A through F)** displayed prominently at the top — A means zero issues, F means 4 or more critical problems.
- **Tip:** Run this at least once a week. It catches problems that are invisible to the naked eye.

#### Action 8: Fix Text-Stored Numbers

- **What it does:** Finds every cell in the workbook where a number is stored as text (a very common Excel problem) and converts it to a real number. This fixes SUM formulas that skip cells, VLOOKUP mismatches, and sorting problems.
- **When to use it:** After Action 7 reports text-stored number issues, or if your SUM formulas are not adding up correctly.
- **What to expect:** A message box will tell you how many cells were fixed. The affected cells will now contain real numbers that formulas can use properly.
- **Important:** This action is safe. It only converts cells that are clearly numbers stored as text. It will never touch dates, names, customer IDs, or any non-numeric data.

#### Action 9: Fix Duplicate Rows

- **What it does:** Scans for duplicate rows in your data and removes the extras, keeping one copy of each unique row.
- **When to use it:** After importing data that may have been pasted twice, or if Action 7 flagged duplicates.
- **What to expect:** A message box will tell you how many duplicate rows were found and removed.
- **Important:** Always review your data after running this to make sure the correct rows were kept.

---

### Reporting (Actions 10–12)

#### Action 10: Export Report Package (PDF)

- **What it does:** Exports a professional multi-sheet PDF containing the 7 key reporting sheets from the workbook. Each sheet is formatted with proper headers, footers, page numbers, and date stamps. The PDF is ready to send to leadership or attach to an email.
- **When to use it:** At month-end when you need to distribute the P&L package to management, or when someone asks for a PDF copy of the reports.
- **What to expect:** A file save dialog will appear. Choose where to save the PDF. The PDF will contain 7 sheets in order, each properly formatted for printing. A message box will confirm the export is complete and tell you the file path.
- **Sheets included:** Report, P&L Monthly Trend, Functional P&L Monthly Trend, Product Line Summary, current month Functional P&L Summary, Checks, and Assumptions.

#### Action 11: Export Active Sheet (PDF)

- **What it does:** Exports only the sheet you are currently looking at as a single PDF file. Same professional formatting as Action 10 but just one sheet.
- **When to use it:** When you only need to share one specific sheet rather than the entire report package.
- **What to expect:** A file save dialog will appear. Choose where to save. The current sheet will be exported as a formatted PDF.

#### Action 12: Build Dashboard Charts

- **What it does:** Creates a set of professional charts on the Executive Dashboard sheet, including revenue trends, expense breakdowns, product comparisons, and waterfall charts. All charts are formatted with iPipeline brand colors.
- **When to use it:** When preparing for a presentation, leadership review, or any time you want a visual summary of the P&L data.
- **What to expect:** The "Executive Dashboard" sheet will be created (or updated if it already exists) with multiple charts. All charts are automatically linked to the workbook data and will update when the data changes.

---

### Utilities (Actions 13–16)

#### Action 13: Refresh Table of Contents

- **What it does:** Creates or updates a Table of Contents sheet that lists every sheet in the workbook with clickable links. This makes it easy to navigate a large workbook.
- **When to use it:** Any time you add or remove sheets and want the Table of Contents to reflect the current state.
- **What to expect:** The Table of Contents sheet will be updated with the current list of all sheets. Each sheet name is a clickable hyperlink.

#### Action 14: Recalculate AWS Allocations

- **What it does:** Validates the AWS allocation table on the "AWS Allocation" sheet and recalculates all allocation amounts to make sure they sum correctly and match the expected totals.
- **When to use it:** After updating AWS cost data, or if reconciliation checks flag an allocation mismatch.
- **What to expect:** The AWS Allocation sheet will be recalculated. A message box will confirm whether the allocations balance or if there are discrepancies.

#### Action 15: Quick Jump to Sheet

- **What it does:** Shows a dialog box where you can type any sheet name (or part of a name) and jump directly to that sheet. No more scrolling through dozens of tabs at the bottom.
- **When to use it:** Any time you want to navigate quickly in a large workbook.
- **What to expect:** A dialog box will appear. Type the name (or partial name) of the sheet you want. Click OK and you will be taken directly to that sheet.

#### Action 16: Go Home (Report Sheet)

- **What it does:** Takes you directly to the "Report-->" sheet, which is the main landing page of the workbook.
- **When to use it:** Any time you want to get back to the main report view quickly.
- **What to expect:** The active sheet will change to "Report-->". You can also use the keyboard shortcut **Ctrl + Shift + H**.

---

### Data & Import (Action 17)

#### Action 17: Import GL Data Pipeline

- **What it does:** Opens a file picker dialog that lets you import new General Ledger (GL) data from a CSV or Excel file into the workbook. The import pipeline validates the data format, checks for required columns, and loads the data into the GL detail sheet.
- **When to use it:** When you receive new GL data exports from the accounting system (e.g., Crossfire) and need to bring them into the P&L model.
- **What to expect:** A file picker dialog will appear. Select your CSV or Excel file. The system will validate the file, show you a preview of what will be imported, and load the data. A message box will confirm how many rows were imported.
- **Important:** The import expects specific column headers (ID, Date, Department, Product, Category, Vendor, Amount). If your file does not match, the system will tell you what is missing.

---

### Forecasting (Actions 18–19)

#### Action 18: Rolling Forecast

- **What it does:** Generates a rolling forecast based on historical trends in the P&L data. It uses the actual numbers from completed months to project the remaining months of the fiscal year.
- **When to use it:** During mid-year planning, when leadership asks "Where are we heading?", or when you need to update the forecast after actual results come in.
- **What to expect:** A "Rolling Forecast" sheet will be created showing projected values for each remaining month, with a comparison to budget.

#### Action 19: Append Month to Trend

- **What it does:** Adds the next month's column to the P&L Monthly Trend and Functional P&L Monthly Trend sheets. It reads today's date, determines the next calendar month automatically, and prepares the column.
- **When to use it:** At the start of each new month, before entering the new month's data.
- **What to expect:** A new column will appear on both trend sheets for the next month. The column header will be highlighted in yellow so you can easily find it. If there is a corresponding Functional P&L Summary tab, a new tab will be cloned for the new month.

---

### Scenarios (Actions 20–23)

#### Action 20: Save Current Scenario

- **What it does:** Saves a snapshot of the current Assumptions sheet values as a named scenario. This lets you preserve a set of assumptions (e.g., "Base Case", "Optimistic", "Worst Case") so you can load them back later.
- **When to use it:** Before making changes to assumptions, or when you want to create multiple what-if scenarios for comparison.
- **What to expect:** A dialog box will ask you to name the scenario. Type a name (e.g., "Q1 Base Case") and click OK. The scenario will be saved to a hidden Scenarios sheet.

#### Action 21: Load Scenario

- **What it does:** Loads a previously saved scenario back into the Assumptions sheet. This overwrites the current assumption values with the saved values.
- **When to use it:** When you want to switch between scenarios (e.g., switch from "Base Case" to "Worst Case" to see how the P&L changes).
- **What to expect:** A dialog box will show you a list of all saved scenarios. Select one and click OK. The Assumptions sheet will be updated with the saved values, and all downstream calculations will update automatically.
- **Warning:** Loading a scenario overwrites current Assumptions values. Save your current scenario first if you want to keep it.

#### Action 22: Compare Scenarios

- **What it does:** Creates a side-by-side comparison of two or more saved scenarios, showing the differences in assumption values and their impact on key P&L metrics.
- **When to use it:** When you need to present multiple scenarios to leadership, or when you want to understand the differences between two sets of assumptions.
- **What to expect:** A comparison sheet will be created showing each scenario's values side by side with the differences highlighted.

#### Action 23: Delete Scenario

- **What it does:** Removes a saved scenario from the Scenarios sheet.
- **When to use it:** When a scenario is no longer relevant and you want to clean up.
- **What to expect:** A dialog box will show you a list of saved scenarios. Select one to delete and confirm.

---

### Allocation (Actions 24–25)

#### Action 24: Run Allocation Engine

- **What it does:** Runs the full cost allocation process — takes shared costs and distributes them across product lines and departments based on the allocation method defined in the Assumptions sheet (revenue share, headcount share, or equal split).
- **When to use it:** After updating allocation percentages in Assumptions, or during month-end close when you need to recalculate allocations.
- **What to expect:** An "Allocation Output" sheet will be created showing the detailed allocation results — how much of each shared cost was assigned to each product line and department, and the method used.

#### Action 25: Allocation Scenario Preview

- **What it does:** Shows you a preview of what the allocation would look like with the current assumptions, WITHOUT actually changing any data. This is a "what-if" preview.
- **When to use it:** Before running the full allocation (Action 24), to make sure the numbers look right before committing.
- **What to expect:** A preview window showing the proposed allocation results. You can review and decide whether to proceed or adjust assumptions first.

---

### Consolidation (Actions 26–30)

#### Action 26: Consolidation Menu

- **What it does:** Opens a sub-menu specifically for multi-entity consolidation tasks. This is the entry point for organizations that need to combine P&L data from multiple entities into one consolidated view.
- **When to use it:** When you are working with multiple entity files and need to consolidate them.
- **What to expect:** A menu window with consolidation options (Add Entity, Generate Consolidated P&L, View Loaded Entities, Add Elimination Entry).

#### Action 27: Add Entity File

- **What it does:** Loads a P&L file from another entity into the consolidation engine. You can add multiple entity files.
- **When to use it:** When you have P&L data from multiple entities (e.g., different business units, subsidiaries) that need to be combined.
- **What to expect:** A file picker dialog will appear. Select the entity's P&L file. The system will validate the format and load the data.

#### Action 28: Generate Consolidated P&L

- **What it does:** Combines all loaded entity files into a single consolidated P&L statement, with proper elimination of intercompany transactions.
- **When to use it:** After loading all entity files (Action 27) and adding any elimination entries (Action 30).
- **What to expect:** A consolidated P&L sheet will be created showing the combined results from all entities.

#### Action 29: View Loaded Entities

- **What it does:** Shows a list of all entity files that have been loaded into the consolidation engine.
- **When to use it:** To verify which entities are loaded before generating the consolidated P&L.
- **What to expect:** A message box listing all loaded entities with their file paths and load dates.

#### Action 30: Add Elimination Entry

- **What it does:** Adds an intercompany elimination entry to the consolidation. This removes internal transactions between entities (e.g., Entity A sold to Entity B — that revenue and cost should not appear in the consolidated view).
- **When to use it:** When you have intercompany transactions that need to be eliminated during consolidation.
- **What to expect:** A dialog box where you enter the elimination details (entity, account, amount, description).

---

### Version Control (Actions 31–35)

#### Action 31: Version Control Menu

- **What it does:** Opens a sub-menu for version control tasks. This lets you save snapshots of the workbook at different points in time and compare or restore them later.
- **When to use it:** When you want to manage versions of the P&L model.
- **What to expect:** A menu window with version control options (Save Version, Compare Versions, Restore Version, List Versions).

#### Action 32: Save Version

- **What it does:** Saves a timestamped snapshot of all key sheet values to a hidden "Version History" sheet. Each version gets a description you provide.
- **When to use it:** Before making significant changes, at month-end close, or any time you want a restore point.
- **What to expect:** A dialog box will ask for a version description (e.g., "Pre-close March 2025"). The snapshot will be saved with a timestamp.
- **Tip:** Think of this like "Save As" but smarter — you can compare versions side by side and restore any previous version.

#### Action 33: Compare Versions

- **What it does:** Creates a side-by-side comparison of two saved versions, highlighting what changed between them (which values increased, decreased, or were added/removed).
- **When to use it:** When you need to understand what changed between two points in time (e.g., before and after month-end adjustments).
- **What to expect:** A comparison sheet showing differences between the two selected versions, with changes highlighted.

#### Action 34: Restore Version

- **What it does:** Restores the workbook values to a previously saved version. This overwrites current values with the saved snapshot.
- **When to use it:** When you need to roll back to a previous state — for example, if adjustments were made incorrectly and you want to start over from a known good point.
- **What to expect:** A list of saved versions will appear. Select the one to restore. The workbook will be updated with the saved values.
- **Warning:** This overwrites current values. Save a new version first (Action 32) if you want to preserve the current state.

#### Action 35: List Versions

- **What it does:** Shows a list of all saved versions with their timestamps and descriptions.
- **When to use it:** To see what versions are available before comparing or restoring.
- **What to expect:** A message box or sheet listing all saved versions chronologically.

---

### Governance (Actions 36–40)

#### Action 36: Auto-Documentation

- **What it does:** Generates a "Tech Documentation" sheet that automatically documents everything in the workbook — all sheet names, all named ranges, all macros, data sources, and key formulas. This is a living document that updates itself.
- **When to use it:** When auditors or IT ask "What does this workbook contain?" or when onboarding a new team member who needs to understand the file.
- **What to expect:** A "Tech Documentation" sheet will be created with a comprehensive inventory of the workbook's contents.

#### Action 37: Change Management Menu

- **What it does:** Opens a sub-menu for tracking changes to the P&L model. This is a formal change request system — like a mini JIRA inside Excel.
- **When to use it:** When your team needs to track who requested what changes and their status.
- **What to expect:** A menu window with options to add, update, and view change requests.

#### Action 38: Add Change Request

- **What it does:** Logs a new change request to the "Change Management Log" sheet. You provide a description, priority, and requester name.
- **When to use it:** When someone requests a change to the P&L model (e.g., "Add a new cost center" or "Change the allocation method for AWS").
- **What to expect:** A dialog box where you fill in the change request details. The request will be added to the Change Management Log with a timestamp and unique ID.

#### Action 39: Update CR Status

- **What it does:** Updates the status of an existing change request (e.g., from "Open" to "In Progress" to "Completed").
- **When to use it:** As work progresses on a change request.
- **What to expect:** A dialog box showing open change requests. Select one and update its status.

#### Action 40: CR Summary Report

- **What it does:** Generates a summary report of all change requests — how many are open, in progress, and completed.
- **When to use it:** During team meetings or when leadership asks about the status of changes.
- **What to expect:** A formatted summary showing change request statistics and details.

---

### Admin & Testing (Actions 41–45)

#### Action 41: View Audit Log

- **What it does:** Opens the hidden "VBA_AuditLog" sheet and makes it visible. This sheet automatically records every action that is run through the Command Center — what was run, when, and whether it succeeded.
- **When to use it:** When you want to see a history of what actions have been run (useful for troubleshooting or audit purposes).
- **What to expect:** The VBA_AuditLog sheet will become visible and active. It contains a table with columns: Timestamp, Action, Message, Status.

#### Action 42: Export Audit Log

- **What it does:** Exports the audit log to a separate file for permanent record-keeping.
- **When to use it:** At the end of each month or quarter, or when auditors request an activity log.
- **What to expect:** A file save dialog will appear. Choose where to save the exported audit log.

#### Action 43: Clear Audit Log

- **What it does:** Clears all entries from the audit log, giving you a fresh start.
- **When to use it:** After exporting the log (Action 42), or at the start of a new fiscal year.
- **What to expect:** A confirmation dialog will appear. If you confirm, the audit log will be cleared. The log starts fresh from the next action you run.
- **Warning:** This cannot be undone. Export the log first (Action 42) if you need to keep the history.

#### Action 44: Full Integration Test

- **What it does:** Runs an 18-test automated test suite that checks every critical function in the workbook — from sheet existence to formula integrity to data quality. This is the ultimate health check.
- **When to use it:** After importing new modules, after making significant changes, or periodically to verify everything is working correctly.
- **What to expect:** An "Integration Test Report" sheet will be created showing PASS or FAIL for each of the 18 tests, along with details for any failures. All tests should show PASS if the workbook is healthy.

#### Action 45: Quick Health Check

- **What it does:** A faster, lighter version of Action 44. Runs the 5 most critical tests instead of all 18. Takes about 5 seconds instead of 30+.
- **When to use it:** When you want a quick sanity check but don't have time for the full test suite.
- **What to expect:** A message box showing the results of 5 key tests (sheet existence, data integrity, reconciliation balance, formula errors, audit log health).

---

### Advanced (Actions 46–50)

#### Action 46: Variance Commentary

- **What it does:** Automatically generates written commentary for each significant variance found in the P&L. Instead of you writing "Revenue decreased 18% due to seasonal factors", the system generates this commentary for you based on the data patterns.
- **When to use it:** When preparing management reports or month-end commentary. This gives you a first draft that you can review and customize.
- **What to expect:** A "Variance Commentary" sheet will be created with auto-generated explanations for each line item with a significant variance. Review and edit the commentary as needed — the system provides the starting point, you provide the context.

#### Action 47: Cross-Sheet Validation

- **What it does:** Performs a detailed cross-sheet validation checking that numbers on one sheet match related numbers on other sheets. For example, does the revenue total on the P&L Monthly Trend match the revenue total on the Product Line Summary? Does the GL detail sum match the report totals?
- **When to use it:** During month-end close, or any time you suspect a sheet may be out of sync.
- **What to expect:** A validation report showing PASS or FAIL for each cross-sheet check, with details on any mismatches found.

#### Action 48: Executive Mode Toggle

- **What it does:** Toggles "Executive Mode" on or off. When Executive Mode is ON, only the key reporting sheets are visible — all technical and working sheets are hidden. This gives a clean, presentation-ready view. When Executive Mode is OFF, all sheets are visible again.
- **When to use it:** Before presenting to leadership or sharing your screen. Turn it ON for presentations, turn it OFF to get back to work.
- **What to expect:** Sheets will hide or unhide depending on the mode. A message box will confirm which mode you are now in.
- **Keyboard shortcut:** **Ctrl + Shift + R**

#### Action 49: Force Recalculate All

- **What it does:** Forces Excel to recalculate every formula in the entire workbook. This is useful when Excel's automatic calculation does not seem to be updating properly.
- **When to use it:** If numbers look stale or formulas are not updating after data changes.
- **What to expect:** All formulas will recalculate. This may take a few seconds on large workbooks. A message box will confirm completion.

#### Action 50: About This Toolkit

- **What it does:** Shows information about the P&L Automation Toolkit — version number, build date, total number of actions, and module count.
- **When to use it:** When someone asks "What version is this?" or when you want to verify the toolkit is fully installed.
- **What to expect:** A message box with toolkit information.

---

### Sheet Tools (Actions 51–62)

#### Action 51: Delete All Blank Rows

- **What it does:** Finds and deletes every completely blank row in the active sheet. A row must be entirely empty (no values, no formulas, no formatting) to be deleted.
- **When to use it:** When your data has gaps from deleted rows or imported data with empty rows scattered throughout.
- **What to expect:** A message box confirming how many blank rows were deleted from the active sheet.

#### Action 52: Unhide All Worksheets

- **What it does:** Makes every hidden sheet in the workbook visible. This includes both "hidden" and "very hidden" sheets.
- **When to use it:** When you need to see or edit sheets that are normally hidden (like VBA_AuditLog, Scenarios, Version History).
- **What to expect:** All sheets will become visible. Sheet tabs that were previously hidden will now appear at the bottom of Excel.

#### Action 53: Sort Sheets Alphabetically

- **What it does:** Rearranges all sheet tabs in alphabetical order (A to Z).
- **When to use it:** When your workbook has many sheets and you want them organized for easy navigation.
- **What to expect:** The sheet tabs at the bottom of Excel will be rearranged alphabetically. No data is changed — only the tab order.

#### Action 54: Toggle Freeze Panes

- **What it does:** Toggles freeze panes on the active sheet. If panes are frozen, it unfreezes them. If panes are not frozen, it freezes at the current cell position.
- **When to use it:** When you want to keep headers visible while scrolling through large data sets.
- **What to expect:** The active sheet's panes will toggle between frozen and unfrozen.

#### Action 55: Convert Formulas to Values

- **What it does:** Converts all formulas in the selected range (or entire active sheet) to their calculated values. The formulas are replaced with the numbers they produce.
- **When to use it:** When you want to "lock in" calculated values — for example, before sharing a file where the recipient should not see your formulas, or to speed up a slow workbook.
- **What to expect:** All formulas in the range will be replaced with their current values. A message box will confirm how many cells were converted.
- **Warning:** This cannot be undone (Ctrl+Z may not work for large operations). Save a version first (Action 32) if you might need the formulas back.

#### Action 56: AutoFit All Columns

- **What it does:** Automatically adjusts the width of every column in the active sheet to fit the content. No more truncated text or columns that are too wide.
- **When to use it:** After importing data, after pasting content, or any time columns are not sized correctly.
- **What to expect:** All columns in the active sheet will resize to fit their content.

#### Action 57: Protect All Sheets

- **What it does:** Applies sheet protection to every sheet in the workbook. This prevents accidental edits to formulas and structure.
- **When to use it:** Before sharing the workbook with other users, or before a presentation.
- **What to expect:** All sheets will be protected. Users will not be able to edit cells unless they are explicitly unlocked. A message box will confirm.

#### Action 58: Unprotect All Sheets

- **What it does:** Removes sheet protection from every sheet in the workbook.
- **When to use it:** When you need to make edits to a protected workbook.
- **What to expect:** All sheets will be unprotected. Editing will be enabled on all sheets.

#### Action 59: Find & Replace (All Sheets)

- **What it does:** Performs a find-and-replace operation across EVERY sheet in the workbook simultaneously. Excel's built-in Find & Replace only works on one sheet — this does all of them at once.
- **When to use it:** When you need to update a company name, correct a label, or change a value that appears on multiple sheets.
- **What to expect:** A dialog box will ask for the Find text and Replace text. The operation will run on all sheets and report how many replacements were made on each sheet.

#### Action 60: Highlight Hardcoded Numbers

- **What it does:** Finds every cell in the active sheet that contains a hardcoded number (not a formula) and highlights it with a colored background. This helps you identify which cells are manual entries vs. calculated values.
- **When to use it:** During audit prep, or when you need to understand which numbers in a report are manually entered vs. formula-driven.
- **What to expect:** Cells with hardcoded numbers will be highlighted. A message box will report how many were found.

#### Action 61: Toggle Presentation Mode

- **What it does:** Switches the active sheet into a clean "presentation mode" — hides gridlines, row/column headers, the formula bar, and the ribbon. Toggle it again to restore the normal view.
- **When to use it:** Before presenting to leadership or sharing your screen. This makes Excel look like a clean dashboard rather than a spreadsheet.
- **What to expect:** The Excel interface will change to a minimal view. Run it again to restore the normal view.

#### Action 62: Unmerge and Fill Down

- **What it does:** Finds all merged cells in the active sheet, unmerges them, and fills the value down into the previously blank cells. This is essential for making data usable in PivotTables and VLOOKUP.
- **When to use it:** When you receive data from another department that uses merged cells, and you need to work with it properly.
- **What to expect:** All merged cells will be unmerged and the values will be filled into every row. A message box will confirm how many merged regions were processed.

---

## 6. Category-by-Category Walkthrough

### Which Category Should I Start With?

Here is a suggested order for first-time users:

| Step | Category | Actions to Try First | Why |
|---|---|---|---|
| 1 | **Data Quality** | Action 7 (Scan) | Always check your data first |
| 2 | **Monthly Operations** | Action 3 (Reconciliation) | Make sure everything ties out |
| 3 | **Analysis** | Action 6 (Variance) | See what changed this month |
| 4 | **Reporting** | Action 12 (Dashboard) then Action 10 (PDF) | Build visuals and export |
| 5 | **Sheet Tools** | Action 56 (AutoFit) | Clean up the look |
| 6 | **Advanced** | Action 48 (Executive Mode) | Get presentation-ready |

### Monthly Close Workflow (Recommended Order)

If you are running the month-end close process, here is the recommended sequence:

1. **Action 17** — Import GL Data Pipeline (bring in the latest GL data)
2. **Action 7** — Scan for Data Quality Issues (check the data is clean)
3. **Action 8** — Fix Text-Stored Numbers (fix any issues found)
4. **Action 19** — Append Month to Trend (set up the new month column)
5. **Action 3** — Run Reconciliation Checks (verify everything ties)
6. **Action 6** — Run Variance Analysis (understand what changed)
7. **Action 46** — Variance Commentary (generate draft commentary)
8. **Action 12** — Build Dashboard Charts (create visual summary)
9. **Action 32** — Save Version (save a snapshot of the close)
10. **Action 10** — Export Report Package (PDF) (create the deliverable)

---

## 7. Keyboard Shortcuts

You do not need to open the Command Center for these — they work any time the workbook is open.

| Shortcut | What It Does | Equivalent Action |
|---|---|---|
| **Ctrl + Shift + M** | Open the Command Center | — |
| **Ctrl + Shift + H** | Go to the Home / Report sheet | Action 16 |
| **Ctrl + Shift + J** | Quick Jump to any sheet | Action 15 |
| **Ctrl + Shift + R** | Toggle Executive Mode | Action 48 |

### How to Use Keyboard Shortcuts

1. Make sure the P&L workbook is the active window (click on it)
2. Hold down the **Ctrl** key
3. While holding Ctrl, also hold down the **Shift** key
4. While holding both Ctrl and Shift, press the letter key (**M**, **H**, **J**, or **R**)
5. Release all keys

---

## 8. Tips for Daily Use

### Tip 1: Start Every Day with Action 7

Run **Action 7 (Scan for Data Quality Issues)** first thing. It takes 10 seconds and catches problems before they snowball. If the letter grade is A, you are good to go. If it is B or lower, fix the issues before doing anything else.

### Tip 2: Save Versions Before Big Changes

Before running allocations, importing data, or making significant updates, run **Action 32 (Save Version)** first. This gives you a restore point if anything goes wrong. It takes 2 seconds and can save you hours of rework.

### Tip 3: Use Search Instead of Browsing

If you know what you want to do but don't remember the action number, just type a keyword in the Command Center search box. It is much faster than scrolling through categories.

### Tip 4: Run & Close for Single Actions

If you only need to run one action, use the **"Run & Close"** button instead of "Run Selected". This closes the Command Center automatically so you can see the results right away.

### Tip 5: Use Executive Mode for Presentations

Before sharing your screen or presenting, run **Action 48 (Executive Mode Toggle)** or press **Ctrl + Shift + R**. This hides all technical sheets and gives you a clean, professional view.

### Tip 6: Check the Audit Log Periodically

Run **Action 41 (View Audit Log)** once a week to review what actions have been run. This is useful for tracking your team's usage and for audit purposes.

### Tip 7: Use the PDF Export for Monthly Packages

Instead of manually formatting and printing each sheet, use **Action 10 (Export Report Package)** to generate a polished 7-sheet PDF in one click. The formatting is already set up with proper headers, footers, and page numbers.

---

## 9. What to Do If Something Goes Wrong

### An Action Gave an Error Message

1. **Read the error message carefully.** Most error messages tell you exactly what is wrong (e.g., "Sheet 'Assumptions' not found" means the Assumptions sheet was renamed or deleted).
2. **Close the error message** and check that all expected sheets exist.
3. **Run Action 45 (Quick Health Check)** to see if the workbook is healthy.
4. **Try the action again.** Some errors are one-time issues (e.g., Excel hadn't finished calculating).
5. **If the error persists**, note the exact error message and contact the Finance Automation Team.

### An Action Ran But the Results Look Wrong

1. **Run Action 7 (Scan for Data Quality Issues)** to check for data problems.
2. **Run Action 3 (Reconciliation Checks)** to verify cross-sheet consistency.
3. **Check the Assumptions sheet** to make sure the input values are correct.
4. **Run Action 49 (Force Recalculate All)** to make sure all formulas are up to date.
5. **If results still look wrong**, run **Action 44 (Full Integration Test)** for a comprehensive check.

### The Command Center Won't Open

1. **Check that macros are enabled.** Look for a yellow "Security Warning" bar at the top of Excel. Click "Enable Content" if you see it.
2. **Try the keyboard shortcut: Ctrl + Shift + M.**
3. **Try Alt + F8**, find "LaunchCommandCenter" in the list, and click Run.
4. **If none of these work**, the VBA modules may not be imported. Contact the Finance Automation Team for help with setup.

### Excel Froze or is Running Slowly

1. **Wait 30 seconds.** Some actions process large amounts of data and may take time. Look at the status bar in the bottom-left corner of Excel — if it says "Processing..." or shows a percentage, the action is still running.
2. **Do not press Ctrl+Break or click anything.** This can cause the action to stop mid-way, leaving the workbook in an incomplete state.
3. **If Excel is still frozen after 2 minutes**, press **Ctrl + Break** to stop the macro, then run **Action 45 (Quick Health Check)** to verify the workbook is okay.

### I Accidentally Changed Something

1. **Press Ctrl + Z immediately** to undo the last change.
2. **If that doesn't work**, check if you saved a version (Action 32) before the change. If so, use **Action 34 (Restore Version)** to go back.
3. **If no version was saved**, close the workbook WITHOUT saving and reopen it. You will lose all changes since the last save, but the workbook will be in its previous state.

---

## 10. Frequently Asked Questions

### Q: Do I need to know VBA or coding to use the Command Center?

**A:** No. The Command Center is designed for Finance and Accounting professionals who work in Excel every day. You just pick an action and click Run. All the code runs behind the scenes.

### Q: Can I break the workbook by running an action?

**A:** The actions are designed to be safe. Most actions create new sheets or reports rather than changing existing data. The few actions that do change data (like Fix Text-Stored Numbers or Restore Version) always tell you what they are about to do and ask for confirmation first. That said, it is always a good practice to save a version (Action 32) before running actions that modify data.

### Q: How long do actions take to run?

**A:** Most actions complete in under 5 seconds. A few actions that process the entire workbook (like Full Integration Test or Export Report Package) may take 15–30 seconds. You will see a progress indicator in the status bar.

### Q: Can I run multiple actions at the same time?

**A:** No. Run one action at a time and wait for it to finish before running the next one. Actions that run simultaneously can conflict with each other.

### Q: What if I run the wrong action?

**A:** Most actions create new sheets or reports, so running the wrong one won't damage anything — you can just delete the unwanted sheet. For actions that modify data, press Ctrl+Z to undo or restore from a saved version.

### Q: How often should I run the Reconciliation Checks?

**A:** At minimum, run them during every month-end close. Ideally, run them after every significant data change (importing new data, updating assumptions, running allocations). They take 10 seconds and can catch errors before they cascade.

### Q: Who do I contact if I have questions?

**A:** Contact the Finance Automation Team. Include the following information:
- What action you were running (number and name)
- The exact error message (screenshot if possible)
- What you were doing before the error

### Q: Does the Command Center work on Mac?

**A:** The Command Center is designed for **Windows Excel Desktop** (Excel 2019 or later, or Microsoft 365). Mac Excel has limitations with VBA UserForms and some features may not work correctly. It does NOT work in Excel Online (browser version).

### Q: What version of the toolkit am I using?

**A:** Run **Action 50 (About This Toolkit)** to see the version number, build date, and module count.

---

## Document Information

| Field | Value |
|---|---|
| **Document Title** | How to Use the Command Center |
| **Version** | 1.0 |
| **Last Updated** | March 5, 2026 |
| **Author** | Finance Automation Team |
| **Audience** | All iPipeline Finance & Accounting Employees |
| **Toolkit Version** | 2.1.0 |
| **Total Actions** | 62 |
| **Total Categories** | 15 + All Actions |

---

*This document is part of the iPipeline P&L Automation Toolkit documentation suite. For setup instructions, see "Getting Started — First Time Setup Guide." For a one-page summary of all actions, see "Quick Reference Card."*
