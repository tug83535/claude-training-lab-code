# iPipeline P&L Automation Toolkit — Quick Reference Card

## Command Center Cheat Sheet | Version 2.1.0

**Open the Command Center:** Press **Ctrl + Shift + M** at any time

---

## Keyboard Shortcuts

| Shortcut | Action |
|---|---|
| **Ctrl + Shift + M** | Open the Command Center |
| **Ctrl + Shift + H** | Go to Home / Report sheet |
| **Ctrl + Shift + J** | Quick Jump to any sheet |
| **Ctrl + Shift + R** | Toggle Executive Mode (hide/show technical sheets) |

---

## All 62 Actions at a Glance

### Monthly Operations

| # | Action | What It Does |
|---|---|---|
| 1 | Generate Monthly Tabs | Creates individual monthly P&L sheets for each month of the fiscal year |
| 2 | Delete Generated Tabs | Removes all generated monthly tabs (keeps original Jan/Feb/Mar) |
| 3 | Run Reconciliation Checks | Validates all cross-sheet totals match — shows PASS/FAIL scorecard |
| 4 | Export Reconciliation Report | Exports the Checks sheet results as a formatted report |

### Analysis

| # | Action | What It Does |
|---|---|---|
| 5 | Run Sensitivity Analysis | Tests how changing assumptions affects the bottom line |
| 6 | Run Variance Analysis | Calculates month-over-month variance, flags items over 15% |

### Data Quality

| # | Action | What It Does |
|---|---|---|
| 7 | Scan for Data Quality Issues | 6-category data scan with A–F letter grade — run this first |
| 8 | Fix Text-Stored Numbers | Converts text-as-numbers to real numbers (safe — never touches dates/names) |
| 9 | Fix Duplicate Rows | Finds and removes duplicate rows in the data |

### Reporting

| # | Action | What It Does |
|---|---|---|
| 10 | Export Report Package (PDF) | Creates a polished 7-sheet PDF ready for leadership |
| 11 | Export Active Sheet (PDF) | Exports just the current sheet as a formatted PDF |
| 12 | Build Dashboard Charts | Creates Executive Dashboard with branded charts and visuals |

### Utilities

| # | Action | What It Does |
|---|---|---|
| 13 | Refresh Table of Contents | Updates the table of contents with clickable sheet links |
| 14 | Recalculate AWS Allocations | Validates and recalculates the AWS allocation table |
| 15 | Quick Jump to Sheet | Type a sheet name and jump directly to it |
| 16 | Go Home (Report-->) | Navigate to the main report sheet instantly |

### Data & Import

| # | Action | What It Does |
|---|---|---|
| 17 | Import GL Data Pipeline | Import GL data from CSV/Excel with format validation |

### Forecasting

| # | Action | What It Does |
|---|---|---|
| 18 | Rolling Forecast | Projects remaining months based on historical trends |
| 19 | Append Month to Trend | Adds next month's column to both trend sheets |

### Scenarios

| # | Action | What It Does |
|---|---|---|
| 20 | Save Current Scenario | Saves current Assumptions as a named scenario |
| 21 | Load Scenario | Loads a saved scenario into Assumptions (overwrites current) |
| 22 | Compare Scenarios | Side-by-side comparison of two or more scenarios |
| 23 | Delete Scenario | Removes a saved scenario |

### Allocation

| # | Action | What It Does |
|---|---|---|
| 24 | Run Allocation Engine | Distributes shared costs across products/departments |
| 25 | Allocation Scenario Preview | Preview allocation results without changing data |

### Consolidation

| # | Action | What It Does |
|---|---|---|
| 26 | Consolidation Menu | Opens the multi-entity consolidation sub-menu |
| 27 | Add Entity File | Load a P&L file from another entity |
| 28 | Generate Consolidated P&L | Combine all entities into one consolidated statement |
| 29 | View Loaded Entities | List all entity files currently loaded |
| 30 | Add Elimination Entry | Add an intercompany elimination entry |

### Version Control

| # | Action | What It Does |
|---|---|---|
| 31 | Version Control Menu | Opens the version control sub-menu |
| 32 | Save Version | Save a timestamped snapshot (your restore point) |
| 33 | Compare Versions | See what changed between two versions |
| 34 | Restore Version | Roll back to a previous version (overwrites current) |
| 35 | List Versions | View all saved versions with timestamps |

### Governance

| # | Action | What It Does |
|---|---|---|
| 36 | Auto-Documentation | Generates a tech doc of everything in the workbook |
| 37 | Change Management Menu | Opens the change request tracking sub-menu |
| 38 | Add Change Request | Log a new change request with description and priority |
| 39 | Update CR Status | Update status of an existing change request |
| 40 | CR Summary Report | View summary of all change requests |

### Admin & Testing

| # | Action | What It Does |
|---|---|---|
| 41 | View Audit Log | Shows the hidden audit log of all actions run |
| 42 | Export Audit Log | Exports audit log to a file for record-keeping |
| 43 | Clear Audit Log | Clears the audit log (export first!) |
| 44 | Full Integration Test | Runs 18 automated tests on the entire workbook |
| 45 | Quick Health Check | 5-point quick health test (fastest way to verify) |

### Advanced

| # | Action | What It Does |
|---|---|---|
| 46 | Variance Commentary | Auto-generates written explanations for each variance |
| 47 | Cross-Sheet Validation | Validates numbers match between related sheets |
| 48 | Executive Mode Toggle | Hides technical sheets for clean presentations |
| 49 | Force Recalculate All | Forces all formulas to recalculate |
| 50 | About This Toolkit | Shows version, build date, and module count |

### Sheet Tools

| # | Action | What It Does |
|---|---|---|
| 51 | Delete All Blank Rows | Removes completely empty rows from the active sheet |
| 52 | Unhide All Worksheets | Makes every hidden sheet visible |
| 53 | Sort Sheets Alphabetically | Rearranges sheet tabs A to Z |
| 54 | Toggle Freeze Panes | Toggles freeze panes on/off at the current position |
| 55 | Convert Formulas to Values | Replaces formulas with their calculated values (irreversible!) |
| 56 | AutoFit All Columns | Auto-sizes all columns to fit their content |
| 57 | Protect All Sheets | Applies protection to every sheet |
| 58 | Unprotect All Sheets | Removes protection from every sheet |
| 59 | Find & Replace (All Sheets) | Find and replace across every sheet at once |
| 60 | Highlight Hardcoded Numbers | Highlights cells with manual values (not formulas) |
| 61 | Toggle Presentation Mode | Clean view: hides gridlines, headers, ribbon |
| 62 | Unmerge and Fill Down | Unmerges cells and fills values down |

---

## Month-End Close — Recommended Workflow

Run these actions in this order during every monthly close:

| Step | Action # | Action | Time |
|---|---|---|---|
| 1 | 17 | Import GL Data Pipeline | 2 min |
| 2 | 7 | Scan for Data Quality Issues | 10 sec |
| 3 | 8 | Fix Text-Stored Numbers (if needed) | 10 sec |
| 4 | 19 | Append Month to Trend | 5 sec |
| 5 | 3 | Run Reconciliation Checks | 10 sec |
| 6 | 6 | Run Variance Analysis | 10 sec |
| 7 | 46 | Variance Commentary | 15 sec |
| 8 | 12 | Build Dashboard Charts | 15 sec |
| 9 | 32 | Save Version | 5 sec |
| 10 | 10 | Export Report Package (PDF) | 15 sec |
| | | **Total** | **~5 min** |

---

## Top 10 Most-Used Actions

If you only remember 10 actions, remember these:

| # | Action | When to Use |
|---|---|---|
| **7** | Scan for Data Quality Issues | First thing — check data health |
| **3** | Run Reconciliation Checks | Verify everything ties |
| **6** | Run Variance Analysis | Understand what changed |
| **12** | Build Dashboard Charts | Visual summary |
| **10** | Export Report Package (PDF) | Create the deliverable |
| **32** | Save Version | Before any major changes |
| **48** | Executive Mode Toggle | Clean up for presentations |
| **46** | Variance Commentary | Auto-generate explanations |
| **45** | Quick Health Check | Fast sanity check |
| **8** | Fix Text-Stored Numbers | Fix the most common data issue |

---

## Workbook Sheets

### Always Visible

| Sheet | Purpose |
|---|---|
| Home | Landing page with Command Center button |
| Report--> | Main P&L report view |
| P&L - Monthly Trend | Revenue/expense by month |
| Functional P&L - Monthly Trend | By department, by month |
| Product Line Summary | By product line |
| Assumptions | Key drivers (the only sheet you edit manually) |
| Data Dictionary | Definitions reference |
| AWS Allocation | AWS cost allocation |
| Checks | Reconciliation PASS/FAIL scorecard |

### Usually Hidden (Visible When Needed)

| Sheet | Purpose | Created/Shown By |
|---|---|---|
| CrossfireHiddenWorksheet | Raw GL transaction data | Hidden by default |
| VBA_AuditLog | Action audit trail | Action 41 |
| Scenarios | Saved scenario data | Actions 20–23 |
| Version History | Saved version snapshots | Actions 31–35 |

### Created on Demand

| Sheet | Created By |
|---|---|
| Data Quality Report | Action 7 |
| Variance Analysis | Action 6 |
| Sensitivity Analysis | Action 5 |
| Executive Dashboard | Action 12 |
| Variance Commentary | Action 46 |
| Rolling Forecast | Action 18 |
| Allocation Output | Action 24 |
| Integration Test Report | Action 44 |
| Tech Documentation | Action 36 |
| Change Management Log | Actions 37–40 |
| Search Results | Action 15 (search) |
| Validation Report | Action 47 |

---

## Quick Troubleshooting

| Problem | Fix |
|---|---|
| Command Center won't open | Click "Enable Content" in yellow bar, then Ctrl+Shift+M |
| "Macros disabled" with no enable button | File > Options > Trust Center > Trust Center Settings > Macro Settings > select "with notification" |
| Action gives an error | Click "End", try again. If it persists, run Action 45 (Health Check) |
| Numbers look wrong | Run Action 7 (Data Scan) then Action 3 (Reconciliation) |
| Sheets are missing | Run Action 52 (Unhide All Worksheets) |
| Need to undo a change | Press Ctrl+Z or use Action 34 (Restore Version) |
| File is slow | Move to local drive, close other Excel files |

---

## Need Help?

- **Run Action 45** — Quick Health Check (5 tests, 5 seconds)
- **Run Action 44** — Full Integration Test (18 tests, 30 seconds)
- **Run Action 50** — About This Toolkit (version info)
- **Contact** the Finance Automation Team with screenshots of any errors

---

**iPipeline P&L Automation Toolkit v2.1.0** | **62 Actions** | **34 VBA Modules** | **14 Python Scripts**

*Print this card and keep it at your desk for quick reference.*
