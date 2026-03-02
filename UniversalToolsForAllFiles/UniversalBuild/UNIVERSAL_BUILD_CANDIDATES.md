# Universal Tools — Build Candidates
**Folder:** `UniversalToolsForAllFiles/UniversalBuild/`
**Status:** AWAITING CONNOR APPROVAL — no code built yet
**Date Created:** 2026-03-02
**Source Files Reviewed:** GrokALL.md, PrelexALL.md (300+ tools reviewed and filtered)

---

## What These Tools Are

These are VBA macros and Python scripts that work on **any** Excel file — not just the P&L demo.
Coworkers could run these on any file they have. They are completely separate from the 32 modules
already built into the main P&L demo workbook.

**VBA tools** → Will be packaged into a Personal Macro Workbook add-in (`KBT_UniversalTools.xlam`)
so coworkers install it once and have these tools available in every Excel session.

**Python scripts** → Standalone scripts coworkers run from command line by passing any file path.

---

## What Was Excluded (and Why)

- Tools already built into the main P&L demo (FindExternalLinks, AuditHiddenSheets,
  FindNegativeAmounts, Cross-Sheet Search, Data Sanitizer — those are in the 32 modules)
- Tools that duplicate OneDrive/SharePoint built-in features (co-authoring, AutoSave, version history)
- Too technical for Finance staff (OCR, machine learning, NLP/LLM, web scraping, regression analysis)
- Low practical value for Finance & Accounting (easter eggs, undo loggers, color standardizers)
- Profiles.md — not a tools file; personal account info only (see security note at bottom)

---

## TIER 1 — Build First
*Highest value. Solve the most common Finance & Accounting pain points. Build these first.*

### Tier 1 VBA Macros (23 tools)

| # | Tool Name | What It Does |
|---|-----------|--------------|
| 1 | Unmerge Cells & Fill Down | Unmerges all merged cells in selection and fills the top value downward. Every Finance export has merged cells — this fixes them instantly. |
| 2 | Fill Blanks Down | Fills every blank cell with the value from the cell above. Fixes stacked layouts from GL exports and pivot tables. |
| 3 | Convert Text to Numbers | Fixes cells storing numbers as text so they actually sum and calculate. One of the most common Finance pain points from data exports. |
| 4 | Remove Leading/Trailing Spaces | Trims all text in selection. Kills lookup failures caused by invisible spaces. |
| 5 | Delete Blank Rows | Removes completely empty rows from selection or used range. Instant cleanup on any raw data export. |
| 6 | Replace Error Values | Replaces all #N/A, #REF!, #VALUE!, #DIV/0! errors with blank or a specified value. Cleans up broken formulas before sharing a file. |
| 7 | Highlight Duplicate Rows | Colors duplicate rows for manual review — does NOT delete. Safe version for invoice and vendor checks. |
| 8 | Remove Duplicate Rows | Deletes confirmed duplicate rows from the active sheet. Use after Highlight version confirms what to remove. |
| 9 | AutoFit All Columns & Rows | Auto-fits every column and row in the entire workbook. One click to make any file readable. |
| 10 | Freeze Top Row on All Sheets | Applies Freeze Panes to row 1 on every worksheet. Finance standard — every file needs this. |
| 11 | Number Format Standardizer | Applies #,##0.00 formatting to all numeric cells across all sheets. Forces consistency. |
| 12 | Currency Format Standardizer | Detects and applies $#,##0.00 to all currency values across all sheets. |
| 13 | Date Format Standardizer | Normalizes all dates to a single format (MM/DD/YYYY) across all sheets. Kills the text-date problem from imports. |
| 14 | Conditional Formatting — Negatives in Red | Highlights all negative numbers in red across the workbook. Finance standard for instant visibility on losses. |
| 15 | Find & Replace Across All Sheets | Performs find-and-replace across every sheet simultaneously. Massive time saver vs doing sheet by sheet. |
| 16 | Unhide All Sheets | Makes every hidden and very-hidden worksheet visible instantly. |
| 17 | Export All Sheets as One Combined PDF | Combines all visible sheets into a single multi-page PDF. Send one clean PDF instead of multiple files. |
| 18 | Save Timestamped Backup Copy | Saves a copy of the workbook with date/time in the filename before making changes. Safe versioning. |
| 19 | External Link Finder | Lists all cells that reference external workbooks with file paths. Critical audit — see what files a workbook depends on. |
| 20 | Circular Reference Detector | Scans the workbook for circular references and reports their exact locations. |
| 21 | Workbook Error Scanner | Lists every cell containing an error value with its sheet name and cell address. One scan, full report. |
| 22 | Search Across All Sheets | Searches for any value across every sheet and returns sheet name + cell address for every match. |
| 23 | Duplicate Invoice Detector | Scans for potential duplicate invoices by matching vendor, amount, date, and invoice number. Flags potentials for review. |

### Tier 1 Python Scripts (5 tools)

| # | Tool Name | What It Does |
|---|-----------|--------------|
| 1 | Universal Data Cleaner | One command: removes duplicates, drops empty rows/columns, trims spaces, converts text-to-numbers, standardizes dates on any file. Run: `python clean_file.py "path\to\file.xlsx"` |
| 2 | Compare Two Excel Files | Cell-level diff report between two workbooks — highlights every difference. Perfect for version comparison or reviewing what changed. |
| 3 | Budget vs. Actual Consolidator | Merges 50+ department budget files from a folder into one master file with dollar and percent variance columns. |
| 4 | AR/AP Aging Report | Auto-buckets invoices by days overdue (0-30, 31-60, 61-90, 90+) from any date column in any file. |
| 5 | Multi-File Data Consolidator | Combines data from hundreds of Excel files in a folder into one master sheet with a "Source_File" column added. |

---

## TIER 2 — Build Later
*Solid value but lower urgency. Build after Tier 1 is complete and tested.*

### Tier 2 VBA Macros

| Tool Name | What It Does |
|-----------|--------------|
| Create Table of Contents Sheet | Generates a clickable index sheet with hyperlinks to every worksheet. |
| Sort Worksheets Alphabetically | Reorders all sheet tabs A–Z. |
| Data Quality Scorecard | Generates a summary: total rows, blank cells, error cells, duplicate rows, data types per column. |
| Formula Auditor — Inconsistent Formulas | Flags cells in a column where the formula differs from the majority pattern. |
| Protect All Sheets | Applies password protection to every sheet at once. |
| Unprotect All Sheets | Removes worksheet protection from every sheet. |
| Lock All Formula Cells | Locks cells with formulas while leaving input cells editable. |
| Export Active Sheet as PDF | Saves the active worksheet as a PDF file. |
| Export All Sheets as Individual PDFs | Saves each worksheet as its own separate PDF. |
| Aging Bucket Calculator | Assigns records to 0-30, 31-60, 61-90, 90+ buckets based on any date column. |
| Variance Analysis Template | Adds dollar and percent variance columns next to Actual and Budget columns. |
| Trial Balance Checker | Verifies total debits equal total credits and highlights any imbalance. |
| Journal Entry Validator | Checks that every journal entry has balanced debits and credits. |
| Flux Analysis / Period Comparison | Flags line items with significant period-over-period changes above a threshold. |
| AP Aging Summary Generator | Generates an accounts payable aging report with standard buckets. |
| AR Aging Summary Generator | Generates an accounts receivable aging report with standard buckets. |
| Ratio Analysis Dashboard Builder | Calculates Current Ratio, ROE, Gross Margin, and other key ratios from statement data. |
| Build Distribution-Ready Copy | Creates a clean copy with formulas as values, metadata stripped, formatting standardized. |
| Reset All Filters | Clears all AutoFilter and Advanced Filter criteria on every sheet. |
| Workbook Health Check | Comprehensive diagnostic: file size, formula count, error count, external links, pivot tables. |

### Tier 2 Python Scripts

| Tool Name | What It Does |
|-----------|--------------|
| Variance Analysis Generator | Compares Actual vs Budget columns across multiple files and creates summary + waterfall chart. |
| GL Reconciliation Engine | Matches two large transaction lists and flags unmatched items. |
| Fuzzy Match / Fuzzy Lookup | Matches records between two datasets using fuzzy string matching — catches vendor/customer name typos. |
| Batch Process Folder of Files | Runs the same cleaning or transformation script on every Excel file in a folder. |
| Forecast Roll-Forward | Takes last month actuals + assumptions and builds the next 12-month forecast. |

---

## Counts

| Category | Count |
|----------|-------|
| Tier 1 VBA Macros | 23 |
| Tier 1 Python Scripts | 5 |
| Tier 2 VBA Macros | 20 |
| Tier 2 Python Scripts | 5 |
| **Total Candidates** | **53** |

---

## Next Steps (After Connor Approves This List)

1. Build all 23 Tier 1 VBA tools as `.bas` modules stored in `UniversalToolsForAllFiles/vba/`
2. Package them into `KBT_UniversalTools.xlam` (Personal Macro Workbook add-in)
3. Build all 5 Tier 1 Python scripts stored in `UniversalToolsForAllFiles/python/`
4. Write coworker install guide for the add-in
5. Write coworker usage guide for the Python scripts
6. Then tackle Tier 2 as backlog

---

## Security Note — Profiles.md

`UniversalToolsForAllFiles/Profiles.md` contains personal account emails and tool credentials.
**This file should not be in the repo.** Connor should delete it or move it out of the repository.
It is not a tools file and is not referenced anywhere in the build plan.
