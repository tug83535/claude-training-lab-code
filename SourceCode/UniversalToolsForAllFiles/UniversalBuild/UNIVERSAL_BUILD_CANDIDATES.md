# Universal Tools — Build Candidates
**Folder:** `UniversalToolsForAllFiles/UniversalBuild/`
**Status:** AWAITING CONNOR APPROVAL — no code built yet
**Date Created:** 2026-03-02
**Last Updated:** 2026-03-02 — added GemAll.md candidates
**Source Files Reviewed:** GrokALL.md, PrelexALL.md, GemAll.md (400+ tools reviewed and filtered)

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
- Too technical for Finance staff (machine learning, NLP/LLM, ARIMA forecasting, regression, web scraping, OCR)
- **Backup Workbook macro** — user declined permanently (2026-02-28)
- **Cell Change Audit Trail** — user declined permanently (2026-02-28)
- **Export Charts to PowerPoint** — user declined permanently (2026-02-28)
- **Email/Outlook integration** — user declined permanently (2026-02-28)
- Low practical value (clipboard clear, zoom standardizer, easter eggs, text-to-speech auditor)
- Profiles.md — not a tools file; personal account info only (already removed from repo)

---

## TIER 1 — Build First
*Highest value. Solve the most common Finance & Accounting pain points. Build these first.*

### Tier 1 VBA Macros (24 tools)

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
| 16 | Unhide All Sheets, Rows & Columns | Makes every hidden worksheet, row, and column visible instantly. |
| 17 | Export All Sheets as One Combined PDF | Combines all visible sheets into a single multi-page PDF. Send one clean PDF instead of multiple files. |
| 18 | External Link Finder | Lists all cells that reference external workbooks with file paths. Critical audit — see what files a workbook depends on. |
| 19 | Circular Reference Detector | Scans the workbook for circular references and reports their exact locations. |
| 20 | Workbook Error Scanner | Lists every cell containing an error value with its sheet name and cell address. One scan, full report. |
| 21 | Search Across All Sheets | Searches for any value across every sheet and returns sheet name + cell address for every match. |
| 22 | Duplicate Invoice Detector | Scans for potential duplicate invoices by matching vendor, amount, date, and invoice number. Flags potentials for review. |
| 23 | Auto-Balancing GL Validator | Sums debit and credit columns, flags any imbalance, and optionally adds a balancing plug line. Instant trial balance check. *(from GemAll)* |
| 24 | Multi-Replace Data Cleaner | Executes a batch Find & Replace based on a two-column mapping table you provide. Perfect for standardizing account names, cost center codes, or category labels. *(from GemAll)* |

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

### Tier 2 VBA Macros (34 tools)

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
| Phantom Hyperlink Purger | Scans and removes all embedded hyperlinks from the selected range or entire sheet. Cleans up files before distribution. *(from GemAll)* |
| Formula-to-Value Hardcoder | Converts all formulas in the selected range to static values. Prevents accidental recalculation when sharing files. *(from GemAll)* |
| Multi-Sheet Batch Renamer | Prompts for a text string and replaces it across all sheet tabs at once. Useful for renaming "2024" to "2025" across an entire workbook. *(from GemAll)* |
| Convert Numbers to Words | Translates numeric values in selected cells into written text (e.g., 1,250.00 → "One Thousand Two Hundred Fifty Dollars"). For formal invoices or check writing. *(from GemAll)* |
| Quick Corkscrew Builder | Automatically constructs a standard financial corkscrew (roll-forward) schedule for a selected balance sheet line item. *(from GemAll)* |
| Financial Number Formatting Suite | Applies standard financial formats — Normal Accounting, Factor (000s), and Percentage — via hotkeys. Instant Finance-standard formatting. *(from GemAll)* |
| General Ledger Journal Mapper | Transforms a raw trial balance into a properly formatted journal entry upload template for accounting systems. *(from GemAll)* |
| Financial Period Roll-Forward | Automates month-end reporting by updating all header dates and clearing prior-period input cells to prepare for new data entry. *(from GemAll)* |
| Multi-Currency Consolidation Aggregator | Consolidates multiple subsidiary sheets with different currencies into a master sheet using a budget FX rate table. *(from GemAll)* |
| Named Range Auditor | Generates a full report of all Named Ranges in the workbook and flags any with hidden or broken references. *(from GemAll)* |
| Conditional Format Purger | Lists all conditional formatting rules and allows batch deletion of redundant rules causing file bloat. *(from GemAll)* |
| Data Validation Checker | Scans for cells with Data Validation dropdowns and flags any with broken source ranges. *(from GemAll)* |
| Print Header/Footer Standardizer | Updates print headers and footers across all selected sheets simultaneously to ensure consistent presentation. *(from GemAll)* |
| External Link Severance Protocol | Replaces all external file references with static values and logs the original formula in a cell comment for reference. *(from GemAll)* |

### Tier 2 Python Scripts (13 tools)

| Tool Name | What It Does |
|-----------|--------------|
| Variance Analysis Generator | Compares Actual vs Budget columns across multiple files and creates summary + waterfall chart. |
| GL Reconciliation Engine | Matches two large transaction lists and flags unmatched items. |
| Fuzzy Match / Fuzzy Lookup | Matches records between two datasets using fuzzy string matching — catches vendor/customer name typos. |
| Batch Process Folder of Files | Runs the same cleaning or transformation script on every Excel file in a folder. |
| Forecast Roll-Forward | Takes last month actuals + assumptions and builds the next 12-month forecast. |
| Unstructured Data Unpivoter | Converts wide-format pivot-style data (one column per month/product) into tall database format (one row per record). Required before most analytics tools can use the data. *(from GemAll)* |
| PDF Tabular Extractor | Scans uneditable PDF documents and extracts tables directly into Excel. Solves the "my report only comes as a PDF" problem. Requires pdfplumber library. *(from GemAll)* |
| Regex Text Extractor | Scours free-text columns to isolate structured patterns like invoice numbers, account codes, or tax IDs into clean separate columns. *(from GemAll)* |
| Word Report Generator | Reads analytical outputs from Excel and programmatically generates a formatted Microsoft Word document — tables, headings, and all. *(from GemAll)* |
| Fuzzy-Match Bank Reconciler | Uses fuzzy string matching to probabilistically match ledger descriptions against bank statement text. Catches near-matches that exact lookups miss. *(from GemAll)* |
| Dynamic Master Data Mapper | Performs SQL-style joins between datasets on common keys (vendor ID, account code, etc.). Replaces complex nested VLOOKUP arrays. *(from GemAll)* |
| Reconciliation Exception Generator | Filters out all matched transactions and outputs only the unmatched exceptions into a new workbook for human review. *(from GemAll)* |
| Variance Decomposition Analyzer | Automatically breaks down financial variances into their price, volume, and mix components — standard FP&A analysis. *(from GemAll)* |

---

## Counts

| Category | Count |
|----------|-------|
| Tier 1 VBA Macros | 24 |
| Tier 1 Python Scripts | 5 |
| Tier 2 VBA Macros | 34 |
| Tier 2 Python Scripts | 13 |
| **Total Candidates** | **76** |

---

## What Was Added from GemAll.md

**Promoted to Tier 1 (2 additions):**
- Auto-Balancing GL Validator — high-value Finance tool, debits = credits check with plug
- Multi-Replace Data Cleaner — batch find-replace from mapping table, very useful for standardizing codes

**Added to Tier 2 VBA (14 additions):**
- Phantom Hyperlink Purger, Formula-to-Value Hardcoder, Multi-Sheet Batch Renamer,
  Convert Numbers to Words, Quick Corkscrew Builder, Financial Number Formatting Suite,
  General Ledger Journal Mapper, Financial Period Roll-Forward, Multi-Currency Consolidation Aggregator,
  Named Range Auditor, Conditional Format Purger, Data Validation Checker,
  Print Header/Footer Standardizer, External Link Severance Protocol

**Added to Tier 2 Python (8 additions):**
- Unstructured Data Unpivoter, PDF Tabular Extractor, Regex Text Extractor,
  Word Report Generator, Fuzzy-Match Bank Reconciler, Dynamic Master Data Mapper,
  Reconciliation Exception Generator, Variance Decomposition Analyzer

**Excluded from GemAll.md (and why):**
- ML/AI/statistical tools (ARIMA, regression, K-Means, FinBERT, Monte Carlo, anomaly detection) — too technical
- Text-to-Speech Auditor, AI Translation, Real-Time Market Data — not practical for Finance staff
- Sparkline Injector, Zoom Standardizer, Clipboard Clear — minor utility, low value
- Assumption Change Log Tracker — user declined cell-change audit trail permanently
- Expand/Collapse Grouping Toggle, Quick Link Generator — too niche
- Excel Environment Logger — IT tool, not Finance use

---

## Next Steps (After Connor Approves This List)

1. Build all 24 Tier 1 VBA tools as `.bas` modules stored in `UniversalToolsForAllFiles/vba/`
2. Package them into `KBT_UniversalTools.xlam` (Personal Macro Workbook add-in)
3. Build all 5 Tier 1 Python scripts stored in `UniversalToolsForAllFiles/python/`
4. Write coworker install guide for the add-in
5. Write coworker usage guide for the Python scripts
6. Then tackle Tier 2 as backlog
