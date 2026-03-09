# Universal Toolkit Guide

## iPipeline Universal Excel & Python Toolkit — Complete User Guide

**Total Tools:** 100+ (79 VBA tools + 22 Python scripts)

---

## Table of Contents

1. [What Is the Universal Toolkit?](#1-what-is-the-universal-toolkit)
2. [How Is This Different from the P&L Command Center?](#2-how-is-this-different-from-the-pl-command-center)
3. [Part A: VBA Tools — Setup](#3-part-a-vba-tools--setup)
4. [Part A: VBA Tools — Complete Reference](#4-part-a-vba-tools--complete-reference)
5. [Part B: Python Scripts — Setup](#5-part-b-python-scripts--setup)
6. [Part B: Python Scripts — Complete Reference](#6-part-b-python-scripts--complete-reference)
7. [Top 20 Most Useful Tools (Start Here)](#7-top-20-most-useful-tools-start-here)
8. [Use Case Playbooks](#8-use-case-playbooks)
9. [Troubleshooting](#9-troubleshooting)
10. [Frequently Asked Questions](#10-frequently-asked-questions)

---

## 1. What Is the Universal Toolkit?

The Universal Toolkit is a collection of **100+ tools** (79 VBA macros and 22+ Python scripts) that work on **any Excel file** — not just the P&L demo workbook.

Think of the P&L Command Center as a specialized tool built for one specific file. The Universal Toolkit is like a toolbox full of general-purpose tools that work on whatever file you open.

### Examples of What You Can Do

| Scenario | Tool |
|---|---|
| You received a GL export with text-stored numbers and your SUM formulas don't work | `ConvertTextToNumbers` — fixes it in one click |
| You need to remove 500 blank rows scattered throughout a report | `DeleteBlankRows` — removes them all instantly |
| You need to compare two versions of a budget file | `compare_files.py` — shows every cell that changed |
| Your boss wants the 50-page PDF report in Excel | `pdf_extractor.py` — extracts the tables directly |
| You need to consolidate 30 department budget files into one | `consolidate_files.py` — combines them with one command |
| You need to match GL entries to bank statements | `bank_reconciler.py` — fuzzy matching catches typos |
| Your invoice data has duplicates you can't find | `DuplicateInvoiceDetector` — catches Vendor + Amount + Date matches |
| You need to prepare a file for external sharing | `BuildDistributionReadyCopy` — formulas to values, everything visible |

### What's Included

| Category | VBA Tools | Python Scripts | Total |
|---|---|---|---|
| Data Cleaning | 15 | 2 | 17 |
| Audit & Compliance | 12 | 1 | 13 |
| Finance & Accounting | 14 | 7 | 21 |
| Formatting | 9 | 1 | 10 |
| Workbook Management | 15 | 3 | 18 |
| Sheet Tools | 3 | 0 | 3 |
| Data Sanitization | 4 | 0 | 4 |
| Branding | 2 | 0 | 2 |
| Data Matching & Lookup | 0 | 4 | 4 |
| Reporting & Export | 0 | 3 | 3 |
| Number Formatting | 2 | 1 | 3 |
| Shared Utilities (Core) | 9 | 0 | 9 |
| **Total** | **79+** | **22+** | **100+** |

---

## 2. How Is This Different from the P&L Command Center?

| Feature | P&L Command Center | Universal Toolkit |
|---|---|---|
| **Works on** | The P&L demo workbook only | ANY Excel file you open |
| **How to use** | Open Command Center (Ctrl+Shift+M), pick action, click Run | Run individual macros via Alt+F8, or Python scripts via command line |
| **Number of tools** | 62 actions | 100+ tools |
| **Specialized for** | Monthly P&L close process | General Excel and data tasks |
| **Setup** | Already built into the P&L file | Import the .bas files into any workbook, or run Python from command line |
| **Target user** | Finance team doing monthly close | Anyone who works with Excel |

**You do NOT need the Universal Toolkit to use the P&L Command Center.** They are separate. The Universal Toolkit is a bonus set of tools for your other work.

---

## 3. Part A: VBA Tools — Setup

### How to Add VBA Tools to Any Excel File

The VBA tools are stored as `.bas` files. To use them, you import the files into whatever Excel workbook you want to work with.

#### Step 1: Open Your Excel File

1. Open the Excel file you want to work with (any `.xlsx` or `.xlsm` file)
2. If the file is `.xlsx`, you will need to save it as `.xlsm` first:
   - Click **File > Save As**
   - Change the file type to **"Excel Macro-Enabled Workbook (.xlsm)"**
   - Click **Save**
   - If Excel asks about compatibility, click **Yes** or **OK**

#### Step 2: Open the VBA Editor

1. Press **Alt + F11** on your keyboard
2. The Visual Basic for Applications (VBA) Editor will open in a new window
3. You will see a "Project Explorer" panel on the left side showing your workbook name

#### Step 3: Import the Tool Modules

1. In the VBA Editor, click **File** in the menu bar
2. Click **Import File...**
3. Navigate to the folder where the Universal Toolkit `.bas` files are saved
4. Select the module you want to import (e.g., `modUTL_DataCleaning.bas`)
5. Click **Open**
6. The module will appear in the Project Explorer under "Modules"
7. **Repeat** for each module you want to import

#### Which Modules to Import

You do not need to import all 13 modules. Import only what you need:

| If You Need To... | Import This Module | Tools Included |
|---|---|---|
| Clean up messy data | `modUTL_Core.bas` + `modUTL_DataCleaning.bas` | 12 cleaning tools |
| Check data quality and audit | `modUTL_Core.bas` + `modUTL_Audit.bas` | 8 audit tools |
| Work with financial data | `modUTL_Core.bas` + `modUTL_Finance.bas` | 14 finance tools |
| Format and style sheets | `modUTL_Core.bas` + `modUTL_Formatting.bas` | 9 formatting tools |
| Manage workbook structure | `modUTL_Core.bas` + `modUTL_WorkbookMgmt.bas` | 15 management tools |
| Fix number/text issues | `modUTL_Core.bas` + `modUTL_DataSanitizer.bas` | 4 sanitizer tools |
| Apply iPipeline branding | `modUTL_Core.bas` + `modUTL_Branding.bas` | 2 branding tools |
| Clone sheets / manage tabs / create folders | `modUTL_Core.bas` + `modUTL_SheetTools.bas` | 4 sheet tools |
| Everything | Import all 13 modules | All 79+ tools |

> **Important:** Always import `modUTL_Core.bas` first — it contains shared utility functions that the other modules depend on.

#### Step 4: Run a Tool

1. Close the VBA Editor (**Alt + Q** or click the X)
2. Go back to Excel
3. Press **Alt + F8** to open the Macro dialog
4. You will see a list of all available macros (tools)
5. Select the one you want to run
6. Click **Run**

> **What you should see:** The tool runs on your active workbook. Most tools show a message box when they finish, telling you what they did (e.g., "Removed 47 blank rows from Sheet1").

#### Step 5: Save Your Workbook

After running a tool, save your workbook (**Ctrl + S**). If the file was originally `.xlsx`, remember to save as `.xlsm` to keep the macros.

---

## 4. Part A: VBA Tools — Complete Reference

Below is every VBA tool organized by module. For each tool:
- **What it does** — Plain English explanation
- **When to use it** — The situation where this tool helps you
- **How to run it** — Alt + F8, select the tool name, click Run

---

### Module: modUTL_Core (Shared Utilities)

> These are helper functions used by the other modules. You do not run these directly — they run automatically behind the scenes. Always import this module first.

| Function | Purpose |
|---|---|
| `UTL_TurboOn` | Speeds up macro execution by turning off screen updates |
| `UTL_TurboOff` | Restores normal Excel behavior after a macro finishes |
| `UTL_SafeDeleteSheet` | Safely deletes a sheet without error prompts |
| `UTL_LastRow` | Finds the last row of data in a column |
| `UTL_LastCol` | Finds the last column of data in a row |
| `UTL_SafeNum` | Converts a value to a number safely (returns 0 if it can't) |
| `UTL_SafeStr` | Converts a value to text safely (returns "" if it can't) |
| `UTL_StyleHeader` | Applies professional iPipeline-branded header formatting |
| `UTL_BackupSheet` | Creates a backup copy of a sheet before making changes |

---

### Module: modUTL_DataCleaning (12 Tools)

#### UnmergeAndFillDown

- **What it does:** Finds every merged cell in the active sheet, unmerges them, and fills the value down into each individual cell. This is the single most useful data cleaning tool for Finance teams.
- **When to use it:** When you receive a report from another department (or from an external system) that uses merged cells. Merged cells break PivotTables, VLOOKUP, sorting, and filtering. This tool fixes them.
- **How to run it:** Alt + F8 > `UnmergeAndFillDown` > Run
- **What to expect:** All merged cells will be unmerged and values filled down. A message box will report how many merged regions were processed.

#### FillBlanksDown

- **What it does:** For every blank cell in the used range, fills it with the value from the cell directly above it. Does not affect cells that already have values.
- **When to use it:** When you have a report where labels only appear once (in the first row of a group) and the rest of the rows are blank. This makes the data usable for PivotTables and filtering.
- **How to run it:** Alt + F8 > `FillBlanksDown` > Run

#### ConvertTextToNumbers

- **What it does:** Scans all cells on the active sheet and converts any text-stored numbers to actual numbers. Handles plain numbers, currency formats ($1,234.56), and percentages.
- **When to use it:** When your SUM formulas return 0 even though there are numbers in the cells, or when VLOOKUP can't find matches because one column has numbers stored as text.
- **How to run it:** Alt + F8 > `ConvertTextToNumbers` > Run
- **What to expect:** A message box showing how many cells were converted.

#### RemoveLeadingTrailingSpaces

- **What it does:** Trims leading and trailing spaces from every text cell in the active sheet. Also removes double spaces within text.
- **When to use it:** When VLOOKUP fails because one list has trailing spaces. When sorting puts items in the wrong order because of leading spaces.
- **How to run it:** Alt + F8 > `RemoveLeadingTrailingSpaces` > Run

#### DeleteBlankRows

- **What it does:** Deletes every completely empty row from the active sheet. A row must have zero values, zero formulas, and zero formatting to be deleted.
- **When to use it:** After deleting data that left empty rows, or when imported data has blank rows scattered throughout.
- **How to run it:** Alt + F8 > `DeleteBlankRows` > Run
- **What to expect:** Message box showing how many rows were deleted.

#### ReplaceErrorValues

- **What it does:** Finds every cell containing a formula error (#N/A, #REF!, #VALUE!, #DIV/0!, #NAME?, #NULL!, #NUM!) and replaces it with a blank cell.
- **When to use it:** When you need to clean up a workbook that has scattered errors — especially before creating PivotTables or charts, which can break on error values.
- **How to run it:** Alt + F8 > `ReplaceErrorValues` > Run

#### HighlightDuplicateRows

- **What it does:** Highlights duplicate rows in yellow so you can review them visually. Does NOT delete anything — review only.
- **When to use it:** When you want to see which rows are duplicated before deciding what to do about them.
- **How to run it:** Alt + F8 > `HighlightDuplicateRows` > Run

#### RemoveDuplicateRows

- **What it does:** Permanently deletes duplicate rows, keeping the first occurrence of each unique row.
- **When to use it:** After reviewing duplicates (using HighlightDuplicateRows) and confirming they should be removed.
- **Warning:** This permanently deletes rows. Run HighlightDuplicateRows first to review.

#### MultiReplaceDataCleaner

- **What it does:** Performs batch find-and-replace using a mapping table. You provide a two-column list (Find | Replace) and it applies every replacement across the active sheet.
- **When to use it:** When you need to standardize hundreds of vendor names, department codes, or account names (e.g., replace "AMZN" with "Amazon", "MSFT" with "Microsoft", etc.).
- **How to run it:** Create a two-column mapping table on a separate sheet first. Then Alt + F8 > `MultiReplaceDataCleaner` > Run. It will ask which sheet contains the mapping.

#### FormulaToValueHardcoder

- **What it does:** Converts every formula in the active sheet to its current calculated value. The formulas are permanently replaced with static numbers.
- **When to use it:** Before sharing a file externally (to hide your formulas), or to speed up a slow workbook with thousands of formulas.
- **Warning:** Irreversible. Save a backup first.

#### PhantomHyperlinkPurger

- **What it does:** Removes every hyperlink from the active sheet while keeping the cell text. This fixes files that are slow because of thousands of embedded hyperlinks.
- **When to use it:** When a workbook is extremely slow and you suspect hyperlinks are the cause. Also useful for cleaning pasted data from web sources.

#### ConvertNumbersToWords

- **What it does:** Converts numbers to written English words (e.g., 1,250 becomes "One Thousand Two Hundred Fifty Dollars"). Useful for generating check amounts, contract values, or legal document figures.
- **When to use it:** When you need a number spelled out for a financial document, check, or legal agreement.

---

### Module: modUTL_DataCleaningPlus (3 Enhanced Tools)

#### UniversalWhitespaceCleaner

- **What it does:** Goes beyond basic TRIM. Removes leading/trailing spaces, double spaces, non-breaking spaces (from web copy-paste), zero-width characters, and other invisible whitespace that causes lookup failures.
- **When to use it:** When ConvertTextToNumbers and RemoveLeadingTrailingSpaces don't fix your lookup problems — there may be invisible characters hiding in the cells.

#### NonPrintableCharStripper

- **What it does:** Removes control characters (ASCII 0–31), soft hyphens, invisible Unicode characters, and other non-printable characters that can break formulas and data matching.
- **When to use it:** When data imported from a database, web scrape, or PDF contains invisible characters that you can't see but are causing problems.

#### TextCaseStandardizer

- **What it does:** Converts text in selected cells to UPPERCASE, lowercase, Title Case, or Sentence case. You choose the format.
- **When to use it:** When you need to standardize text formatting across a column (e.g., all vendor names in Title Case, all codes in UPPERCASE).

---

### Module: modUTL_Audit (8 Tools)

#### ExternalLinkFinder

- **What it does:** Scans the entire workbook and lists every cell that references an external file. Shows the cell address, formula, and the external file path being referenced.
- **When to use it:** Before sharing a workbook externally, or when you get #REF! errors because the linked file was moved or deleted.
- **What to expect:** A report sheet listing every external reference with its location and target file.

#### CircularReferenceDetector

- **What it does:** Finds every circular reference in the workbook (cells that directly or indirectly reference themselves) and lists them in a report.
- **When to use it:** When Excel shows a "Circular Reference" warning and you can't find where it is. This tool finds them all.

#### WorkbookErrorScanner

- **What it does:** Scans every cell in every sheet for formula errors (#REF!, #VALUE!, #DIV/0!, #N/A, etc.) and creates a comprehensive report.
- **When to use it:** During audit prep, or when you need to clean up a workbook that has accumulated errors over time.

#### DataQualityScorecard

- **What it does:** Creates a column-by-column data quality report showing: how many blanks, how many errors, how many duplicates, what data types are present, and what percentage of each column is complete.
- **When to use it:** When you receive a new dataset and need to understand its quality before working with it.

#### NamedRangeAuditor

- **What it does:** Lists every named range in the workbook, shows what it points to, and flags any broken references (named ranges that point to deleted sheets or cells).
- **When to use it:** During workbook cleanup, or when formulas using named ranges are returning errors.

#### DataValidationChecker

- **What it does:** Finds every cell with a data validation dropdown and checks if the source range is still valid. Reports any dropdowns with broken source references.
- **When to use it:** When dropdown lists are showing old data or are blank.

#### InconsistentFormulasAuditor

- **What it does:** Checks each column of formulas and identifies any formula that differs from the majority pattern. For example, if 99 rows use `=B2*C2` and one row uses `=B2+C2`, it flags the outlier.
- **When to use it:** When you suspect a formula was accidentally changed in one row. This finds those "needle in a haystack" formula errors.

#### ExternalLinkSeveranceProtocol

- **What it does:** Replaces all external link formulas with their current values (breaks the links) and saves the original formulas as cell comments for reference. Creates a backup sheet first.
- **When to use it:** When you need to permanently break external links but want to keep a record of what the formulas were.

---

### Module: modUTL_AuditPlus (4 Enhanced Tools)

#### DataBoundaryDetector

- **What it does:** Detects the actual rectangular data area on a sheet and reports any gaps — blank rows, blank columns, or inconsistencies in the data boundary.
- **When to use it:** When imported data has irregular boundaries and you need to understand the exact shape of the dataset.

#### HeaderValidator

- **What it does:** You provide a list of expected column headers, and the tool checks the actual headers against your list. Reports exact matches, fuzzy matches (close but not exact), and completely missing headers.
- **When to use it:** When validating that an imported file has the correct column structure before processing it.

#### FormulaErrorFinder

- **What it does:** Scans every sheet for formula errors and creates a detailed report with the sheet name, cell address, error type, and the formula that caused it.
- **When to use it:** Quick way to find all formula errors across an entire multi-sheet workbook.

#### FormulaConsistencyChecker

- **What it does:** Checks if all formulas in a column follow the same pattern. If one row has a different formula structure, it flags it as an outlier.
- **When to use it:** Auditing large worksheets where one inconsistent formula can cause major calculation errors.

---

### Module: modUTL_DataSanitizer (4 Tools)

#### RunFullSanitize

- **What it does:** Runs all three sanitization fixes in one click: (1) convert text-stored numbers to real numbers, (2) fix floating-point tails (e.g., 9412.300000001 becomes 9412.30), (3) normalize integer formatting.
- **When to use it:** After importing data from any external source. This is the "fix everything numeric" button.
- **Important:** Safe for financial data — never touches dates, names, customer IDs, or text-labeled columns. Uses smart header detection to skip non-numeric columns.

#### PreviewSanitizeChanges

- **What it does:** Shows you exactly what RunFullSanitize WOULD change, without actually changing anything. Dry-run only.
- **When to use it:** Before running RunFullSanitize, to review the changes and make sure nothing unexpected will be touched.

#### FixFloatingPointTails

- **What it does:** Fixes floating-point precision errors where numbers like 100.00 show as 100.0000000001 or 99.9999999999. Rounds to the appropriate number of decimal places.
- **When to use it:** When you see numbers with excessive decimal places that don't match the expected precision.

#### ConvertTextStoredNumbers

- **What it does:** Converts cells where numbers are stored as text to actual numeric values. Same as ConvertTextToNumbers in the DataCleaning module but with header-aware safety.
- **When to use it:** When SUM formulas return 0, when VLOOKUP can't find numbers, or when sorting puts numbers in the wrong order (1, 10, 2, 20 instead of 1, 2, 10, 20).

---

### Module: modUTL_Formatting (9 Tools)

#### AutoFitAllColumnsRows

- **What it does:** Auto-sizes every column and every row on ALL sheets in the workbook to fit their content perfectly.
- **When to use it:** After importing data or pasting content when columns are too narrow or too wide.

#### FreezeTopRowAllSheets

- **What it does:** Applies freeze panes to row 1 on every sheet so headers stay visible while scrolling.
- **When to use it:** After setting up a multi-sheet workbook where you want headers frozen everywhere.

#### NumberFormatStandardizer

- **What it does:** Applies a consistent number format (#,##0.00) to all numeric cells on the active sheet.
- **When to use it:** When numbers have inconsistent formatting (some with decimals, some without, some with commas, some without).

#### CurrencyFormatStandardizer

- **What it does:** Applies currency format ($#,##0.00) to a selected range of cells.
- **When to use it:** When you need dollar signs on financial data.

#### DateFormatStandardizer

- **What it does:** Normalizes all dates on the active sheet to a consistent format (MM/DD/YYYY).
- **When to use it:** When dates are in mixed formats (some as 3/5/2026, some as 2026-03-05, some as March 5, 2026).

#### HighlightNegativesRed

- **What it does:** Applies red font or red fill to every negative number on the active sheet.
- **When to use it:** When you need negative numbers to stand out visually in a financial report.

#### FinancialNumberFormattingSuite

- **What it does:** Opens a dialog box where you choose a financial formatting style: Accounting (parentheses for negatives), Factor 000s (divide by 1,000 with "K" suffix), Percentage, Plain Number, or Integer. Applies to the selected range.
- **When to use it:** When formatting financial reports for different audiences (board wants 000s, detail wants full numbers, etc.).

#### ConditionalFormatPurger

- **What it does:** Removes ALL conditional formatting rules from the active sheet. Does not change the current cell colors — only removes the rules.
- **When to use it:** When conditional formatting has accumulated over time and is making the workbook slow or causing unexpected color changes.

#### PrintHeaderFooterStandardizer

- **What it does:** Applies a consistent header and footer (filename, sheet name, page number, date) to all sheets in the workbook for printing.
- **When to use it:** Before printing or exporting to PDF, to ensure every page has a professional header and footer.

---

### Module: modUTL_Finance (14 Tools)

#### DuplicateInvoiceDetector

- **What it does:** Scans invoice data and flags potential duplicates by matching Vendor + Amount + Date within a 3-day window. Catches duplicate payments before they happen.
- **When to use it:** Before processing a batch of invoices, or during month-end AP review.
- **What to expect:** Duplicate matches highlighted, with a summary report.

#### AutoBalancingGLValidator

- **What it does:** Checks that debit and credit columns balance. If they don't, it calculates the difference and optionally inserts a balancing "plug" entry.
- **When to use it:** When validating a GL export or journal entry batch.

#### TrialBalanceChecker

- **What it does:** Sums all debits and all credits and verifies they are equal. Reports the difference if they don't balance.
- **When to use it:** During month-end close when verifying the trial balance ties.

#### JournalEntryValidator

- **What it does:** Groups journal entries by entry number and verifies that each individual entry balances (total debits = total credits for each entry).
- **When to use it:** When reviewing a batch of journal entries before posting.

#### FluxAnalysis

- **What it does:** Compares two columns (e.g., Actual vs. Prior Month) row by row and calculates the dollar and percentage change. Flags changes above a threshold you specify.
- **When to use it:** During variance analysis, budget review, or any two-column comparison.

#### APAgingSummaryGenerator

- **What it does:** Takes AP invoice data and generates an aging summary bucketed by days overdue: Current, 0–30, 31–60, 61–90, 90+.
- **When to use it:** For AP aging reports during month-end close or cash flow planning.

#### ARAgingSummaryGenerator

- **What it does:** Same as AP Aging but for Accounts Receivable — buckets outstanding invoices by days since invoice date.
- **When to use it:** For AR aging reports and collections prioritization.

#### AgingBucketCalculator

- **What it does:** Adds an "Aging Bucket" column to any dataset with a date column. Calculates days from the date to today and assigns a bucket (Current, 0–30, 31–60, 61–90, 90+).
- **When to use it:** When you have any date-based data and need to categorize by age.

#### VarianceAnalysisTemplate

- **What it does:** Automatically adds "$ Variance" and "% Variance" columns next to existing Actual and Budget columns. Calculates the formulas for you.
- **When to use it:** When building a variance report from scratch — this saves you from typing the formulas manually.

#### QuickCorkscrewBuilder

- **What it does:** Builds a roll-forward schedule (Beginning Balance + Additions - Deductions = Ending Balance) from your data. Creates the structure and formulas automatically.
- **When to use it:** For building roll-forward schedules for any account (inventory, prepaid, accrued liabilities, etc.).

#### FinancialPeriodRollForward

- **What it does:** Updates month-end column headers to the next period and clears input cells for new data entry. Shifts the previous month's actuals into a "Prior" column.
- **When to use it:** At the start of each new month to prepare the workbook for new data.

#### MultiCurrencyConsolidationAggregator

- **What it does:** Consolidates amounts from multiple currencies using an FX rate table. Converts all amounts to a base currency before summing.
- **When to use it:** When consolidating international data with multiple currencies.

#### RatioAnalysisDashboard

- **What it does:** Calculates key financial ratios (gross margin, operating margin, current ratio, quick ratio, ROE, ROA, debt-to-equity, etc.) and displays them in a formatted dashboard.
- **When to use it:** For quick financial analysis or board reporting.

#### GeneralLedgerJournalMapper

- **What it does:** Transforms a trial balance into a journal entry upload template formatted for your GL system.
- **When to use it:** When preparing journal entries for system upload.

---

### Module: modUTL_WorkbookMgmt (15 Tools)

#### UnhideAllSheetsRowsColumns

- **What it does:** Makes every hidden sheet, row, and column visible across the entire workbook. Includes "very hidden" sheets.
- **When to use it:** When you receive a workbook with hidden content and need to see everything.

#### ExportAllSheetsCombinedPDF

- **What it does:** Exports all visible sheets into a single multi-page PDF file.
- **When to use it:** When you need a complete workbook PDF with all sheets combined.

#### FindReplaceAcrossAllSheets

- **What it does:** Performs find-and-replace across every sheet in the workbook simultaneously.
- **When to use it:** When you need to update a company name, account code, or label that appears on multiple sheets.

#### SearchAcrossAllSheets

- **What it does:** Searches for any text or number across all sheets and returns every match with the sheet name and cell address.
- **When to use it:** When you need to find where a specific value appears in a large multi-sheet workbook.

#### MultiSheetBatchRenamer

- **What it does:** Performs find-and-replace on sheet tab names. For example, replace "2024" with "2025" in every sheet name.
- **When to use it:** When rolling forward a workbook to a new year or period.

#### SortWorksheetsAlphabetically

- **What it does:** Rearranges all sheet tabs in alphabetical order (A to Z).
- **When to use it:** When your workbook has many tabs and you want them organized.

#### CreateTableOfContents

- **What it does:** Creates a new "Table of Contents" sheet with clickable hyperlinks to every sheet in the workbook.
- **When to use it:** For large workbooks to make navigation easy.

#### ProtectAllSheets / UnprotectAllSheets

- **What they do:** Apply or remove sheet protection across all sheets in one action.
- **When to use them:** Before distributing a file (protect), or when you need to edit a protected file (unprotect).

#### LockAllFormulaCells

- **What it does:** Locks all cells containing formulas while leaving input cells (plain values) editable. Then applies sheet protection.
- **When to use it:** When you want users to be able to enter data but not accidentally overwrite formulas.

#### ExportActiveSheetPDF / ExportAllSheetsIndividualPDFs

- **What they do:** Export the current sheet as a PDF, or export every visible sheet as a separate PDF file.
- **When to use them:** For creating individual PDF files from a multi-sheet workbook.

#### ResetAllFilters

- **What it does:** Clears all AutoFilter criteria on every sheet, showing all rows.
- **When to use it:** When you want to quickly clear all filters across a workbook to see all data.

#### BuildDistributionReadyCopy

- **What it does:** Creates a clean copy of the workbook with all formulas converted to values, all hidden sheets made visible, and the filename appended with "_DIST". Ready for external distribution.
- **When to use it:** Before sharing a workbook externally where you don't want formulas visible.

#### WorkbookHealthCheck

- **What it does:** Creates a diagnostic report showing: file size, total sheets, total cells with data, total formulas, total errors, total external links, total named ranges, total blanks in data areas, and total duplicates.
- **When to use it:** To get a complete picture of a workbook's health and complexity.

---

### Module: modUTL_Branding (2 Tools)

#### ApplyiPipelineBranding

- **What it does:** Detects headers and total rows on the active sheet and applies iPipeline brand styling: iPipeline Blue header row with white bold text, Navy Blue total row, alternating row colors, Arial font throughout.
- **When to use it:** When you want any workbook to look like an official iPipeline document.

#### SetiPipelineThemeColors

- **What it does:** Sets the workbook's theme colors to the iPipeline brand palette so the standard Excel color picker shows iPipeline colors.
- **When to use it:** Before building charts or formatting — the iPipeline colors will be the default choices in the color picker.

---

### Module: modUTL_SheetTools (4 Tools)

#### ListAllSheetsWithLinks

- **What it does:** Creates a "UTL_SheetIndex" sheet listing every sheet in the workbook with clickable hyperlinks and visibility status (Visible/Hidden/Very Hidden).
- **When to use it:** For navigation in large workbooks.

#### TemplateCloner

- **What it does:** Pick any sheet, type how many copies you want (1–50), and get instant clones. Handles name conflicts and the 31-character sheet name limit.
- **When to use it:** When you need to create multiple copies of a template sheet (e.g., one per month, one per department, one per region).

#### GenerateUniqueCustomerIDs

- **What it does:** Scans existing IDs in a column, finds the maximum, and fills blank cells with sequential IDs in CUST-00001 format. Custom prefix supported.
- **When to use it:** When you need to assign unique IDs to new records without duplicating existing ones.

#### CreateFoldersFromSelection

- **What it does:** Highlight a column of cell values (names, projects, clients, months, etc.) and this tool creates a Windows folder for each value. It shows you a preview first, then asks where to create the folders using a folder picker dialog.
- **Safety features:** Skips blanks, skips duplicates, cleans illegal folder characters (\ / : * ? " < > |), won't overwrite existing folders, preview before creating.
- **When to use it:** When you need to create folders from a list — client folders, project folders, monthly folders, vendor folders, department folders, etc.

---

### Module: modUTL_DuplicateDetection (1 Tool)

#### ExactDuplicateFinder

- **What it does:** You pick a key column, and the tool highlights all duplicate values in that column in yellow. Creates a summary report with duplicate counts.
- **When to use it:** When you need to find duplicates based on a specific key (invoice number, customer ID, transaction ID, etc.).

---

### Module: modUTL_NumberFormat (2 Tools)

#### EnhancedTextToNumberConverter

- **What it does:** Advanced text-to-number conversion that handles currency strings ("$1,234.56"), parenthetical negatives ("(500)"), percentages ("45%"), and European number formats.
- **When to use it:** When ConvertTextToNumbers doesn't handle your specific number format (e.g., imported data with currency symbols embedded in the text).

#### WorkbookMetadataReporter

- **What it does:** Creates a summary report with: file information (name, path, size, dates), sheet inventory with cell counts, named ranges list, external link list, and metadata.
- **When to use it:** For documentation, auditing, or understanding an unfamiliar workbook.

---

## 5. Part B: Python Scripts — Setup

### What Are the Python Scripts?

The Python scripts handle tasks that are too complex or too slow for VBA — like processing 50 files at once, fuzzy-matching thousands of records, extracting data from PDFs, or building interactive dashboards.

### Do I Need Python?

**No.** Python is optional. The VBA tools cover most common tasks. Python scripts are for advanced use cases:
- Processing 50+ files at once
- Fuzzy matching (finding "Amzon Web Srvcs" matches "Amazon Web Services")
- PDF table extraction
- Interactive web dashboards
- Variance decomposition into Price/Volume/Mix effects

If you don't need these, skip Part B entirely.

### Python Setup (One Time)

If you do want to use the Python scripts:

#### Step 1: Install Python

1. Go to **python.org/downloads** (or ask IT to install it)
2. Download Python 3.11 or later
3. **During installation, check the box that says "Add Python to PATH"** — this is critical
4. Click "Install Now"
5. When installation finishes, open Command Prompt (search for "cmd" in Start Menu)
6. Type `python --version` and press Enter
7. You should see something like `Python 3.11.5` — if you do, Python is installed

#### Step 2: Install Required Libraries

1. Open Command Prompt
2. Navigate to the Universal Toolkit Python folder:
   ```
   cd C:\path\to\UniversalToolsForAllFiles\python
   ```
3. Install the required libraries:
   ```
   pip install -r requirements.txt
   ```
4. Wait for all libraries to install (this may take 1–2 minutes)

#### Step 3: Run a Script

1. Open Command Prompt
2. Navigate to the script folder
3. Run a script with:
   ```
   python script_name.py --input "your_file.xlsx" --output "result.xlsx"
   ```
4. Each script has a `--help` flag that shows all available options:
   ```
   python script_name.py --help
   ```

---

## 6. Part B: Python Scripts — Complete Reference

### Data Cleaning Scripts

#### clean_data.py — Universal Data Cleaner

- **What it does:** Cleans any Excel or CSV file: removes blank rows/columns, trims spaces, converts text-to-numbers, standardizes dates, removes duplicates.
- **When to use it:** First step after receiving any messy data file.
- **Command:** `python clean_data.py --input "messy_data.xlsx" --output "clean_data.xlsx"`

#### batch_process.py — Batch File Cleaner

- **What it does:** Runs the data cleaner on every Excel file in an entire folder. Processes 50+ files automatically.
- **When to use it:** When you receive a folder full of files from different departments that all need cleaning.
- **Command:** `python batch_process.py --folder "C:\Budget_Files\" --output "C:\Cleaned\"`

### File Comparison & Reconciliation Scripts

#### compare_files.py — Cell-by-Cell File Comparison

- **What it does:** Compares two Excel workbooks cell by cell. Creates a diff report highlighting every difference with color coding (Green = Added, Red = Removed, Yellow = Changed).
- **When to use it:** When you have two versions of a file and need to know exactly what changed.
- **Command:** `python compare_files.py --file1 "v1.xlsx" --file2 "v2.xlsx" --output "diff.xlsx"`

#### bank_reconciler.py — Bank Statement Reconciler

- **What it does:** Fuzzy-matches ledger entries against bank statement entries. Catches typos and abbreviations (e.g., "AMZN WEB SRVCS" matches "Amazon Web Services").
- **When to use it:** Monthly bank reconciliation. Saves hours of manual matching.
- **Command:** `python bank_reconciler.py --ledger "gl.xlsx" --bank "statement.xlsx" --output "recon.xlsx"`

#### gl_reconciliation.py — GL to Sub-Ledger Reconciler

- **What it does:** Matches GL transactions to sub-ledger entries by Amount + Date (within 3 days) + optional reference number. Reports matched pairs and unmatched items.
- **When to use it:** Reconciling GL to AP, AR, or any sub-ledger.
- **Command:** `python gl_reconciliation.py --gl "gl_data.xlsx" --subledger "ap_data.xlsx" --output "recon.xlsx"`

#### reconciliation_exceptions.py — Exception Report Generator

- **What it does:** Compares two datasets, filters out matched items, and outputs ONLY the unmatched exceptions. Supports composite key matching.
- **When to use it:** When you only want to see items that DON'T match — the exceptions that need attention.
- **Command:** `python reconciliation_exceptions.py --source "file1.xlsx" --target "file2.xlsx" --output "exceptions.xlsx"`

#### two_file_reconciler.py — Two-File Reconciler

- **What it does:** Reconciles any two Excel/CSV files on shared key columns. Reports matches, mismatches, and items only in one file.
- **When to use it:** General-purpose reconciliation between any two datasets.
- **Command:** `python two_file_reconciler.py --file1 "data1.xlsx" --file2 "data2.xlsx" --key "InvoiceNo" --output "recon.xlsx"`

### Data Consolidation Scripts

#### consolidate_files.py — Multi-File Consolidator

- **What it does:** Combines hundreds of Excel files in a folder into one master sheet. Adds a "Source_File" column so you know which file each row came from.
- **When to use it:** Combining budget files, expense reports, or any collection of similarly structured files.
- **Command:** `python consolidate_files.py --folder "C:\Department_Budgets\" --output "consolidated.xlsx"`

#### consolidate_budget.py — Budget Consolidation with Variance

- **What it does:** Merges 50+ department budget files into a master with dollar and percentage variance columns. Creates a summary by file.
- **When to use it:** Annual budget consolidation — combining all department budget submissions.
- **Command:** `python consolidate_budget.py --folder "C:\Budgets\" --output "master_budget.xlsx"`

#### multi_file_consolidator.py — Advanced Multi-File Consolidator

- **What it does:** Consolidates files with column mismatch handling (union mode keeps all columns, intersection mode keeps only common columns). Supports recursive directory search.
- **When to use it:** When files from different sources have different column structures.
- **Command:** `python multi_file_consolidator.py --folder "C:\Data\" --mode union --output "combined.xlsx"`

### Data Matching & Lookup Scripts

#### fuzzy_lookup.py — Fuzzy Match Lookups

- **What it does:** Like VLOOKUP but with fuzzy matching. Matches "John Smth" to "John Smith", "Microsft Corp" to "Microsoft Corp". Returns match scores.
- **When to use it:** Matching customer lists, vendor lists, or any two lists with inconsistent naming.
- **Command:** `python fuzzy_lookup.py --source "our_list.xlsx" --lookup "master_list.xlsx" --key "Company" --output "matched.xlsx"`

#### master_data_mapper.py — SQL-Style Data Joins

- **What it does:** SQL-style JOIN between two Excel files on a common key. Supports LEFT, INNER, OUTER, and RIGHT joins.
- **When to use it:** When you need VLOOKUP but across two separate files, with control over which records to keep.
- **Command:** `python master_data_mapper.py --left "transactions.xlsx" --right "master.xlsx" --key "AcctNo" --join inner --output "joined.xlsx"`

### Financial Analysis Scripts

#### variance_analysis.py — Multi-File Variance Analysis

- **What it does:** Consolidates Actual vs. Budget data from multiple files. Creates summary grouped by department with dollar and percentage variance plus a bar chart.
- **When to use it:** Budget-to-actual analysis across multiple departments or entities.
- **Command:** `python variance_analysis.py --folder "C:\Actuals_vs_Budget\" --output "variance_report.xlsx"`

#### variance_decomposition.py — Price/Volume/Mix Decomposition

- **What it does:** Breaks down financial variances into three effects: Price Effect, Volume Effect, and Mix Effect. Standard FP&A methodology.
- **When to use it:** When leadership asks "WHY did revenue change?" — was it price, volume, or product mix?
- **Command:** `python variance_decomposition.py --input "revenue_data.xlsx" --output "decomposition.xlsx"`

#### forecast_rollforward.py — Rolling Forecast Builder

- **What it does:** Builds a 12-month rolling forecast from historical data. Methods: moving average, growth rate, or flat. Creates forecast sheet with actuals-vs-forecast chart.
- **When to use it:** For forward-looking revenue or expense projections.
- **Command:** `python forecast_rollforward.py --input "historical.xlsx" --method moving_avg --periods 12 --output "forecast.xlsx"`

#### aging_report.py — AR/AP Aging Report

- **What it does:** Generates aging reports from invoice data. Buckets: Current, 0–30, 31–60, 61–90, 90+ days. Creates detail + summary + by-vendor pivot.
- **When to use it:** For AP aging during month-end or AR collections prioritization.
- **Command:** `python aging_report.py --input "invoices.xlsx" --output "aging.xlsx"`

### Utility Scripts

#### pdf_extractor.py — PDF Table Extractor

- **What it does:** Extracts tables from PDF documents directly into Excel. Solves the "my report only comes as a PDF" problem.
- **When to use it:** When you receive financial data as a PDF and need it in Excel.
- **Command:** `python pdf_extractor.py --input "report.pdf" --output "extracted.xlsx" --pages "1-5"`

#### unpivot_data.py — Wide-to-Tall Data Converter

- **What it does:** Converts wide-format data (one column per month: Jan, Feb, Mar...) to tall database format (one row per month per record). Essential for PivotTables and BI tools.
- **When to use it:** When your data is in "spreadsheet format" (wide) and you need it in "database format" (tall).
- **Command:** `python unpivot_data.py --input "wide_data.xlsx" --output "tall_data.xlsx" --id-cols "Department,Account"`

#### regex_extractor.py — Pattern Extractor

- **What it does:** Extracts structured data from free text using pattern matching. Presets: invoice numbers, email addresses, phone numbers, dates, currency amounts, account numbers, zip codes, SSNs (masked).
- **When to use it:** When you have unstructured text with data buried in it (e.g., extracting invoice numbers from email text).
- **Command:** `python regex_extractor.py --input "text_data.xlsx" --column "Description" --preset invoice --output "extracted.xlsx"`

#### word_report.py — Excel-to-Word Report Generator

- **What it does:** Reads Excel data and generates a formatted Word document with tables, headings, and styling.
- **When to use it:** When you need to turn Excel analysis into a Word document for distribution.
- **Command:** `python word_report.py --input "analysis.xlsx" --output "report.docx"`

#### date_format_unifier.py — Date Format Standardizer

- **What it does:** Detects all date columns and converts every date to a consistent format. Handles 13+ date formats including Excel serial dates.
- **When to use it:** When dates are in mixed formats across a file (some MM/DD/YYYY, some YYYY-MM-DD, some text dates).
- **Command:** `python date_format_unifier.py --input "data.xlsx" --format "YYYY-MM-DD" --output "standardized.xlsx"`

#### sql_query_tool.py — SQL Query Tool for Excel/CSV

- **What it does:** Run SQL queries directly against Excel/CSV files. Each file becomes a table you can SELECT from, JOIN, filter, and aggregate.
- **When to use it:** When you want to query Excel data using SQL instead of formulas.
- **Command:** `python sql_query_tool.py --input "data.xlsx"` (then type SQL queries interactively)

---

## 7. Top 20 Most Useful Tools (Start Here)

If you are new to the Universal Toolkit, start with these 20 tools. They solve the most common problems.

| # | Tool | Type | Solves This Problem |
|---|---|---|---|
| 1 | `ConvertTextToNumbers` | VBA | SUM formulas returning 0, VLOOKUP failures |
| 2 | `DeleteBlankRows` | VBA | Scattered empty rows from deleted data |
| 3 | `UnmergeAndFillDown` | VBA | Merged cells breaking PivotTables and VLOOKUP |
| 4 | `RunFullSanitize` | VBA | All numeric data problems at once |
| 5 | `RemoveLeadingTrailingSpaces` | VBA | Invisible spaces causing lookup failures |
| 6 | `FindReplaceAcrossAllSheets` | VBA | Need to change something on every sheet |
| 7 | `AutoFitAllColumnsRows` | VBA | Columns too narrow or too wide |
| 8 | `ExportAllSheetsCombinedPDF` | VBA | Need the whole workbook as one PDF |
| 9 | `DuplicateInvoiceDetector` | VBA | Catching duplicate payments |
| 10 | `WorkbookHealthCheck` | VBA | Understanding a new/unfamiliar workbook |
| 11 | `ExternalLinkFinder` | VBA | Tracking down #REF errors from broken links |
| 12 | `ApplyiPipelineBranding` | VBA | Making any sheet look professional |
| 13 | `BuildDistributionReadyCopy` | VBA | Preparing files for external sharing |
| 14 | `FluxAnalysis` | VBA | Quick two-column variance comparison |
| 15 | `clean_data.py` | Python | Cleaning messy data files |
| 16 | `compare_files.py` | Python | Finding differences between file versions |
| 17 | `consolidate_files.py` | Python | Combining multiple files into one |
| 18 | `fuzzy_lookup.py` | Python | Matching lists with inconsistent naming |
| 19 | `bank_reconciler.py` | Python | Monthly bank reconciliation |
| 20 | `pdf_extractor.py` | Python | Getting PDF tables into Excel |

---

## 8. Use Case Playbooks

### Playbook 1: "I Just Received a Messy Data File"

1. Open the file in Excel
2. Import `modUTL_Core.bas` + `modUTL_DataCleaning.bas` + `modUTL_DataSanitizer.bas`
3. Run `PreviewSanitizeChanges` — see what needs fixing (dry run)
4. Run `RunFullSanitize` — fix all numeric issues
5. Run `RemoveLeadingTrailingSpaces` — fix text issues
6. Run `DeleteBlankRows` — remove empty rows
7. Run `UnmergeAndFillDown` — fix merged cells (if any)
8. Save the cleaned file

### Playbook 2: "I Need to Consolidate 30 Budget Files"

1. Put all 30 files in one folder
2. Open Command Prompt
3. Run: `python consolidate_files.py --folder "C:\Budget_Files\" --output "master.xlsx"`
4. Open the resulting `master.xlsx` — all 30 files combined with a Source_File column

### Playbook 3: "I Need to Reconcile GL to Bank Statement"

1. Open Command Prompt
2. Run: `python bank_reconciler.py --ledger "gl_export.xlsx" --bank "bank_statement.xlsx" --output "recon.xlsx"`
3. Open `recon.xlsx` — matched items, unmatched GL entries, and unmatched bank entries on separate sheets

### Playbook 4: "I Need to Find What Changed Between Two File Versions"

1. Open Command Prompt
2. Run: `python compare_files.py --file1 "Budget_v1.xlsx" --file2 "Budget_v2.xlsx" --output "changes.xlsx"`
3. Open `changes.xlsx` — every difference highlighted with color coding

### Playbook 5: "I Need to Prepare a File for External Sharing"

1. Open the file in Excel
2. Import `modUTL_Core.bas` + `modUTL_WorkbookMgmt.bas`
3. Run `BuildDistributionReadyCopy` — creates a clean copy with formulas as values
4. Send the `_DIST` copy, keep the original

### Playbook 6: "I Need to Match Two Customer Lists with Different Spelling"

1. Open Command Prompt
2. Run: `python fuzzy_lookup.py --source "our_customers.xlsx" --lookup "vendor_list.xlsx" --key "CompanyName" --output "matched.xlsx"`
3. Open `matched.xlsx` — each record matched with a confidence score

---

## 9. Troubleshooting

### VBA Tool Troubleshooting

| Problem | Solution |
|---|---|
| Alt+F8 shows no macros | Make sure macros are enabled (File > Options > Trust Center) and you imported the .bas files |
| "Sub or Function not defined" error | You need to import `modUTL_Core.bas` — the other modules depend on it |
| Tool runs but nothing happens | Check that you have data on the active sheet. Most tools work on the active sheet only |
| Tool is very slow | Large files (100,000+ rows) take longer. Wait for it to finish — check the status bar |
| "Object variable not set" error | The tool couldn't find the expected data structure. Check that your data has headers in row 1 |

### Python Script Troubleshooting

| Problem | Solution |
|---|---|
| "python is not recognized" | Python is not in your PATH. Reinstall Python and check "Add to PATH" |
| "ModuleNotFoundError" | Run `pip install -r requirements.txt` to install missing libraries |
| "FileNotFoundError" | Check the file path — use full path with quotes if there are spaces |
| Script runs but output is empty | Check that the input file has data on the first sheet |
| "Permission denied" on output file | Close the output file in Excel before running the script |

---

## 10. Frequently Asked Questions

### Q: Do I need to import the VBA modules into every file I want to use them on?

**A:** Yes. The VBA modules live inside the workbook file. When you import them, they are added to that specific file. If you want to use them on a different file, you need to import them again. (In the future, we plan to package these as an Excel Add-In that loads automatically with every file.)

### Q: Will the VBA tools change my original data?

**A:** Some tools modify data (like ConvertTextToNumbers, DeleteBlankRows) and some are read-only (like WorkbookHealthCheck, ExternalLinkFinder). The tool descriptions above indicate which ones modify data. When in doubt, save a backup of your file first.

### Q: Can I undo a VBA tool action?

**A:** Ctrl+Z may work for small changes, but it is unreliable for large VBA operations. The best practice is to save your file before running any tool that modifies data.

### Q: Do the Python scripts modify my original file?

**A:** No. Every Python script creates a NEW output file. Your original file is never changed.

### Q: Can I use the Python scripts on CSV files?

**A:** Yes. Most Python scripts accept both `.xlsx` and `.csv` files as input.

### Q: How do I know which tool to use?

**A:** Check the "Top 20 Most Useful Tools" section or the "Use Case Playbooks" section above. If you still aren't sure, describe your problem to the Finance Automation Team and they can recommend the right tool.

### Q: I don't have Python and don't want to install it. Can I still use the toolkit?

**A:** Absolutely. The 79 VBA tools cover most common scenarios. Python is optional and only needed for advanced tasks (50+ file processing, fuzzy matching, PDF extraction). In the future, we plan to convert the Python scripts to standalone .exe files so you can run them without installing Python.

---

## Document Information

| Field | Value |
|---|---|
| **Document Title** | Universal Toolkit Guide |
| **Version** | 1.0 |
| **Last Updated** | March 5, 2026 |
| **Author** | Finance Automation Team |
| **Audience** | All iPipeline Employees |
| **VBA Tools** | 79+ (13 modules) |
| **Python Scripts** | 22+ |
| **Total Tools** | 100+ |

---

*This document is part of the iPipeline P&L Automation Toolkit documentation suite. For the P&L-specific Command Center, see "How to Use the Command Center." For initial setup, see "Getting Started — First Time Setup Guide."*
