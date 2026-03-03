# KBT Universal Tools — Complete How-To Guide

**For:** Finance & Accounting Staff at iPipeline
**Written for:** Non-technical users — every step is spelled out
**Last Updated:** 2026-03-03

---

## Table of Contents

1. [How to Install the VBA Tools](#how-to-install-the-vba-tools)
2. [How to Run Any VBA Tool](#how-to-run-any-vba-tool)
3. [How to Install the Python Tools](#how-to-install-the-python-tools)
4. [How to Run Any Python Tool](#how-to-run-any-python-tool)
5. [VBA Tool Reference — Data Cleaning (12 tools)](#vba-data-cleaning-module)
6. [VBA Tool Reference — Formatting (9 tools)](#vba-formatting-module)
7. [VBA Tool Reference — Workbook Management (15 tools)](#vba-workbook-management-module)
8. [VBA Tool Reference — Finance (14 tools)](#vba-finance-module)
9. [VBA Tool Reference — Audit (8 tools)](#vba-audit-module)
10. [Python Tool Reference (18 tools)](#python-tools)

---

## How to Install the VBA Tools

These 58 VBA tools live in 5 module files (.bas). You install them once and they work on any Excel file you open.

### Option A — Install into Your Personal Macro Workbook (Recommended)

This makes the tools available in every Excel file you open, forever.

1. Open Excel (any file, or a blank workbook)
2. Press **Alt + F11** to open the VBA Editor
3. In the left panel, look for **VBAProject (PERSONAL.XLSB)**
   - If you do NOT see PERSONAL.XLSB, you need to create it first:
     - Close the VBA Editor (click X or press Alt+F11 again)
     - In Excel, go to **View > Macros > Record Macro**
     - In the "Store macro in" dropdown, choose **Personal Macro Workbook**
     - Click **OK**, then immediately click **View > Macros > Stop Recording**
     - Press **Alt + F11** again — you should now see PERSONAL.XLSB
4. Right-click on **VBAProject (PERSONAL.XLSB)** in the left panel
5. Click **Import File...**
6. Navigate to the folder where the .bas files are saved
7. Select **modUTL_DataCleaning.bas** and click **Open**
8. Repeat steps 4-7 for each of the other 4 files:
   - modUTL_Formatting.bas
   - modUTL_WorkbookMgmt.bas
   - modUTL_Finance.bas
   - modUTL_Audit.bas
9. Press **Ctrl + S** to save
10. Close the VBA Editor (Alt + F11)
11. You're done — the tools are now available in every workbook

### Option B — Install into a Specific Workbook

This makes the tools available only in that one file.

1. Open the Excel file you want the tools in
2. Press **Alt + F11** to open the VBA Editor
3. Right-click on **VBAProject (YourFileName.xlsm)** in the left panel
4. Click **Import File...**
5. Import all 5 .bas files (same steps as Option A, steps 6-8)
6. Press **Ctrl + S** to save
7. If the file was .xlsx, you will be asked to save as .xlsm — click **Yes**

---

## How to Run Any VBA Tool

After installing, here is how to run any tool:

1. Open the Excel file you want to work on
2. Press **Alt + F8** (this opens the Macro dialog box)
3. You will see a list of all available macros
4. Type part of the tool name to filter (e.g., type "Duplicate" to find DuplicateInvoiceDetector)
5. Click the tool name to select it
6. Click **Run**
7. Follow any prompts that appear (the tool will ask you questions like "Which column has the Amount?")
8. When the tool finishes, it will show a message with results

**Important notes:**
- Always save your file before running any tool (Ctrl + S)
- Most tools create results on a new sheet or highlight cells — they do not delete your data
- If a tool asks "Which column?" it wants the column letter (A, B, C, etc.)
- You can press Ctrl + Z to undo most changes if something looks wrong

---

## How to Install the Python Tools

### Step 1 — Install Python (one time only)

1. Go to python.org
2. Click **Downloads** and download the latest version for Windows
3. Run the installer
4. **IMPORTANT:** Check the box that says **"Add Python to PATH"** at the bottom of the installer
5. Click **Install Now**
6. Wait for it to finish, then click **Close**

### Step 2 — Install Required Packages (one time only)

1. Open **Command Prompt** (press the Windows key, type "cmd", press Enter)
2. Type this command and press Enter:
   ```
   pip install -r "C:\path\to\UniversalToolsForAllFiles\python\requirements.txt"
   ```
   (Replace the path with the actual folder location on your computer)
3. Wait for all packages to download and install
4. You should see "Successfully installed" messages — you're done

### Step 3 — Verify It Worked

1. In Command Prompt, type: `python --version`
2. You should see something like "Python 3.12.0" — that means Python is installed correctly

---

## How to Run Any Python Tool

Every Python tool follows the same pattern:

1. Open **Command Prompt** (press the Windows key, type "cmd", press Enter)
2. Type the command shown in each tool's section below
3. Replace the file paths with your actual file locations
4. Press **Enter**
5. The tool will print progress to the screen
6. When done, it will tell you where the output file was saved
7. Open the output file in Excel

**Tip:** You can drag and drop a file from File Explorer into the Command Prompt window to paste its full path instead of typing it out.

---

# VBA Tool Reference

---

## VBA Data Cleaning Module
**File:** modUTL_DataCleaning.bas | **Tools:** 12

---

### 1. UnmergeAndFillDown

**What it does:** Finds all merged cells in your selection, unmerges them, and fills the blank cells with the value that was in the merged cell. Fixes the #1 problem when working with data copied from reports.

**When to use it:** You received a report where someone merged cells (e.g., a department name spans 5 rows). You need each row to have its own value so you can sort, filter, or use VLOOKUP.

**How to run it:**
1. Select the range of cells that contains merged cells
2. Press Alt + F8
3. Type "UnmergeAndFillDown" and click Run
4. The tool will unmerge everything and fill blanks with the value above

**Example:** Before: "Marketing" is merged across rows 2-6. After: Each row 2-6 has "Marketing" in its own cell.

---

### 2. FillBlanksDown

**What it does:** Fills every blank cell in your selection with the value from the cell directly above it. Does NOT touch merged cells — just regular blank cells.

**When to use it:** You have a column where only the first row of each group has a value (like a GL account number) and the rest are blank. You need every row filled in.

**How to run it:**
1. Select the column or range with blanks you want filled
2. Press Alt + F8
3. Type "FillBlanksDown" and click Run
4. Every blank cell now has the value from the cell above it

---

### 3. ConvertTextToNumbers

**What it does:** Finds cells that look like numbers but are stored as text (you can tell because they have a green triangle in the corner). Converts them to actual numbers so formulas like SUM work correctly.

**When to use it:** You imported data from a CSV or another system and your SUM formulas return 0 even though you can see numbers in the cells.

**How to run it:**
1. Select the range that has text-stored numbers
2. Press Alt + F8
3. Type "ConvertTextToNumbers" and click Run
4. The tool tells you how many cells were converted

---

### 4. RemoveLeadingTrailingSpaces

**What it does:** Strips invisible spaces from the beginning and end of every text cell in your selection. Also removes extra spaces between words (turns "John   Smith" into "John Smith").

**When to use it:** VLOOKUPs are failing because the lookup value has a hidden space. Or your data has inconsistent spacing from a system export.

**How to run it:**
1. Select the range to clean
2. Press Alt + F8
3. Type "RemoveLeadingTrailingSpaces" and click Run

---

### 5. DeleteBlankRows

**What it does:** Deletes every completely blank row on the active sheet. Only removes rows where every single cell is empty — will not delete a row that has even one value in it.

**When to use it:** Your data has random blank rows scattered throughout from a copy-paste or system export.

**How to run it:**
1. Make sure the sheet you want to clean is the active sheet
2. Press Alt + F8
3. Type "DeleteBlankRows" and click Run
4. The tool tells you how many blank rows were removed

---

### 6. ReplaceErrorValues

**What it does:** Finds every cell with an error (#N/A, #REF!, #VALUE!, #DIV/0!, etc.) and replaces it with a value you choose (0, blank, or custom text).

**When to use it:** You need to clean up a report before sharing it. Error values look unprofessional and can break downstream calculations.

**How to run it:**
1. Make sure the sheet you want to clean is active
2. Press Alt + F8
3. Type "ReplaceErrorValues" and click Run
4. A box will ask: "Replace errors with what?" — type your replacement (e.g., 0) and click OK
5. The tool tells you how many errors were replaced

---

### 7. HighlightDuplicateRows

**What it does:** Checks every row in the selected column for duplicate values. Highlights duplicates in orange so you can see them instantly.

**When to use it:** You want to quickly see which invoice numbers, vendor names, or account codes appear more than once.

**How to run it:**
1. Select the column you want to check for duplicates
2. Press Alt + F8
3. Type "HighlightDuplicateRows" and click Run
4. Duplicate values are highlighted orange

---

### 8. RemoveDuplicateRows

**What it does:** Removes entire rows where the value in the selected column is a duplicate. Keeps the first occurrence and deletes the rest.

**When to use it:** You need to de-duplicate a list — for example, a vendor list that has the same vendor entered multiple times.

**How to run it:**
1. Select the column to check for duplicates
2. Press Alt + F8
3. Type "RemoveDuplicateRows" and click Run
4. The tool tells you how many duplicate rows were removed

**Important:** This permanently deletes rows. Save your file first (Ctrl + S).

---

### 9. MultiReplaceDataCleaner

**What it does:** Performs multiple find-and-replace operations at once using a mapping table. You set up a sheet with "Find" values in column A and "Replace" values in column B, and the tool runs all replacements across your data.

**When to use it:** You need to standardize dozens of variations (e.g., "NY", "New York", "N.Y." all become "New York"). Instead of doing 30 separate find-and-replace operations, you do it in one click.

**How to run it:**
1. Create a new sheet in your workbook
2. In column A, put the values you want to find (starting in row 1)
3. In column B, put the replacement values
4. Go back to the sheet with your data
5. Press Alt + F8
6. Type "MultiReplaceDataCleaner" and click Run
7. A box will ask for the name of your mapping sheet — type it and click OK
8. The tool tells you how many replacements were made

---

### 10. FormulaToValueHardcoder

**What it does:** Converts all formulas in your selection to their current values. The numbers stay the same, but they are no longer formulas — they are plain values.

**When to use it:** You need to share a file and don't want formulas that reference other sheets or workbooks to break. Or you want to "lock in" calculated values before making changes.

**How to run it:**
1. Select the range of cells you want to convert
2. Press Alt + F8
3. Type "FormulaToValueHardcoder" and click Run
4. All formulas in the selection are now values

**Important:** This cannot be undone after saving. Save a backup first.

---

### 11. PhantomHyperlinkPurger

**What it does:** Removes all hyperlinks from the active sheet. The text stays — only the clickable link is removed.

**When to use it:** You pasted data from a website or email and every cell has a blue underlined hyperlink. You just want the text without the links.

**How to run it:**
1. Make sure the correct sheet is active
2. Press Alt + F8
3. Type "PhantomHyperlinkPurger" and click Run

---

### 12. ConvertNumbersToWords

**What it does:** Takes a number in the selected cell and converts it to written English words. For example: 1,234.56 becomes "One Thousand Two Hundred Thirty-Four Dollars and 56/100".

**When to use it:** You need to write out a dollar amount in words for a check, contract, or legal document.

**How to run it:**
1. Select the cell containing the number you want to convert
2. Press Alt + F8
3. Type "ConvertNumbersToWords" and click Run
4. The word version appears in the cell to the right

---

## VBA Formatting Module
**File:** modUTL_Formatting.bas | **Tools:** 9

---

### 13. AutoFitAllColumnsRows

**What it does:** Automatically resizes every column and row on every sheet in the workbook so all content is visible. No more "####" in cells.

**When to use it:** You just received a file and half the columns are too narrow to read. Instead of manually dragging 50 columns, run this.

**How to run it:**
1. Press Alt + F8
2. Type "AutoFitAllColumnsRows" and click Run
3. Every sheet in the workbook is now auto-fitted

---

### 14. FreezeTopRowAllSheets

**What it does:** Freezes the top row (header row) on every sheet so it stays visible when you scroll down.

**When to use it:** You have a workbook with 20 sheets and want the header row frozen on all of them without doing it one by one.

**How to run it:**
1. Press Alt + F8
2. Type "FreezeTopRowAllSheets" and click Run

---

### 15. NumberFormatStandardizer

**What it does:** Applies a clean number format (#,##0.00) to all numeric cells in your selection. Skips text, dates, and cells with formulas.

**When to use it:** Your data has inconsistent number formatting — some cells show 2 decimals, some show 6, some have no commas.

**How to run it:**
1. Select the range to format
2. Press Alt + F8
3. Type "NumberFormatStandardizer" and click Run

---

### 16. CurrencyFormatStandardizer

**What it does:** Applies standard currency format ($#,##0.00) to all numeric cells in your selection.

**When to use it:** You need dollar signs and consistent decimal places across a financial report.

**How to run it:**
1. Select the range to format
2. Press Alt + F8
3. Type "CurrencyFormatStandardizer" and click Run

---

### 17. DateFormatStandardizer

**What it does:** Converts all date cells in your selection to a standard format (MM/DD/YYYY).

**When to use it:** Your dates are all over the place — some say "3/1/26", some say "March 1, 2026", some say "2026-03-01". This makes them all consistent.

**How to run it:**
1. Select the range containing dates
2. Press Alt + F8
3. Type "DateFormatStandardizer" and click Run

---

### 18. HighlightNegativesRed

**What it does:** Applies conditional formatting to turn all negative numbers red in your selection.

**When to use it:** You want negative variances, losses, or credits to stand out visually in a financial report.

**How to run it:**
1. Select the range to format
2. Press Alt + F8
3. Type "HighlightNegativesRed" and click Run

---

### 19. FinancialNumberFormattingSuite

**What it does:** Applies professional financial formatting to the entire active sheet: thousands separators, negative numbers in red parentheses, header row bold with dark background. Makes any data look CFO-ready.

**When to use it:** You need to present a sheet to leadership and want it formatted professionally in one click.

**How to run it:**
1. Make sure the sheet you want to format is active
2. Press Alt + F8
3. Type "FinancialNumberFormattingSuite" and click Run

---

### 20. ConditionalFormatPurger

**What it does:** Removes ALL conditional formatting rules from every sheet in the workbook. The cell values stay — only the conditional formatting rules are deleted.

**When to use it:** Your file is slow because someone added hundreds of conditional formatting rules over time. Or you want a clean slate before applying new formatting.

**How to run it:**
1. Press Alt + F8
2. Type "ConditionalFormatPurger" and click Run
3. Confirm when asked

**Important:** This removes ALL conditional formatting on ALL sheets. Cannot be undone after saving.

---

### 21. PrintHeaderFooterStandardizer

**What it does:** Sets professional print headers and footers on every sheet: file name in the header, page numbers and date in the footer. Also sets landscape orientation and fit-to-page.

**When to use it:** You need to print a workbook and want every page to look professional with consistent headers and footers.

**How to run it:**
1. Press Alt + F8
2. Type "PrintHeaderFooterStandardizer" and click Run

---

## VBA Workbook Management Module
**File:** modUTL_WorkbookMgmt.bas | **Tools:** 15

---

### 22. UnhideAllSheetsRowsColumns

**What it does:** Makes every hidden sheet, hidden row, and hidden column visible across the entire workbook. One click to reveal everything.

**When to use it:** You received a file from someone and suspect there are hidden sheets or rows you cannot see.

**How to run it:**
1. Press Alt + F8
2. Type "UnhideAllSheetsRowsColumns" and click Run
3. The tool tells you how many hidden sheets were revealed

---

### 23. ExportAllSheetsCombinedPDF

**What it does:** Exports all visible sheets into a single multi-page PDF file.

**When to use it:** You need to send the entire workbook as a PDF to someone who does not have Excel.

**How to run it:**
1. Press Alt + F8
2. Type "ExportAllSheetsCombinedPDF" and click Run
3. Choose where to save the PDF and what to name it
4. The PDF opens automatically after saving

---

### 24. FindReplaceAcrossAllSheets

**What it does:** Performs a find-and-replace across every sheet in the workbook at once.

**When to use it:** You need to change "FY2025" to "FY2026" everywhere in the workbook, not just on one sheet.

**How to run it:**
1. Press Alt + F8
2. Type "FindReplaceAcrossAllSheets" and click Run
3. Enter the text to find, then the replacement text
4. Confirm to proceed
5. The tool tells you how many replacements were made

---

### 25. SearchAcrossAllSheets

**What it does:** Searches every sheet for a value and creates a clickable results sheet. Click any result to jump directly to that cell.

**When to use it:** You need to find where a specific account number, vendor, or value appears in a large workbook.

**How to run it:**
1. Press Alt + F8
2. Type "SearchAcrossAllSheets" and click Run
3. Type the value to search for
4. A new sheet appears with all results — click the cell address links to navigate

---

### 26. MultiSheetBatchRenamer

**What it does:** Finds and replaces text in sheet tab names across all tabs at once.

**When to use it:** You have tabs named "Jan 2025", "Feb 2025", etc. and need to change them all to "Jan 2026", "Feb 2026".

**How to run it:**
1. Press Alt + F8
2. Type "MultiSheetBatchRenamer" and click Run
3. Enter the text to find in tab names (e.g., "2025")
4. Enter the replacement text (e.g., "2026")

---

### 27. SortWorksheetsAlphabetically

**What it does:** Rearranges all sheet tabs in alphabetical order (A to Z).

**When to use it:** Your workbook has 30+ tabs and they're in random order. This sorts them instantly.

**How to run it:**
1. Press Alt + F8
2. Type "SortWorksheetsAlphabetically" and click Run
3. Confirm when asked

---

### 28. CreateTableOfContents

**What it does:** Creates a new sheet at the front of the workbook with a clickable list of every sheet. Click any sheet name to jump directly to it.

**When to use it:** Your workbook has many tabs and you want an easy way to navigate.

**How to run it:**
1. Press Alt + F8
2. Type "CreateTableOfContents" and click Run
3. A new "Table of Contents" sheet appears at the front

---

### 29. ProtectAllSheets

**What it does:** Applies worksheet protection to every sheet in the workbook at once, with a password you choose.

**When to use it:** You are sharing a workbook and want to prevent people from accidentally changing formulas or layouts.

**How to run it:**
1. Press Alt + F8
2. Type "ProtectAllSheets" and click Run
3. Enter a password (or leave blank for no-password protection)
4. Confirm to apply

---

### 30. UnprotectAllSheets

**What it does:** Removes worksheet protection from every sheet at once.

**When to use it:** You need to edit a protected workbook and know the password.

**How to run it:**
1. Press Alt + F8
2. Type "UnprotectAllSheets" and click Run
3. Enter the password

---

### 31. LockAllFormulaCells

**What it does:** Locks cells that contain formulas and leaves input cells (values) editable. Run this before protecting the sheet to let people enter data without breaking formulas.

**When to use it:** You built a model with formulas and want coworkers to fill in numbers without accidentally overwriting calculations.

**How to run it:**
1. Go to the sheet you want to lock
2. Press Alt + F8
3. Type "LockAllFormulaCells" and click Run
4. Then protect the sheet (Review tab > Protect Sheet) to activate the lock

---

### 32. ExportActiveSheetPDF

**What it does:** Exports just the current sheet as a PDF file.

**When to use it:** You want to send one specific sheet as a PDF, not the whole workbook.

**How to run it:**
1. Go to the sheet you want to export
2. Press Alt + F8
3. Type "ExportActiveSheetPDF" and click Run
4. Choose where to save and what to name it

---

### 33. ExportAllSheetsIndividualPDFs

**What it does:** Exports every visible sheet as its own separate PDF file into a folder you choose.

**When to use it:** You need each sheet as a separate PDF — for example, one per department or one per month.

**How to run it:**
1. Press Alt + F8
2. Type "ExportAllSheetsIndividualPDFs" and click Run
3. Choose the folder where PDFs should be saved
4. Each sheet becomes its own PDF file

---

### 34. ResetAllFilters

**What it does:** Clears all AutoFilter criteria on every sheet so all data is visible again.

**When to use it:** Someone left filters applied on multiple sheets and you can't see all the data.

**How to run it:**
1. Press Alt + F8
2. Type "ResetAllFilters" and click Run

---

### 35. BuildDistributionReadyCopy

**What it does:** Creates a clean copy of the workbook with all formulas converted to values, all hidden sheets visible, and metadata stripped. Your original file is not touched.

**When to use it:** You need to share a file externally (with a client, auditor, or partner) and don't want formulas, hidden sheets, or internal references visible.

**How to run it:**
1. Press Alt + F8
2. Type "BuildDistributionReadyCopy" and click Run
3. Confirm when asked
4. A new file with "_DIST" in the name is saved in the same folder

---

### 36. WorkbookHealthCheck

**What it does:** Generates a full diagnostic report on the workbook: total sheets, total cells, formula cells, error cells, hyperlinks, and external links. Flags anything that needs attention.

**When to use it:** You want a quick overview of a workbook's health before working on it or sharing it.

**How to run it:**
1. Press Alt + F8
2. Type "WorkbookHealthCheck" and click Run
3. A summary report appears in a message box

---

## VBA Finance Module
**File:** modUTL_Finance.bas | **Tools:** 14

---

### 37. DuplicateInvoiceDetector

**What it does:** Scans for potential duplicate invoices by matching on Vendor + Amount + Date (within 3 days) or Invoice Number. Highlights suspected duplicates in orange.

**When to use it:** Before running a payment batch, check for duplicate invoices that could cause double-payments.

**How to run it:**
1. Go to the sheet with your invoice data
2. Press Alt + F8
3. Type "DuplicateInvoiceDetector" and click Run
4. Enter the column letters when asked (Vendor, Amount, Date, Invoice#)
5. Suspected duplicates are highlighted orange

---

### 38. AutoBalancingGLValidator

**What it does:** Sums the Debit and Credit columns and checks if they balance. If they don't, it can insert a plug entry for you to investigate.

**When to use it:** You are reviewing a GL extract or journal entry and need to confirm debits equal credits.

**How to run it:**
1. Go to the sheet with your GL data
2. Press Alt + F8
3. Type "AutoBalancingGLValidator" and click Run
4. Enter the Debit and Credit column letters
5. If balanced, you get a green confirmation. If not, you can choose to insert a plug entry.

---

### 39. TrialBalanceChecker

**What it does:** Adds up all debit balances and all credit balances and reports whether the trial balance is in balance.

**When to use it:** Quick check at month-end to confirm your trial balance ties.

**How to run it:**
1. Go to your trial balance sheet
2. Press Alt + F8
3. Type "TrialBalanceChecker" and click Run
4. Enter the Debit and Credit column letters

---

### 40. JournalEntryValidator

**What it does:** Groups all rows by journal entry number, sums debits and credits within each JE, and reports how many JEs are out of balance.

**When to use it:** You have a batch of journal entries and need to verify each one balances individually before posting.

**How to run it:**
1. Go to the sheet with journal entries
2. Press Alt + F8
3. Type "JournalEntryValidator" and click Run
4. Enter the column letters for JE Number, Debit, and Credit

---

### 41. FluxAnalysis

**What it does:** Compares two columns (Current Period vs Prior Period) row by row. Adds $ Variance and % Variance columns. Highlights any row where the change exceeds a threshold you set.

**When to use it:** Month-end flux analysis — quickly see which P&L or BS lines had significant changes.

**How to run it:**
1. Go to the sheet with your period data
2. Press Alt + F8
3. Type "FluxAnalysis" and click Run
4. Enter the Current Period column, Prior Period column, and threshold %
5. Rows exceeding the threshold are highlighted yellow

---

### 42. APAgingSummaryGenerator

**What it does:** Buckets AP invoices by days overdue (Current, 0-30, 31-60, 61-90, 90+) and creates a summary sheet with totals per bucket.

**When to use it:** You need an AP aging report from raw invoice data.

**How to run it:**
1. Go to the sheet with your AP invoices
2. Press Alt + F8
3. Type "APAgingSummaryGenerator" and click Run
4. Enter the column letters for Due Date, Amount, and Vendor

---

### 43. ARAgingSummaryGenerator

**What it does:** Same as the AP version but for Accounts Receivable. Buckets by days since invoice date.

**How to run it:** Same as APAgingSummaryGenerator — just enter your AR column letters.

---

### 44. AgingBucketCalculator

**What it does:** Adds a new column to your data that labels each row with its aging bucket (Current, 0-30 Days, 31-60 Days, etc.) based on a date column you specify.

**When to use it:** You want to add aging buckets to your data without creating a separate summary — useful for pivot tables.

**How to run it:**
1. Go to your data sheet
2. Press Alt + F8
3. Type "AgingBucketCalculator" and click Run
4. Enter the date column letter

---

### 45. VarianceAnalysisTemplate

**What it does:** Inserts two new columns ($ Variance and % Variance) next to your Actual and Budget columns. Calculates the difference for every row.

**When to use it:** You have Actual and Budget columns side by side and want variance columns added automatically.

**How to run it:**
1. Go to your data sheet
2. Press Alt + F8
3. Type "VarianceAnalysisTemplate" and click Run
4. Enter the Actual and Budget column letters

---

### 46. QuickCorkscrewBuilder

**What it does:** Creates a standard roll-forward schedule on a new sheet: Beginning Balance + Additions - Deductions +/- Adjustments = Ending Balance. Pre-formatted with formulas.

**When to use it:** You need to build a roll-forward for fixed assets, reserves, deferred revenue, or any balance sheet account.

**How to run it:**
1. Press Alt + F8
2. Type "QuickCorkscrewBuilder" and click Run
3. Enter the name (e.g., "Fixed Assets") and beginning balance
4. A new formatted sheet is created with the schedule

---

### 47. FinancialPeriodRollForward

**What it does:** Updates period header labels (e.g., "March 2026" to "April 2026") across a row and optionally clears all input cells below to prepare for new period data.

**When to use it:** Monthly close — roll your model forward to the next period.

**How to run it:**
1. Go to your model sheet
2. Press Alt + F8
3. Type "FinancialPeriodRollForward" and click Run
4. Enter the new period label, header row number, and old period label to replace

---

### 48. MultiCurrencyConsolidationAggregator

**What it does:** Converts foreign currency amounts to USD using an FX rate table in your workbook. Adds a "USD Equivalent" column to your data.

**When to use it:** You are consolidating data from multiple countries and need everything in USD.

**How to run it:**
1. Create a sheet in your workbook with Currency Code in column A and Rate to USD in column B (e.g., EUR = 1.08, GBP = 1.27)
2. Go to the sheet with your foreign currency data
3. Press Alt + F8
4. Type "MultiCurrencyConsolidationAggregator" and click Run
5. Enter the FX rate sheet name, currency column, and amount column

---

### 49. RatioAnalysisDashboard

**What it does:** Scans your data for standard financial labels (Revenue, Net Income, Total Assets, etc.) and calculates key ratios: Gross Margin, Net Profit Margin, Current Ratio, ROE, ROA, and EBITDA Margin.

**When to use it:** You have a financial statement and want a quick ratio dashboard without manual calculation.

**How to run it:**
1. Go to the sheet with your financial data (labels in column A, values in column B)
2. Press Alt + F8
3. Type "RatioAnalysisDashboard" and click Run
4. A new "UTL Ratios" sheet is created with calculated ratios

---

### 50. GeneralLedgerJournalMapper

**What it does:** Transforms a trial balance or GL extract into a journal entry upload format: Date | Account | Description | Debit | Credit. Positive balances go to Debit, negative to Credit.

**When to use it:** You need to create a journal entry upload file from a trial balance.

**How to run it:**
1. Go to the sheet with your trial balance
2. Press Alt + F8
3. Type "GeneralLedgerJournalMapper" and click Run
4. Enter the Account, Description, and Balance column letters, plus the JE date

---

## VBA Audit Module
**File:** modUTL_Audit.bas | **Tools:** 8

---

### 51. ExternalLinkFinder

**What it does:** Scans every cell in the workbook and lists all cells that reference external files (formulas with [brackets]). Creates a report sheet with the file path and cell address of each link.

**When to use it:** You want to know if your workbook depends on other files that might break if moved or deleted.

**How to run it:**
1. Press Alt + F8
2. Type "ExternalLinkFinder" and click Run
3. If links are found, a report sheet is created. If none, you get a clean confirmation.

---

### 52. CircularReferenceDetector

**What it does:** Checks every sheet for circular references (a formula that refers to itself, directly or indirectly) and reports their locations.

**When to use it:** Excel is showing a circular reference warning and you need to find exactly which cell(s) are causing it.

**How to run it:**
1. Press Alt + F8
2. Type "CircularReferenceDetector" and click Run

---

### 53. WorkbookErrorScanner

**What it does:** Finds every cell with an error value (#N/A, #REF!, #VALUE!, #DIV/0!, etc.) across all sheets. Creates a report sheet listing each error's location, type, and formula.

**When to use it:** You need to clean up a workbook before sharing it and want to find every error quickly.

**How to run it:**
1. Press Alt + F8
2. Type "WorkbookErrorScanner" and click Run
3. A report sheet is created listing all errors (or you get a clean confirmation)

---

### 54. DataQualityScorecard

**What it does:** Analyzes every column on the active sheet and creates a scorecard showing: total rows, blank count, error count, duplicate count, and data type breakdown (numeric, text, date) for each column.

**When to use it:** You received a new dataset and want a quick quality assessment before working with it.

**How to run it:**
1. Go to the sheet you want to analyze
2. Press Alt + F8
3. Type "DataQualityScorecard" and click Run
4. A new "UTL Data Quality" sheet is created with the scorecard

---

### 55. NamedRangeAuditor

**What it does:** Lists every Named Range in the workbook and checks if each one still points to a valid location. Flags broken references in red.

**When to use it:** Your workbook has old named ranges that might be broken (#REF!) and you need to clean them up.

**How to run it:**
1. Press Alt + F8
2. Type "NamedRangeAuditor" and click Run
3. A report sheet is created with all named ranges and their status

---

### 56. DataValidationChecker

**What it does:** Finds all cells with data validation rules (dropdowns, number restrictions, etc.) across every sheet. Flags any dropdown whose source list is broken.

**When to use it:** Your dropdowns stopped working after someone edited the workbook. This finds which ones are broken.

**How to run it:**
1. Press Alt + F8
2. Type "DataValidationChecker" and click Run
3. A report sheet is created listing all validation rules and their status

---

### 57. InconsistentFormulasAuditor

**What it does:** Checks a column of formulas and flags any cell where the formula pattern differs from the majority. Also catches hardcoded values hiding in formula columns.

**When to use it:** You suspect someone manually overrode a formula in a column and you want to find it. Or you want to verify an entire column uses the same formula consistently.

**How to run it:**
1. Select the column range to audit (e.g., select D2:D100)
2. Press Alt + F8
3. Type "InconsistentFormulasAuditor" and click Run
4. Orange = different formula. Red = hardcoded value in a formula column.

---

### 58. ExternalLinkSeveranceProtocol

**What it does:** Replaces every external link formula with its current value. The original formula is saved as a cell comment so you have a record of what was there.

**When to use it:** You are finalizing a file and want to break all external links permanently so the file is self-contained.

**How to run it:**
1. Press Alt + F8
2. Type "ExternalLinkSeveranceProtocol" and click Run
3. Confirm when asked
4. Affected cells are highlighted yellow. Hover over them to see the original formula in the comment.

**Important:** This changes formulas to values permanently. Save a backup first.

---

# Python Tools

**Location:** UniversalToolsForAllFiles/python/
**Requirement:** Python must be installed (see installation section above)

---

### 59. Universal Data Cleaner (clean_data.py)

**What it does:** Cleans any Excel file in one command: removes empty rows/columns, trims spaces, converts text-stored numbers, standardizes dates, and removes duplicates.

**How to run it:**
```
python clean_data.py "C:\path\to\your_file.xlsx"
```

**Options:**
- Clean only one sheet: `python clean_data.py "file.xlsx" --sheet "Sheet1"`
- Skip duplicate removal: `python clean_data.py "file.xlsx" --no-dedupe`

**Output:** Saves a cleaned copy as "your_file_CLEANED.xlsx" in the same folder.

---

### 60. Excel File Comparison (compare_files.py)

**What it does:** Compares two Excel files cell by cell. Produces a color-coded diff report showing every difference: added, removed, or changed values.

**How to run it:**
```
python compare_files.py "C:\path\file1.xlsx" "C:\path\file2.xlsx"
```

**Options:**
- Compare only one sheet: `python compare_files.py "file1.xlsx" "file2.xlsx" --sheet "Sheet1"`

**Output:** Saves "COMPARISON_REPORT.xlsx" — Green = Added, Red = Removed, Yellow = Changed.

---

### 61. Budget vs Actual Consolidator (consolidate_budget.py)

**What it does:** Merges multiple department budget files from a folder into one master file with $ and % variance columns.

**How to run it:**
```
python consolidate_budget.py "C:\path\to\folder"
```

**Options:**
- Custom column names: `python consolidate_budget.py "folder" --actual "Actuals" --budget "Plan"`
- Specific sheet: `python consolidate_budget.py "folder" --sheet "P&L"`

**Output:** Saves "CONSOLIDATED_BUDGET_VS_ACTUAL.xlsx" in the folder.

---

### 62. AR/AP Aging Report (aging_report.py)

**What it does:** Generates a full aging report with buckets (Current, 0-30, 31-60, 61-90, 90+ days) from any Excel file with dates and amounts.

**How to run it:**
```
python aging_report.py "C:\path\to\invoices.xlsx"
```

**Options:**
- Custom columns: `python aging_report.py "file.xlsx" --date "Due Date" --amount "Balance" --name "Vendor"`
- AR instead of AP: `python aging_report.py "file.xlsx" --type AR`

**Output:** Saves "AP_AGING_REPORT.xlsx" or "AR_AGING_REPORT.xlsx" with detail, bucket summary, and by-vendor breakdown.

---

### 63. Multi-File Data Consolidator (consolidate_files.py)

**What it does:** Combines data from dozens or hundreds of Excel files in a folder into one master sheet with a "Source_File" column added.

**How to run it:**
```
python consolidate_files.py "C:\path\to\folder"
```

**Options:**
- Specific sheet: `python consolidate_files.py "folder" --sheet "Data"`
- File pattern: `python consolidate_files.py "folder" --pattern "*Q1*"`

**Output:** Saves "MASTER_CONSOLIDATED.xlsx" in the folder.

---

### 64. Variance Analysis Generator (variance_analysis.py)

**What it does:** Compares Actual vs Budget columns across files in a folder. Creates a summary with bar chart, $ and % variance, and Favorable/Unfavorable labels.

**How to run it:**
```
python variance_analysis.py "C:\path\to\folder"
```

**Options:**
- Custom columns: `python variance_analysis.py "folder" --actual "Actuals" --budget "Plan" --label "Department"`

**Output:** Saves "VARIANCE_ANALYSIS.xlsx" with Detail, Summary, and chart.

---

### 65. GL Reconciliation Engine (gl_reconciliation.py)

**What it does:** Matches two transaction lists (GL vs Sub-ledger) using Amount + Date (within 3 days) + optional reference number. Flags unmatched items.

**How to run it:**
```
python gl_reconciliation.py "C:\path\gl.xlsx" "C:\path\subledger.xlsx"
```

**Options:**
- Custom columns: `python gl_reconciliation.py "gl.xlsx" "sub.xlsx" --amount "Amt" --date "Post Date" --ref "Invoice"`
- Adjust date tolerance: `python gl_reconciliation.py "gl.xlsx" "sub.xlsx" --tolerance 5`

**Output:** Saves "GL_RECONCILIATION_REPORT.xlsx" with matched pairs and unmatched items.

---

### 66. Fuzzy Match / Fuzzy Lookup (fuzzy_lookup.py)

**What it does:** Matches records between two datasets using fuzzy string matching. Catches typos and variations (e.g., "ACME Corp" vs "Acme Corporation").

**How to run it:**
```
python fuzzy_lookup.py "C:\data.xlsx" "C:\master.xlsx" --lookup-col "Vendor" --match-col "Vendor Name"
```

**Options:**
- Adjust sensitivity: `python fuzzy_lookup.py "data.xlsx" "master.xlsx" --threshold 80`
- Pull extra columns: `python fuzzy_lookup.py "data.xlsx" "master.xlsx" --return-cols "Category" "Region"`

**Output:** Saves "FUZZY_MATCH_RESULTS.xlsx" — Green = Exact, Yellow = Fuzzy, Red = No Match.

---

### 67. Batch File Processor (batch_process.py)

**What it does:** Runs the data cleaner on every Excel file in a folder automatically. Saves cleaned copies and a processing log.

**How to run it:**
```
python batch_process.py "C:\path\to\folder"
```

**Options:**
- File pattern: `python batch_process.py "folder" --pattern "*Q1*"`
- Output to different folder: `python batch_process.py "folder" --output "C:\cleaned"`

**Output:** Saves "_CLEANED" copies of each file plus "BATCH_PROCESSING_LOG.xlsx".

---

### 68. Forecast Roll-Forward (forecast_rollforward.py)

**What it does:** Takes historical data and builds a 12-month rolling forecast using one of three methods: moving average, fixed growth rate, or flat (repeat last value).

**How to run it:**
```
python forecast_rollforward.py "C:\path\data.xlsx" --method moving_avg
```

**Options:**
- Growth method: `python forecast_rollforward.py "data.xlsx" --method growth --rate 0.05`
- Custom periods: `python forecast_rollforward.py "data.xlsx" --periods 6`
- Custom columns: `python forecast_rollforward.py "data.xlsx" --date "Month" --value-col "Revenue"`

**Output:** Saves "FORECAST_ROLLFORWARD.xlsx" with actuals + forecast and a line chart.

---

### 69. Data Unpivoter (unpivot_data.py)

**What it does:** Converts wide-format data (one column per month) into tall database format (one row per record). The opposite of a pivot table.

**How to run it:**
```
python unpivot_data.py "C:\path\data.xlsx"
```

**Options:**
- Specify ID columns: `python unpivot_data.py "data.xlsx" --id-cols "Name" "Department"`
- Custom names: `python unpivot_data.py "data.xlsx" --value-name "Amount" --var-name "Period"`

**Output:** Saves "UNPIVOTED.xlsx".

---

### 70. PDF Table Extractor (pdf_extractor.py)

**What it does:** Scans a PDF document and extracts any tables it finds into an Excel file. Each table gets its own sheet.

**How to run it:**
```
python pdf_extractor.py "C:\path\report.pdf"
```

**Options:**
- Specific pages: `python pdf_extractor.py "report.pdf" --pages "1-5"`

**Output:** Saves "PDF_EXTRACTED_TABLES.xlsx" with one sheet per table found.

**Note:** Works best on PDFs with structured grid tables. Scanned/image PDFs will not work — the text must be selectable in the PDF.

---

### 71. Regex Text Extractor (regex_extractor.py)

**What it does:** Extracts structured patterns from free-text columns — invoice numbers, email addresses, phone numbers, dates, dollar amounts, and more. No coding needed — just pick a preset.

**How to run it:**
```
python regex_extractor.py "C:\path\data.xlsx" --col "Description" --pattern invoice
```

**Available presets:** invoice, email, phone, date, currency, account, zipcode, ssn

**Options:**
- Custom pattern: `python regex_extractor.py "data.xlsx" --col "Notes" --custom "[A-Z]{2}-\d{4}"`

**Output:** Saves "REGEX_EXTRACTED.xlsx" — Green = Match Found, Red = No Match.

---

### 72. Word Report Generator (word_report.py)

**What it does:** Reads data from an Excel file and generates a professionally formatted Microsoft Word document with tables, headings, and footer.

**How to run it:**
```
python word_report.py "C:\path\data.xlsx"
```

**Options:**
- Custom title: `python word_report.py "data.xlsx" --title "Q1 Finance Report" --author "Connor"`
- Specific sheets: `python word_report.py "data.xlsx" --sheets "Summary" "Detail"`

**Output:** Saves "WORD_REPORT.docx" in the same folder.

---

### 73. Fuzzy-Match Bank Reconciler (bank_reconciler.py)

**What it does:** Matches ledger descriptions against bank statement text using fuzzy string matching. Catches near-matches that exact lookups miss (e.g., "AMAZON MKTP" vs "Amazon").

**How to run it:**
```
python bank_reconciler.py "C:\path\ledger.xlsx" "C:\path\bank.xlsx"
```

**Options:**
- Custom columns: `python bank_reconciler.py "ledger.xlsx" "bank.xlsx" --desc "Description" --amount "Debit" --bank-desc "Memo"`
- Adjust threshold: `python bank_reconciler.py "ledger.xlsx" "bank.xlsx" --threshold 70`

**Output:** Saves "BANK_RECONCILIATION.xlsx" — Green = Matched, Yellow = Fuzzy, Red = Unmatched.

---

### 74. Master Data Mapper (master_data_mapper.py)

**What it does:** Performs SQL-style joins between two Excel datasets on a common key column. Replaces complex nested VLOOKUP formulas with a simple command.

**How to run it:**
```
python master_data_mapper.py "C:\data.xlsx" "C:\master.xlsx" --key "Vendor ID"
```

**Options:**
- Join type: `python master_data_mapper.py "data.xlsx" "master.xlsx" --key "ID" --join inner`
- Specific columns: `python master_data_mapper.py "data.xlsx" "master.xlsx" --key "ID" --cols "Name" "Category"`

**Output:** Saves "MASTER_DATA_MAPPED.xlsx" — Green = Matched, Red = No master record found.

---

### 75. Reconciliation Exception Generator (reconciliation_exceptions.py)

**What it does:** Compares two datasets on one or more key columns and outputs ONLY the unmatched exceptions. Matched items are excluded — you only see what needs investigation.

**How to run it:**
```
python reconciliation_exceptions.py "C:\list1.xlsx" "C:\list2.xlsx" --key "Invoice No"
```

**Options:**
- Multiple keys: `python reconciliation_exceptions.py "list1.xlsx" "list2.xlsx" --key "Invoice No" "Amount"`

**Output:** Saves "RECONCILIATION_EXCEPTIONS.xlsx" — Red = in File 1 only, Yellow = in File 2 only.

---

### 76. Variance Decomposition Analyzer (variance_decomposition.py)

**What it does:** Breaks down financial variances into Price Effect, Volume Effect, and Mix Effect — standard FP&A analysis. Creates a bridge chart data sheet.

**How to run it:**
```
python variance_decomposition.py "C:\path\data.xlsx"
```

**Options:**
- Custom columns: `python variance_decomposition.py "data.xlsx" --product "Product" --act-vol "Act Units" --bud-vol "Bud Units" --act-price "Act Price" --bud-price "Bud Price"`

**Output:** Saves "VARIANCE_DECOMPOSITION.xlsx" — Green = Favorable, Red = Unfavorable.

---

## Quick Reference — Which Tool Do I Need?

| I need to... | Use this tool |
|---|---|
| Clean up messy data from a report | UnmergeAndFillDown, FillBlanksDown, RemoveLeadingTrailingSpaces |
| Fix numbers stored as text | ConvertTextToNumbers (VBA) or clean_data.py (Python) |
| Find and remove duplicates | HighlightDuplicateRows, RemoveDuplicateRows |
| Compare two files side by side | compare_files.py |
| Combine data from many files | consolidate_files.py or consolidate_budget.py |
| Create an aging report | APAgingSummaryGenerator, ARAgingSummaryGenerator, or aging_report.py |
| Do variance analysis | VarianceAnalysisTemplate, FluxAnalysis, or variance_analysis.py |
| Match data with fuzzy/approximate names | fuzzy_lookup.py or bank_reconciler.py |
| Find external links or errors | ExternalLinkFinder, WorkbookErrorScanner |
| Make a file look professional | FinancialNumberFormattingSuite, PrintHeaderFooterStandardizer |
| Export to PDF | ExportActiveSheetPDF, ExportAllSheetsCombinedPDF |
| Build a forecast | forecast_rollforward.py |
| Extract data from a PDF | pdf_extractor.py |
| Create a Word report from Excel | word_report.py |
| Reconcile GL vs sub-ledger | gl_reconciliation.py |
| Find exceptions between two lists | reconciliation_exceptions.py |
| Break down variance into Price/Volume/Mix | variance_decomposition.py |

---

## Need Help?

If a tool does not work as expected:
1. Make sure your file is saved before running the tool
2. Check that your column headers match what the tool is asking for
3. For Python tools, make sure Python is installed and packages are up to date (`pip install -r requirements.txt`)
4. Contact Connor for assistance
