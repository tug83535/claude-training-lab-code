Grok

**# Giant List of Universal Excel Automation Tools (VBA Macros + Python Scripts)**
These tools work on **any** Excel file (.xlsx, .xlsm, etc.) with no file-specific setup required.
**VBA Macros** → Store once in your **Personal Macro Workbook** (Personal.xlsb) so they appear in every Excel session. Run via Alt+F8, Quick Access Toolbar, or ribbon buttons.
**Python Scripts** → Run from command line / Jupyter / batch file; pass any file path as argument. Use pandas + openpyxl (install once: pip install pandas openpyxl).
All are generic, productivity-focused, and finance-team friendly (data cleaning, reporting, financial calcs, reconciliation, etc.). Nothing that overlaps with OneDrive/SharePoint auto-features.
### VBA Macros (Personal Macro Workbook – 70+ ready-to-use)
### Core Navigation & Workbook Management
* **Unhide All Worksheets** – Instantly unhides every hidden sheet in the active workbook.
* **Hide All Selected Sheets** – Hides multiple sheets you have selected at once.
* **Create Table of Contents** – Builds a new sheet with clickable hyperlinks to every worksheet.
* **Sort Worksheets Alphabetically** – Reorders all sheets A–Z (or Z–A).
* **List All Sheet Names** – Creates a new sheet listing every worksheet name with hyperlinks.
* **Backup Workbook with Timestamp** – Saves a copy named “FileName_YYYYMMDD_HHMM.xlsx”.
* **Save & Close All Open Workbooks** – Saves and closes every open file with one click.
* **Protect All Sheets** – Applies password protection to every sheet (or unprotects).
* **Insert Standard Header/Footer** – Adds company logo, page numbers, filename, date to all sheets.
### Data Cleaning & Formatting (Finance Favorites)
* **Unmerge Cells & Fill Down** – Unmerges selection and fills the top value downward.
* **Fill Blanks Down in Selection** – Fills every blank cell with the value from the cell above.
* **AutoFit All Columns & Rows** – Auto-fits entire workbook or selection.
* **Remove All Hyperlinks in Workbook** – Strips every hyperlink.
* **Convert All Formulas to Values** – Selection or entire workbook.
* **Highlight Duplicates in Column** – Colors duplicates red (perfect for invoice checks).
* **Delete Blank Rows in Selection** – Removes completely empty rows.
* **Delete Blank Columns in Selection** – Removes completely empty columns.
* **Remove Leading/Trailing Spaces** – Trims all text in selection.
* **Standardize Text Case** – Forces Proper, UPPER, or lower case on selection.
* **Convert Text to Numbers** – Fixes “text” numbers that won’t sum.
* **Clear All Formatting** – Resets selection to default (great before pasting finance exports).
### Reporting & Pivot Automation
* **Refresh All Pivot Tables & Charts** – Refreshes every pivot in the workbook.
* **Format All Pivot Charts** – Applies company-standard colors, fonts, and layout.
* **Create Summary Dashboard** – Inserts a new sheet with key stats from active data.
* **Export All Sheets as Separate PDFs** – One PDF per sheet, named automatically.
* **Export Selection as PDF** – Quick PDF of any highlighted range.
* **Generate Monthly Variance Report** – Assumes standard “Actual | Budget” columns and adds variance %/$.
* **Aging Analysis Buckets** – Auto-creates 0-30, 31-60, 61-90, 90+ buckets from date column.
### Finance-Specific Calculations & Modeling
* **Auto-Calculate Key Ratios** – From labeled columns (Revenue, Net Income, Assets, etc.) adds ROE, Gross Margin, Current Ratio, etc.
* **NPV / IRR Calculator on Selection** – Runs NPV or IRR on any cash-flow range you select.
* **Loan Amortization Generator** – Creates full schedule from principal, rate, term inputs.
* **Break-Even Analysis** – Calculates break-even point and adds chart.
* **Scenario Manager Runner** – Loops through multiple “What-If” input sets and outputs summary table.
* **Depreciation Scheduler** – Straight-line / double-declining on asset list.
* **Currency Conversion on Selection** – Multiplies selection by exchange rate you enter.
* **Consolidate Multiple Open Workbooks** – Sums same-named sheets across every open file.
* **Reconcile Two Sheets** – Highlights differences between “Book” and “Bank” columns.
### Advanced Productivity
* **Find & Replace Across All Sheets** – Global search/replace.
* **List All Formulas in Workbook** – New sheet showing every formula location.
* **Highlight Cells with Formulas vs Values** – Colors formulas blue, values green.
* **Random Data Generator** – Fills selection with realistic test numbers/dates/names (great for model testing).
* **Insert Timestamp** – Puts “Last Updated: date time” in a cell.
* **Cycle Through Open Workbooks** – Keyboard shortcut to flip between files.
* **Toggle Gridlines & Headings** – One-click clean view.
* **Add Company Watermark** – Semi-transparent text on every sheet.
* **Export Charts as Images** – Saves every chart as PNG to a folder.
### Security & Audit
* **Lock All Cells Except Input Cells** – Quick protection setup.
* **Show All Hidden Rows/Columns** – Reveals everything.
* **Remove Password from Workbook** – (if you know the password) unlocks protected files.
* **Audit Trail Logger** – Records every change with username & timestamp on a hidden sheet.
### Python Scripts (Run on ANY Excel file – 40+ ideas)
Run example: python clean_file.py "C:\reports\Q1_data.xlsx"
### Data Cleaning & Preparation
* **Universal Data Cleaner** – Removes duplicates, drops fully empty rows/columns, trims spaces, converts text-to-numbers, standardizes dates.
* **Fill Missing Values** – Forward-fill, mean-fill, or zero-fill any column you choose.
* **Standardize Column Names** – Makes headers snake_case or Title Case automatically.
* **Remove Special Characters** – Cleans phone numbers, IDs, currency symbols.
* **Split One Sheet into Multiple Files** – By unique value in any column (e.g., by department or region).
* **Merge All Sheets in Workbook** – Flattens every tab into one master sheet with “Source_Sheet” column.
### Finance Reporting & Analysis
* **Auto Financial Ratios Dashboard** – Reads any sheet with standard labels and outputs ROE, EBITDA margin, liquidity ratios + charts.
* **Variance Analysis Generator** – Compares “Actual vs Budget” columns across multiple files and creates summary + waterfall chart.
* **AR/AP Aging Report** – Buckets invoices by days overdue and adds totals.
* **Cash Flow Categorizer** – Auto-tags transactions by keywords and pivots monthly summary.
* **Portfolio Return Calculator** – Computes total return, IRR, Sharpe ratio from holdings + dates.
* **Monthly Close Packager** – From raw GL export creates P&L, BS, Cash Flow, and variance tabs in one click.
* **Budget vs Actual Consolidator** – Merges 50 department files and adds % variance.
* **Depreciation Roll-Forward** – Tracks asset additions, disposals, accumulated dep.
### Bulk Processing
* **Batch Process Folder of Files** – Runs the same cleaning/ratio script on every Excel file in a folder and saves “_processed” versions.
* **Compare Two Excel Files** – Highlights differences side-by-side or outputs diff report.
* **Convert All Excel in Folder to CSV** – Or to PDF (with formatting preserved via openpyxl).
* **Extract Specific Ranges from Multiple Files** – Pulls cell A1:J100 from every file and combines.
* **Password Remover / Protector Batch** – Removes or adds protection to dozens of files.
### Visualization & Output
* **Auto Pivot & Chart Creator** – Creates 5 standard pivots + charts and saves to new “Dashboard” sheet.
* **Conditional Formatting Applier** – Adds red/amber/green rules for variances, negatives, etc.
* **Styled Report Exporter** – Applies company branding (fonts, colors, logos) to any raw data file.
* **Generate PDF Report Pack** – One beautiful PDF with cover, summary tables, and charts.
### Advanced Finance & Automation
* **Sensitivity / Scenario Table Runner** – Varies 3–5 inputs and outputs full data table + tornado chart.
* **GL Reconciliation Engine** – Matches two large transaction lists and flags unmatched items.
* **Forecast Roll-Forward** – Takes last month actual + assumptions and builds next 12 months.
* **Invoice Number Validator** – Checks for duplicates or gaps in invoice sequences.
* **Tax Schedule Builder** – Applies tax rates by jurisdiction and summarizes.
* **Cap Table Waterfall** – From share issuance data creates ownership % waterfall.
### Utility Scripts
* **File Metadata Reporter** – Lists every sheet name, row count, last modified date for a folder of files.
* **Large File Splitter** – Splits files >1M rows into 100k-row chunks.
* **Transpose Any Sheet** – Flips rows ↔ columns.
* **Add Unique ID Column** – Inserts UUID or sequential ID.
* **Email Report Sender** – (Outlook integration) attaches processed file and sends to distribution list.
