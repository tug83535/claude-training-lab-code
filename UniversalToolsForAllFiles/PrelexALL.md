# Universal Excel Macros & Python Scripts — Company-Wide Utility List

> **Purpose:** A master catalog of VBA macros and Python scripts that can be executed against **any** Excel file regardless of structure. Organized by category. Each entry includes a name and purpose only — no code. Intended for review by all teams with additional emphasis on Finance use cases. Excludes functionality already built into OneDrive for Business (co-authoring, version history, AutoSave, sharing, access control, cloud sync).

---

## 🔹 1. FORMATTING & PRESENTATION

| # | Name | Purpose |
|---|------|---------|
| 1 | AutoFit All Columns | Automatically resize every column width to fit its content |
| 2 | AutoFit All Rows | Automatically resize every row height to fit its content |
| 3 | Standardize Font Across Workbook | Set every cell in the workbook to a single company-standard font, size, and color |
| 4 | Zebra-Stripe Rows | Apply alternating row shading to any data range for readability |
| 5 | Apply Corporate Header Style | Format row 1 of every sheet with company branding (colors, bold, font, borders) |
| 6 | Freeze Top Row on All Sheets | Apply Freeze Panes to row 1 on every worksheet in the workbook |
| 7 | Freeze Top Row and First Column on All Sheets | Apply Freeze Panes to row 1 + column A on every worksheet |
| 8 | Remove All Cell Colors | Strip all fill/background colors from every cell in the active sheet or workbook |
| 9 | Reset All Font Colors to Black | Set font color to black on every cell to remove rainbow formatting |
| 10 | Border All Data Cells | Apply thin borders to every cell that contains data |
| 11 | Remove All Borders | Strip all cell borders from the active sheet |
| 12 | Set Print Area to Used Range | Automatically set the print area to only the range with data on each sheet |
| 13 | Add Page Header/Footer to All Sheets | Insert standardized headers and footers (file name, date, page number) on every sheet |
| 14 | Force Landscape Orientation on All Sheets | Set every sheet to landscape printing |
| 15 | Fit All Sheets to One Page Wide | Set page setup so content fits to one page wide across all sheets |
| 16 | Number Format Standardizer | Apply consistent number formatting (e.g., #,##0.00) to all numeric cells across all sheets |
| 17 | Currency Format Standardizer | Detect and convert all currency values to a consistent format ($#,##0.00) |
| 18 | Percentage Format Fixer | Identify values between 0-1 that should be percentages and apply % formatting |
| 19 | Date Format Standardizer | Normalize all dates to a single format (e.g., MM/DD/YYYY or YYYY-MM-DD) across all sheets |
| 20 | Conditional Formatting — Negative Values in Red | Highlight all negative numbers in red across the workbook |
| 21 | Conditional Formatting — Top/Bottom N Values | Highlight the top 10 and bottom 10 values in a selected column |
| 22 | Conditional Formatting — Duplicates Highlighter | Highlight duplicate values in a selected column with a fill color |
| 23 | Remove All Conditional Formatting | Strip conditional formatting rules from every cell in the workbook |
| 24 | Text Alignment Standardizer | Left-align text, right-align numbers, and center headers across all sheets |
| 25 | Wrap Text on All Cells | Enable text wrapping on all cells that contain text longer than the column width |
| 26 | Merge Cell Remover | Unmerge all merged cells across the entire workbook and fill values down |

---

## 🔹 2. DATA CLEANING & TRANSFORMATION

| # | Name | Purpose |
|---|------|---------|
| 27 | Remove Duplicate Rows | Delete duplicate rows from the active sheet based on all columns or selected columns |
| 28 | Highlight Duplicate Rows | Color-code duplicate rows instead of deleting them for manual review |
| 29 | Trim All Whitespace | Remove leading, trailing, and excess interior spaces from every text cell |
| 30 | Clean Non-Printable Characters | Remove non-printable and special characters (CHAR(0)–CHAR(31)) from all cells |
| 31 | Convert Text-Numbers to Actual Numbers | Fix cells storing numbers as text by converting them to true numeric values |
| 32 | Convert Text-Dates to Date Values | Parse text strings that represent dates and convert to Excel date serial numbers |
| 33 | Remove All Blank Rows | Delete every entirely blank row within the used range |
| 34 | Remove All Blank Columns | Delete every entirely blank column within the used range |
| 35 | Highlight Blank Cells in Data Range | Color blank/empty cells within a populated range to identify missing data |
| 36 | Fill Blank Cells with Value Above | Fill empty cells downward with the value from the cell above (useful for merged-style layouts) |
| 37 | Fill Blank Cells with Custom Value | Replace all blank cells in a range with a specified default value (0, "N/A", "TBD", etc.) |
| 38 | Remove All Hyperlinks | Strip every hyperlink from the workbook while preserving the display text |
| 39 | Convert All Formulas to Values | Replace every formula in the workbook with its current calculated value |
| 40 | Convert Formulas to Values — Active Sheet Only | Same as above but scoped to the active sheet |
| 41 | Standardize Case — UPPER | Convert all text cells to UPPER CASE |
| 42 | Standardize Case — lower | Convert all text cells to lower case |
| 43 | Standardize Case — Proper/Title | Convert all text cells to Proper Case / Title Case |
| 44 | Find & Replace Across Workbook | Perform a find-and-replace operation across every sheet in the workbook simultaneously |
| 45 | Split Delimited Cell Values into Columns | Split comma-separated (or other delimiter) values in a column into multiple columns |
| 46 | Split Delimited Cell Values into Rows | Explode delimited values in a column into separate rows (one value per row) |
| 47 | Concatenate Columns | Merge values from multiple columns into a single column with a specified separator |
| 48 | Transpose Data Blocks | Flip rows to columns or columns to rows for a selected range |
| 49 | Remove Line Breaks Within Cells | Strip carriage returns and line feeds (Alt+Enter content) from all cells |
| 50 | Standardize Phone Numbers | Detect phone number patterns and reformat to a consistent format (e.g., (XXX) XXX-XXXX) |
| 51 | Standardize Email Addresses | Trim and lowercase all email addresses in a specified column |
| 52 | Extract Emails from Text Column | Scan a text column and extract valid email addresses into a new column |
| 53 | Extract Numbers from Text | Pull numeric values out of cells that contain mixed text and numbers |
| 54 | Remove All Comments/Notes | Delete every comment and note from all cells in the workbook |
| 55 | Convert Smart Quotes to Straight Quotes | Replace curly/smart quotes with standard straight quotes throughout the workbook |
| 56 | Replace #N/A and Error Values | Replace all error values (#N/A, #REF!, #VALUE!, #DIV/0!) with blank or a specified value |
| 57 | Unpivot / Melt Data | Transform wide-format crosstab data into long-format (database-friendly) rows |
| 58 | Pivot / Reshape Data | Aggregate and pivot long-format data into a summary crosstab table |
| 59 | Deduplicate and Aggregate | Remove duplicate keys while summing, averaging, or concatenating related values |
| 60 | Regex Pattern Search and Extract | Use regular expressions to find and extract patterns (IDs, codes, amounts) from text cells |

---

## 🔹 3. DATA QUALITY & VALIDATION

| # | Name | Purpose |
|---|------|---------|
| 61 | Data Quality Scorecard | Generate a summary report: total rows, blank cells, error cells, duplicate rows, data types per column |
| 62 | Column Data Type Audit | Report the data type (text, number, date, blank, error) breakdown for every column |
| 63 | Outlier Detector (IQR Method) | Flag numeric values that fall outside 1.5× the interquartile range as potential outliers |
| 64 | Outlier Detector (Z-Score Method) | Flag numeric values with a Z-score > 3 as statistical outliers |
| 65 | Missing Data Heatmap | Create a color-coded matrix showing where data is missing across all columns/rows |
| 66 | Validate Email Format | Check a column for improperly formatted email addresses and flag them |
| 67 | Validate Date Ranges | Flag dates that fall outside an expected range (e.g., future dates, dates before 2000) |
| 68 | Validate Numeric Ranges | Flag values outside a user-defined min/max range in a specified column |
| 69 | Cross-Column Consistency Check | Verify logical rules between columns (e.g., Start Date < End Date, Subtotal + Tax = Total) |
| 70 | Spell Check All Sheets | Run spell check across all text cells in all sheets |
| 71 | Identify Formula Cells | Highlight or list all cells that contain formulas rather than hard-coded values |
| 72 | Formula Auditor — Identify Inconsistent Formulas | Flag cells in a column where the formula pattern differs from the majority |
| 73 | Circular Reference Detector | Scan the workbook for circular references and report their locations |
| 74 | External Link Finder | List all cells that reference external workbooks with file paths and sheet references |
| 75 | Broken Link Reporter | Identify all external links that are broken or point to missing files |
| 76 | Named Range Auditor | List every named range in the workbook with its scope, reference, and whether it's valid |
| 77 | Data Validation Rule Auditor | List all data validation rules applied in the workbook with cell locations and criteria |
| 78 | Check for Hidden Content | Report any hidden rows, columns, sheets, or cell content (white font on white) |
| 79 | Sensitive Data Scanner (PII Detector) | Scan for patterns that match SSNs, credit card numbers, or account numbers and flag them |
| 80 | Leading Zero Check | Identify numeric columns where leading zeros have been dropped (ZIP codes, account numbers) |

---

## 🔹 4. SHEET & WORKBOOK MANAGEMENT

| # | Name | Purpose |
|---|------|---------|
| 81 | Create Table of Contents Sheet | Generate a clickable index sheet listing every worksheet name with hyperlinks |
| 82 | Alphabetize Sheet Tabs | Sort all worksheet tabs in alphabetical order |
| 83 | Sort Sheet Tabs by Color | Rearrange worksheet tabs grouped by their tab color |
| 84 | Unhide All Sheets | Make every hidden and very-hidden worksheet visible |
| 85 | Hide All Sheets Except Active | Hide every sheet except the one currently active |
| 86 | Delete All Blank Sheets | Remove any worksheets that have no data |
| 87 | Rename Sheets from Cell Values | Rename each sheet tab using the value from a specified cell (e.g., A1) on that sheet |
| 88 | Batch Rename Sheets with Prefix/Suffix | Add a consistent prefix or suffix to every sheet name |
| 89 | Copy Active Sheet to New Workbook | Duplicate the active sheet into a brand-new workbook file |
| 90 | Split Each Sheet into Separate Workbooks | Save every worksheet as its own individual Excel file |
| 91 | Merge Multiple Workbooks into One | Combine all sheets from multiple selected Excel files into a single workbook |
| 92 | Merge All Sheets into One Master Sheet | Consolidate data from all sheets (with identical column headers) into one unified sheet |
| 93 | Compare Two Sheets Side by Side | Identify and highlight cell-level differences between two worksheets |
| 94 | Protect All Sheets with Password | Apply worksheet protection with a specified password to every sheet |
| 95 | Unprotect All Sheets | Remove worksheet protection from every sheet (requires password if set) |
| 96 | Protect Workbook Structure | Lock the workbook structure to prevent adding, deleting, or moving sheets |
| 97 | Lock All Formula Cells | Lock cells containing formulas while leaving input cells editable |
| 98 | Timestamp Sheet — Last Modified | Add a cell to each sheet showing when it was last modified and by whom |
| 99 | Sheet Row/Column Count Summary | Generate a summary table showing row count and column count for every sheet |
| 100 | Archive Completed Sheets | Move sheets flagged as "complete" to a designated Archive section of the workbook |

---

## 🔹 5. EXPORT, IMPORT & FILE OPERATIONS

| # | Name | Purpose |
|---|------|---------|
| 101 | Export Active Sheet as PDF | Save the active worksheet as a PDF file |
| 102 | Export All Sheets as Individual PDFs | Save each worksheet as its own separate PDF file |
| 103 | Export All Sheets as One Combined PDF | Combine all visible sheets into a single multi-page PDF |
| 104 | Export Active Sheet as CSV | Save the active sheet as a CSV file with UTF-8 encoding |
| 105 | Export All Sheets as Individual CSVs | Save each worksheet as its own CSV file |
| 106 | Export Selected Range as Image | Capture a selected range and save it as a PNG/JPG image |
| 107 | Import CSV to New Sheet | Load a selected CSV file into a new worksheet with auto-detected delimiters |
| 108 | Import Multiple CSVs into Sheets | Batch-import a folder of CSV files, each into its own worksheet |
| 109 | Import All Excel Files from Folder | Open all .xlsx/.xls files in a specified folder and consolidate into one workbook |
| 110 | Export Data as JSON | Convert the active sheet's data to a JSON file |
| 111 | Export Data as XML | Convert the active sheet's data to an XML file |
| 112 | Import JSON into Sheet | Parse a JSON file and load it into a structured Excel worksheet |
| 113 | Batch File Renamer from List | Rename files in a folder based on a mapping list in the workbook (old name → new name) |
| 114 | Email Active Workbook via Outlook | Attach the current workbook to a new Outlook email with pre-filled recipients and subject |
| 115 | Email Active Sheet as PDF via Outlook | Export the active sheet to PDF and attach it to a new Outlook email |
| 116 | Batch Email Sheets to Different Recipients | Send each sheet as a PDF to a different recipient based on a distribution list in the workbook |
| 117 | Save Timestamped Backup Copy | Save a copy of the workbook with a timestamp appended to the filename |
| 118 | Archive Workbook to Specified Folder | Copy the current workbook to a designated archive/backup folder |
| 119 | Convert Workbook to .xlsx (Remove Macros) | Save a macro-enabled workbook as a macro-free .xlsx file |
| 120 | Flatten Workbook — Single Sheet Export | Export only the data (values, no formulas) from the active sheet into a clean new workbook |

---

## 🔹 6. NAVIGATION & USABILITY

| # | Name | Purpose |
|---|------|---------|
| 121 | Go to First Blank Row | Jump the cursor to the first empty row at the bottom of the dataset |
| 122 | Go to Last Cell with Data | Navigate to the last used cell in the sheet |
| 123 | Add Hyperlinked Sheet Navigator | Insert dropdown or button-based navigation to jump between sheets |
| 124 | Highlight Active Row and Column | Dynamically highlight the entire row and column of the selected cell for easy tracking |
| 125 | Search Across All Sheets | Search for a value or text string across every sheet and return sheet name + cell address |
| 126 | Toggle Gridlines on All Sheets | Turn gridlines on or off across every worksheet at once |
| 127 | Toggle Sheet Tab Visibility | Show or hide the sheet tab bar |
| 128 | Toggle Formula View on All Sheets | Switch between showing formulas and showing calculated values on all sheets |
| 129 | Zoom All Sheets to Specified Level | Set a consistent zoom level (e.g., 85%, 100%) across every sheet |
| 130 | Collapse All Grouped Rows/Columns | Collapse every outline group in the workbook |
| 131 | Expand All Grouped Rows/Columns | Expand every outline group in the workbook |
| 132 | Auto-Add Dropdown Lists from Unique Values | Create data validation dropdowns populated by the unique values found in a column |
| 133 | Create Dynamic Named Ranges | Automatically create named ranges that expand as data is added |
| 134 | Reset All Filters | Clear all AutoFilter and Advanced Filter criteria on every sheet |
| 135 | Remove All Filters | Remove AutoFilter dropdowns from every sheet |

---

## 🔹 7. ANALYSIS & REPORTING

| # | Name | Purpose |
|---|------|---------|
| 136 | Summary Statistics Generator | Create a new sheet with count, sum, average, min, max, median, stdev for every numeric column |
| 137 | Frequency Distribution Builder | Generate frequency distribution tables and bins for numeric columns |
| 138 | Correlation Matrix Generator | Calculate and display a correlation matrix for all numeric columns |
| 139 | Pivot Table Auto-Builder | Automatically create a pivot table from the active dataset with user-selected row/column/value fields |
| 140 | Year-Over-Year Comparison Builder | Align two periods of data side by side and calculate absolute and percentage change |
| 141 | Month-Over-Month Trend Table | Generate a table showing MoM values and changes for a selected metric |
| 142 | Variance Analysis Template | Calculate actual vs. budget variances ($ and %) for all line items |
| 143 | Rank Column Values | Add a rank column based on ascending or descending sort of a numeric column |
| 144 | Running Total / Cumulative Sum | Add a column with running totals for a specified numeric column |
| 145 | Moving Average Calculator | Calculate N-period moving averages for a time series column |
| 146 | Pareto Analysis (80/20) | Identify which items account for 80% of the total and flag them |
| 147 | ABC Classification | Classify items into A (top 80%), B (next 15%), C (bottom 5%) categories by value |
| 148 | Aging Bucket Calculator | Assign records to aging buckets (0-30, 31-60, 61-90, 90+ days) based on a date column |
| 149 | Waterfall Data Preparer | Structure data for waterfall chart visualization (opening, increases, decreases, closing) |
| 150 | Auto-Chart Generator | Create a chart (bar, line, pie) from any selected data range with one click |
| 151 | Sparklines Inserter | Add sparkline mini-charts next to each row of data |
| 152 | Cross-Tab Report Generator | Create a cross-tabulation summary from two categorical columns and a value column |
| 153 | Top N / Bottom N Report | Extract and display the top N and bottom N records by a specified metric |
| 154 | Data Profiling Report | Generate a full profile of every column: unique values, nulls, min/max length, patterns |

---

## 🔹 8. FINANCE-SPECIFIC UTILITIES

| # | Name | Purpose |
|---|------|---------|
| 155 | Account Number Validator | Validate account number format/length against a chart of accounts or expected pattern |
| 156 | GL Account Mapper | Map general ledger codes to account names/descriptions using a reference table |
| 157 | Trial Balance Checker | Verify that total debits equal total credits across all accounts |
| 158 | Balance Sheet Balancer | Confirm Assets = Liabilities + Equity and highlight any imbalance |
| 159 | Intercompany Eliminations Checker | Identify and flag intercompany transactions that should net to zero |
| 160 | Journal Entry Validator | Check that every journal entry has balanced debits and credits |
| 161 | Duplicate Invoice Detector | Scan for potential duplicate invoices by matching vendor, amount, date, and invoice number |
| 162 | Three-Way Invoice Match | Compare PO, receipt, and invoice data to flag mismatches |
| 163 | Bank Reconciliation Matcher | Match bank statement lines to GL entries by amount and date and flag unmatched items |
| 164 | Outstanding Check Lister | Extract a list of all outstanding/uncleared checks from reconciliation data |
| 165 | AP Aging Summary Generator | Generate an accounts payable aging report with standard buckets (current, 30, 60, 90, 120+) |
| 166 | AR Aging Summary Generator | Generate an accounts receivable aging report with standard buckets |
| 167 | DSO Calculator (Days Sales Outstanding) | Calculate DSO from revenue and receivables data |
| 168 | DPO Calculator (Days Payable Outstanding) | Calculate DPO from COGS and payables data |
| 169 | Cash Flow Categorizer | Classify transactions as operating, investing, or financing cash flows based on account codes |
| 170 | Revenue Recognition Scheduler | Spread a revenue amount over a specified recognition period in monthly buckets |
| 171 | Depreciation Schedule Generator | Calculate straight-line, declining balance, or sum-of-years depreciation for asset lists |
| 172 | Amortization Schedule Generator | Build loan amortization tables with principal, interest, and balance for each period |
| 173 | NPV Calculator (Batch) | Calculate Net Present Value for multiple projects/scenarios using provided cash flows and discount rate |
| 174 | IRR Calculator (Batch) | Calculate Internal Rate of Return for multiple project cash flow series |
| 175 | Budget vs. Actual Variance Report | Compare budget line items against actuals and calculate dollar and percentage variances |
| 176 | Forecast Variance Tracker | Compare latest forecast to prior forecast and budget, showing directional changes |
| 177 | Rolling Forecast Builder | Generate a rolling 12-month or 18-month forecast template from historical data |
| 178 | Cost Allocation Spreader | Distribute a pool of costs across departments/projects based on allocation percentages |
| 179 | FX Currency Converter | Convert amounts from one currency to another using a rate table in the workbook |
| 180 | Multi-Currency Consolidator | Aggregate financial data from sheets in different currencies into a single reporting currency |
| 181 | Tax Rate Applier | Apply tax rates from a reference table to transaction data based on jurisdiction or category |
| 182 | Accrual / Reversal Generator | Create accrual journal entries and their corresponding reversal entries for the next period |
| 183 | Expense Report Validator | Check expense reports for policy violations (missing receipts, over-limit amounts, duplicate entries) |
| 184 | Vendor Spend Analyzer | Summarize total spend by vendor with transaction counts and average amounts |
| 185 | Customer Revenue Analyzer | Summarize total revenue by customer with transaction counts and growth calculations |
| 186 | Margin Calculator | Calculate gross margin, operating margin, and net margin from financial statement data |
| 187 | Break-Even Analysis | Calculate break-even point from fixed costs, variable costs, and price data |
| 188 | Ratio Analysis Dashboard Builder | Calculate key financial ratios (current ratio, quick ratio, debt-to-equity, ROE, ROA, etc.) from statement data |
| 189 | Flux Analysis / Period Comparison | Identify line items with significant period-over-period changes exceeding a threshold |
| 190 | Check Digit Validator (Modulus) | Validate account/ID numbers using check digit algorithms (Mod 10, Mod 11) |
| 191 | 1099 / Tax Form Data Preparer | Extract and format vendor payment data for 1099 reporting thresholds |
| 192 | Payment Terms Calculator | Calculate due dates from invoice dates based on payment terms (Net 30, 2/10 Net 30, etc.) |
| 193 | Late Payment Interest Calculator | Calculate interest charges on overdue invoices based on rates and days past due |
| 194 | Capitalization Threshold Checker | Flag items above/below the capitalization threshold to determine expense vs. asset treatment |

---

## 🔹 9. AUDIT, COMPLIANCE & CHANGE TRACKING

| # | Name | Purpose |
|---|------|---------|
| 195 | Change Log Generator | Create a detailed log of every cell change with old value, new value, user, and timestamp |
| 196 | Sheet Snapshot / Baseline Capture | Save a snapshot of all current values as a baseline for future comparison |
| 197 | Baseline vs. Current Comparison | Compare current data against a saved baseline and highlight all changes |
| 198 | Who-Changed-What Tracker | Log the username, cell address, old value, and new value every time a cell is edited |
| 199 | Workbook Metadata Reporter | Generate a report with file name, path, size, author, creation date, last modified date, and sheet count |
| 200 | Formula Complexity Scorer | Score each formula by complexity (nesting depth, function count) to identify risk areas |
| 201 | Workbook Error Scanner | Scan all sheets and list every cell containing an error value with its location and error type |
| 202 | Hardcoded Number Finder | Identify numeric constants embedded in formulas instead of referencing cells |
| 203 | Access Control Summary | Report which cells/sheets are protected and what the protection settings are |
| 204 | Pre-Submission Checklist Runner | Run a configurable list of quality checks before a report is finalized (no blanks, totals balance, etc.) |
| 205 | Regulatory Rounding Enforcer | Apply specified rounding rules (e.g., round to nearest cent, nearest thousand) consistently |
| 206 | Footnote / Annotation Manager | Add, view, and manage numbered footnotes linked to specific cells |

---

## 🔹 10. SECURITY & CLEANUP

| # | Name | Purpose |
|---|------|---------|
| 207 | Remove Personal Information / Metadata | Strip author name, comments, revision history, and other metadata from the workbook |
| 208 | Redact Sensitive Columns | Mask or redact specified columns (e.g., SSN, account numbers) before sharing |
| 209 | Password Protect Specific Sheets | Selectively apply password protection to named sheets |
| 210 | Clear All Sheet Contents (Keep Structure) | Delete all data from every sheet while preserving sheet names and formatting |
| 211 | Clear All Formatting (Keep Data) | Remove all formatting while preserving raw data values |
| 212 | Remove All VBA Code | Strip all macros/VBA modules from the workbook for a clean file |
| 213 | Remove All Pictures and Objects | Delete all embedded images, shapes, charts, and ActiveX controls |
| 214 | Compress / Optimize File Size | Remove unused styles, empty cells beyond data range, and metadata to reduce file size |
| 215 | Reset Used Range | Fix the "phantom row/column" issue where Excel thinks the used range extends far beyond actual data |

---

## 🔹 11. AUTOMATION & WORKFLOW

| # | Name | Purpose |
|---|------|---------|
| 216 | Auto-Run Macro on File Open | Trigger a specified set of actions (formatting, refresh, validation) every time the workbook opens |
| 217 | Scheduled Report Refresh | Refresh all data connections, pivot tables, and calculations on a timed schedule |
| 218 | Refresh All Pivot Tables | Refresh every pivot table and pivot chart in the workbook with one click |
| 219 | Refresh All Data Connections | Refresh all external data connections (databases, web queries, Power Query) |
| 220 | Button-Based Macro Runner | Create clickable buttons on a control sheet that execute specific macros |
| 221 | UserForm Input Dialog | Create a popup form for guided data entry with validation rules |
| 222 | Progress Bar for Long-Running Macros | Display a visual progress indicator during lengthy macro operations |
| 223 | Batch Process Files in Folder | Run a specified macro or transformation against every Excel file in a selected folder |
| 224 | Auto-Email on Cell Value Change | Trigger an Outlook email when a specific cell value meets a condition (e.g., threshold exceeded) |
| 225 | Auto-Email on Workbook Save | Send a notification email every time the workbook is saved |
| 226 | Custom Ribbon Tab Builder | Add a custom ribbon tab with grouped buttons for your most-used macros |
| 227 | Right-Click Context Menu Customizer | Add custom options to the right-click context menu |
| 228 | Keyboard Shortcut Mapper | Assign custom keyboard shortcuts to frequently used macros |
| 229 | Task Scheduler Integration (Python) | Schedule a Python script to run Excel automations at specific times via Windows Task Scheduler |
| 230 | Multi-Step Workflow Chain | Execute a sequence of macros in a defined order with error handling between steps |

---

## 🔹 12. CROSS-APPLICATION INTEGRATION

| # | Name | Purpose |
|---|------|---------|
| 231 | Excel to PowerPoint — Paste Ranges as Images | Copy specified ranges from Excel and paste them as images into a PowerPoint template |
| 232 | Excel to PowerPoint — Update Linked Charts | Push updated charts from Excel into an existing PowerPoint deck |
| 233 | Excel to Word — Mail Merge Data Prep | Format and validate Excel data for use as a Word mail merge data source |
| 234 | Excel to Word — Paste Tables into Template | Insert formatted Excel tables into specified locations in a Word document |
| 235 | Outlook Calendar Creator from Sheet | Create Outlook calendar events from a list of dates, times, and descriptions in Excel |
| 236 | Outlook Contact Importer | Import a list of contacts from Excel into Outlook |
| 237 | SQL Query Runner | Execute a SQL query against a database and return results directly into a worksheet |
| 238 | API Data Fetcher (Python) | Call a REST API and load the response data into an Excel sheet |
| 239 | SharePoint List Uploader | Push Excel data to a SharePoint list |
| 240 | SharePoint List Downloader | Pull data from a SharePoint list into an Excel sheet |
| 241 | Database Table Exporter | Write worksheet data directly to a SQL Server / MySQL / PostgreSQL table |
| 242 | PDF Table Extractor (Python) | Extract tables from PDF files and load them into Excel worksheets |
| 243 | Web Scraper to Excel (Python) | Scrape structured data from a webpage and populate an Excel sheet |
| 244 | Teams / Slack Notification Sender (Python) | Send a message to a Teams or Slack channel when a workbook process completes |

---

## 🔹 13. TEXT & CONTENT UTILITIES

| # | Name | Purpose |
|---|------|---------|
| 245 | Word Count per Cell | Count words in each cell and place the count in an adjacent column |
| 246 | Character Count per Cell | Count characters in each cell for length validation |
| 247 | Extract Initials from Names | Pull initials from full name cells (e.g., "John Smith" → "JS") |
| 248 | Split Full Name into First/Last | Separate a full name column into First Name and Last Name columns |
| 249 | Address Parser | Split a full address string into Street, City, State, ZIP components |
| 250 | Lookup Value Across All Sheets | Search for a value across all sheets and return every match with sheet and cell location |
| 251 | Batch Find and Highlight | Highlight all cells matching any value from a supplied list of search terms |
| 252 | Text-to-Columns Batch | Apply text-to-columns splitting to a specified column across all sheets |
| 253 | Generate Random Sample | Extract a random sample of N rows from a dataset for testing or auditing |
| 254 | Encrypt / Obfuscate Column Values | Apply simple obfuscation (e.g., hash, mask) to column values for data anonymization |
| 255 | Unique Values Extractor | Extract a list of unique/distinct values from a column into a new column or sheet |
| 256 | VLOOKUP/XLOOKUP Batch Builder | Automatically insert lookup formulas that pull data from a reference sheet for every row |

---

## 🔹 14. PERFORMANCE & DIAGNOSTICS

| # | Name | Purpose |
|---|------|---------|
| 257 | Workbook Health Check | Generate a comprehensive diagnostic: file size, sheet count, formula count, error count, external links, named ranges, pivot tables |
| 258 | Calculation Speed Timer | Measure how long a full workbook recalculation takes to identify performance bottlenecks |
| 259 | Volatile Function Finder | Identify cells using volatile functions (NOW, TODAY, INDIRECT, OFFSET, RAND) that slow recalculation |
| 260 | Large Range Reference Finder | Identify formulas referencing unnecessarily large ranges (e.g., A:A instead of A1:A1000) |
| 261 | Unused Named Range Cleaner | Find and delete named ranges that are not referenced anywhere in the workbook |
| 262 | Style / Cell Format Bloat Fixer | Remove excess cell styles that cause "Too many cell formats" errors |
| 263 | Used Range vs. Actual Data Comparison | Compare Excel's perceived used range versus the actual data range and fix discrepancies |
| 264 | Memory Usage Estimator | Estimate workbook memory consumption based on cell count, formula count, and object count |

---

## 🔹 15. COLLABORATION & DISTRIBUTION

| # | Name | Purpose |
|---|------|---------|
| 265 | Build Distribution-Ready Copy | Create a clean copy with formulas converted to values, metadata stripped, and formatting standardized |
| 266 | Sheet-Level Access Matrix Builder | Generate a matrix showing which users/roles should have access to which sheets |
| 267 | Dynamic Filter by User | Automatically filter data on the active sheet based on the logged-in Windows username |
| 268 | Create User-Specific Views | Save and restore named views (filter settings, hidden columns) for different users/roles |
| 269 | Data Entry Template Generator | Build a blank data entry template from the headers and validation rules of an existing sheet |
| 270 | Changelog Summarizer | Read the change log and produce a human-readable summary of recent modifications |
| 271 | Sign-Off / Approval Stamp | Insert a cell stamp with the current user's name and datetime as a review/approval indicator |
| 272 | Version Number Incrementer | Automatically increment a version number cell each time the workbook is saved |

---

## 🔹 16. PYTHON-SPECIFIC ADVANCED UTILITIES

| # | Name | Purpose |
|---|------|---------|
| 273 | Multi-File Data Consolidator | Read and combine data from hundreds of Excel files in a folder into one master DataFrame/sheet |
| 274 | Fuzzy Match / Fuzzy Lookup | Match records between two datasets using fuzzy string matching (vendor names, customer names) |
| 275 | Anomaly Detection (Statistical) | Apply statistical anomaly detection (Isolation Forest, DBSCAN) to flag unusual transactions |
| 276 | Trend Line and Forecast Generator | Fit trend lines and generate forecasts for time-series data using statistical models |
| 277 | Benford's Law Analyzer | Test a numeric column against Benford's Law distribution to detect potential fraud or data manipulation |
| 278 | Sentiment Analysis on Text Column | Analyze text feedback or notes columns for positive/negative/neutral sentiment scores |
| 279 | Cluster Analysis | Group records into clusters based on numeric attributes using K-Means or similar |
| 280 | Regression Analysis Report | Run linear or multiple regression on selected columns and output coefficients, R², and p-values |
| 281 | Automated Chart Suite | Generate a full suite of charts (histogram, box plot, scatter, heatmap) for every numeric column |
| 282 | Data Dictionary Generator | Auto-generate a data dictionary from the workbook: column names, types, sample values, descriptions |
| 283 | Schema Comparator | Compare the column structure (headers, order, types) of two Excel files and report differences |
| 284 | Large File Handler (Chunked Processing) | Process very large Excel files in chunks to avoid memory errors |
| 285 | Workbook Differ (Cell-Level) | Compare two versions of a workbook cell by cell and produce a detailed diff report |
| 286 | Auto-Documentation Generator | Generate a markdown or HTML document describing the workbook's structure, formulas, and logic |
| 287 | OCR Image-to-Excel (Python) | Extract text/tables from images or scanned documents and load into Excel |
| 288 | Natural Language Query (Python + LLM) | Allow users to ask plain-English questions about the data and return filtered/summarized answers |
| 289 | Automated Email Report with Inline Charts (Python) | Generate a formatted HTML email with embedded charts from Excel data and send via SMTP |
| 290 | Scheduled Data Pull and Refresh (Python) | Automatically pull fresh data from APIs/databases on a schedule and update the Excel file |
| 291 | Excel to Dashboard HTML Exporter (Python) | Convert Excel data into a standalone interactive HTML dashboard |
| 292 | Cross-File VLOOKUP (Python) | Perform lookups across multiple Excel files without opening them all simultaneously |
| 293 | Bulk Watermark Adder (Python) | Add a watermark (e.g., "DRAFT", "CONFIDENTIAL") to every sheet when exporting to PDF |

---

## 🔹 17. MISCELLANEOUS POWER UTILITIES

| # | Name | Purpose |
|---|------|---------|
| 294 | Undo/Redo History Logger | Maintain a persistent log of undo/redo actions for session recovery |
| 295 | Color Palette Standardizer | Replace all non-standard colors in the workbook with the nearest company-approved color |
| 296 | Chart Style Standardizer | Update all charts in the workbook to match a corporate style (colors, fonts, gridlines) |
| 297 | Cell Comment-to-Column Converter | Extract all cell comments/notes into a new column next to the source data |
| 298 | Workbook Table of Formulas | List every unique formula in the workbook with its sheet, cell, and formula text |
| 299 | Dependency Tree / Precedent-Dependent Map | Map out which cells feed into which for a selected formula (precedent/dependent tree) |
| 300 | Easter Egg: Random Motivational Quote | Insert a random motivational quote into a cell — because sometimes your team needs it |

---

*Generated for company-wide distribution. Each item is tool-agnostic (VBA Macro or Python script) and designed to run on any Excel file without requiring file-specific customization.*
