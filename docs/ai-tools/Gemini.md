# Gemini

# Excel VBA Automation: Comparison & Build Specifications

**SYSTEM INSTRUCTIONS FOR AI:**
The user will provide an existing Excel file, a description of their workflow, or their current VBA code. Your role is to act as an Expert VBA Developer. Please execute the following steps:
1. **Analyze:** Review the user's provided file structure, current code, or workflow description.
2. **Compare:** Cross-reference the user's current setup against the "VBA Macro Catalog" below. 
3. **Recommend:** Identify gaps and recommend which macros from this catalog will provide the highest return on investment (ROI) for their specific file.
4. **Build:** When the user selects a macro to build, use the exact Logic Requirements and Edge Cases provided below to generate highly optimized, heavily commented VBA code. Ensure all generated code includes robust error handling (`On Error GoTo`), screen updating toggles (`Application.ScreenUpdating = False`), and calculation toggles (`Application.Calculation = xlCalculationManual`) for maximum performance.

---

## VBA Macro Catalog

### LEVEL 1: EASY (QUICK WINS)

**1. Unmerge and Fill Down**
* **Target Audience:** Finance teams dealing with messy ERP/system exports.
* **Logic Requirements:** Act only on the user's currently selected range. Identify any merged cells within the selection and unmerge them. Loop through the selection; if a cell is blank, fill it with the value of the cell directly above it.
* **Edge Cases:** If the selection is only one cell, prompt the user to select a larger range.

**2. Auto-Format Financial Reports**
* **Target Audience:** Analysts standardizing raw data dumps.
* **Logic Requirements:** Apply Accounting number format (no decimals, $ symbol) to all cells containing numbers. Identify the first row with data and make it Bold with a bottom border (Header row). Auto-fit all columns in the active worksheet.
* **Edge Cases:** Ignore completely blank columns/rows during the auto-fit process.

**3. Highlight Hardcoded Numbers**
* **Target Audience:** Financial modelers and auditors.
* **Logic Requirements:** Scan the active worksheet's used range. Use `SpecialCells(xlCellTypeConstants, xlNumbers)` to identify cells containing typed numbers (not formulas). Change the font color of those cells to blue (RGB 0,0,255).
* **Edge Cases:** Do not change the color of dates, text, or formulas. 

**4. Toggle Presentation Mode**
* **Target Audience:** Dashboard creators and executive reporting teams.
* **Logic Requirements:** Create a toggle switch that checks the current state of the application. Hide/unhide `ActiveWindow.DisplayGridlines`, `ActiveWindow.DisplayHeadings`, `Application.DisplayFormulaBar`, and collapse/expand the Excel Ribbon.
* **Edge Cases:** Ensure the toggle safely restores all default views when turned off.

### LEVEL 2: MEDIUM (DATA WRANGLING & NAVIGATION)

**5. Delete Completely Blank Rows**
* **Target Audience:** Data cleaners dealing with fragmented datasets.
* **Logic Requirements:** Define the used range of the active worksheet. Loop backwards from the last used row to row 1 (to prevent shifting errors during deletion). Use `WorksheetFunction.CountA` to check if the row is empty. If 0, delete the entire row.
* **Edge Cases:** Trim data first to ensure rows containing only spaces or hidden characters are accurately identified as blank and deleted.

**6. Generate a Table of Contents (TOC)**
* **Target Audience:** Users navigating large, 30+ tab financial models.
* **Logic Requirements:** Check if a sheet named "TOC" exists. If yes, delete it and create a new one at the front of the workbook. Loop through all worksheets. In the "TOC" sheet, list each sheet's name starting in cell A2. Convert each name into a working hyperlink jumping to cell A1 of that sheet.
* **Edge Cases:** Skip very hidden or hidden worksheets. Do not hyperlink the TOC sheet to itself.

**7. Consolidate Multiple Workbooks**
* **Target Audience:** Managers aggregating regional data or departmental budgets.
* **Logic Requirements:** Prompt the user to select a folder path. Use the `Dir` function to loop through all `.xlsx` files in that folder. Open each file, find the last used row, copy the data (excluding headers if appending), paste to the master workbook's next blank row, and close the source file without saving.
* **Edge Cases:** Prevent the macro from opening the master file if it is stored in the same folder as the source files.

**8. Protect/Unprotect All Sheets**
* **Target Audience:** Financial controllers securing models for distribution.
* **Logic Requirements:** Utilize an InputBox to ask the user for a password. Loop through all `Worksheets` in the `ActiveWorkbook`. Apply `.Protect` or `.Unprotect` using the provided password string.
* **Edge Cases:** Include error handling to gracefully inform the user if the "Unprotect" password entered is incorrect, rather than throwing a debug error.

### LEVEL 3: HARD (ADVANCED AUTOMATION)

**9. Batch Email Generation via Outlook**
* **Target Audience:** Accounts Receivable, Billing, and Client Communication teams.
* **Logic Requirements:** Assume a structured table (e.g., Col A: Name, Col B: Email, Col C: Amount Due). Use late binding (`CreateObject("Outlook.Application")`) to prevent broken references. Loop through the rows and generate an `.HTMLBody` email draft for each row, dynamically inserting the Name and Amount Due variables.
* **Edge Cases:** Skip rows where the Email column is blank or lacks an "@" symbol.

**10. Automated Monthly Roll-Forward**
* **Target Audience:** Accounting teams performing month-end close.
* **Logic Requirements:** Duplicate the active worksheet and place it at the end of the workbook. Parse the current sheet name (e.g., "Jan_2024") and mathematically increment the month to rename the new sheet ("Feb_2024"). Clear contents of any cell featuring a specific "Input" interior color (e.g., light yellow) while preserving formulas and formatting.
* **Edge Cases:** If the newly calculated sheet name already exists, prompt the user with a MsgBox to resolve the duplication.

**11. Batch PDF Generator**
* **Target Audience:** HR, Payroll, or Client Reporting teams.
* **Logic Requirements:** Loop through a predefined named range containing a list of unique identifiers (e.g., employee names). Paste each identifier into the dashboard's primary lookup cell. Trigger calculation. Export the resulting dashboard as a PDF using `ExportAsFixedFormat` to a specific folder path.
* **Edge Cases:** Strip out illegal file path characters (like /, \, *, ?, <, >) from the identifiers before saving the PDF names.

**12. Extract Unique Values to New Tabs**
* **Target Audience:** Financial analysts generating variance reports by department.
* **Logic Requirements:** Identify a target column in a master dataset. Extract all unique values from that column using a Scripting Dictionary or Advanced Filter. Loop through the unique values, filter the master dataset for that value, create a new worksheet, name the sheet after the value, and paste the filtered data.
* **Edge Cases:** Excel sheet names cannot exceed 31 characters. Truncate long unique values to 31 characters to prevent sheet-naming errors.

---

## Macro Summaries (Quick Reference Guide)

1. **Unmerge and Fill Down:** Removes merged cells and fills blank spaces with data from the cell directly above. It is incredibly useful for instantly prepping messy system exports for pivot tables and lookup formulas.
2. **Auto-Format Reports:** Applies a pre-defined set of visual styles, such as bolding headers and applying accounting formats. It guarantees visual consistency across financial reports and saves minutes of repetitive clicking.
3. **Highlight Hardcoded Numbers:** Scans a worksheet and changes manually typed numbers to a blue font while keeping formulas black. It allows analysts to instantly audit a file and see which inputs are driving the math.
4. **Toggle Presentation Mode:** Hides the Excel ribbon, gridlines, formula bar, and headers with a single click. It turns a chaotic working spreadsheet into a clean, professional dashboard ready for a presentation.
5. **Delete Completely Blank Rows:** Scans a dataset and removes any row containing absolutely no data. It sanitizes large data dumps so formulas and charts process without returning errors.
6. **Generate a Table of Contents:** Creates a master list of all worksheet tabs and links directly to them. It is highly useful for navigating massive, multi-department budget files without endlessly scrolling.
7. **Consolidate Multiple Workbooks:** Searches a folder, opens every Excel file inside, copies the data, and stacks it sequentially into one master sheet. It saves hours of copy-pasting when aggregating monthly reports from different teams.
8. **Protect/Unprotect All Sheets:** Applies or removes password protection across every worksheet simultaneously. It secures financial models before distributing them to external clients without locking individual tabs manually.
9. **Batch Email Generation:** Connects Excel to Outlook to automatically draft personalized messages based on spreadsheet data. It scales communications, allowing users to generate dozens of individualized client emails in one click.
10. **Automated Monthly Roll-Forward:** Duplicates a financial file, clears historical input data, and updates timeline headers for the new period. It eliminates the tedious, error-prone manual setup required at the start of every month-end close.
11. **Batch PDF Generator:** Cycles through a list of names, recalculates a dashboard for each, and automatically saves a PDF copy. It is perfect for generating hundreds of individualized employee or vendor statements in seconds.
12. **Extract Unique Values to New Tabs:** Identifies categories within a large master table and splits the data into separate, newly created tabs for each category. It is incredibly helpful for instantly generating departmental reports from a company-wide data dump.
