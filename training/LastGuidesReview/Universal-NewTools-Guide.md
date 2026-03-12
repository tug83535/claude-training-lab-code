# Universal Toolkit — New Tools Guide (9 Modules)

> **Audience:** Finance & Accounting staff using the Universal Toolkit
> **Version:** 1.0 | **Date:** 2026-03-12
> **Total new tools:** 38 tools across 9 modules

---

## Quick Reference — All 9 New Modules

| Module | Tools | What It Does |
|--------|-------|-------------|
| Compare | 3 | Compare two sheets or ranges cell-by-cell |
| Consolidate | 2 | Combine multiple sheets into one |
| Pivot Tools | 4 | Refresh, list, and manage pivot tables |
| Comments | 4 | Extract, delete, and count cell comments |
| Column Ops | 4 | Split, combine, extract, and swap columns |
| Highlights | 5 | Quick conditional highlighting tools |
| Tab Organizer | 6 | Color, hide/unhide, reorder, rename tabs |
| Validation Builder | 6 | Create dropdowns, number/date rules, find violations |
| Lookup Builder | 4 | Auto-build VLOOKUP and INDEX-MATCH formulas |

**How to run any tool:** Press **Alt+F8**, type the tool name, click **Run**.
Or use the **Universal Command Center** (`LaunchCommandCenter`) to browse by category.

---

## 1. Sheet & Range Comparison (modUTL_Compare)

### CompareSheets
**What it does:** Picks two sheets and compares every cell. Shows differences with optional red highlighting and a summary report.

**Step by step:**
1. Press **Alt+F8** → type `CompareSheets` → click **Run**
2. A numbered list of all sheets appears
3. Type the **number** of the first sheet → press **Enter**
4. Type the **number** of the second sheet → press **Enter**
5. Choose whether to highlight differences in red on the sheets:
   - **YES** = Red highlights appear directly on both sheets
   - **NO** = Only creates the summary report (no changes to your data)
6. Wait for the comparison to finish
7. The results show: total cells compared, matches, differences, match rate %
8. A report sheet (`UTL_CompareReport`) is created with every difference listed

**What the report shows:**
- Source 1 and Source 2 names
- Total cells, matches, differences, match rate %
- Detail table: Cell address, Value in Sheet 1, Value in Sheet 2

### CompareRanges
**What it does:** Same as CompareSheets but you select specific ranges instead of entire sheets. Ranges can be on different sheets.

**Step by step:**
1. Press **Alt+F8** → type `CompareRanges` → click **Run**
2. Click OK, then **select the first range** (click and drag)
3. Click OK, then **select the second range**
4. If sizes don't match, you'll be asked whether to compare anyway
5. Choose whether to highlight differences
6. See the results

### ClearCompareHighlights
**What it does:** Removes the red highlighting from a previous comparison.

**Options:**
- **YES** = Clear from active sheet only
- **NO** = Clear from all sheets

---

## 2. Sheet Consolidation (modUTL_Consolidate)

### ConsolidateSheets
**What it does:** Combines data from multiple sheets into one master sheet. Adds a "Source Sheet" column so you know where each row came from.

**Step by step:**
1. Press **Alt+F8** → type `ConsolidateSheets` → click **Run**
2. A numbered list of all sheets appears
3. Type the sheet numbers separated by commas → press **Enter**
   - Example: `1,3,5,6`
4. Choose whether sheets have headers:
   - **YES** = Headers from the first sheet are kept; headers on other sheets are skipped
   - **NO** = All rows from every sheet are copied
5. Choose whether to add a "Source Sheet" column:
   - **YES** = Adds a column showing which sheet each row came from (recommended)
   - **NO** = No source tracking
6. Results appear on the `UTL_Consolidated` sheet

**Important:** All sheets being combined should have the same column structure.

### ConsolidateByPattern
**What it does:** Same as ConsolidateSheets but finds sheets by keyword instead of number.

**Step by step:**
1. Press **Alt+F8** → type `ConsolidateByPattern` → click **Run**
2. Type a keyword (e.g., `Q1`, `2025`, `Jan`)
3. All matching sheet names are shown — confirm to proceed
4. Choose header handling
5. Data is combined on the `UTL_Consolidated` sheet

---

## 3. Pivot Table Utilities (modUTL_PivotTools)

### RefreshAllPivots
**What it does:** Refreshes every pivot table in the workbook. Unlike Excel's "Refresh All" button, this only touches pivot tables — it does NOT trigger external connections or Power Query refreshes.

**Step by step:**
1. Press **Alt+F8** → type `RefreshAllPivots` → click **Run**
2. Shows how many pivots were found — confirm to proceed
3. All pivots are refreshed — shows count of successful vs failed

### RefreshSelectedPivots
**What it does:** Shows a numbered list of all pivot tables. You pick which ones to refresh.

**Step by step:**
1. Press **Alt+F8** → type `RefreshSelectedPivots` → click **Run**
2. A numbered list shows every pivot with its sheet name
3. Type the numbers to refresh (comma-separated), or type `ALL`
4. Only the selected pivots are refreshed

### ListAllPivots
**What it does:** Creates a detailed inventory sheet of all pivot tables including their source ranges, row/column/data fields, and locations.

### ClearOldPivotCache
**What it does:** Reports on orphaned pivot caches (old caches from deleted pivots that bloat file size). Explains how to clean them up safely via Save As.

---

## 4. Comment & Note Manager (modUTL_Comments)

### ExtractAllComments
**What it does:** Exports every comment in the workbook to a styled summary sheet. Shows the sheet name, cell address, cell value, author, and comment text.

**When to use:** Before cleaning up comments, or when you need to review all comments at once (instead of clicking through each cell).

### DeleteSheetComments
**What it does:** Shows all sheets with comment counts. You pick which sheets to clear.

**Step by step:**
1. Run the tool
2. See the list with comment counts per sheet
3. Type sheet numbers (comma-separated) or `ALL`
4. Confirm deletion — this cannot be undone

**Tip:** Run `ExtractAllComments` first to save a backup of all comments.

### DeleteAllComments
**What it does:** Deletes every comment in the entire workbook. Requires double confirmation (button click + type "DELETE").

### CountComments
**What it does:** Quick popup showing comment counts per sheet. No changes are made.

---

## 5. Column Operations (modUTL_ColumnOps)

### SplitColumn
**What it does:** Splits one column into multiple by a delimiter. New columns are inserted to the right — your existing data is NOT overwritten.

**Step by step:**
1. Run the tool
2. Select the data cells to split (not the header)
3. Choose a delimiter:
   - Comma, semicolon, space, dash, pipe, or type your own
4. The tool finds the maximum number of parts and inserts that many new columns
5. The first part stays in the original cell; remaining parts go into new columns

**Example:** "Smith, John" with comma delimiter → "Smith" | "John"

### CombineColumns
**What it does:** Merges multiple columns into one new column with a separator you choose.

**Step by step:**
1. Run the tool
2. Select all columns to combine (multi-column selection)
3. Choose a separator: comma+space, space, dash, pipe, none, or custom
4. A new column is inserted to the right with the combined values

**Example:** "John" + "Smith" with space separator → "John Smith"

### ExtractPattern
**What it does:** Extracts specific content from text cells into a new column.

**Options:**
1. First number found (e.g., "Invoice #12345" → "12345")
2. All numbers concatenated (e.g., "A1B2C3" → "123")
3. Text before a delimiter (e.g., "Smith, John" before comma → "Smith")
4. Text after a delimiter (e.g., "Smith, John" after comma → "John")
5. First N characters (e.g., first 3 chars of "ABCDEF" → "ABC")
6. Last N characters (e.g., last 4 chars of "ABCDEF" → "CDEF")

### SwapColumns
**What it does:** Swaps the contents of two columns. Both must have the same number of rows.

---

## 6. Conditional Highlighting (modUTL_Highlights)

### HighlightByThreshold
**What it does:** Highlights cells above, below, or equal to a value you type.

**Colors:**
- Above threshold = light green
- Below threshold = light red
- Equal to threshold = yellow

**Step by step:**
1. Select the range of numbers
2. Type the threshold value (e.g., 1000)
3. Choose: above only, below only, both, or equal to
4. Matching cells are highlighted immediately

### HighlightTopBottom
**What it does:** Highlights the top N or bottom N values in a range.

**Step by step:**
1. Select the range
2. Choose: top N, bottom N, or both
3. Type how many (e.g., 5, 10, 20)
4. Green = top values, Red = bottom values

### HighlightDuplicateValues
**What it does:** Highlights cells with duplicate values in orange.

### ApplyColorScale
**What it does:** Applies a red-yellow-green gradient across a range of numbers.

**Options:**
1. Low=Red, High=Green (higher is better — revenue, scores)
2. Low=Green, High=Red (lower is better — costs, errors)

### ClearHighlights
**What it does:** Removes cell background colors from selection, active sheet, or all sheets.

---

## 7. Tab Organizer (modUTL_TabOrganizer)

### ColorTabsByKeyword
**What it does:** Type a keyword — all tabs containing that text get colored.

**Example:** Type "Q1" → all tabs with "Q1" in the name turn blue.

**Colors available:** Blue, Green, Red, Orange, Purple, Yellow, or remove color.

### ColorTabsInteractive
**What it does:** Pick specific sheets from a numbered list and assign a color.

### BulkHideSheets
**What it does:** Hide multiple sheets at once from a numbered list. At least 1 sheet must remain visible.

### BulkUnhideSheets
**What it does:** Shows all hidden sheets (including Very Hidden) and lets you pick which to unhide. Type `ALL` to unhide everything.

### ReorderTabs
**What it does:** Move selected sheets to the front, back, or after a specific sheet.

### BulkRenameTabs
**What it does:** Find and replace text in tab names. Shows a preview before applying.

**Example:** Replace "2025" with "2026" in all tab names.

---

## 8. Data Validation Builder (modUTL_ValidationBuilder)

### CreateDropdownList
**What it does:** Creates a dropdown list on selected cells.

**Two source options:**
1. **Select a range** — values from a cell range become the dropdown options
2. **Type the options** — comma-separated list (e.g., "Yes,No,Maybe")

### ApplyNumberValidation
**What it does:** Restricts cells to numbers only.

**Rule options:**
1. Any number (no limits)
2. Between min and max
3. Greater than a minimum
4. Less than a maximum
5. Whole numbers only (no decimals)

### ApplyDateValidation
**What it does:** Restricts cells to dates only.

**Rule options:**
1. Any date
2. Between two dates
3. After a specific date
4. Before a specific date

### CopyValidationRules
**What it does:** Copies validation rules from one cell to a target range. No need to recreate rules from scratch.

### FindValidationViolations
**What it does:** Scans for cells that break their validation rules. Checks active sheet or all sheets.

### RemoveAllValidation
**What it does:** Removes data validation from selection, active sheet, or all sheets. Requires confirmation.

---

## 9. Lookup Builder (modUTL_LookupBuilder)

### BuildVLOOKUP
**What it does:** Builds VLOOKUP formulas step by step — no formula knowledge needed.

**Step by step:**
1. Select your **lookup keys** (the values you want to look up)
2. Select the **source table** (first column must contain matching keys)
3. Pick which column has the value you want (shown with column headers)
4. Select where to put the results
5. Formulas are written automatically, wrapped in IFERROR (shows blank if no match)

**When to use:** When you have IDs/names on one sheet and want to pull data from another sheet.

### BuildINDEXMATCH
**What it does:** Same concept as VLOOKUP but more flexible — the lookup column doesn't need to be first, and you can look left.

**Step by step:**
1. Select your lookup keys
2. Select the match column (where to search in the source data)
3. Select the return column (the values to pull)
4. Select where to put the results

**When to use VLOOKUP vs INDEX-MATCH:**
- VLOOKUP: Simpler setup, lookup key must be in the first column
- INDEX-MATCH: More flexible, lookup key can be anywhere, can look left

### MatchAndPull
**What it does:** Compares two lists and pulls matching values — no formulas involved, just values.

**Example:** You have employee IDs on Sheet1 and a table with IDs + departments on Sheet2. This tool pulls the department onto Sheet1 where IDs match.

**Step by step:**
1. Select your key column (the IDs/names you have)
2. Select the source key column (matching IDs in the source data)
3. Select the source value column (data to pull)
4. Select where to put the results
5. Values are written directly (no formulas)

### FindUnmatched
**What it does:** Finds values in List A that don't exist in List B.

**Example:** Find customers in your list that aren't in the master customer list.

**Options for unmatched values:**
1. Highlight them yellow on List A
2. List them on a new sheet (`UTL_Unmatched`)
3. Both

---

## Troubleshooting

### "Run-time error '1004'" on any tool
**Cause:** Usually a protected sheet or a range you don't have access to.
**Fix:** Unprotect the sheet first, then re-run the tool.

### Compare shows 100% match but data looks different
**Cause:** The tool uses a floating-point tolerance of 0.0001. Very small differences (like 100.00001 vs 100.00002) are treated as matches.
**Fix:** This is by design to avoid false positives from Excel's floating-point math.

### Consolidate is missing rows
**Cause:** If "Skip headers" is YES, the first row of sheets 2+ is skipped.
**Fix:** Choose "NO" for headers if your data starts on row 1 with no header.

### VLOOKUP returns blanks for everything
**Cause:** The lookup key column must be the **first column** of the source table range. If you selected columns B-F but the key is in column A, VLOOKUP can't find it.
**Fix:** Expand your source table selection to include the key column as the first column. Or use **INDEX-MATCH** instead (no first-column requirement).

### Tab rename fails for some sheets
**Cause:** The new name is a duplicate of an existing sheet, or contains illegal characters (\, /, *, ?, [, ]).
**Fix:** Try a different replacement text.
