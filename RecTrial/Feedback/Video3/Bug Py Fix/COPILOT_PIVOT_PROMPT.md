# COPILOT PIVOT TABLE PROMPT
# Use this in Excel Copilot after running fix_sample_file.py
# Open Sample_Quarterly_ReportV2_FIXED.xlsx in Excel first, then paste this prompt.

---

## STEP 1 — Open the file in Excel
Open `Sample_Quarterly_ReportV2_FIXED.xlsx` in Excel.

---

## STEP 2 — Paste this into Excel Copilot

Create two pivot tables using the data on the Q1 Revenue sheet (columns A through I, rows 1 through 42).

Pivot table 1:
- Place it on a new sheet named "Pivot_SalesByRegion"
- Rows: Region
- Values: Sum of Amount, formatted as currency with no decimal places
- Sort by Sum of Amount descending
- Give it a title in cell A1: "Q1 Sales by Region"

Pivot table 2:
- Place it on a new sheet named "Pivot_SalesByRep"
- Rows: Sales Rep
- Values: Sum of Amount, formatted as currency with no decimal places
- Sort by Sum of Amount descending
- Give it a title in cell A1: "Q1 Sales by Rep"

Name the pivot tables "PivotSalesByRegion" and "PivotSalesByRep" respectively.

---

## STEP 3 — Verify

After Copilot finishes:
- Click on any cell inside each pivot table
- Confirm the PivotTable Fields pane appears on the right
- Confirm both new sheet tabs exist (Pivot_SalesByRegion, Pivot_SalesByRep)
- Save the file as Sample_Quarterly_ReportV2_FIXED.xlsx (overwrite)

---

## STEP 4 — Test with VBA tool

In the .xlsm file, run the "List All Pivot Tables" tool manually (Alt+F8) to confirm
it detects both pivot tables before re-recording.

---

## WHY THIS MATTERS

The demo video shows "List All Pivot Tables" finding zero results because the sample
file had no pivot tables. For a financial demo file this looks broken on camera.
These two pivot tables give the tool real findings to display, making the demo
convincing and professional.
