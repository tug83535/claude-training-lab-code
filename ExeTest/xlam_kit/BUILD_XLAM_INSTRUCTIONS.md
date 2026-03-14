# How to Build KBT_UniversalTools.xlam

## What This Is
This folder contains 23 VBA modules (~140+ tools) that work on ANY Excel file.
When packaged as a `.xlam` add-in, coworkers can install it once and have all
tools available from a menu in every workbook they open.

## The 23 Modules Included

| # | Module | Tools | What It Does |
|---|--------|-------|-------------|
| 1 | modUTL_Core | 5 | Shared utilities (styling, logging, performance) |
| 2 | modUTL_CommandCenter | 1 | Master menu launcher for all toolkit tools |
| 3 | modUTL_Audit | 6 | External link finder/fixer, hidden sheet audit, error scan |
| 4 | modUTL_Branding | 2 | iPipeline brand colors + theme setup |
| 5 | modUTL_ColumnOps | 7 | Column insert/delete/move/split/merge/fill/swap |
| 6 | modUTL_Comments | 3 | Extract/clear/convert comments and notes |
| 7 | modUTL_Compare | 1 | Sheet comparison with color-coded diff report |
| 8 | modUTL_Consolidate | 1 | Multi-sheet data consolidation with source tracking |
| 9 | modUTL_DataCleaning | 8 | Remove blanks, trim, dedupe, standardize dates |
| 10 | modUTL_DataSanitizer | 4 | Fix text-numbers, floating-point tails, sanitize |
| 11 | modUTL_ExecBrief | 1 | Auto-generate executive workbook brief |
| 12 | modUTL_Finance | 12 | Aging, amortization, depreciation, IRR, NPV, payback |
| 13 | modUTL_Formatting | 8 | Auto-format, alternating rows, freeze panes, print setup |
| 14 | modUTL_Highlights | 3 | Threshold, top/bottom N, duplicate highlighting |
| 15 | modUTL_LookupBuilder | 2 | VLOOKUP/INDEX-MATCH formula builder with preview |
| 16 | modUTL_PivotTools | 4 | PivotTable create/refresh/style/drill-down |
| 17 | modUTL_ProgressBar | 3 | Status bar progress indicator (ASCII visual) |
| 18 | modUTL_SheetTools | 3 | Sheet index with hyperlinks, template cloner, custom IDs |
| 19 | modUTL_SplashScreen | 1 | Branded welcome screen |
| 20 | modUTL_TabOrganizer | 6 | Sort/color/group/reorder/rename tabs in bulk |
| 21 | modUTL_ValidationBuilder | 5 | Data validation builder: lists, numbers, dates, custom |
| 22 | modUTL_WhatIf | 1 | What-If scenario analysis tool |
| 23 | modUTL_WorkbookMgmt | 5 | Backup, health check, sheet protection, cleanup |

**Total: ~140+ tools across 23 modules**

---

## Step-by-Step Build Instructions

### Step 1: Create a New Blank Workbook
1. Open Excel
2. Click **File > New > Blank Workbook**
3. Press **Alt+F11** to open the VBA Editor

### Step 2: Import All 23 Modules
1. In the VBA Editor, right-click on **VBAProject (Book1)** in the Project Explorer
2. Click **Import File...**
3. Navigate to this `xlam_kit` folder
4. Select **modUTL_Core.bas** and click **Open**
5. Repeat for ALL 23 `.bas` files in this folder

**Import order recommendation** (Core first, CommandCenter last):
1. modUTL_Core.bas (import FIRST — other modules reference it)
2. modUTL_ProgressBar.bas
3. modUTL_SplashScreen.bas
4. modUTL_Audit.bas
5. modUTL_Branding.bas
6. modUTL_ColumnOps.bas
7. modUTL_Comments.bas
8. modUTL_Compare.bas
9. modUTL_Consolidate.bas
10. modUTL_DataCleaning.bas
11. modUTL_DataSanitizer.bas
12. modUTL_ExecBrief.bas
13. modUTL_Finance.bas
14. modUTL_Formatting.bas
15. modUTL_Highlights.bas
16. modUTL_LookupBuilder.bas
17. modUTL_PivotTools.bas
18. modUTL_SheetTools.bas
19. modUTL_TabOrganizer.bas
20. modUTL_ValidationBuilder.bas
21. modUTL_WhatIf.bas
22. modUTL_WorkbookMgmt.bas
23. modUTL_CommandCenter.bas (import LAST — references all other modules)

### Step 3: Verify — Compile the Project
1. In the VBA Editor menu, click **Debug > Compile VBAProject**
2. If no errors appear, you're good
3. If you see an error, note which module/line and check that all 23 modules are imported

### Step 4: Save as .xlam Add-In
1. Close the VBA Editor (click X or press Alt+Q)
2. Click **File > Save As**
3. Change the file type dropdown to **Excel Add-In (*.xlam)**
4. Name it: **KBT_UniversalTools.xlam**
5. Save it to your Add-Ins folder:
   - Default location: `C:\Users\YourName\AppData\Roaming\Microsoft\AddIns\`
   - Or save anywhere and note the path
6. Click **Save**

### Step 5: Install the Add-In
1. Close all Excel workbooks
2. Open Excel fresh
3. Click **File > Options > Add-Ins**
4. At the bottom, make sure "Excel Add-ins" is selected in the dropdown
5. Click **Go...**
6. Click **Browse...**
7. Navigate to where you saved `KBT_UniversalTools.xlam`
8. Select it and click **OK**
9. Make sure the checkbox next to **KBT_UniversalTools** is checked
10. Click **OK**

### Step 6: Test It
1. Open any Excel file
2. Press **Alt+F8** (Macro dialog)
3. You should see all the KBT tools listed
4. Try running **UTL_ShowCommandCenter** — the master menu should appear
5. Pick any tool and confirm it works on your data

---

## How Coworkers Install It (Quick Version)
1. Copy `KBT_UniversalTools.xlam` to their computer
2. Open Excel > File > Options > Add-Ins > Go > Browse
3. Select the `.xlam` file > OK > check the box > OK
4. Done — tools available in every workbook from now on

## Uninstalling
1. File > Options > Add-Ins > Go
2. Uncheck **KBT_UniversalTools**
3. Click OK
4. Delete the `.xlam` file if desired
