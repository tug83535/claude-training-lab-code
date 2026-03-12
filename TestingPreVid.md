# Pre-Video Testing Guide — iPipeline P&L Demo File

> **For:** Connor (the user) + any Claude assistant with no prior context
> **Purpose:** Step-by-step guide to set up, import, build, and test the entire demo file before recording the video
> **File:** ExcelDemoFile_adv.xlsm
> **Date:** 2026-03-12
> **Repo:** github.com/tug83535/claude-training-lab-code
> **Branch:** claude/resume-ipipeline-demo-qKRHn

---

## CONTEXT FOR A NEW CLAUDE ASSISTANT

Connor works in Finance & Accounting at iPipeline. He is NOT a developer. He built a P&L demo Excel file with 39 VBA modules, 14 Python scripts, and 14 universal toolkit VBA modules. The demo will be presented to 2,000+ employees and the CFO/CEO. Your job is to guide Connor through every testing step below — one at a time, in plain English, no shortcuts. If something fails, help him fix it before moving on.

The VBA source code lives in the GitHub repo as `.bas` text files. They must be manually imported into the Excel workbook before they work. The Excel file is a binary `.xlsm` — Claude cannot read it directly.

**Key references in the repo:**
- `CLAUDE.md` — Full project context, module list, session history
- `qa/TEST_PLAN.md` — Official test plan with all 69 tests
- `qa/BUG_LOG.md` — All 36 bugs found and fixed to date
- `tasks/lessons.md` — Known bug patterns to watch for
- `tasks/todo.md` — Current task list and status

---

## PART 1: OPEN THE EXCEL FILE

### Step 1 — Open the file
1. Navigate to where you saved `ExcelDemoFile_adv.xlsm` on your computer
2. Double-click to open it in Excel
3. If you see a yellow "SECURITY WARNING — Macros have been disabled" bar at the top, click **Enable Content**
4. If you do NOT see the yellow bar, you may need to enable macros manually — see Step 2

### Step 2 — Enable macros (if not already)
1. Go to **File** > **Options** > **Trust Center** > **Trust Center Settings**
2. Click **Macro Settings** on the left
3. Select **Enable all macros** (or at minimum "Disable all macros with notification")
4. Check the box: **Trust access to the VBA project object model** (IMPORTANT — needed for BuildCommandCenter)
5. Click **OK** twice to close
6. **Close and reopen** the file for settings to take effect

### Step 3 — Verify the file opened correctly
1. Look at the sheet tabs at the bottom — you should see at minimum:
   - CrossfireHiddenWorksheet (may be hidden)
   - Assumptions
   - Data Dictionary
   - AWS Allocation
   - Report-->
   - P&L - Monthly Trend
   - Product Line Summary
   - Functional P&L - Monthly Trend
   - Functional P&L Summary - Jan 25
   - Functional P&L Summary - Feb 25
   - Functional P&L Summary - Mar 25
   - US January 2025 Natural P&L
   - Checks
2. That's 13 sheets total. If any are missing, stop and investigate.

---

## PART 2: IMPORT ALL 39 VBA MODULES

You need to import the `.bas` files from the `vba/` folder in the GitHub repo into the Excel workbook.

### Step 4 — Open the VBA Editor
1. Press **Alt + F11** to open the VBA Editor
2. In the left panel (Project Explorer), you should see your workbook name (e.g., "VBAProject (ExcelDemoFile_adv.xlsm)")

### Step 5 — Delete any old modules (if re-importing)
If you previously imported modules and are updating them:
1. In Project Explorer, right-click each module you need to update
2. Click **Remove [module name]**
3. When asked "Do you want to export the module before removing?" click **No**

### Step 6 — Import all 39 .bas files
1. Download all `.bas` files from the `vba/` folder in the repo to a local folder on your computer
2. In the VBA Editor, go to **File** > **Import File** (or right-click the project > **Import File**)
3. Navigate to your local folder
4. Select a `.bas` file and click **Open**
5. Repeat for ALL 39 files listed below

**Complete list of 39 .bas files to import (in recommended order):**

**Foundation (import these first):**
1. modConfig_v2.1.bas
2. modPerformance_v2.1.bas
3. modLogger_v2.1.bas

**Core features:**
4. modNavigation_v2.1.bas
5. modFormBuilder_v2.1.bas
6. modMasterMenu_v2.1.bas
7. modDashboard_v2.1.bas
8. modDashboardAdvanced_v2.1.bas
9. modDataQuality_v2.1.bas
10. modReconciliation_v2.1.bas
11. modVarianceAnalysis_v2.1.bas
12. modPDFExport_v2.1.bas
13. modMonthlyTabGenerator_v2.1.bas
14. modSearch_v2.1.bas
15. modUtilities_v2.1.bas

**Advanced features:**
16. modSensitivity_v2.1.bas
17. modAWSRecompute_v2.1.bas
18. modImport_v2.1.bas
19. modForecast_v2.1.bas
20. modScenario_v2.1.bas
21. modAllocation_v2.1.bas
22. modConsolidation_v2.1.bas
23. modVersionControl_v2.1.bas
24. modAdmin_v2.1.bas
25. modIntegrationTest_v2.1.bas

**New v2.1 modules:**
26. modDemoTools_v2.1.bas
27. modDataGuards_v2.1.bas
28. modDrillDown_v2.1.bas
29. modAuditTools_v2.1.bas
30. modETLBridge_v2.1.bas
31. modTrendReports_v2.1.bas
32. modDataSanitizer_v2.1.bas
33. modSheetIndex_v2.1.bas

**Optional add-ins:**
34. modTimeSaved_v2.1.bas
35. modSplashScreen_v2.1.bas
36. modProgressBar_v2.1.bas
37. modWhatIf_v2.1.bas
38. modExecBrief_v2.1.bas

**Note:** That's 38 files. The 39th "module" is the frmCommandCenter UserForm built in Part 3.

### Step 7 — Compile the project
1. In the VBA Editor, go to **Debug** > **Compile VBAProject**
2. If you get ZERO errors, you're good — move to Part 3
3. If you get an error, note the module name and line number and fix it before continuing

---

## PART 3: BUILD THE COMMAND CENTER

The Command Center is a UserForm (popup window) that lets you run all 62 actions. It must be built programmatically.

### Step 8 — Build frmCommandCenter (Mode A — automatic)
1. In the VBA Editor, press **Ctrl + G** to open the Immediate Window (bottom panel)
2. Type this and press Enter:
   ```
   modFormBuilder.BuildCommandCenter
   ```
3. You should see a message like "frmCommandCenter created successfully"
4. In Project Explorer, you should now see **frmCommandCenter** under the Forms folder

### Step 9 — Test the Command Center
1. Press **Ctrl + G** (Immediate Window) and type:
   ```
   modFormBuilder.LaunchCommandCenter
   ```
2. OR press **Ctrl+Shift+M** from the spreadsheet
3. The Command Center popup should appear showing all 62 actions
4. Try selecting a category from the dropdown — the list should filter
5. Try typing "variance" in the search box — should show relevant actions
6. Close the form for now

**If BuildCommandCenter fails (Mode B — manual fallback):**
1. In the Immediate Window, type: `modFormBuilder.GetFormInstallGuide`
2. Follow the printed instructions to manually create the UserForm
3. OR just use the InputBox fallback — press Ctrl+Shift+M, it will show a text menu instead

---

## PART 4: ASSIGN KEYBOARD SHORTCUTS

### Step 10 — Set up shortcuts
1. In the Immediate Window, type:
   ```
   modNavigation.AssignShortcuts
   ```
2. This assigns:
   - **Ctrl+Shift+H** — Go to Home/Report sheet
   - **Ctrl+Shift+J** — Jump to any sheet (shows list)
   - **Ctrl+Shift+R** — Go to Reconciliation Checks sheet
   - **Ctrl+Shift+M** — Open Command Center

---

## PART 5: RUN ALL TESTS — DEMO FILE

**How to run macros:** Either use the Command Center (Ctrl+Shift+M) or the Immediate Window (Alt+F11, then Ctrl+G, then type the macro name and press Enter).

For each test, record: PASS or FAIL + any notes.

---

### TEST CATEGORY T1 — Compilation & Load (8 tests)

| # | Test | What to Do | Expected Result | PASS/FAIL |
|---|------|-----------|-----------------|-----------|
| T1.01 | VBA compiles | Debug > Compile VBAProject | Zero errors | |
| T1.02 | 39 modules present | Count modules in Project Explorer | 39 modules (38 .bas + frmCommandCenter) | |
| T1.03 | Option Explicit | Search modules for "Option Explicit" | Found in every module | |
| T1.04 | modConfig loads | Immediate Window: `?APP_VERSION` | Returns "2.1.0" | |
| T1.05 | Python config | Command prompt: `python pnl_config.py` | Prints config, no errors | |
| T1.06 | Python imports | Command prompt: `python -c "import pnl_config; import pnl_runner; import pnl_forecast; import pnl_tests; import pnl_month_end; import pnl_snapshot; import pnl_allocation_simulator; import pnl_ap_matcher; import pnl_dashboard; import pnl_monte_carlo; import pnl_cli; print('All imports OK')"` | Prints "All imports OK" | |
| T1.07 | UTF-8 clean | Already verified — PASS from prior session | PASS (carry forward) | |
| T1.08 | requirements.txt | Command prompt: `pip install -r requirements.txt` | All packages install | |

**STOP if T1.01 fails.** Fix all compile errors before continuing.

---

### TEST CATEGORY T2 — Foundation Issues (7 tests)

| # | Test | What to Do | Expected Result | PASS/FAIL |
|---|------|-----------|-----------------|-----------|
| T2.01 | Config constants | Immediate Window: `?SH_GL` then `?SH_TECH_DOC` etc. | All 13 constants return values | |
| T2.02 | SafeDeleteSheet | Immediate Window: `Call SafeDeleteSheet("NonExistent")` | No error, no prompt | |
| T2.03 | StyleHeader | Immediate Window: `Call StyleHeader(ActiveSheet, 1, Array("Col A","Col B","Col C"))` | Row 1 gets navy background, white bold text | |
| T2.04 | UpdateHeaderText | Immediate Window: `Call TestUpdateHeaderText` | MsgBox: A1="Margin", A2="Market", A3="Apr 25" | |
| T2.05 | FixTextNumbers guard | Immediate Window: `Call modDataQuality.FixTextNumbers` (without running ScanAll first) | Message: "Run Scan Data Quality first" | |
| T2.06 | Shortcuts safe | Run AssignShortcuts, then press Ctrl+H | Excel's Find & Replace opens (not overridden) | |
| T2.07 | Timer rollover | Immediate Window: `modPerformance.m_StartTime = 86390` then `?modPerformance.ElapsedSeconds` | Returns positive number (not negative) | |

---

### TEST CATEGORY T3 — Menu & Command Center (5 tests)

| # | Test | What to Do | Expected Result | PASS/FAIL |
|---|------|-----------|-----------------|-----------|
| T3.01 | 62 items | Ctrl+Shift+M, select "All Actions" | 62 actions listed | |
| T3.02 | Form launches | Ctrl+Shift+M | frmCommandCenter popup appears | |
| T3.03 | Category filter | Select each category dropdown | List filters correctly | |
| T3.04 | Search filter | Type "variance" in search box | Shows relevant actions | |
| T3.05 | PDF dynamic | Run Command 10 (Export Report Package) | PDF includes all existing monthly tabs | |

---

### TEST CATEGORY T4 — Python Ecosystem (4 tests)

| # | Test | What to Do | Expected Result | PASS/FAIL |
|---|------|-----------|-----------------|-----------|
| T4.01 | UTF-8 clean | Already verified — PASS from prior session | PASS (carry forward) | |
| T4.02 | Config self-test | `python pnl_config.py` | Shares sum to 1.0, version 2.1.0 | |
| T4.03 | Runner help | `python pnl_runner.py --help` | Shows all 8 commands | |
| T4.04 | Pytest passes | `python -m pytest pnl_tests.py -v` | All non-skip tests pass (expect 99 pass, 15 skip) | |

---

### TEST CATEGORY T5 — Advanced VBA Features (6 tests)

| # | Test | What to Do | Expected Result | PASS/FAIL |
|---|------|-----------|-----------------|-----------|
| T5.01 | Exec Dashboard | Immediate Window: `modDashboardAdvanced.CreateExecutiveDashboard` | Dashboard sheet created with charts | |
| T5.02 | Waterfall chart | Immediate Window: `modDashboardAdvanced.WaterfallChart` | Chart on Dashboard sheet | |
| T5.03 | Product comparison | Immediate Window: `modDashboardAdvanced.ProductComparison` | Chart with 4 product series | |
| T5.04 | Commentary | Run Variance Analysis (Command 6) first, then: `modVarianceAnalysis.GenerateCommentary` | "Variance Commentary" sheet created | |
| T5.05 | Cross-sheet validation | Immediate Window: `modReconciliation.ValidateCrossSheet` | "Cross-Sheet Validation" sheet with PASS/FAIL | |
| T5.06 | Search cap | Immediate Window: `modSearch.CrossSheetSearch` then search for "a" | Shows "Showing first 200 of N total matches" | |

---

### TEST CATEGORY T6 — Data Integrity (6 tests)

| # | Test | What to Do | Expected Result | PASS/FAIL |
|---|------|-----------|-----------------|-----------|
| T6.01 | GL row count | Unhide CrossfireHiddenWorksheet, count rows | 510 data rows (511 with header) | |
| T6.02 | GL total | SUM the Amount column (column G) | $3,721,942.88 | |
| T6.03 | Products | Check unique values in Product column (D) | iGO, Affirm, InsureSight, DocFast only | |
| T6.04 | Departments | Check unique values in Department column (C) | 7 departments only | |
| T6.05 | Revenue shares | Immediate Window: `?REVENUE_SHARES(0) + REVENUE_SHARES(1) + REVENUE_SHARES(2) + REVENUE_SHARES(3)` | 1.000 | |
| T6.06 | Recon checks | Run Command 3 (Run Reconciliation) | All 9 checks show PASS | |

---

### TEST CATEGORY T7 — Integration (4 tests)

| # | Test | What to Do | Expected Result | PASS/FAIL |
|---|------|-----------|-----------------|-----------|
| T7.01 | Full test | Run Command 44 | Report generated on "Integration Test Report" sheet | |
| T7.02 | Quick health | Run Command 45 | Summary with PASS/FAIL/WARN counts | |
| T7.03 | Month close | Follow OPERATIONS_RUNBOOK steps 3.1-3.12 | All steps complete without error | |
| T7.04 | Python month-end | `python pnl_runner.py month-end --month 1` | CloseReport generated | |

---

### TEST CATEGORY T8 — New v2.1 Modules (36 tests)

**modDataGuards:**

| # | Test | What to Do | Expected | PASS/FAIL |
|---|------|-----------|----------|-----------|
| T8.01 | ValidateAssumptions | `modDataGuards.ValidateAssumptionsPresence` | True if all filled; False + list if blanks | |
| T8.02 | FindNegativeAmounts | `modDataGuards.FindNegativeAmounts` | Red highlights on negative GL cells + count | |
| T8.03 | FindZeroAmounts | `modDataGuards.FindZeroAmounts` | Yellow highlights on zero GL cells + count | |
| T8.04 | FindSuspiciousRound | `modDataGuards.FindSuspiciousRoundNumbers` | Orange highlights on round thousands | |

**modDataSanitizer:**

| # | Test | What to Do | Expected | PASS/FAIL |
|---|------|-----------|----------|-----------|
| T8.05 | Preview | `modDataSanitizer.PreviewSanitizeChanges` | Preview sheet created, no data changed | |
| T8.06 | Date safety | `modDataSanitizer.RunFullSanitize` | Dates untouched, text-numbers converted | |
| T8.07 | Header skip | Same as above | "Customer ID", "Date", "Name" columns skipped | |

**modAuditTools:**

| # | Test | What to Do | Expected | PASS/FAIL |
|---|------|-----------|----------|-----------|
| T8.08 | ChangeLog | `modAuditTools.AppendChangeLogEntry` (enter a note) | Entry with timestamp in Change Log sheet | |
| T8.09 | External links | `modAuditTools.FindExternalLinks` | "No external links found" message | |
| T8.10 | Hidden sheets | `modAuditTools.AuditHiddenSheets` | Lists all hidden/very-hidden sheets | |

**modMonthlyTabGenerator:**

| # | Test | What to Do | Expected | PASS/FAIL |
|---|------|-----------|----------|-----------|
| T8.11 | Next month | `modMonthlyTabGenerator.AddNextMonthToModel` | Popup shows correct next month | |
| T8.12 | Trend yellow | Confirm the popup | Next month column highlighted yellow | |
| T8.13 | New tab | Same | New "Functional P&L Summary - [Month] 25" tab | |

**modDemoTools:**

| # | Test | What to Do | Expected | PASS/FAIL |
|---|------|-----------|----------|-----------|
| T8.14 | Control buttons | `modDemoTools.AddControlSheetButtons` | 5 buttons on Report--> sheet | |
| T8.15 | Buttons work | Click each of the 5 buttons | Each runs correct macro, no errors | |
| T8.16 | Print area | `modDemoTools.SetParameterizedPrintArea` | Print area set, fits 1 page | |
| T8.17 | Exec summary | `modDemoTools.CreatePrintableExecSummary` | "Exec Summary - Print" sheet created | |

**modDrillDown:**

| # | Test | What to Do | Expected | PASS/FAIL |
|---|------|-----------|----------|-----------|
| T8.18 | Drill links | `modDrillDown.AddReconciliationDrillLinks` | Blue "View Data" links on Checks sheet | |
| T8.19 | Links navigate | Click any "View Data" link | Jumps to GL sheet, no error | |
| T8.20 | Auto-populate | `modDrillDown.AutoPopulateReconciliationChecks` | Checks refreshed with timestamp | |
| T8.21 | Heatmap | `modDrillDown.ApplyReconciliationHeatmap` | Color-coded Difference/Status columns | |
| T8.22 | Golden save | `modDrillDown.RunGoldenFileCompare` (first time) | Saves baseline, creates hidden sheet | |
| T8.23 | Golden compare | `modDrillDown.RunGoldenFileCompare` (second time) | Compare report with MATCH/CHANGED rows | |

**modETLBridge:**

| # | Test | What to Do | Expected | PASS/FAIL |
|---|------|-----------|----------|-----------|
| T8.24 | ETL no script | `modETLBridge.TriggerETLLocally` (no script file) | File browser opens, Cancel exits cleanly | |
| T8.25 | Import no file | `modETLBridge.ImportETLOutput` (no output file) | File browser opens, Cancel exits cleanly | |

**modTrendReports:**

| # | Test | What to Do | Expected | PASS/FAIL |
|---|------|-----------|----------|-----------|
| T8.26 | Rolling 12 | `modTrendReports.CreateRolling12MonthView` | Sheet + line chart created | |
| T8.27 | Archive recon | `modTrendReports.ArchiveReconciliationResults` > Yes | "Recon Archive" sheet with timestamps | |
| T8.28 | Trend no data | `modTrendReports.CreateReconciliationTrendChart` (before archive) | "No archive found" message | |
| T8.29 | Trend chart | `modTrendReports.CreateReconciliationTrendChart` (after archive) | Column chart with green/red bars | |

**modTimeSaved (Optional Add-In):**

| # | Test | What to Do | Expected | PASS/FAIL |
|---|------|-----------|----------|-----------|
| T8.34 | Time report | `modTimeSaved.TimeSavedReport` | "Time Saved Analysis" sheet with 62 actions | |

**modSplashScreen (Optional Add-In):**

| # | Test | What to Do | Expected | PASS/FAIL |
|---|------|-----------|----------|-----------|
| T8.35 | Splash | `modSplashScreen.ShowSplash` | Branded splash appears (form or MsgBox) | |

**modProgressBar (Optional Add-In):**

| # | Test | What to Do | Expected | PASS/FAIL |
|---|------|-----------|----------|-----------|
| T8.36 | Progress bar | `modProgressBar.StartProgress "Test", 10` then `modProgressBar.UpdateProgress 5, "Halfway"` then `modProgressBar.EndProgress` | Shows %, clears cleanly | |

**modWhatIf (Optional Add-In):**

| # | Test | What to Do | Expected | PASS/FAIL |
|---|------|-----------|----------|-----------|
| T8.37 | WhatIf preset | `modWhatIf.RunWhatIfDemo` > "Revenue +15%" | Baseline saved, impact report created | |
| T8.38 | Restore | `modWhatIf.RestoreBaseline` | All values restored, impact sheets removed | |

**modExecBrief (Optional Add-In):**

| # | Test | What to Do | Expected | PASS/FAIL |
|---|------|-----------|----------|-----------|
| T8.39 | Exec brief | `modExecBrief.GenerateExecBrief` | "Executive Brief" sheet with 5 sections | |
| T8.40 | Missing sheets | Hide a sheet, run GenerateExecBrief again | No crash, shows "N/A" for missing section | |

**Bug Fix Verifications:**

| # | Test | What to Do | Expected | PASS/FAIL |
|---|------|-----------|----------|-----------|
| T8.30 | Button macros | Run AddControlSheetButtons, check OnAction | Correct macro names (ScanAll, ExportReportPackage) | |
| T8.31 | ClearShortcuts | `modNavigation.ClearShortcuts` | Runs with no error | |
| T8.32 | Chart range | Run CreateRolling12MonthView, right-click chart > Select Data | Values match revenue row | |
| T8.33 | StyleHeader | Run GenerateCommentary (after Variance Analysis) | Navy/white header, no error | |

---

## PART 6: UNIVERSAL TOOLKIT TESTING

The Universal Toolkit is a separate set of 16 VBA modules in `UniversalToolsForAllFiles/vba/`. These are designed to work on ANY Excel file, not just the demo. To test them:

### Step 11 — Open a separate test workbook
1. Open a blank new Excel workbook OR the sample file `videodraft/Sample_Quarterly_Report.xlsx`
2. Import the universal modules into THIS workbook (not the demo file)

### Step 12 — Import universal modules
Import these 16 `.bas` files from `UniversalToolsForAllFiles/vba/`:

**Core:**
1. modUTL_Core.bas

**Main modules:**
2. modUTL_Branding.bas
3. modUTL_Finance.bas
4. modUTL_Formatting.bas
5. modUTL_DataCleaning.bas
6. modUTL_Audit.bas
7. modUTL_DataSanitizer.bas
8. modUTL_WorkbookMgmt.bas
9. modUTL_SheetTools.bas
10. modUTL_ProgressBar.bas
11. modUTL_ExecBrief.bas
12. modUTL_SplashScreen.bas

**NewTools (extras):**
13. modUTL_DataCleaningPlus.bas
14. modUTL_AuditPlus.bas
15. modUTL_DuplicateDetection.bas
16. modUTL_NumberFormat.bas

### Step 13 — Compile
Debug > Compile VBAProject — should be zero errors.

### Step 14 — Test key universal tools
These are spot-checks on the most important tools. Run each from the Immediate Window:

| # | Tool | What to Do | Expected | PASS/FAIL |
|---|------|-----------|----------|-----------|
| U1 | Branding | `modUTL_Branding.ApplyiPipelineBranding` | Active sheet gets iPipeline colors/fonts | |
| U2 | Sheet index | `modUTL_SheetTools.ListAllSheetsWithLinks` | UTL_SheetIndex sheet with clickable links | |
| U3 | Template clone | `modUTL_SheetTools.TemplateCloner` | Prompts for sheet + count, creates copies | |
| U4 | Sanitize preview | `modUTL_DataSanitizer.PreviewSanitizeChanges` | Preview sheet, no data changed | |
| U5 | Full sanitize | `modUTL_DataSanitizer.RunFullSanitize` | Text-numbers converted, dates untouched | |
| U6 | External links | `modUTL_Audit.ExternalLinkFinder` | Lists external links or "none found" | |
| U7 | Hidden sheets | `modUTL_Audit.HiddenSheetAuditor` | Lists hidden sheets | |
| U8 | Progress bar | `modUTL_ProgressBar.UTL_StartProgress "Test", 5` then `modUTL_ProgressBar.UTL_UpdateProgress 3, "Step 3"` then `modUTL_ProgressBar.UTL_EndProgress` | Status bar shows ASCII progress | |
| U9 | Splash screen | `modUTL_SplashScreen.UTL_ShowSplash` | MsgBox splash appears | |
| U10 | Exec brief | `modUTL_ExecBrief.UTL_GenerateExecBrief` | Report sheet with workbook overview | |
| U11 | Workbook backup | `modUTL_WorkbookMgmt.BuildDistributionReadyCopy` | Clean copy saved (no macros) | |
| U12 | Duplicate detect | `modUTL_DuplicateDetection.HighlightDuplicatesInColumn` | Duplicates highlighted in selected column | |

---

## PART 7: POST-TEST CHECKLIST

After all tests pass:

- [ ] All T1-T8 tests recorded (PASS/FAIL)
- [ ] All universal tool spot-checks done
- [ ] No compile errors remain
- [ ] Save the workbook (Ctrl+S)
- [ ] Close and reopen — confirm Command Center still works (Ctrl+Shift+M)
- [ ] Check for any leftover test sheets (delete: Sanitizer Preview, Golden Compare Report, etc. if you don't want them in the demo)
- [ ] File is ready for video recording

---

## TROUBLESHOOTING

| Problem | Solution |
|---------|----------|
| "Compile error" on import | One module references another that isn't imported yet. Import modConfig first, then modPerformance, then modLogger, then the rest. |
| "Sub or Function not defined" | The module containing that sub isn't imported. Check the module list above. |
| "Subscript out of range" | A sheet name doesn't match what the code expects. Check modConfig constants (SH_GL, SH_ASSUMPTIONS, etc.) against your actual sheet names. |
| Command Center doesn't appear | Run `modFormBuilder.BuildCommandCenter` first. If that fails, check that "Trust access to the VBA project object model" is enabled in Trust Center. |
| Macro runs but nothing happens | Check if screen updating is frozen: in Immediate Window type `Application.ScreenUpdating = True` |
| "Expected array" error in modConfig | The `lastRow` variable was renamed to `lRow` in AddNamedRanges. Make sure you have the latest version of modConfig_v2.1.bas from the repo. |
| Python scripts won't import | Make sure you're running from the `python/` folder in the repo, and `pip install -r requirements.txt` was run first. |

---

## KNOWN TEST EXCEPTIONS

These are acceptable to SKIP or may need special handling:
- **T1.07** — Already PASS from prior session (UTF-8 verified)
- **T4.01** — Same as T1.07, carry forward
- **T8.22** — Can only run once (saves baseline). Second run becomes T8.23.
- **T8.24/T8.25** — These test "file not found" behavior. If you HAVE the ETL files, the test behaves differently (it runs the ETL instead).
- **T8.28** — Must be run BEFORE T8.27 (archive). If you already ran archive, this test is superseded.
- **T8.40** — Edge case test. Hide a sheet first, then run ExecBrief.

---

## PASS CRITERIA FOR VIDEO READINESS

To be demo-ready, you need:
- **T1:** ALL 8 PASS
- **T2:** ALL 7 PASS
- **T3:** At least 4 of 5 PASS
- **T4:** ALL 4 PASS
- **T5:** At least 5 of 6 PASS
- **T6:** ALL 6 PASS
- **T7:** At least 3 of 4 PASS
- **T8:** At least 30 of 36 PASS
- **Universal:** At least 10 of 12 spot-checks PASS

**Total minimum: 62 of 69 demo tests + 10 of 12 universal tests**

If you hit this bar, the file is ready to record.
