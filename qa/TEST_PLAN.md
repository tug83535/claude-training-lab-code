# KBT P&L Toolkit — Test Plan

> **Version:** 2.1.0
> **Date:** 2026-03-01
> **Scope:** Verification of all 15 pre-audit issues, integration testing, regression testing, and 8 new v2.1 modules.

---

## 1. Test Scope

### In Scope
- All 15 pre-audit issues (ISSUE-001 through ISSUE-015)
- VBA module compilation and cross-module dependency verification
- Python script import chain and UTF-8 encoding verification
- Workbook data integrity (GL totals, reconciliation checks, formula errors)
- Command Center functionality (all 62 actions reachable)
- Documentation completeness

### Out of Scope
- Performance benchmarking (covered separately if needed)
- Multi-user concurrent access testing
- Network/SharePoint deployment testing
- Non-Windows platforms (Mac Excel has known VBA limitations)

---

## 2. Test Environments

### Primary: Excel for Microsoft 365 (Windows)

| Component | Version |
|-----------|---------|
| OS | Windows 10/11 |
| Excel | Microsoft 365 (latest channel) |
| VBA | 7.1 |
| Python | 3.11+ |
| Workbook | ExcelDemoFile_adv.xlsm |

### Secondary: Excel 2019 (Windows)

| Component | Version |
|-----------|---------|
| OS | Windows 10 |
| Excel | Excel 2019 (16.0) |
| VBA | 7.1 |

### Known Excel 2019 Limitations
- Dynamic arrays (`FILTER`, `SORT`, `UNIQUE`) not available — toolkit does not use these
- `XLOOKUP` not available — toolkit uses `INDEX/MATCH` instead
- Power Query may have fewer connectors — not relevant for core VBA functionality

---

## 3. Test Categories

### Category T1 — Compilation & Load (8 tests)

Verifies all 32 VBA modules compile without error and all Python scripts import cleanly.

| Test ID | Test Name | Procedure | Pass Criteria |
|---------|-----------|-----------|---------------|
| T1.01 | VBA project compiles | Open VBA Editor → Debug → Compile VBAProject | Zero compile errors |
| T1.02 | All 32 modules present | Check Project Explorer modules list | All 32 module names visible |
| T1.03 | Option Explicit on all | Search all modules for "Option Explicit" | Found in every module |
| T1.04 | modConfig loads | Immediate Window: `?APP_VERSION` | Returns "2.1.0" |
| T1.05 | Python pnl_config imports | `python pnl_config.py` | Prints config summary, no errors |
| T1.06 | All Python scripts import | `python -c "import pnl_config; import pnl_runner; import pnl_forecast; import pnl_tests; import pnl_month_end; import pnl_snapshot; import pnl_allocation_simulator; import pnl_ap_matcher; import pnl_dashboard; import pnl_monte_carlo; import pnl_cli; print('All imports OK')"` | Prints "All imports OK" — Zero ImportError |
| T1.07 | Python UTF-8 clean | Run scan_chars.py on all .py files | Zero mojibake codepoints (Ã/â garbled sequences). **NOTE — 2026-03-02:** All 14 files confirmed valid UTF-8. Non-ASCII bytes present are legitimate intentional Unicode: em dashes (—), arrows (→ ↑ ↓ ↔), check/warning marks (✓ ✗ ⚠), box-drawing chars (─ ║ ╔), Greek letters (Δ α), and emoji (📊). No mojibake sequences detected. **This test is PASS.** A scan that flags ALL non-ASCII bytes is too strict — only garbled Latin-1 misread sequences (Ã©, â€", etc.) constitute failure. Update any scan script to check for mojibake patterns specifically, not just any non-ASCII byte. |
| T1.08 | requirements.txt complete | `pip install -r requirements.txt` | All packages install |

### Category T2 — Foundation Issues (7 tests)

Verifies ISSUE-001 through ISSUE-007 fixes.

| Test ID | Test Name | Issue | Procedure | Pass Criteria |
|---------|-----------|-------|-----------|---------------|
| T2.01 | modConfig has all constants | ISSUE-001 | `?SH_GL`, `?SH_TECH_DOC`, etc. in Immediate Window | All 13 new constants return values |
| T2.02 | SafeDeleteSheet works | ISSUE-001 | Call `SafeDeleteSheet("NonExistent")` | No error, no prompt |
| T2.03 | StyleHeader works | ISSUE-001 | `Call StyleHeader(ActiveSheet, 1, Array("Col A","Col B","Col C"))` — note: requires all 3 arguments (ws, headerRow, headers array) | Navy background, white bold text |
| T2.04 | UpdateHeaderText safe | ISSUE-002 | In Immediate Window run `Call TestUpdateHeaderText` — MsgBox shows results | A1 = "Margin" (unchanged), A2 = "Market" (unchanged), A3 = "Apr 25" (replaced) |
| T2.05 | FixTextNumbers requires scan | ISSUE-003 | Call FixTextNumbers without running ScanAll first | Shows "Run Scan Data Quality first" message |
| T2.06 | Shortcuts use OnKey | ISSUE-004 | Run AssignShortcuts, press Ctrl+H | Excel Find & Replace opens (not overridden) |
| T2.07 | Timer midnight rollover | ISSUE-005 | Set m_StartTime = 86390 (near midnight), call ElapsedSeconds after Timer < m_StartTime | Returns positive value (not negative) |

### Category T3 — Menu & Command Center (5 tests)

| Test ID | Test Name | Issue | Procedure | Pass Criteria |
|---------|-----------|-------|-----------|---------------|
| T3.01 | 62 items in menu | ISSUE-006 | Press Ctrl+Shift+M, count items in "All Actions" | Shows 62 actions |
| T3.02 | UserForm launches | ISSUE-006 | Press Ctrl+Shift+M | frmCommandCenter appears (or InputBox fallback) |
| T3.03 | Category filtering | — | Select each category in the form | Action list filters correctly |
| T3.04 | Search filtering | — | Type "variance" in search box | Shows relevant actions only |
| T3.05 | PDF uses dynamic names | ISSUE-007 | Run Command 10 (Export Report Package) | PDF includes all existing monthly tabs |

### Category T4 — Python Ecosystem (4 tests)

| Test ID | Test Name | Issue | Procedure | Pass Criteria |
|---------|-----------|-------|-----------|---------------|
| T4.01 | UTF-8 clean across all files | ISSUE-008 | Scan for mojibake sequences: Latin-1 misread patterns (Ã©, â€", Ã¢, etc.) | Zero mojibake hits. Non-ASCII Unicode (em dashes, arrows, check marks, box-drawing, Greek letters, emoji) is intentional and acceptable. See T1.07 note. |
| T4.02 | pnl_config self-test | ISSUE-008 | `python pnl_config.py` | All shares sum to 1.0, version 2.1.0 |
| T4.03 | pnl_runner dispatches | — | `python pnl_runner.py --help` | Shows all 8 commands |
| T4.04 | pytest suite passes | — | `python -m pytest pnl_tests.py -v` | All non-skip tests pass |

### Category T5 — Advanced VBA Features (6 tests)

| Test ID | Test Name | Issue | Procedure | Pass Criteria |
|---------|-----------|-------|-----------|---------------|
| T5.01 | Executive Dashboard renders | ISSUE-009 | Run CreateExecutiveDashboard | Dashboard sheet created with charts |
| T5.02 | Waterfall chart renders | ISSUE-009 | Run WaterfallChart | Chart created on Dashboard sheet |
| T5.03 | Product comparison renders | ISSUE-009 | Run ProductComparison | Chart with 4 product series |
| T5.04 | GenerateCommentary output | ISSUE-010 | Run GenerateCommentary (after Variance Analysis) | "Variance Commentary" sheet with narrative text |
| T5.05 | ValidateCrossSheet output | ISSUE-011 | Run ValidateCrossSheet | "Cross-Sheet Validation" sheet with PASS/FAIL rows |
| T5.06 | Search cap warning | ISSUE-012 | Search for a common term (e.g., "a") that exceeds 200 results | Shows "Showing first 200 of N total matches" |

### Category T6 — Data Integrity (6 tests)

| Test ID | Test Name | Procedure | Pass Criteria |
|---------|-----------|-----------|---------------|
| T6.01 | GL row count | Count rows in CrossfireHiddenWorksheet | 510 data rows (511 including header) |
| T6.02 | GL total amount | SUM(Amount column) | $3,721,942.88 |
| T6.03 | All products known | DISTINCT products in GL | iGO, Affirm, InsureSight, DocFast only |
| T6.04 | All departments known | DISTINCT departments in GL | 7 known departments only |
| T6.05 | Revenue shares sum | SUM of REVENUE_SHARES values | 1.000 |
| T6.06 | Reconciliation checks | Run Command 3 | PASS count documented; FAIL items have known explanations |

### Category T7 — Integration (4 tests)

| Test ID | Test Name | Procedure | Pass Criteria |
|---------|-----------|-----------|---------------|
| T7.01 | RunFullTest | Run Command 44 | Report generated, all categories tested |
| T7.02 | QuickHealthCheck | Run Command 45 | Summary with PASS/FAIL/WARN counts |
| T7.03 | End-to-end month close | Execute all OPERATIONS_RUNBOOK month-close steps (3.1–3.12) | All steps complete without error |
| T7.04 | Python month-end close | `python pnl_runner.py month-end --month 1` | CloseReport generated with check results |

### Category T8 — New v2.1 Modules (29 tests)

Tests for the 8 modules added in the 2026-03-01 session, plus verification of 4 bug fixes from the pre-import code review.

**modDataGuards**

| Test ID | Test Name | Procedure | Pass Criteria |
|---------|-----------|-----------|---------------|
| T8.01 | ValidateAssumptionsPresence | Run ValidateAssumptionsPresence | Returns True if all drivers filled; returns False + lists blanks if any are empty |
| T8.02 | FindNegativeAmounts | Run FindNegativeAmounts | GL Amount cells with values < 0 highlighted red; message box shows count |
| T8.03 | FindZeroAmounts | Run FindZeroAmounts | GL Amount cells equal to 0 highlighted yellow; message box shows count |
| T8.04 | FindSuspiciousRoundNumbers | Run FindSuspiciousRoundNumbers | GL Amount cells ≥ $1,000 and exactly divisible by 1,000 highlighted orange |

**modDataSanitizer**

| Test ID | Test Name | Procedure | Pass Criteria |
|---------|-----------|-----------|---------------|
| T8.05 | PreviewSanitizeChanges | Run PreviewSanitizeChanges | "Sanitizer Preview" sheet created with color-coded cells; no values changed in source data |
| T8.06 | RunFullSanitize — date safety | Run RunFullSanitize on a sheet with date cells | Date cells unchanged; numeric text cells converted |
| T8.07 | RunFullSanitize — header skip | Run RunFullSanitize | Columns with headers like "Customer ID", "Date", "Name" are entirely skipped |

**modAuditTools**

| Test ID | Test Name | Procedure | Pass Criteria |
|---------|-----------|-----------|---------------|
| T8.08 | AppendChangeLogEntry | Run AppendChangeLogEntry, enter a note | Entry appears in Change Log sheet with timestamp, user, and version |
| T8.09 | FindExternalLinks | Run FindExternalLinks on a workbook with no external links | Message says "No external links found" or report shows zero rows |
| T8.10 | AuditHiddenSheets | Run AuditHiddenSheets | Message box lists all hidden and very-hidden sheets by name |

**modMonthlyTabGenerator — AddNextMonthToModel**

| Test ID | Test Name | Procedure | Pass Criteria |
|---------|-----------|-----------|---------------|
| T8.11 | Correct next month detected | Run AddNextMonthToModel | Confirmation popup shows the correct next calendar month |
| T8.12 | Trend columns marked yellow | Run AddNextMonthToModel → confirm | Next month's column on P&L Monthly Trend and Functional P&L Monthly Trend is highlighted yellow |
| T8.13 | New summary tab created | Run AddNextMonthToModel → confirm | New "Functional P&L Summary - [Month] 25" tab created with green tab color and [NEW - DATA NEEDED] stamp |

**modDemoTools**

| Test ID | Test Name | Procedure | Pass Criteria |
|---------|-----------|-----------|---------------|
| T8.14 | AddControlSheetButtons | Run AddControlSheetButtons | 5 buttons appear on the Report--> sheet with correct labels: "Run Reconciliation", "Build Dashboard", "Data Quality Check", "Export PDF", "Validate Assumptions". Message box confirms "5 control buttons added" |
| T8.15 | Demo buttons call correct macros | Click each of the 5 buttons created by T8.14 | Each button runs the correct macro (e.g., "Data Quality Check" runs modDataQuality.ScanAll, "Export PDF" runs modPDFExport.ExportReportPackage). No "macro not found" errors |
| T8.16 | SetParameterizedPrintArea | Run SetParameterizedPrintArea | Print area is set on a Functional P&L Summary sheet. Message says "Print area set on '[sheet name]'". Go to File > Print Preview — the sheet fits to 1 page portrait with company header |
| T8.17 | CreatePrintableExecSummary | Run CreatePrintableExecSummary | New "Exec Summary - Print" sheet created with: company title, FY label, date, 4 KPI cells (Revenue, Gross Margin %, OpEx, Net Income), product breakdown table with 4 products, navy tab color |

**modDrillDown**

| Test ID | Test Name | Procedure | Pass Criteria |
|---------|-----------|-----------|---------------|
| T8.18 | AddReconciliationDrillLinks | Run AddReconciliationDrillLinks | Column F on the Checks sheet gets a "Drill To Data" header and blue "View Data" hyperlinks on every check row. Message box shows count of links added |
| T8.19 | Drill links navigate to GL | Click any "View Data" hyperlink created by T8.18 | Excel jumps to the CrossfireHiddenWorksheet (GL sheet), which is made visible. No error |
| T8.20 | AutoPopulateReconciliationChecks | Run AutoPopulateReconciliationChecks | Checks sheet is fully recalculated. Cell E1 shows "Last Refreshed:" with current timestamp. Message says "All named ranges verified" (or lists any missing named ranges) |
| T8.21 | ApplyReconciliationHeatmap | Run ApplyReconciliationHeatmap | Checks sheet column D (Difference) is color-coded: green for < $1, yellow for $1–$100, red for > $100. Column E (Status) is green for PASS, red for FAIL |
| T8.22 | RunGoldenFileCompare — first run | Run RunGoldenFileCompare for the first time (no GoldenBaseline sheet exists) | Confirmation prompt asks "Save current FY Total values as the baseline now?" → click Yes → message confirms baseline saved with row count. A very-hidden "GoldenBaseline" sheet is created |
| T8.23 | RunGoldenFileCompare — compare run | Run RunGoldenFileCompare a second time (after GoldenBaseline exists) | "Golden Compare Report" sheet created with columns: Row Label, Golden Value, Current Value, Difference, Status. Each row shows MATCH (green) or CHANGED (red). Message shows count of changed lines |

**modETLBridge**

| Test ID | Test Name | Procedure | Pass Criteria |
|---------|-----------|-----------|---------------|
| T8.24 | TriggerETLLocally — no script | Run TriggerETLLocally when kbt_etl_pipeline.py is NOT in the workbook folder | A file browser dialog opens asking you to locate the script. Clicking Cancel exits cleanly with no error |
| T8.25 | ImportETLOutput — no output file | Run ImportETLOutput when KBT_Cleaned.xlsx does NOT exist in the workbook folder | A file browser dialog opens asking you to locate the file. Clicking Cancel exits cleanly with no error |

**modTrendReports**

| Test ID | Test Name | Procedure | Pass Criteria |
|---------|-----------|-----------|---------------|
| T8.26 | CreateRolling12MonthView | Run CreateRolling12MonthView | "Rolling 12-Month P&L" sheet created with: company title, generated date, navy headers for each month, P&L line items with dollar values, and a line chart titled "Revenue — Rolling [N] Months". Tab color is navy |
| T8.27 | ArchiveReconciliationResults | Run ArchiveReconciliationResults → click Yes to confirm | "Recon Archive" sheet created (or appended to if it exists). Each check row from the Checks sheet is copied with a timestamp. A bold SUMMARY row is added at the bottom with PASS/FAIL counts. Message shows the pass/fail totals |
| T8.28 | CreateReconciliationTrendChart — no archive | Run CreateReconciliationTrendChart when NO "Recon Archive" sheet exists | Message says "No reconciliation archive found" and tells you to run ArchiveReconciliationResults first. No error |
| T8.29 | CreateReconciliationTrendChart — with archive | Run CreateReconciliationTrendChart AFTER running ArchiveReconciliationResults at least once | "Recon Trend Chart" sheet created with a data table (Run Date, PASS, FAIL) and a clustered column chart showing green bars (PASS) and red bars (FAIL). Tab color is green |

**Bug Fix Verifications (commit c232ca7)**

| Test ID | Test Name | Procedure | Pass Criteria |
|---------|-----------|-----------|---------------|
| T8.30 | Demo buttons use correct macro names | Run AddControlSheetButtons → in the VBA Editor, check each button's OnAction property | "Data Quality Check" button calls `modDataQuality.ScanAll` (not RunAllChecks). "Export PDF" button calls `modPDFExport.ExportReportPackage` (not ExportAllSheets) |
| T8.31 | ClearShortcuts runs without error | In the Immediate Window, run `modNavigation.ClearShortcuts` | Runs without compile error or "Sub or Function not defined" error. All Ctrl+Shift shortcuts are unbound |
| T8.32 | Rolling 12 chart range correct | Run CreateRolling12MonthView → right-click the chart → Select Data | The series Values range points to the Total Revenue row (the same row used by revRow), not an offset or overcomplicated formula. Data points match the revenue values in the table |
| T8.33 | GenerateCommentary StyleHeader works | Run GenerateCommentary (after running Variance Analysis via Command 6 first) | "Variance Commentary" sheet is created with navy/white header row. No "Type mismatch" or "argument not optional" error from StyleHeader |

---

## 4. Test Execution Procedure

### Pre-Test Setup
1. Start with a fresh copy of `ExcelDemoFile_adv.xlsm`
2. Enable macros and Trust Access per IMPLEMENTATION_GUIDE.md
3. Import all 32 VBA modules from the `vba/` folder
4. Build the frmCommandCenter (Mode A or B)
5. Verify Python environment: `pip install -r requirements.txt`

### Execution Order
1. **T1 (Compilation)** — Must all pass before proceeding
2. **T2 (Foundation)** — Tests the critical bug fixes
3. **T3 (Menu)** — Tests user interface
4. **T4 (Python)** — Tests parallel analytics layer
5. **T5 (Advanced)** — Tests Phase 3 features
6. **T6 (Data Integrity)** — Tests workbook data
7. **T7 (Integration)** — Full system tests
8. **T8 (New v2.1 Modules)** — Tests the 8 new modules + 4 bug fix verifications

### Recording Results
For each test, record: Test ID, Date, Tester, Result (PASS/FAIL/SKIP), Notes.

---

## 5. Pass/Fail Summary Criteria

### Phase Gate: Ready for Production
- **All T1 tests:** PASS (hard gate — blocks all further testing)
- **All T2 tests:** PASS (critical fixes must be verified)
- **T3 tests:** 4 of 5 PASS minimum (PDF export may vary by printer setup)
- **T4 tests:** All PASS
- **T5 tests:** 5 of 6 PASS minimum
- **T6 tests:** All PASS
- **T7 tests:** 3 of 4 PASS minimum (Python month-end may skip if no Python env)
- **T8 tests:** 25 of 29 PASS minimum (T8.24, T8.25 may SKIP if you have the ETL script/output file present; T8.28 is superseded if you already ran ArchiveReconciliationResults; T8.22 can only run once per baseline)

### Known Acceptable Failures
- T6.06 (Reconciliation): The existing workbook has 6 FAIL checks on the Checks sheet. These are pre-existing data discrepancies between Natural P&L and Functional P&L Summary sheets, not toolkit bugs. See VALIDATION_REPORT.md for details.
- T5.06 (Search cap): Depends on data volume — may need to use a single-character search to trigger the cap.

---

## 6. Test Execution Results

### Run Date: 2026-03-03 (initial) / 2026-03-04 (updated)

#### Category T1 — Compilation & Load

| Test ID | Result | Notes |
|---------|--------|-------|
| T1.01 | PASS | Debug > Compile — zero errors |
| T1.02 | PASS | All 32 modules present in Project Explorer |
| T1.03 | PASS | Option Explicit found in every module |
| T1.04 | PASS | `?APP_VERSION` returns "2.1.0" |
| T1.05 | PASS | pnl_config.py prints config summary, no errors |
| T1.06 | PASS | All Python scripts import cleanly |
| T1.07 | PASS | All 14 files valid UTF-8. Non-ASCII bytes are intentional Unicode (em dashes, arrows, check marks, etc.) — no mojibake detected |
| T1.08 | PASS | `pip install -r requirements.txt` — all packages installed successfully |

#### Category T2 — Foundation Issues

| Test ID | Result | Notes |
|---------|--------|-------|
| T2.01 | **PASS (after fix)** | 9 missing sheet-name constants added to modConfig (commit af44453). Re-tested — all 13 constants return values. |
| T2.02 | PASS | SafeDeleteSheet("NonExistent") — no error, no prompt |
| T2.03 | **PASS (after fix)** | CLR_NAVY and CLR_ALT_ROW had wrong hex-to-decimal conversion (VBA BGR byte order). Fixed in commit 19320db. Re-tested — navy background, white bold text confirmed. |
| T2.04 | **PASS (after fix)** | Added TestUpdateHeaderText wrapper (commit 6f40f91) + NumberFormat text fix (commit ed3276f). Re-tested — A1="Margin" unchanged, A2="Market" unchanged, A3="Apr 25" replaced. |
| T2.05 | — | Not yet run |
| T2.06 | — | Not yet run |
| T2.07 | — | Not yet run |

#### Category T4 — Python Ecosystem (partial)

| Test ID | Result | Notes |
|---------|--------|-------|
| T4.01 | — | Not yet run |
| T4.02 | — | Not yet run |
| T4.03 | — | Not yet run |
| T4.04 | **PASS (after fix)** | Windows PermissionError on temp file cleanup fixed + email report feature removed (commit 3024c44). pytest: 99 passed, 15 skipped, 0 failures. |

#### Category T5 — Advanced VBA Features (partial)

| Test ID | Result | Notes |
|---------|--------|-------|
| T5.01 | **PASS (after fix)** | CreateExecutiveDashboard read row 1 instead of row 4 for headers + Error 5 crash + row/column detection failures. Fixed in commits 6c17bd5, 847a982. Re-tested — Dashboard sheet created with charts. |
| T5.02 | **PASS (after fix)** | WaterfallChart row label fallbacks — searches for multiple label variants. Fixed in commit 304743b. Re-tested — chart created on Dashboard sheet. |
| T5.03 | — | Not yet run |
| T5.04 | — | Not yet run |
| T5.05 | — | Not yet run |
| T5.06 | — | Not yet run |

#### Categories T3, T6, T7, T8 — Not yet started

### Bugs Found During Testing

| Bug # | Test | Module | Description | Fix | Commit |
|-------|------|--------|-------------|-----|--------|
| BUG-T2.01 | T2.01 | modConfig_v2.1.bas | 9 sheet-name constants missing (SH_GL, SH_TECH_DOC, etc.) | Added all 9 constants | af44453 |
| BUG-T2.03a | T2.03 | modConfig_v2.1.bas | `CLR_NAVY = 2050943` decodes to RGB(127,75,31) = tan/brown, not navy. Hex-to-decimal did not use VBA BGR byte order. | Changed to `7949855` = RGB(31,78,121) | 19320db |
| BUG-T2.03b | T2.03 | modConfig_v2.1.bas | `CLR_ALT_ROW = 15651567` decodes to RGB(239,210,238) = pink/lavender, not light blue. Same hex conversion error. | Changed to `16380653` = RGB(237,242,249) | 19320db |
| BUG-T2.04a | T2.04 | modMonthlyTabGenerator_v2.1.bas | `UpdateHeaderText` declared as `Private Sub` — cannot be called from Immediate Window for testing. | Added `Public Sub TestUpdateHeaderText()` wrapper | 6f40f91 |
| BUG-T2.04b | T2.04 | modMonthlyTabGenerator_v2.1.bas | Test wrapper wrote `"Mar 25"` without Text format — Excel auto-converted to date. | Added `NumberFormat = "@"` before writing values | ed3276f |
| BUG-T4.04a | T4.04 | pnl_tests.py | Windows PermissionError on temp file cleanup during pytest | Fixed temp file handling | 3024c44 |
| BUG-T5.01a | T5.01 | modDashboard_v2.1.bas | CreateExecutiveDashboard read row 1 (company title) instead of row 4 (HDR_ROW_REPORT) for column headers | Changed to HDR_ROW_REPORT | 6c17bd5 |
| BUG-T5.01b | T5.01 | modDashboard_v2.1.bas | Error 5 (Invalid procedure call) crash in CreateExecutiveDashboard | Fixed row/column detection | 847a982 |
| BUG-T5.02 | T5.02 | modDashboard_v2.1.bas | WaterfallChart hardcoded "Total Revenue" label — P&L Trend sheet may use "Revenue" or "Net Revenue" | Added multi-variant fallback search | 304743b |

### Bugs Found During Self-Review (2026-03-04, commit 22ba831)

These bugs were found by self-reviewing all remaining untested code against the test plan pass criteria BEFORE the user tested them.

| Bug # | Severity | Module | Line | Description | Fix |
|-------|----------|--------|------|-------------|-----|
| SR-01 | CRITICAL | modReconciliation_v2.1.bas | 292 | `dateCol = 5` (Category column E) instead of `COL_GL_DATE = 2` (Date column B). Check 2 would never find January GL rows. | Changed to `COL_GL_DATE` |
| SR-02 | INFO | modReconciliation_v2.1.bas | 293 | `amtCol = 7` hardcoded instead of using `COL_GL_AMOUNT` constant | Changed to `COL_GL_AMOUNT` |
| SR-03 | CRITICAL | modVarianceAnalysis_v2.1.bas | 221 | GenerateCommentary read row 1 (company title) for `tLastCol` + column header search. Row 1 has only col A, so `tLastCol=1`, FY/Budget loops never execute, all variances = 0. | Changed to `HDR_ROW_REPORT` (row 4) |
| SR-04 | MODERATE | modDashboard_v2.1.bas | 99 | LogAction: `elapsed` (Double) passed as 4th arg (status field expects String "OK"). Audit log Status column corrupted with "0.547". | Moved elapsed into message string |
| SR-05 | MODERATE | modDashboard_v2.1.bas | 369 | Same LogAction issue — CreateExecutiveDashboard | Moved elapsed into message string |
| SR-06 | MODERATE | modDashboard_v2.1.bas | 534 | Same LogAction issue — WaterfallChart | Moved elapsed into message string |
| SR-07 | MODERATE | modDashboard_v2.1.bas | 669 | Same LogAction issue — ProductComparison | Moved elapsed into message string |
| SR-08 | MODERATE | modDashboard_v2.1.bas | 1226 | Same LogAction issue — CreateSmallMultiplesGrid | Moved elapsed into message string |
| SR-09 | MODERATE | modDemoTools_v2.1.bas | — | Same LogAction issue — CreatePrintableExecSummary | Moved elapsed into message string |
| SR-10 | MODERATE | modTrendReports_v2.1.bas | 153 | Same LogAction issue — CreateRolling12MonthView | Moved elapsed into message string |
| SR-11 | MODERATE | modMonthlyTabGenerator_v2.1.bas | 110 | Same LogAction issue — GenerateMonthlyTabs | Moved elapsed into message string |
| SR-12 | MODERATE | modMonthlyTabGenerator_v2.1.bas | 230 | Same LogAction issue — GenerateNextMonthOnly | Moved elapsed into message string |
