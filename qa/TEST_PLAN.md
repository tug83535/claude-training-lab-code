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
- Command Center functionality (all 50 actions reachable)
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
| Workbook | KeystoneBenefitTech_PL_Model.xlsm |

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
| T1.06 | All Python scripts import | `python -c "import pnl_X"` for each script | Zero ImportError |
| T1.07 | Python UTF-8 clean | Run scan_chars.py on all .py files | Zero mojibake codepoints |
| T1.08 | requirements.txt complete | `pip install -r requirements.txt` | All packages install |

### Category T2 — Foundation Issues (7 tests)

Verifies ISSUE-001 through ISSUE-007 fixes.

| Test ID | Test Name | Issue | Procedure | Pass Criteria |
|---------|-----------|-------|-----------|---------------|
| T2.01 | modConfig has all constants | ISSUE-001 | `?SH_GL`, `?SH_TECH_DOC`, etc. in Immediate Window | All 13 new constants return values |
| T2.02 | SafeDeleteSheet works | ISSUE-001 | Call `SafeDeleteSheet("NonExistent")` | No error, no prompt |
| T2.03 | StyleHeader works | ISSUE-001 | Call on test sheet with sample headers | Navy background, white bold text |
| T2.04 | UpdateHeaderText safe | ISSUE-002 | Create test sheet with "Margin", "Market", "Mar 25" cells. Run UpdateHeaderText("Mar","Apr") | "Margin" and "Market" unchanged; "Mar 25" → "Apr 25" |
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
| T4.01 | UTF-8 clean across all files | ISSUE-008 | Scan for suspect codepoints (U+00E2, U+00C3, etc.) | Zero hits |
| T4.02 | pnl_config self-test | ISSUE-008 | `python pnl_config.py` | All shares sum to 1.0, version 2.1.0 |
| T4.03 | pnl_runner dispatches | — | `python pnl_runner.py --help` | Shows all 9 commands |
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

### Category T8 — New v2.1 Modules (13 tests)

Tests for the 8 modules added in the 2026-03-01 session.

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

---

## 4. Test Execution Procedure

### Pre-Test Setup
1. Start with a fresh copy of `KeystoneBenefitTech_PL_Model.xlsx`
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

### Known Acceptable Failures
- T6.06 (Reconciliation): The existing workbook has 6 FAIL checks on the Checks sheet. These are pre-existing data discrepancies between Natural P&L and Functional P&L Summary sheets, not toolkit bugs. See VALIDATION_REPORT.md for details.
- T5.06 (Search cap): Depends on data volume — may need to use a single-character search to trigger the cap.
