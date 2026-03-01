# KBT P&L Toolkit — Integration Test Execution Guide

> **Module:** modIntegrationTest
> **Commands:** #44 (Full Integration Test), #45 (Quick Health Check)

---

## Overview

The integration test suite (`modIntegrationTest`) exercises all major modules and cross-module dependencies in a single automated run. It produces a structured report on the "Integration Test Report" sheet.

There are two test modes:

| Mode | Command # | Duration | Tests | Use Case |
|------|-----------|----------|-------|----------|
| **Full Integration Test** | 44 | 30-60 sec | All 7 categories | After module updates, quarterly review |
| **Quick Health Check** | 45 | 5-10 sec | Sheets + formulas only | Daily spot check |

---

## Pre-Test Checklist

Before running tests, verify:

```
□  Workbook is open and macros are enabled
□  All 32 VBA modules are imported
□  No other macros are currently running
□  The workbook has not been filtered or grouped
□  Calculation mode is set to Automatic (Formulas → Calculation → Automatic)
```

---

## Running the Full Integration Test (Command 44)

### Step 1 — Launch

Press **Ctrl+Shift+M** → select **#44 Full Integration Test** → click **Run**

Or from the VBA Immediate Window: `modIntegrationTest.RunFullTest`

### Step 2 — Wait for Completion

The test runs through 7 categories:
1. **Core Sheets** — Verifies all required sheets exist
2. **GL Integrity** — Checks GL data has rows, columns, headers, numeric amounts
3. **VBA Modules** — Confirms all expected modules are installed
4. **Named Ranges** — Checks dynamic named ranges are defined
5. **Formula Errors** — Scans all sheets for #REF!, #NAME?, etc.
6. **Reconciliation** — Reads the Checks sheet for FAIL results
7. **Cross-Module** — Tests audit log, assumptions data, AWS data, product references

The status bar shows progress. Do not interact with Excel until complete.

### Step 3 — Review the Report

When complete, Excel navigates to the "Integration Test Report" sheet.

**Report layout:**

```
Row 1: "Keystone BenefitTech, Inc."
Row 2: "Integration Test Report — 2026-02-20 14:30:00"
Row 3: "Results: 42 tests | PASS: 38 | FAIL: 3 | WARN: 1 | Duration: 12.5s"
Row 4: [Headers] Test # | Category | Test Name | Expected | Actual | Status | Detail
Row 5+: [Test results — one row per test]
```

### Step 4 — Interpret Results

**Color coding:**
- Green row = PASS
- Red row = FAIL
- Yellow row = WARN

**Status meanings:**

| Status | Meaning | Action Required |
|--------|---------|-----------------|
| PASS | Test succeeded | None |
| FAIL | Test failed — investigate | Fix the issue and retest |
| WARN | Ambiguous result | Review the Detail column for context |

### Step 5 — Address Failures

For each FAIL row:
1. Read the **Category** and **Test Name** columns to identify what failed
2. Read the **Expected** vs **Actual** columns to understand the discrepancy
3. Read the **Detail** column for diagnostic information
4. Fix the root cause (see Common Failures below)
5. Re-run the test to verify the fix

### Step 6 — Confirm and Log

A summary message box appears after completion showing PASS/FAIL/WARN counts.

The test execution is also logged in the audit trail (VBA_AuditLog sheet) with:
- Module: `modIntegrationTest`
- Sub: `RunFullTest`
- Detail: PASS/FAIL/WARN counts and duration

---

## Running the Quick Health Check (Command 45)

### Step 1 — Launch

Press **Ctrl+Shift+M** → select **#45 Quick Health Check** → click **Run**

Or: `modIntegrationTest.QuickHealthCheck`

### Step 2 — Review

The Quick Health Check tests only:
- Core sheet existence
- Formula error scan

Report appears on the same "Integration Test Report" sheet (overwritten each run).

Row 3 will show "Quick Check: N tests | PASS: X | FAIL: Y | WARN: Z"

### When to Use Quick vs Full

| Scenario | Recommended Mode |
|----------|-----------------|
| Start of workday | Quick Health Check (#45) |
| After importing new data | Quick Health Check (#45) |
| After importing new VBA modules | Full Integration Test (#44) |
| After changing modConfig constants | Full Integration Test (#44) |
| Monthly close sign-off | Full Integration Test (#44) |
| Quarterly audit preparation | Full Integration Test (#44) |
| Something "feels off" | Quick Health Check (#45) first, then Full if issues found |

---

## Test Category Details

### Category: Core Sheets

Tests that all 13+ required sheets exist by name (using modConfig constants).

**Common failure:** A sheet was renamed or deleted.
**Fix:** Rename the sheet to match the modConfig constant, or recreate it.

### Category: GL Integrity

Tests: GL has data rows (>1), has 7+ columns, header row exists, no blank rows in first 100, Amount column is numeric.

**Common failure:** GL data was cleared or the sheet was reformatted.
**Fix:** Re-import GL data using Command 17.

### Category: VBA Modules

Tests that all expected module names are present in the VBA project.

**Common failure:** A module was not imported or was deleted.
**Fix:** Re-import the missing .bas file from the 03_Code/VBA/ folder.

### Category: Named Ranges

Tests that dynamic named ranges (created by modSetup.DynamicNamedRanges) are defined.

**Common failure:** Named ranges were never created.
**Fix:** Run `modSetup.DynamicNamedRanges` from the Immediate Window.

### Category: Formula Errors

Scans every cell on every report sheet for error values.

**Common failure:** A formula reference broke due to sheet/column changes.
**Fix:** Identify the error cells from the Detail column, trace the formula, and fix the reference.

### Category: Reconciliation

Reads the Checks sheet and counts FAIL results.

**Common failure:** Data discrepancies between sheets (see VALIDATION_REPORT.md for known issues).
**Fix:** Investigate each FAIL check individually. Some may be pre-existing data issues.

### Category: Cross-Module

Tests cross-module dependencies: audit log sheet exists, Assumptions has data, AWS Allocation has data, all 4 products found in GL.

**Common failure:** A dependent sheet lost its data.
**Fix:** Restore from a saved scenario (Command 21) or re-import data.

---

## Archiving Test Results

After each Full Integration Test:

1. Save a screenshot or PDF of the report sheet (Command 11)
2. Export the audit log entry (Command 42)
3. Store in the project's QA records folder with the date

For audit trail purposes, retain test reports for at least 4 quarters.

---

## Troubleshooting

| Symptom | Cause | Resolution |
|---------|-------|------------|
| "Subscript out of range" during test | Required sheet missing | Check the error detail; the test was trying to access a sheet by its modConfig constant name |
| Test hangs or takes >120 seconds | Large workbook or calculation mode issue | Press Esc to interrupt; set Calculation to Manual, run test, then restore to Automatic |
| Report sheet is blank | modPerformance.TurboOff failed to restore | Press Ctrl+Break, then manually set `Application.ScreenUpdating = True` in Immediate Window |
| All tests pass but numbers "look wrong" | Tests check structure, not business logic | Run Command 3 (Reconciliation) and Command 47 (Cross-Sheet Validation) separately for data verification |
| Python test equivalent | Use `python pnl_runner.py test` | Runs the pytest suite; covers pnl_config 100%, pnl_month_end 80%, pnl_allocation_simulator 80% |
