# KBT P&L Toolkit — Operations Runbook

> **Audience:** Finance team members responsible for the monthly P&L close cycle.
> Covers month-open, mid-month, and month-close procedures with exact command sequences and failure resolution.

---

## Monthly Calendar Overview

```
 Week 1 (Days 1-5)         Week 2 (Days 6-12)       Week 3-4 (Days 13-EOM)
 ─────────────────          ─────────────────        ─────────────────────
 MONTH-OPEN                 MID-MONTH                MONTH-CLOSE
 ✓ Save opening snapshot    ✓ Import partial GL      ✓ Final GL import
 ✓ Validate assumptions     ✓ Run forecast update    ✓ Full data quality scan
 ✓ Import opening GL data   ✓ Variance spot-check    ✓ Run all reconciliation
 ✓ Data quality scan        ✓ Allocation preview     ✓ Variance analysis
 ✓ Preliminary checks                                ✓ Generate commentary
                                                     ✓ Cross-sheet validation
                                                     ✓ Build dashboard
                                                     ✓ Export report package
                                                     ✓ Send exec summary
                                                     ✓ Save closing snapshot
```

---

## Phase 1 — Month-Open (Days 1–5)

### Step 1.1 — Save Opening Snapshot

**Purpose:** Create a restore point before any new data enters the workbook.

1. Open Command Center: **Ctrl+Shift+M**
2. Run **Command 20** (Save Current Scenario)
3. Name it: `FY25_MonXX_Opening` (e.g., `FY25_Mon04_Opening` for April)
4. Verify: the confirmation message shows the snapshot was saved

### Step 1.2 — Validate Assumptions

**Purpose:** Ensure allocation shares and drivers are correct for the new month.

1. Navigate to the **Assumptions** sheet (Ctrl+Shift+J → select Assumptions)
2. Verify:
   - Revenue share percentages sum to 100%
   - AWS allocation shares sum to 100%
   - Headcount allocations reflect any org changes
3. Run **Command 14** (Recalculate AWS Allocations) to validate share sums
4. If shares changed, document the reason in a change request (**Command 38**)

### Step 1.3 — Import Opening GL Data

**Purpose:** Load the new month's general ledger extract.

1. Obtain the GL data extract from the accounting system (CSV or XLSX format)
2. Verify the file has 7 columns: ID, Date, Department, Product, Expense Category, Vendor, Amount
3. Run **Command 17** (Import GL Data Pipeline)
4. Select the source file when prompted
5. Verify the import count matches your source record count

### Step 1.4 — Data Quality Scan

**Purpose:** Catch data issues before they propagate into reports.

1. Run **Command 7** (Scan Data Quality)
2. Review the Data Quality Report sheet:
   - **Duplicates:** If found, verify whether legitimate, then fix with Command 9 if needed
   - **Text-stored numbers:** Fix with Command 8
   - **Blank fields:** Return to source system for correction
   - **Unknown products/departments:** Fix spelling or add to master lists
3. Re-run Command 7 to confirm all issues resolved

### Step 1.5 — Preliminary Reconciliation

**Purpose:** Baseline check that data loaded correctly.

1. Run **Command 3** (Run Reconciliation Checks) — shortcut: **Ctrl+Shift+R**
2. Review the Checks sheet
3. Expected: Most checks PASS; some may show WARN if the month's data is still partial
4. Investigate any FAIL results immediately

---

## Phase 2 — Mid-Month (Days 6–12)

### Step 2.1 — Import Updated GL Data

If additional GL extracts arrive mid-month:

1. Run **Command 17** (Import GL Data Pipeline) with the updated file
2. Run **Command 7** (Data Quality Scan)
3. Run **Command 3** (Reconciliation Checks)

### Step 2.2 — Forecast Update

**Purpose:** Update the rolling forecast with the latest actuals.

1. Run **Command 18** (Rolling Forecast)
2. Review the forecast output for reasonableness
3. Note any significant changes from last month's forecast

**Python alternative (more detailed):**
```bash
python pnl_runner.py forecast --months 3 --export forecast_output.xlsx
```

### Step 2.3 — Variance Spot-Check

**Purpose:** Early identification of emerging variance trends.

1. Run **Command 6** (Variance Analysis)
2. Scan for any line items flagged at >15% variance
3. If material variances exist, alert the responsible department head
4. No formal action required — this is a monitoring step

### Step 2.4 — Allocation Preview (if assumptions changed)

If allocation shares or methods were modified:

1. Run **Command 25** (Allocation Scenario Preview)
2. Review the projected impact before committing
3. If acceptable, run **Command 24** (Run Allocation Engine)

---

## Phase 3 — Month-Close (Days 13–End of Month)

### Step 3.1 — Final GL Import

1. Obtain the final, complete GL extract for the month
2. Run **Command 17** (Import GL Data Pipeline)
3. Verify record count matches the source system's month-end count

### Step 3.2 — Full Data Quality Scan

1. Run **Command 7** (Scan Data Quality)
2. Resolve ALL issues — no outstanding quality flags should remain
3. Run **Command 8** (Fix Text Numbers) and **Command 9** (Fix Duplicates) as needed
4. Re-run Command 7 to confirm zero issues

### Step 3.3 — Run All Reconciliation

1. Run **Command 3** (Run Reconciliation Checks)
2. **All checks must show PASS** before proceeding
3. If any checks FAIL:
   - Document the discrepancy
   - Trace to source (see Failure Scenarios below)
   - Fix and re-run until all PASS

### Step 3.4 — Variance Analysis

1. Run **Command 6** (Variance Analysis)
2. Review all flagged items (>15% threshold)
3. For each material variance, prepare an explanation

### Step 3.5 — Generate Variance Commentary

1. Run **Command 46** (Variance Commentary)
2. Review the auto-generated narrative on the "Variance Commentary" sheet
3. Edit as needed for accuracy and tone
4. This becomes the basis for the executive summary

### Step 3.6 — Cross-Sheet Validation

1. Run **Command 47** (Cross-Sheet Validation)
2. Review the "Cross-Sheet Validation" sheet
3. All items should show PASS or REVIEW (investigate any FAIL)
4. This is the final validation gate before reporting

### Step 3.7 — Run Allocation Engine

1. Run **Command 24** (Run Allocation Engine)
2. Verify the Allocation Output sheet shows expected Department × Product distribution

### Step 3.8 — Build Dashboard

1. Run **Command 12** (Build Dashboard Charts)
2. Visually verify:
   - Revenue trend shows the new month
   - CM% trend is reasonable
   - Revenue mix pie chart reflects current allocation

### Step 3.9 — Export Report Package

1. Run **Command 10** (Export Report Package PDF)
2. Save to the designated month-end folder
3. Verify the PDF includes all report sheets with correct page numbers

### Step 3.10 — Append Month to Trend

1. Run **Command 19** (Append Month to Trend)
2. Verify the P&L Trend sheet has the new month column

### Step 3.11 — Save Closing Snapshot

1. Run **Command 20** (Save Current Scenario)
2. Name it: `FY25_MonXX_Close` (e.g., `FY25_Mon04_Close`)

### Step 3.12 — Executive Summary (Optional)

**VBA path:**
1. Navigate to a summary view (Dashboard or P&L Trend)
2. Use **modEmailSummary** to generate a formatted summary for clipboard/email

**Python path:**
```bash
python pnl_runner.py email --month 4 --output april_exec_summary.html
```

---

## Month-Close Checklist

Print this and check off each step:

```
□  1. Final GL import complete
□  2. Data quality scan — zero issues remaining
□  3. Reconciliation — all checks PASS
□  4. Variance analysis — all flags reviewed
□  5. Variance commentary — narrative approved
□  6. Cross-sheet validation — all PASS
□  7. Allocation engine — run and verified
□  8. Dashboard charts — built and reviewed
□  9. Report package PDF — exported and filed
□ 10. Month appended to trend sheet
□ 11. Closing snapshot saved
□ 12. Executive summary distributed
```

---

## Common Failure Scenarios

### F1 — Reconciliation Check Fails: "GL Total ≠ P&L Trend Total"

**Symptoms:** Command 3 shows FAIL for the GL-to-Trend total comparison.

**Root Cause:** Usually a partial import, missing transactions, or a manual edit on the P&L Trend sheet.

**Resolution:**
1. Note the discrepancy amount from the Checks sheet detail column
2. Open the GL tab (CrossfireHiddenWorksheet) and sum the Amount column
3. Open P&L Trend and check the FY Total column
4. Common fixes:
   - Re-import GL if records are missing
   - Undo manual edits on P&L Trend (it should be formula-driven)
   - Check for filtered rows hiding data
5. Re-run Command 3 to verify PASS

---

### F2 — Variance Analysis Shows Unrealistic Percentages

**Symptoms:** Variance percentages of 500%+ or -100% on major line items.

**Root Cause:** Usually a data type issue (text vs. number), a missing prior month, or an allocation that wasn't run.

**Resolution:**
1. Run Command 7 (Data Quality Scan) to check for text-stored numbers
2. Verify both the current and prior month have data on the P&L Trend sheet
3. Ensure Command 24 (Allocation Engine) was run for both months
4. Re-run Command 6

---

### F3 — Import Fails: "Column Mismatch"

**Symptoms:** Command 17 shows an error about missing or extra columns.

**Root Cause:** The source file doesn't have the expected 7-column format.

**Resolution:**
1. Open the source file and verify columns: ID, Date, Department, Product, Expense Category, Vendor, Amount
2. Column headers must match exactly (case-sensitive)
3. Remove any extra columns or blank columns
4. Save and retry Command 17

---

### F4 — "Subscript Out of Range" Error

**Symptoms:** VBA error dialog with "Run-time error 9: Subscript out of range."

**Root Cause:** A sheet that the code expects doesn't exist or has been renamed.

**Resolution:**
1. Check the error message for which sheet is missing
2. Verify all sheet tab names match modConfig constants:
   - CrossfireHiddenWorksheet, Assumptions, P&L - Monthly Trend, etc.
3. Rename the sheet to match, or check if it was accidentally deleted
4. Run Command 45 (Quick Health Check) to verify all sheets exist

---

### F5 — Dashboard Charts Are Blank

**Symptoms:** Command 12 creates the Dashboard sheet but charts show no data.

**Root Cause:** No populated months detected, or the P&L Trend sheet structure changed.

**Resolution:**
1. Open P&L Trend and verify data exists in row 5+ under month columns
2. Check that Row 4 has the expected column headers
3. Verify the FY Total column exists
4. Re-run Command 12

---

### F6 — PDF Export Has Wrong Pages or Missing Sheets

**Symptoms:** The PDF is missing sheets or includes unwanted sheets.

**Root Cause:** Sheet visibility or print area issues.

**Resolution:**
1. Ensure all report sheets are visible (not hidden) — use Command 48 to toggle Executive Mode off
2. Run Command 13 (Refresh TOC) to reset sheet references
3. Re-run Command 10

---

### F7 — Python Script Fails: "FileNotFoundError"

**Symptoms:** Any Python command fails with "FileNotFoundError."

**Resolution:**
1. Verify the Excel file is in the same directory as the Python scripts
2. Or specify the path explicitly:
   ```bash
   python pnl_runner.py month-end --file "C:\path\to\workbook.xlsx"
   ```
3. Or set the environment variable:
   ```bash
   set KBT_SOURCE_FILE=C:\path\to\workbook.xlsx
   ```

---

### F8 — "Run Scan First" Error on Fix Commands

**Symptoms:** Commands 8 or 9 refuse to run, showing "Run Scan Data Quality first."

**Root Cause:** These commands only operate on cells pre-flagged by Command 7's scanner. This is a safety feature.

**Resolution:**
1. Run Command 7 (Scan Data Quality) first
2. Then run Command 8 or 9 as needed
3. The scan populates an internal tracking list that the fix commands use

---

## Annual Procedures

### Fiscal Year Rollover

See the **Implementation Guide**, Section 6 for the complete FY rollover procedure. Summary:

1. Save a final closing snapshot for the old FY
2. Update `FISCAL_YEAR`, `FISCAL_YEAR_4`, `FY_LABEL` in modConfig and pnl_config.py
3. Create new monthly tabs (Command 1)
4. Run full reconciliation (Command 3)
5. Run integration test (Command 44)
6. Save an opening snapshot for the new FY

### Quarterly Audit Prep

1. Run Command 44 (Full Integration Test) — save the report
2. Run Command 36 (Auto-Documentation) — generates system documentation
3. Export the audit log (Command 42) — save as CSV
4. Run Command 47 (Cross-Sheet Validation) — save results
5. Compile the report package PDFs for all 3 months of the quarter
