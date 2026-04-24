# CoPilot Prompt Guide

## 1) Who this guide is for

This guide is for iPipeline coworkers who want to use the **demo-specific VBA features** on their own workbook.

If your workbook layout is different from the demo workbook, this guide shows you how to ask Microsoft 365 CoPilot to adapt the code safely.

---

## 2) What you need before starting

- Excel desktop app with macros enabled.
- Access to Microsoft 365 CoPilot.
- A **copy** of your workbook (never test on your production file first).
- The original VBA module you want to adapt.

Recommended safety setup:
1. Save your workbook copy as `YourFile_copilot_test.xlsm`.
2. In Excel, create a sheet called `Change_Log`.
3. Paste all CoPilot prompts and responses into `Change_Log`.

---

## 3) How to describe your workbook to CoPilot

CoPilot works best when you give it a simple map of your workbook first.

Use this exact format:

```text
Workbook map:
- Sheet: <name> | Purpose: <what it does> | Header row: <row number>
- Key columns on this sheet: <column name>, <column name>, <column name>
- Named ranges used: <name>=<ref> (or "none")

Business rules:
- Materiality absolute threshold: <$ amount>
- Materiality percent threshold: <percent>
- PASS/FAIL logic: <plain language rules>

Output requirements:
- New output sheet names I want: <list>
- Existing sheets that must not be edited: <list>
- Audit log sheet name: <name>
```

---

## 4) Starter prompt template (copy/paste)

```text
You are adapting VBA code for my workbook.

Task:
1) Keep the original behavior and intent.
2) Update sheet names, header detection, and column references based on my workbook map.
3) Do not use InputBox or SendKeys.
4) Keep output non-destructive: write to new sheets unless I explicitly say otherwise.
5) Add clear success/failure messages and logging calls.
6) Return only:
   - changed procedures
   - brief list of assumptions
   - test steps I should run in Excel

Workbook map:
<paste workbook map>

Original VBA:
<paste module code>
```

---

## 5) Standard adaptation workflow

1. Paste your workbook map into CoPilot.
2. Paste one VBA module at a time (do not paste everything at once).
3. Ask CoPilot to return only changed procedures.
4. Paste returned code into a test copy of the workbook.
5. Run the macro once.
6. Check outputs and audit logs.
7. If needed, send one correction prompt:
   - “Only fix sheet-name mapping and do not change logic.”

---

## 6) Worked examples by feature

### Example A — Material Variance Narrative Engine

**Goal:** adapt narrative generation to your own trend sheet.

```text
Adapt this variance narrative macro.
My trend sheet is 'Monthly Finance Trend'.
Header row is 5.
Line item label is column A.
Monthly values start in column C.
Output sheet should be 'Exec_Variance_Narrative'.
Keep materiality thresholds at $10000 and 15%.
Do not edit the source trend sheet values.
Return only changed procedures.
```

### Example B — Reconciliation Command Run

**Goal:** adapt checks for your source transaction sheet.

```text
Adapt this reconciliation macro.
My source transaction sheet is 'GL_Source'.
Amount column is 'Net Amount'.
Date column is 'Posting Date'.
Checks output sheet should be 'Checks'.
If a check fails, write FAIL and a clear reason.
Keep all checks in one run.
Return only changed procedures and the expected output columns.
```

### Example C — Executive Brief Pack Builder

**Goal:** adapt KPI extraction and PDF output.

```text
Adapt this executive brief pack macro.
My KPI sheet is 'Executive Dashboard'.
Revenue KPI cell is C12.
Margin KPI cell is E12.
Top product KPI cell is G12.
Output brief sheet should be 'Exec_Brief'.
Export one PDF in the same folder as the workbook.
Use Arial and iPipeline colors only.
Return only changed procedures.
```

### Example D — AWS Allocation Sensitivity Runner

**Goal:** adapt scenario logic to your assumptions sheet.

```text
Adapt this scenario runner.
My assumptions sheet is 'Drivers'.
Driver names are in column A and values in column B.
Allocation output sheet should be 'Scenario_Compare'.
Run scenarios: Base, Growth Push, Margin Protection, Stress Case.
Write one narrative sentence per scenario.
Return only changed procedures.
```

### Example E — Period Close Storyboard Generator

**Goal:** adapt close narrative to your period tabs.

```text
Adapt this close storyboard macro.
My period summary sheets are:
- Close_Jan
- Close_Feb
- Close_Mar
Use these plus 'Checks' and 'VBA_AuditLog'.
Output sheet should be 'Close_Storyboard'.
Do not modify source period sheets.
Return only changed procedures.
```

### Example F — Demo Command Center

**Goal:** adapt command center buttons for your macro set.

```text
Adapt this command center module.
Build sheet name: 'Demo_CommandCenter'.
Buttons to include:
- Run Reconciliation
- Generate Variance Narrative
- Build Executive Brief Pack
- Run Scenario Comparison
Use Arial font and iPipeline color palette.
Do not add extra buttons.
Return only changed procedures.
```

---

## 7) Validation checklist after CoPilot edits

Run this checklist after every adaptation:

1. Does the macro run without compile errors?
2. Does it create only the expected output sheet(s)?
3. Does it avoid editing protected source sheets?
4. Does it write clear PASS/FAIL status messages?
5. Does it write audit log entries?
6. Do numeric outputs match manual spot checks for 3 rows?
7. If exporting PDFs, does the output file open correctly?

---

## 8) Troubleshooting when CoPilot gets it wrong

### Problem: Wrong sheet name
Use this correction prompt:

```text
Do not change any logic.
Only replace sheet references as follows:
- Old: <old sheet>
- New: <new sheet>
Return only changed lines.
```

### Problem: Wrong column mapping

```text
Do not change formulas or business rules.
Only remap these columns:
- Date -> <column name>
- Amount -> <column name>
- Status -> <column name>
Return only changed procedures.
```

### Problem: CoPilot changed too much

```text
Revert to original logic.
Only apply the workbook map updates.
No new features.
No UI changes.
Return only changed procedures.
```

### Problem: Macro edits source sheets

```text
Update the macro to be non-destructive.
Write outputs to a new sheet named <sheet name>.
Do not edit source sheets.
Return only changed procedures.
```

---

## 9) Safe rollback process

If a run fails:

1. Close workbook without saving.
2. Reopen your test copy.
3. Paste back the last known-good VBA version.
4. Re-run with smaller scope (one feature only).
5. Save that stable version before testing anything else.

---

## 10) Recommended operating pattern for teams

- One owner runs CoPilot prompts.
- One reviewer checks output and calculations.
- One approver signs off before production use.

Use this handoff note after each adaptation:

```text
Feature adapted:
Workbook:
Tester:
Reviewer:
Result (PASS/FAIL):
Open issues:
```

This keeps the process audit-ready for Finance and leadership demos.
