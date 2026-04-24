# Troubleshooting Reference

## 1) Macro does not run

### Symptoms
- "Macros are disabled" message
- Button click does nothing

### Fix
1. Save workbook as `.xlsm`.
2. Enable macros in Trust Center.
3. Reopen file and try again.

## 2) Missing required sheet

### Symptoms
- Error says required sheet is missing.

### Fix
1. Check sheet names match expected names exactly.
2. Unhide sheets and verify spelling.
3. Re-run macro.

## 3) PDF export fails

### Symptoms
- Export procedure errors or no file created.

### Fix
1. Save workbook first.
2. Confirm output folder path exists.
3. Retry export.

## 4) Compare/consolidate outputs look wrong

### Symptoms
- Missing columns or mismatched rows.

### Fix
1. Confirm header row is populated.
2. Remove blank top rows.
3. Run profile tool to verify detected header row.

## 5) CoPilot adaptation produced broken code

### Fix
1. Revert to last working code.
2. Re-run CoPilot with a narrower prompt.
3. Ask for only sheet/column mapping changes.

## 6) Escalation template

Use this when sending support request:

```text
Workbook:
Feature:
Error message:
What I expected:
What happened:
Last step completed successfully:
```
