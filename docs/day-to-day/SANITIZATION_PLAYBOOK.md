# KBT P&L Toolkit — Sanitization Playbook

> **Audience:** Anyone preparing the workbook for demos, presentations, training, or external sharing.
> Covers which fields to mask, how to mask them, verification, and reversal.

---

## When to Sanitize

Sanitize the workbook before:
- External presentations or conference demos
- Sharing with vendors, consultants, or partners
- Training materials for new hires (before NDA clearance)
- Public-facing documentation screenshots
- Any scenario where real financial data should not be visible

**Never share the unsanitized workbook outside the finance team.**

---

## What Gets Masked

### Tier 1 — Always Mask (Sensitive Financial Data)

| Field | Location | Masking Method |
|-------|----------|---------------|
| **Dollar amounts** | GL (Amount column), P&L Trend, Product Summary, Functional P&L, Allocation Output | Scale by random factor (see below) |
| **Vendor names** | GL (Vendor column) | Replace with `Vendor_001` through `Vendor_NNN` |
| **Transaction IDs** | GL (ID column) | Replace with sequential `TXN-000001` through `TXN-NNNNNN` |

### Tier 2 — Mask If Sharing Externally

| Field | Location | Masking Method |
|-------|----------|---------------|
| **Company name** | Row 1 headers on all report sheets | Replace with "Acme Corp" or "Sample Company Inc." |
| **Specific dates** | GL (Date column) | Shift all dates by a fixed offset (e.g., +90 days) |
| **Department names** | GL, Functional P&L tabs | Replace with `Dept_A` through `Dept_G` |

### Tier 3 — Usually Safe to Keep

| Field | Rationale |
|-------|-----------|
| Product names (iGO, Affirm, etc.) | Publicly known product lines |
| Expense category names | Generic accounting categories |
| Sheet structure and formulas | The automation framework is the demo subject |
| Allocation percentages | Can be replaced with round numbers if sensitive |

---

## Sanitization Procedures

### Procedure S1 — Scale Dollar Amounts

**Purpose:** Preserve the mathematical relationships (ratios, trends, variances) while making absolute values unrecognizable.

**Method:**
1. Choose a random scaling factor between 0.3 and 3.0 (e.g., 1.47)
2. Apply uniformly to ALL dollar amounts across ALL sheets
3. This preserves: percentage margins, growth rates, relative comparisons
4. This obscures: actual revenue, actual costs, real P&L values

**VBA approach (run in Immediate Window):**
```vba
' Scale all amounts on GL sheet by factor
Sub SanitizeAmounts()
    Dim factor As Double: factor = 1.47  ' change this
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("CrossfireHiddenWorksheet")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 7).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If IsNumeric(ws.Cells(r, 7).Value) Then
            ws.Cells(r, 7).Value = Round(ws.Cells(r, 7).Value * factor, 2)
        End If
    Next r
    ' Recalculate all formulas
    Application.CalculateFull
End Sub
```

**Python approach:**
```python
import openpyxl
factor = 1.47
wb = openpyxl.load_workbook("ExcelDemoFile_adv.xlsm")
ws = wb["CrossfireHiddenWorksheet"]
for row in ws.iter_rows(min_row=2, min_col=7, max_col=7):
    for cell in row:
        if isinstance(cell.value, (int, float)):
            cell.value = round(cell.value * factor, 2)
wb.save("ExcelDemoFile_adv_SANITIZED.xlsm")
```

**Important:** Record the scaling factor securely. You need it for reversal.

---

### Procedure S2 — Mask Vendor Names

**Method:** Replace each unique vendor name with a sequential identifier.

**VBA approach:**
```vba
Sub SanitizeVendors()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("CrossfireHiddenWorksheet")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
    Dim vendorMap As Object: Set vendorMap = CreateObject("Scripting.Dictionary")
    Dim counter As Long: counter = 0
    Dim r As Long
    For r = 2 To lastRow
        Dim v As String: v = CStr(ws.Cells(r, 6).Value)
        If Not vendorMap.Exists(v) Then
            counter = counter + 1
            vendorMap(v) = "Vendor_" & Format(counter, "000")
        End If
        ws.Cells(r, 6).Value = vendorMap(v)
    Next r
End Sub
```

---

### Procedure S3 — Mask Transaction IDs

```vba
Sub SanitizeIDs()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("CrossfireHiddenWorksheet")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        ws.Cells(r, 1).Value = "TXN-" & Format(r - 1, "000000")
    Next r
End Sub
```

---

### Procedure S4 — Replace Company Name

**Method:** Find and replace the company name across all sheets.

1. Press **Ctrl+H** (Find & Replace)
2. Set **Within:** Workbook
3. Find: `Keystone BenefitTech, Inc.`
4. Replace: `Acme Corp`
5. Click **Replace All**
6. Also replace: `Keystone BenefitTech` → `Acme Corp` (without "Inc.")
7. Also replace: `KBT` → `AC` (if used in headers or codes)

---

### Procedure S5 — Shift Dates

**Method:** Add a fixed number of days to all dates to obscure the real reporting period.

```vba
Sub SanitizeDates()
    Dim offset As Long: offset = 90  ' shift by 90 days
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("CrossfireHiddenWorksheet")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If IsDate(ws.Cells(r, 2).Value) Then
            ws.Cells(r, 2).Value = ws.Cells(r, 2).Value + offset
        End If
    Next r
End Sub
```

**Note:** Date shifting will change month assignments. If month integrity is important for the demo, use a multiple-of-365 offset instead (e.g., +365 shifts to the same relative month in a different year).

---

### Procedure S6 — Mask Department Names

```vba
Sub SanitizeDepartments()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("CrossfireHiddenWorksheet")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    Dim deptMap As Object: Set deptMap = CreateObject("Scripting.Dictionary")
    deptMap("NetOps") = "Dept_A"
    deptMap("Security") = "Dept_B"
    deptMap("Support") = "Dept_C"
    deptMap("Partners") = "Dept_D"
    deptMap("Content") = "Dept_E"
    deptMap("R&D") = "Dept_F"
    deptMap("Product Management") = "Dept_G"
    Dim r As Long
    For r = 2 To lastRow
        Dim d As String: d = CStr(ws.Cells(r, 3).Value)
        If deptMap.Exists(d) Then ws.Cells(r, 3).Value = deptMap(d)
    Next r
End Sub
```

---

## Full Sanitization Script

Run all procedures in sequence for complete sanitization:

```vba
Sub FullSanitize()
    ' SAVE ORIGINAL FIRST
    Dim origPath As String
    origPath = ThisWorkbook.FullName
    ThisWorkbook.SaveCopyAs Replace(origPath, ".xls", "_ORIGINAL.xls")

    ' Run all sanitization procedures
    SanitizeAmounts      ' S1 - scale dollars
    SanitizeVendors      ' S2 - mask vendors
    SanitizeIDs          ' S3 - mask IDs
    SanitizeDates        ' S5 - shift dates
    SanitizeDepartments  ' S6 - mask departments

    ' S4 - company name (manual Ctrl+H or:)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Cells.Replace "Keystone BenefitTech, Inc.", "Acme Corp"
        ws.Cells.Replace "Keystone BenefitTech", "Acme Corp"
    Next ws

    ' Recalculate everything
    Application.CalculateFull

    ' Save sanitized version
    ThisWorkbook.SaveAs Replace(origPath, ".xls", "_SANITIZED.xls")
    MsgBox "Sanitization complete. Original saved as _ORIGINAL.", vbInformation
End Sub
```

---

## Verification Checklist

After sanitization, verify the output is safe to share:

```
□  1. Search for company name — zero hits across all sheets
□  2. GL vendor column — only "Vendor_NNN" patterns, no real names
□  3. GL ID column — only "TXN-NNNNNN" patterns
□  4. Dollar amounts — different from originals (spot-check 5 random cells)
□  5. Dates — shifted from original reporting period
□  6. Department names — only "Dept_X" patterns (if Tier 2 applied)
□  7. Formulas still work — reconciliation checks still PASS
□  8. Charts regenerate correctly — run Command 12
□  9. File name — does not contain real company name
□ 10. File properties — check File → Info for author/company metadata
```

### Metadata Cleanup

Excel files contain hidden metadata. Before sharing:

1. Go to **File → Info → Check for Issues → Inspect Document**
2. Check all categories and click **Inspect**
3. Click **Remove All** for: Document Properties, Personal Information, Hidden Rows/Columns (if desired)
4. Save the file

---

## Reversal Procedure

To restore the original data from a sanitized copy:

### If You Have the Original

Simply use the `_ORIGINAL` backup file created by the FullSanitize script.

### If You Only Have the Sanitized Copy + Parameters

**Amounts:** Divide all amounts by the scaling factor.
```vba
' Reverse: divide by the same factor used for sanitization
factor = 1.47
ws.Cells(r, 7).Value = Round(ws.Cells(r, 7).Value / factor, 2)
```

**Dates:** Subtract the same offset.
```vba
ws.Cells(r, 2).Value = ws.Cells(r, 2).Value - 90
```

**Vendor names, IDs, department names, company name:** These are one-way replacements. The original values cannot be recovered from the sanitized copy alone. Always maintain the `_ORIGINAL` backup.

---

## Sanitization Parameters Record

Keep this record secure and separate from the sanitized file:

```
Date sanitized:     _______________
Sanitized by:       _______________
Scaling factor:     _______________
Date offset (days): _______________
Company replaced:   Keystone BenefitTech → _______________
Original file:      _______________
Sanitized file:     _______________
```
