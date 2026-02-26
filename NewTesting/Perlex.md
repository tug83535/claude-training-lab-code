# Excel VBA Macro Library & Teaching Outline (Complete)

**Purpose:** Provide a complete list of Excel VBA macros (with code where applicable) ranging from easy to medium-hard, focused on general office and finance-heavy workflows. Designed so another AI can understand, compare to an existing workbook, and generate/modify code.

**Document Version:** 1.0  
**Date:** February 25, 2026

---

## Global Conventions

- Use `Option Explicit` in every module.
- Avoid `.Select` / `.Activate` where possible; use fully-qualified references.
- For heavy macros, use this performance wrapper:

```vba
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

' ... main logic here ...

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
```

---

## 🟢 EASY – Anyone Can Use These Day 1

### 1. AutoFit All Column Widths

**What it does / Why it's useful**  
Automatically resizes every column on the active sheet so all text is visible, removing the need to manually drag column borders. Very handy for data dumps, CSV imports, and quick report cleanup.

```vba
Option Explicit

Sub AutoFitAll()
    Application.ScreenUpdating = False
    ActiveSheet.Cells.EntireColumn.AutoFit
    Application.ScreenUpdating = True
End Sub
```

---

### 2. Freeze/Unfreeze Panes Toggle

**What it does / Why it's useful**  
Turns freeze panes on or off with a single click. When enabling, it freezes row 1 and column A (via cell B2) so headers stay visible while you scroll.

```vba
Option Explicit

Sub ToggleFreezePanes()
    If ActiveWindow.FreezePanes Then
        ActiveWindow.FreezePanes = False
    Else
        Cells(2, 2).Select
        ActiveWindow.FreezePanes = True
    End If
End Sub
```

---

### 3. Convert Formulas to Values (Selection)

**What it does / Why it's useful**  
Replaces formulas in the selected range with their calculated values. Essential when finalizing finance files for sharing so formulas don't break or expose internal logic.

```vba
Option Explicit

Sub ConvertToValues()
    If TypeName(Selection) = "Range" Then
        Selection.Value = Selection.Value
    Else
        MsgBox "Please select a range first."
    End If
End Sub
```

---

### 4. Clear All Hyperlinks on a Sheet

**What it does / Why it's useful**  
Removes every hyperlink object on the active sheet while leaving the display text intact. Great for cleaning data pasted from websites, bank portals, or emails.

```vba
Option Explicit

Sub ClearHyperlinks()
    ActiveSheet.Hyperlinks.Delete
End Sub
```

---

### 5. Highlight Duplicates in a Selection

**What it does / Why it's useful**  
Prompts the user to select a range, then highlights any values that appear more than once. Very useful for catching duplicate invoices, IDs, or customer records.

```vba
Option Explicit

Sub HighlightDuplicates()
    Dim rng As Range, cell As Range

    On Error Resume Next
    Set rng = Application.InputBox("Select a range:", Type:=8)
    On Error GoTo 0

    If rng Is Nothing Then
        MsgBox "No range selected."
        Exit Sub
    End If

    For Each cell In rng
        If WorksheetFunction.CountIf(rng, cell.Value) > 1 Then
            cell.Interior.ColorIndex = 6 ' Yellow
        End If
    Next cell
End Sub
```

---

### 6. Unhide All Worksheets

**What it does / Why it's useful**  
Loops through the workbook and unhides every worksheet. Saves a ton of time when you inherit a workbook with many hidden or very hidden tabs.

```vba
Option Explicit

Sub UnhideAllSheets()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws
End Sub
```

---

### 7. Quick Format Header Row

**What it does / Why it's useful**  
Applies a standard header style (bold, blue fill, white text, bottom border) to row 1 and auto-fits the columns. Instantly makes any report look more professional.

```vba
Option Explicit

Sub FormatHeaderRow()
    With Rows(1)
        .Font.Bold = True
        .Font.Size = 11
        .Interior.Color = RGB(68, 114, 196) ' Blue
        .Font.Color = RGB(255, 255, 255)    ' White
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .EntireColumn.AutoFit
    End With
End Sub
```

---

## 🟡 EASY–MEDIUM – Starting to Save Real Time

### 8. Delete All Blank Rows in a Selection

**What it does / Why it's useful**  
Looks at each row in your selected range and deletes it if the row is completely empty. Perfect for cleaning messy exports from ERPs, bank files, or CSVs.

```vba
Option Explicit

Sub DeleteBlankRows()
    Dim rng As Range
    Dim i As Long

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first."
        Exit Sub
    End If

    Set rng = Selection

    For i = rng.Rows.Count To 1 Step -1
        If Application.WorksheetFunction.CountA(rng.Rows(i)) = 0 Then
            rng.Rows(i).EntireRow.Delete
        End If
    Next i
End Sub
```

---

### 9. Protect / Unprotect All Sheets (with Password)

**What it does / Why it's useful**  
Protects or unprotects every worksheet in the workbook with a single password prompt. This is critical for locking down budget or payroll files before sending them out.

```vba
Option Explicit

Sub ProtectAllSheets()
    Dim ws As Worksheet
    Dim pwd As String

    pwd = InputBox("Enter password to protect all sheets:")
    If pwd = "" Then Exit Sub

    For Each ws In ActiveWorkbook.Worksheets
        ws.Protect Password:=pwd
    Next ws

    MsgBox "All sheets protected."
End Sub

Sub UnprotectAllSheets()
    Dim ws As Worksheet
    Dim pwd As String

    pwd = InputBox("Enter password to unprotect all sheets:")
    If pwd = "" Then Exit Sub

    For Each ws In ActiveWorkbook.Worksheets
        ws.Unprotect Password:=pwd
    Next ws

    MsgBox "All sheets unprotected."
End Sub
```

---

### 10. Save Active Sheet as PDF (Dated Filename)

**What it does / Why it's useful**  
Exports the current sheet as a PDF in the workbook's folder, naming it with the sheet name and current date. Ideal for invoices, financial statements, and one-off reports.

```vba
Option Explicit

Sub SaveSheetAsPDF()
    Dim filePath As String

    If ThisWorkbook.Path = "" Then
        MsgBox "Please save the workbook first."
        Exit Sub
    End If

    filePath = ThisWorkbook.Path & "\" & _
               ActiveSheet.Name & "_" & Format(Date, "YYYY-MM-DD") & ".pdf"

    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=filePath
    MsgBox "Saved as PDF: " & filePath
End Sub
```

---

### 11. Backup Workbook with Timestamp

**What it does / Why it's useful**  
Creates a copy of the current workbook in the same folder with a timestamp in the filename. This gives you an instant rollback point before running risky macros or structural changes.

```vba
Option Explicit

Sub BackupWorkbook()
    Dim backupName As String

    If ThisWorkbook.Path = "" Then
        MsgBox "Please save the workbook first."
        Exit Sub
    End If

    backupName = ThisWorkbook.Path & "\BACKUP_" & _
                 Format(Now, "YYYY-MM-DD_HHMMSS") & "_" & ThisWorkbook.Name

    ThisWorkbook.SaveCopyAs backupName
    MsgBox "Backup saved: " & backupName
End Sub
```

---

### 12. Sort Sheets Alphabetically

**What it does / Why it's useful**  
Reorders all worksheet tabs alphabetically by sheet name. Super helpful when you have many tabs by client, region, or month and want consistent navigation.

```vba
Option Explicit

Sub SortSheetsAlphabetically()
    Dim i As Long, j As Long

    For i = 1 To Sheets.Count - 1
        For j = i + 1 To Sheets.Count
            If UCase$(Sheets(j).Name) < UCase$(Sheets(i).Name) Then
                Sheets(j).Move Before:=Sheets(i)
            End If
        Next j
    Next i
End Sub
```

---

## 🟠 MEDIUM – Real Workflow Automation

### 13. Create a Table of Contents (Hyperlinked)

**What it does / Why it's useful**  
Builds a "TOC" sheet that lists every worksheet in the workbook with a clickable hyperlink to cell A1 of each sheet. This massively improves navigation in big finance workbooks and is loved by auditors.

```vba
Option Explicit

Sub CreateTOC()
    Dim ws As Worksheet, tocSheet As Worksheet
    Dim r As Long

    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("TOC").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set tocSheet = Sheets.Add(Before:=Sheets(1))
    tocSheet.Name = "TOC"
    tocSheet.Range("A1").Value = "Table of Contents"
    tocSheet.Range("A1").Font.Bold = True

    r = 3
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "TOC" Then
            tocSheet.Hyperlinks.Add _
                Anchor:=tocSheet.Cells(r, 1), _
                Address:="", _
                SubAddress:="'" & ws.Name & "'!A1", _
                TextToDisplay:=ws.Name
            r = r + 1
        End If
    Next ws

    tocSheet.Columns("A").AutoFit
End Sub
```

---

### 14. Consolidate All Sheets into One "Master" Sheet

**What it does / Why it's useful**  
Creates a "Master" sheet and appends the used range from every other sheet into a single stacked table. This kills the manual copy-paste grind when consolidating monthly P&Ls, regional reports, or department data.

```vba
Option Explicit

Sub ConsolidateSheets()
    Dim ws As Worksheet, master As Worksheet
    Dim pasteRow As Long

    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("Master").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set master = Sheets.Add(After:=Sheets(Sheets.Count))
    master.Name = "Master"
    pasteRow = 1

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Master" Then
            ws.UsedRange.Copy
            master.Cells(pasteRow, 1).PasteSpecial xlPasteValues
            pasteRow = master.Cells(master.Rows.Count, 1).End(xlUp).Row + 1
        End If
    Next ws

    Application.CutCopyMode = False
End Sub
```

---

### 15. Bulk Find and Replace Across Entire Workbook

**What it does / Why it's useful**  
Prompts for a "find" and "replace" string and performs that replacement on every worksheet. Ideal for updating fiscal years, cost center names, or account label changes across massive workbooks.

```vba
Option Explicit

Sub FindReplaceAllSheets()
    Dim ws As Worksheet
    Dim findText As String, replaceText As String

    findText = InputBox("Find what:")
    If findText = "" Then Exit Sub

    replaceText = InputBox("Replace with:")
    If replaceText = "" And _
       MsgBox("Replace with blank?", vbYesNo) = vbNo Then Exit Sub

    For Each ws In ActiveWorkbook.Worksheets
        ws.Cells.Replace What:=findText, Replacement:=replaceText, _
                         LookAt:=xlPart, MatchCase:=False
    Next ws

    MsgBox "Done! Replaced across all sheets."
End Sub
```

---

### 16. Export Each Sheet as a Separate PDF

**What it does / Why it's useful**  
Loops through every sheet and exports each one as its own PDF file in the workbook folder. Perfect for sending each department or client their own individual report or budget tab.

```vba
Option Explicit

Sub ExportEachSheetAsPDF()
    Dim ws As Worksheet
    Dim folderPath As String

    If ThisWorkbook.Path = "" Then
        MsgBox "Please save the workbook first."
        Exit Sub
    End If

    folderPath = ThisWorkbook.Path & "\"

    For Each ws In ActiveWorkbook.Worksheets
        ws.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=folderPath & ws.Name & ".pdf"
    Next ws

    MsgBox "All sheets exported as PDFs!"
End Sub
```

---

### 17. Email the Active Workbook via Outlook

**What it does / Why it's useful**  
Creates a new Outlook email with the current workbook already attached and a pre-populated subject/body. It removes the tedious step of going to Outlook, drafting, and manually attaching the file each time.

```vba
Option Explicit

Sub EmailWorkbook()
    Dim outApp As Object, outMail As Object

    On Error Resume Next
    Set outApp = CreateObject("Outlook.Application")
    On Error GoTo 0

    If outApp Is Nothing Then
        MsgBox "Outlook is not available."
        Exit Sub
    End If

    Set outMail = outApp.CreateItem(0)

    With outMail
        .To = ""
        .Subject = "Report - " & Format(Date, "yyyy-mm-dd")
        .Body = "Hi," & vbNewLine & vbNewLine & _
                "Please find the attached report." & vbNewLine & "Thanks."
        .Attachments.Add ThisWorkbook.FullName
        .Display ' or .Send
    End With

    Set outMail = Nothing
    Set outApp = Nothing
End Sub
```

---

### 18. Data Entry UserForm (Specification)

**What it does / Why it's useful**  
Provides a simple pop-up form where users can enter data (e.g., expenses, AP entries) without touching the raw sheet, which reduces errors and keeps structure consistent.

**Specification (for AI to build code)**  
- Create a UserForm with fields: Date, Vendor, Amount, GL Code, Description, plus Submit and Cancel buttons.  
- On Submit: validate required fields, check Amount is numeric, then write a new row to Sheet("Data") in the next empty row (e.g., columns A:E).  
- On Cancel: close the form without writing anything.  

---

## 🔴 MEDIUM–HARD – High-Impact Finance & Dashboards

### 19. Automated Overdue Invoice Email Reminders (Spec)

**What it does / Why it's useful**  
Scans an invoice table for past-due items and sends polite reminder emails through Outlook, while tracking when reminders were sent so you don't spam clients.

**Specification**  
- Sheet "Invoices" columns: InvoiceID, ClientName, ClientEmail, Amount, DueDate, Status, LastReminderDate.  
- For each row where: Status = "Open", DueDate < Today, and LastReminderDate is blank or older than X days (e.g., 7), do:  
  - Create Outlook email to ClientEmail with invoice details and a reminder message.  
  - After sending/displaying, update LastReminderDate to Today.  
  - Optionally append a record to a "RemindersLog" sheet.

---

### 20. Financial Statement Generator from Trial Balance (Spec)

**What it does / Why it's useful**  
Takes a raw trial balance and a mapping table, then automatically builds formatted Income Statement and Balance Sheet tabs. This can turn a manual month-end process into a repeatable macro.

**Specification**  
- Inputs:  
  - Sheet "TB": AccountNumber, AccountName, Debit, Credit.  
  - Sheet "Mapping": AccountNumber, StatementType (IS/BS), Section (Revenue, COGS, OpEx, Assets, etc.), DisplayName.  
- Outputs:  
  - Sheet "IncomeStatement", Sheet "BalanceSheet".  
- Logic:  
  - Join TB to Mapping by AccountNumber.  
  - For IS: aggregate by Section, compute subtotals (Revenue, Gross Profit, Operating Income, Net Income).  
  - For BS: aggregate by Assets, Liabilities, Equity, and check that Assets = Liabilities + Equity.  
  - Write clean, labeled, and formatted statements.

---

### 21. Export All Charts to PowerPoint (Spec)

**What it does / Why it's useful**  
Creates a new PowerPoint presentation and sends each Excel chart onto its own slide automatically, eliminating repetitive copy-paste when preparing decks.

**Specification**  
- Open or create PowerPoint and a new presentation.  
- Loop through all worksheets and all ChartObjects.  
- For each chart:  
  - Add a new slide.  
  - Copy the chart and paste it into the slide.  
  - Optionally set slide title to chart name or sheet name.

---

### 22. Auto-Refresh Pivot Tables on Workbook Open

**What it does / Why it's useful**  
Automatically refreshes all pivot tables when the workbook opens, so dashboards and summary reports always reflect the latest data.

```vba
Option Explicit

Private Sub Workbook_Open()
    Dim ws As Worksheet
    Dim pt As PivotTable

    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws
End Sub
```

*(Place this code in the ThisWorkbook module.)*

---

### 23. Dynamic Progress Bar (KPI Shape – Spec)

**What it does / Why it's useful**  
Visually shows progress toward a goal using a bar that grows and changes color based on a percentage. Useful for dashboards tracking sales, budget utilization, or project progress.

**Specification**  
- Assume a cell, e.g. Dashboard!B2, holds a percentage (0–1 or 0–100%).  
- Assume a rectangular shape named "ProgressBar" on Dashboard.  
- Macro should:  
  - Set ProgressBar.Width proportional to the percentage.  
  - Set color:  
    - <50% = red  
    - 50–79% = orange  
    - ≥80% = green.  
- Trigger via a button or via Worksheet_Change monitoring that KPI cell.

---

### 24. Timestamp Audit Trail on Cell Changes (Spec)

**What it does / Why it's useful**  
Logs changes to critical cells, including user, timestamp, and old/new values, providing an audit trail for sensitive financial models.

**Specification**  
- Create an "AuditLog" sheet with columns: DateTime, UserName, SheetName, CellAddress, OldValue, NewValue.  
- On key input sheets, use Worksheet_Change (and optionally SelectionChange) to detect edits to a defined monitored range.  
- When a monitored cell changes, append a row to AuditLog with all fields populated.

---

## Additional Best Practices

### Performance Optimization
- Always disable ScreenUpdating, EnableEvents, and set Calculation to Manual during long-running macros
- Use arrays for bulk data operations instead of cell-by-cell loops
- Read ranges into memory, process, then write back once

### Error Handling
- Use On Error GoTo ErrorHandler in all production code
- Log errors to a hidden ErrorLog sheet for finance/audit purposes
- Always restore Application settings in error handler

### Code Structure
- Use Option Explicit at top of every module
- Keep procedures under 50-75 lines (break into smaller functions if longer)
- Use descriptive variable names: totalAmount not ta
- Avoid .Select and .Activate - use direct object references

### Testing
- Always test on a copy of production data first
- Test edge cases: empty cells, text in numeric columns, special characters
- Use Debug.Print to verify logic before full execution

---

## How Another AI Should Use This File

- Treat each numbered item (1–24) as a distinct feature.  
- For items with code blocks, compare to existing VBA modules to find overlaps and refactor opportunities.  
- For spec-only items, generate new VBA wired to the described sheet layouts and behaviors.  
- Prioritize high-impact finance items:  
  - PDFs, backups, consolidation, emailing (10, 11, 14, 16, 17)  
  - Finance automations (19, 20, 22, 24).

---

**End of Document**