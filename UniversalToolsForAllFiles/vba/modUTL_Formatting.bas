Attribute VB_Name = "modUTL_Formatting"
Option Explicit

' ============================================================
' KBT Universal Tools — Formatting Module
' Works on ANY Excel file — no project-specific setup required
' Install in Personal.xlsb to use across all Excel sessions
' Tools: 9 | Tier 1: 6 | Tier 2: 3
' ============================================================

Private Sub UTL_TurboOn()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

Private Sub UTL_TurboOff()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

' ============================================================
' TOOL 1 — AutoFit All Columns & Rows               [TIER 1]
' Auto-fits every column and row across the entire workbook
' Run: no selection needed — works on all sheets
' ============================================================
Sub AutoFitAllColumnsRows()
    Dim scopeChoice As Integer
    scopeChoice = MsgBox("AutoFit ALL sheets in this workbook?" & Chr(10) & _
                         "Click Yes for all sheets, No for active sheet only.", _
                         vbQuestion + vbYesNo, "UTL Formatting")

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim ws As Worksheet
    If scopeChoice = vbYes Then
        For Each ws In ActiveWorkbook.Worksheets
            ws.Cells.EntireColumn.AutoFit
            ws.Cells.EntireRow.AutoFit
        Next ws
        MsgBox "Done! All columns and rows auto-fitted on every sheet.", vbInformation, "UTL Formatting"
    Else
        ActiveSheet.Cells.EntireColumn.AutoFit
        ActiveSheet.Cells.EntireRow.AutoFit
        MsgBox "Done! All columns and rows auto-fitted on active sheet.", vbInformation, "UTL Formatting"
    End If

    UTL_TurboOff
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Formatting"
End Sub

' ============================================================
' TOOL 2 — Freeze Top Row on All Sheets             [TIER 1]
' Applies Freeze Panes to row 1 on every worksheet
' Finance standard — makes every file easier to read
' ============================================================
Sub FreezeTopRowAllSheets()
    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        ActiveWindow.FreezePanes = False
        ws.Rows(2).Select
        ActiveWindow.FreezePanes = True
    Next ws

    ActiveWorkbook.Sheets(1).Activate
    UTL_TurboOff
    MsgBox "Done! Top row frozen on every sheet.", vbInformation, "UTL Formatting"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Formatting"
End Sub

' ============================================================
' TOOL 3 — Number Format Standardizer               [TIER 1]
' Applies #,##0.00 to all numeric cells across all sheets
' Skips text, dates, and cells with existing currency formats
' ============================================================
Sub NumberFormatStandardizer()
    If MsgBox("Apply standard number format (#,##0.00) to all numeric cells across all sheets?", _
              vbQuestion + vbYesNo, "UTL Formatting") = vbNo Then Exit Sub

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim count As Long
    Dim ws As Worksheet
    Dim c As Range

    For Each ws In ActiveWorkbook.Worksheets
        For Each c In ws.UsedRange
            If Not c.HasFormula And IsNumeric(c.Value) And Not IsEmpty(c) Then
                If Not IsDate(c.Value) Then
                    If InStr(c.NumberFormat, "$") = 0 And _
                       InStr(c.NumberFormat, "d/") = 0 And _
                       InStr(c.NumberFormat, "m/") = 0 Then
                        c.NumberFormat = "#,##0.00"
                        count = count + 1
                    End If
                End If
            End If
        Next c
    Next ws

    UTL_TurboOff
    MsgBox "Done! " & count & " cells formatted as numbers.", vbInformation, "UTL Formatting"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Formatting"
End Sub

' ============================================================
' TOOL 4 — Currency Format Standardizer             [TIER 1]
' Applies $#,##0.00 to all cells in the selected range
' Run: select the currency columns first
' ============================================================
Sub CurrencyFormatStandardizer()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select the cells to format as currency.", vbExclamation, "UTL Formatting"
        Exit Sub
    End If

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim count As Long
    Dim c As Range
    For Each c In Selection
        If IsNumeric(c.Value) And Not IsEmpty(c) And Not IsDate(c.Value) Then
            c.NumberFormat = "$#,##0.00"
            count = count + 1
        End If
    Next c

    UTL_TurboOff
    MsgBox "Done! " & count & " cells formatted as currency ($#,##0.00).", vbInformation, "UTL Formatting"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Formatting"
End Sub

' ============================================================
' TOOL 5 — Date Format Standardizer                 [TIER 1]
' Normalizes all date cells to MM/DD/YYYY across all sheets
' Kills the mixed date format problem from imports
' ============================================================
Sub DateFormatStandardizer()
    If MsgBox("Standardize all date cells to MM/DD/YYYY across all sheets?", _
              vbQuestion + vbYesNo, "UTL Formatting") = vbNo Then Exit Sub

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim count As Long
    Dim ws As Worksheet
    Dim c As Range

    For Each ws In ActiveWorkbook.Worksheets
        For Each c In ws.UsedRange
            If IsDate(c.Value) And Not IsEmpty(c) Then
                c.NumberFormat = "MM/DD/YYYY"
                count = count + 1
            End If
        Next c
    Next ws

    UTL_TurboOff
    MsgBox "Done! " & count & " date cells standardized to MM/DD/YYYY.", vbInformation, "UTL Formatting"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Formatting"
End Sub

' ============================================================
' TOOL 6 — Highlight Negatives in Red               [TIER 1]
' Applies conditional formatting to show all negatives in red
' Finance standard — instant loss visibility across all sheets
' ============================================================
Sub HighlightNegativesRed()
    Dim scopeChoice As Integer
    scopeChoice = MsgBox("Apply red highlighting for negative numbers?" & Chr(10) & _
                         "Yes = all sheets | No = active sheet only", _
                         vbQuestion + vbYesNo, "UTL Formatting")

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim ws As Worksheet
    Dim rng As Range
    Dim fc As FormatCondition

    If scopeChoice = vbYes Then
        For Each ws In ActiveWorkbook.Worksheets
            Set rng = ws.UsedRange
            rng.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
            Set fc = rng.FormatConditions(rng.FormatConditions.Count)
            fc.Font.Color = RGB(192, 0, 0)
            fc.Font.Bold = True
        Next ws
        MsgBox "Done! Negative values will display red on all sheets.", vbInformation, "UTL Formatting"
    Else
        Set rng = ActiveSheet.UsedRange
        rng.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        Set fc = rng.FormatConditions(rng.FormatConditions.Count)
        fc.Font.Color = RGB(192, 0, 0)
        fc.Font.Bold = True
        MsgBox "Done! Negative values will display red on this sheet.", vbInformation, "UTL Formatting"
    End If

    UTL_TurboOff
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Formatting"
End Sub

' ============================================================
' TOOL 7 — Financial Number Formatting Suite        [TIER 2]
' Applies standard Finance formats via choice menu
' Options: Accounting | Factor (000s) | Percentage | Plain
' ============================================================
Sub FinancialNumberFormattingSuite()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first.", vbExclamation, "UTL Formatting"
        Exit Sub
    End If

    Dim choice As String
    choice = InputBox("Choose a format to apply to your selection:" & Chr(10) & Chr(10) & _
                      "1 — Accounting   ($  1,250.00)" & Chr(10) & _
                      "2 — Factor 000s  (    1,250  )" & Chr(10) & _
                      "3 — Percentage   (   12.50%  )" & Chr(10) & _
                      "4 — Plain Number ( 1,250.00  )" & Chr(10) & _
                      "5 — Integer      (    1,250  )" & Chr(10) & Chr(10) & _
                      "Type the number (1-5):", "UTL Formatting", "1")

    If choice = "" Then Exit Sub

    Dim fmt As String
    Select Case choice
        Case "1": fmt = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        Case "2": fmt = "#,##0,;[Red](#,##0,)"
        Case "3": fmt = "0.00%"
        Case "4": fmt = "#,##0.00"
        Case "5": fmt = "#,##0"
        Case Else
            MsgBox "Invalid choice. Please enter 1, 2, 3, 4, or 5.", vbExclamation, "UTL Formatting"
            Exit Sub
    End Select

    On Error GoTo ErrHandler
    Selection.NumberFormat = fmt
    MsgBox "Done! Format applied to " & Selection.Cells.Count & " cells.", vbInformation, "UTL Formatting"
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Formatting"
End Sub

' ============================================================
' TOOL 8 — Conditional Format Purger                [TIER 2]
' Lists and removes redundant conditional formatting rules
' Reduces file bloat caused by excessive CF rules
' ============================================================
Sub ConditionalFormatPurger()
    Dim ws As Worksheet
    Dim totalRules As Long

    For Each ws In ActiveWorkbook.Worksheets
        totalRules = totalRules + ws.Cells.FormatConditions.Count
    Next ws

    If totalRules = 0 Then
        MsgBox "No conditional formatting rules found in this workbook.", vbInformation, "UTL Formatting"
        Exit Sub
    End If

    Dim choice As Integer
    choice = MsgBox("Found " & totalRules & " conditional formatting rules across all sheets." & Chr(10) & Chr(10) & _
                    "Click YES to delete ALL rules (clean slate)." & Chr(10) & _
                    "Click NO to cancel.", _
                    vbExclamation + vbYesNo, "UTL Formatting")

    If choice = vbNo Then Exit Sub

    On Error GoTo ErrHandler
    UTL_TurboOn

    For Each ws In ActiveWorkbook.Worksheets
        ws.Cells.FormatConditions.Delete
    Next ws

    UTL_TurboOff
    MsgBox "Done! All " & totalRules & " conditional formatting rules removed.", vbInformation, "UTL Formatting"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Formatting"
End Sub

' ============================================================
' TOOL 9 — Print Header/Footer Standardizer         [TIER 2]
' Applies consistent print headers/footers across all sheets
' Header: company name left, filename center | Footer: page numbers
' ============================================================
Sub PrintHeaderFooterStandardizer()
    Dim companyName As String
    companyName = InputBox("Enter your company name for the header:", _
                           "UTL Formatting", "iPipeline")
    If companyName = "" Then Exit Sub

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        With ws.PageSetup
            .LeftHeader   = "&B" & companyName
            .CenterHeader = "&B&14" & ActiveWorkbook.Name
            .RightHeader  = "Date Printed: &D"
            .LeftFooter   = "Confidential — For Internal Use Only"
            .CenterFooter = "Page &P of &N"
            .RightFooter  = "&F — &A"
        End With
    Next ws

    UTL_TurboOff
    MsgBox "Done! Print headers and footers standardized on all " & _
           ActiveWorkbook.Sheets.Count & " sheets.", vbInformation, "UTL Formatting"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Formatting"
End Sub
