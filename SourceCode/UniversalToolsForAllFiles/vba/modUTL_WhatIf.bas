Attribute VB_Name = "modUTL_WhatIf"
Option Explicit

'===============================================================================
' modUTL_WhatIf - Universal What-If Scenario Tool
' Universal Toolkit - Works on ANY Excel file
'===============================================================================
' PURPOSE:  Let users run "What If" scenarios on any worksheet with numeric
'           data. User selects a range of driver cells, enters a % change,
'           and instantly sees the impact. A baseline is saved so the original
'           values can be restored with one click.
'
'           Works on any file — no assumptions about sheet structure, column
'           names, or workbook layout.
'
' PUBLIC SUBS:
'   RunWhatIf          - Apply a % change to user-selected cells
'   RunWhatIfPresets    - Quick menu: +/-5%, +/-10%, +/-25%, custom
'   RestoreBaseline     - Undo the last what-if and restore originals
'   ViewBaseline        - Show what baseline values are currently saved
'
' DEPENDENCIES: None (fully standalone)
' VERSION:  1.0.0
'===============================================================================

Private Const SH_BASELINE As String = "UTL_WhatIf_Backup"
Private Const SH_IMPACT   As String = "UTL_WhatIf_Impact"

'===============================================================================
' RunWhatIfPresets - Quick preset menu for common scenarios
'===============================================================================
Public Sub RunWhatIfPresets()
    On Error GoTo ErrHandler

    Dim msg As String
    msg = "WHAT-IF QUICK SCENARIOS" & vbCrLf & vbCrLf
    msg = msg & "Select a preset % change to apply to your selected cells:" & vbCrLf & vbCrLf
    msg = msg & "1.  Increase  5%" & vbCrLf
    msg = msg & "2.  Increase 10%" & vbCrLf
    msg = msg & "3.  Increase 25%" & vbCrLf
    msg = msg & "4.  Decrease  5%" & vbCrLf
    msg = msg & "5.  Decrease 10%" & vbCrLf
    msg = msg & "6.  Decrease 25%" & vbCrLf
    msg = msg & "7.  Custom %" & vbCrLf & vbCrLf
    msg = msg & "First SELECT the cells you want to change, then pick a preset."

    Dim choice As String
    choice = InputBox(msg, "Universal What-If Tool")
    If choice = "" Then Exit Sub
    If Not IsNumeric(choice) Then Exit Sub

    Dim pct As Double
    Select Case CLng(choice)
        Case 1: pct = 0.05
        Case 2: pct = 0.1
        Case 3: pct = 0.25
        Case 4: pct = -0.05
        Case 5: pct = -0.1
        Case 6: pct = -0.25
        Case 7
            Dim customPct As String
            customPct = InputBox("Enter percentage change:" & vbCrLf & vbCrLf & _
                "Examples:" & vbCrLf & _
                "  10  = increase by 10%" & vbCrLf & _
                "  -15 = decrease by 15%" & vbCrLf & _
                "  25  = increase by 25%", _
                "Universal What-If Tool - Custom %")
            If customPct = "" Then Exit Sub
            If Not IsNumeric(customPct) Then
                MsgBox "Enter a number (e.g., 10 or -15).", vbExclamation, "What-If"
                Exit Sub
            End If
            pct = CDbl(customPct) / 100
        Case Else
            MsgBox "Invalid choice. Pick 1-7.", vbExclamation, "What-If"
            Exit Sub
    End Select

    ApplyWhatIf pct
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "What-If Presets error: " & Err.Description, vbCritical, "What-If"
End Sub

'===============================================================================
' RunWhatIf - Apply a custom % change to selected cells
'===============================================================================
Public Sub RunWhatIf()
    On Error GoTo ErrHandler

    Dim pctStr As String
    pctStr = InputBox("Enter percentage change to apply to selected cells:" & vbCrLf & vbCrLf & _
        "Examples:" & vbCrLf & _
        "  10  = increase by 10%" & vbCrLf & _
        "  -15 = decrease by 15%" & vbCrLf & _
        "  25  = increase by 25%" & vbCrLf & vbCrLf & _
        "SELECT YOUR CELLS FIRST, then run this.", _
        "Universal What-If Tool")
    If pctStr = "" Then Exit Sub
    If Not IsNumeric(pctStr) Then
        MsgBox "Enter a number (e.g., 10 or -15).", vbExclamation, "What-If"
        Exit Sub
    End If

    ApplyWhatIf CDbl(pctStr) / 100
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "What-If error: " & Err.Description, vbCritical, "What-If"
End Sub

'===============================================================================
' ApplyWhatIf - Core engine: save baseline, apply %, build impact report
'===============================================================================
Private Sub ApplyWhatIf(ByVal pctChange As Double)
    On Error GoTo ErrHandler

    ' Validate selection
    Dim sel As Range
    On Error Resume Next
    Set sel = Selection
    On Error GoTo ErrHandler

    If sel Is Nothing Then
        MsgBox "Select the cells you want to change first.", vbExclamation, "What-If"
        Exit Sub
    End If

    ' Count numeric cells only
    Dim cell As Range
    Dim numCount As Long: numCount = 0
    For Each cell In sel.Cells
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            numCount = numCount + 1
        End If
    Next cell

    If numCount = 0 Then
        MsgBox "No numeric values found in selection." & vbCrLf & _
               "Select cells containing numbers.", vbExclamation, "What-If"
        Exit Sub
    End If

    ' Confirm
    Dim pctLabel As String
    If pctChange >= 0 Then
        pctLabel = "+" & Format(pctChange, "0%")
    Else
        pctLabel = Format(pctChange, "0%")
    End If

    If MsgBox("Apply " & pctLabel & " to " & numCount & " numeric cell(s)?" & vbCrLf & _
              "Sheet: " & sel.Worksheet.Name & vbCrLf & _
              "Range: " & sel.Address(False, False) & vbCrLf & vbCrLf & _
              "A baseline will be saved so you can undo this later.", _
              vbYesNo + vbQuestion, "Confirm What-If") = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    Application.StatusBar = "Saving baseline..."

    ' Save baseline
    SaveBaseline sel

    Application.StatusBar = "Applying " & pctLabel & "..."

    ' Collect before/after data and apply changes
    Dim sourceSheet As String: sourceSheet = sel.Worksheet.Name
    Dim addresses() As String
    Dim origVals() As Double
    Dim newVals() As Double
    Dim labels() As String
    ReDim addresses(1 To numCount)
    ReDim origVals(1 To numCount)
    ReDim newVals(1 To numCount)
    ReDim labels(1 To numCount)

    Dim idx As Long: idx = 0
    For Each cell In sel.Cells
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            idx = idx + 1
            addresses(idx) = cell.Address(False, False)
            origVals(idx) = CDbl(cell.Value)
            newVals(idx) = origVals(idx) * (1 + pctChange)
            cell.Value = newVals(idx)

            ' Try to get a label from column A or the row header
            Dim lbl As String: lbl = ""
            If cell.Column > 1 Then
                lbl = Trim(CStr(cell.Worksheet.Cells(cell.Row, 1).Value))
            End If
            If lbl = "" Then lbl = "Row " & cell.Row
            labels(idx) = lbl
        End If
    Next cell

    Application.Calculate
    DoEvents

    ' Build impact report
    Application.StatusBar = "Building impact report..."
    BuildImpactReport sourceSheet, pctLabel, addresses, labels, origVals, newVals, numCount

    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "What-If Applied: " & pctLabel & vbCrLf & vbCrLf & _
           numCount & " cell(s) changed on '" & sourceSheet & "'." & vbCrLf & _
           "Impact report created on '" & SH_IMPACT & "' sheet." & vbCrLf & vbCrLf & _
           "Run 'RestoreBaseline' to undo.", vbInformation, "What-If Complete"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "What-If error: " & Err.Description, vbCritical, "What-If"
End Sub

'===============================================================================
' RestoreBaseline - Undo the last what-if by restoring saved values
'===============================================================================
Public Sub RestoreBaseline()
    On Error GoTo ErrHandler

    If Not SheetExists(SH_BASELINE) Then
        MsgBox "No baseline saved." & vbCrLf & _
               "Run a What-If scenario first.", vbInformation, "What-If"
        Exit Sub
    End If

    If MsgBox("Restore all original values from the last What-If?" & vbCrLf & _
              "This will undo the changes and delete the backup.", _
              vbYesNo + vbQuestion, "Restore Baseline") = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    Application.StatusBar = "Restoring baseline..."

    Dim wsBL As Worksheet: Set wsBL = ThisWorkbook.Worksheets(SH_BASELINE)
    Dim lastRow As Long: lastRow = wsBL.Cells(wsBL.Rows.Count, 1).End(xlUp).Row

    Dim restored As Long: restored = 0
    Dim targetSheet As String
    Dim r As Long

    For r = 2 To lastRow  ' Row 1 = headers
        targetSheet = Trim(CStr(wsBL.Cells(r, 1).Value))
        Dim addr As String: addr = Trim(CStr(wsBL.Cells(r, 2).Value))
        Dim origVal As Double: origVal = CDbl(wsBL.Cells(r, 3).Value)

        If targetSheet <> "" And addr <> "" Then
            On Error Resume Next
            Dim wsTarget As Worksheet
            Set wsTarget = ThisWorkbook.Worksheets(targetSheet)
            If Not wsTarget Is Nothing Then
                wsTarget.Range(addr).Value = origVal
                restored = restored + 1
            End If
            Set wsTarget = Nothing
            On Error GoTo ErrHandler
        End If
    Next r

    Application.Calculate
    DoEvents

    ' Clean up
    SafeDeleteSheet SH_BASELINE
    SafeDeleteSheet SH_IMPACT

    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Baseline restored!" & vbCrLf & _
           restored & " cell(s) set back to original values." & vbCrLf & vbCrLf & _
           "The workbook has been recalculated.", vbInformation, "What-If Restored"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Restore error: " & Err.Description, vbCritical, "What-If"
End Sub

'===============================================================================
' ViewBaseline - Show what values are currently saved in the baseline
'===============================================================================
Public Sub ViewBaseline()
    On Error GoTo ErrHandler

    If Not SheetExists(SH_BASELINE) Then
        MsgBox "No baseline is currently saved." & vbCrLf & _
               "Run a What-If scenario first to create one.", _
               vbInformation, "What-If"
        Exit Sub
    End If

    Dim wsBL As Worksheet: Set wsBL = ThisWorkbook.Worksheets(SH_BASELINE)
    wsBL.Visible = xlSheetVisible
    wsBL.Activate

    MsgBox "Baseline sheet is now visible." & vbCrLf & _
           "This shows the original values saved before your last What-If." & vbCrLf & vbCrLf & _
           "Run 'RestoreBaseline' to put these values back.", _
           vbInformation, "What-If Baseline"
    Exit Sub

ErrHandler:
    MsgBox "View Baseline error: " & Err.Description, vbCritical, "What-If"
End Sub

'===============================================================================
' SaveBaseline - Save current values of selected cells to hidden sheet
'===============================================================================
Private Sub SaveBaseline(ByRef sel As Range)
    ' If baseline already exists, don't overwrite (prevents double-run issues)
    If SheetExists(SH_BASELINE) Then
        SafeDeleteSheet SH_BASELINE
    End If

    Dim wsBL As Worksheet
    Set wsBL = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsBL.Name = SH_BASELINE

    ' Headers
    wsBL.Cells(1, 1).Value = "Sheet"
    wsBL.Cells(1, 2).Value = "Cell Address"
    wsBL.Cells(1, 3).Value = "Original Value"
    wsBL.Cells(1, 4).Value = "Label"
    wsBL.Range("A1:D1").Font.Bold = True

    Dim outRow As Long: outRow = 2
    Dim cell As Range
    Dim sourceSheet As String: sourceSheet = sel.Worksheet.Name

    For Each cell In sel.Cells
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            wsBL.Cells(outRow, 1).Value = sourceSheet
            wsBL.Cells(outRow, 2).Value = cell.Address(False, False)
            wsBL.Cells(outRow, 3).Value = cell.Value

            ' Try to grab a label
            Dim lbl As String: lbl = ""
            If cell.Column > 1 Then
                lbl = Trim(CStr(cell.Worksheet.Cells(cell.Row, 1).Value))
            End If
            If lbl = "" Then lbl = "Row " & cell.Row
            wsBL.Cells(outRow, 4).Value = lbl

            outRow = outRow + 1
        End If
    Next cell

    wsBL.Columns("A:D").AutoFit
    wsBL.Visible = xlSheetVeryHidden
End Sub

'===============================================================================
' BuildImpactReport - Create styled before/after report sheet
'===============================================================================
Private Sub BuildImpactReport(ByVal sourceSheet As String, _
                               ByVal pctLabel As String, _
                               ByRef addresses() As String, _
                               ByRef labels() As String, _
                               ByRef origVals() As Double, _
                               ByRef newVals() As Double, _
                               ByVal cnt As Long)
    SafeDeleteSheet SH_IMPACT

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = SH_IMPACT

    ' Title block
    ws.Cells(1, 1).Value = "What-If Impact Report"
    ws.Cells(1, 1).Font.Size = 14: ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Color = RGB(11, 71, 121)

    ws.Cells(2, 1).Value = "Scenario: " & pctLabel & " applied to " & cnt & " cell(s)"
    ws.Cells(2, 1).Font.Size = 11: ws.Cells(2, 1).Font.Bold = True

    ws.Cells(3, 1).Value = "Source: " & sourceSheet & " | Generated: " & Format(Now, "mmmm d, yyyy h:mm AM/PM")
    ws.Cells(3, 1).Font.Italic = True: ws.Cells(3, 1).Font.Color = RGB(100, 100, 100)

    ' Header row
    Dim hdrRow As Long: hdrRow = 5
    Dim hdrCols As Variant
    hdrCols = Array("Label", "Cell", "Original Value", "New Value", "Change", "Change %")

    Dim c As Long
    For c = 0 To 5
        ws.Cells(hdrRow, c + 1).Value = hdrCols(c)
    Next c

    With ws.Range(ws.Cells(hdrRow, 1), ws.Cells(hdrRow, 6))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(11, 71, 121)  ' iPipeline Blue
        .Font.Name = "Arial"
    End With

    ' Column widths
    ws.Columns(1).ColumnWidth = 30
    ws.Columns(2).ColumnWidth = 10
    ws.Columns(3).ColumnWidth = 18
    ws.Columns(4).ColumnWidth = 18
    ws.Columns(5).ColumnWidth = 18
    ws.Columns(6).ColumnWidth = 12

    ' Data rows
    Dim r As Long: r = hdrRow + 1
    Dim totalOrig As Double: totalOrig = 0
    Dim totalNew As Double: totalNew = 0
    Dim i As Long

    For i = 1 To cnt
        ws.Cells(r, 1).Value = labels(i)
        ws.Cells(r, 2).Value = addresses(i)
        ws.Cells(r, 3).Value = origVals(i)
        ws.Cells(r, 3).NumberFormat = "#,##0.00"
        ws.Cells(r, 4).Value = newVals(i)
        ws.Cells(r, 4).NumberFormat = "#,##0.00"

        Dim chg As Double: chg = newVals(i) - origVals(i)
        ws.Cells(r, 5).Value = chg
        ws.Cells(r, 5).NumberFormat = "#,##0.00"

        If origVals(i) <> 0 Then
            ws.Cells(r, 6).Value = chg / origVals(i)
        End If
        ws.Cells(r, 6).NumberFormat = "+0.0%;-0.0%"

        ' Color coding
        If chg > 0 Then
            ws.Cells(r, 5).Font.Color = RGB(0, 128, 0)
            ws.Cells(r, 6).Font.Color = RGB(0, 128, 0)
        ElseIf chg < 0 Then
            ws.Cells(r, 5).Font.Color = RGB(200, 0, 0)
            ws.Cells(r, 6).Font.Color = RGB(200, 0, 0)
        End If

        ' Alternating row shading
        If r Mod 2 = 0 Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).Interior.Color = RGB(242, 242, 242)
        End If

        totalOrig = totalOrig + origVals(i)
        totalNew = totalNew + newVals(i)
        r = r + 1
    Next i

    ' Totals row
    r = r + 1
    ws.Cells(r, 1).Value = "TOTAL"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 3).Value = totalOrig
    ws.Cells(r, 3).NumberFormat = "#,##0.00"
    ws.Cells(r, 4).Value = totalNew
    ws.Cells(r, 4).NumberFormat = "#,##0.00"
    ws.Cells(r, 5).Value = totalNew - totalOrig
    ws.Cells(r, 5).NumberFormat = "#,##0.00"
    If totalOrig <> 0 Then
        ws.Cells(r, 6).Value = (totalNew - totalOrig) / totalOrig
    End If
    ws.Cells(r, 6).NumberFormat = "+0.0%;-0.0%"

    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 6))
        .Font.Bold = True
        .Interior.Color = RGB(17, 46, 81)  ' Navy
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Instructions
    r = r + 2
    ws.Cells(r, 1).Value = "HOW TO UNDO"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Color = RGB(11, 71, 121)

    ws.Cells(r + 1, 1).Value = "Run modUTL_WhatIf.RestoreBaseline to put all values back to their originals."
    ws.Cells(r + 1, 1).Font.Bold = True
    ws.Cells(r + 1, 1).Font.Color = RGB(180, 0, 0)

    ws.Cells(r + 3, 1).Value = "This report was generated by the Universal What-If Tool."
    ws.Cells(r + 3, 1).Font.Italic = True
    ws.Cells(r + 3, 1).Font.Color = RGB(150, 150, 150)

    ws.Activate
    ws.Cells(1, 1).Select
End Sub

'===============================================================================
' Helper: SheetExists
'===============================================================================
Private Function SheetExists(ByVal shName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(shName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

'===============================================================================
' Helper: SafeDeleteSheet
'===============================================================================
Private Sub SafeDeleteSheet(ByVal shName As String)
    If Not SheetExists(shName) Then Exit Sub
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(shName).Delete
    Application.DisplayAlerts = True
End Sub
