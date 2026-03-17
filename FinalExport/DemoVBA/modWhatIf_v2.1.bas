Attribute VB_Name = "modWhatIf"
Option Explicit

'===============================================================================
' modWhatIf - Live "What If" Scenario Demo
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  One-click demo scenarios that speak the CFO's language.
'           "What if revenue drops 15%?" or "What if AWS costs increase 10%?"
'           Modifies Assumptions, recalculates the model, and shows instant
'           P&L impact on a styled summary sheet. Includes a "Restore Original"
'           option so the demo can be reset cleanly.
'
' PUBLIC SUBS:
'   RunWhatIfDemo        - Show menu of preset scenarios and run the selected one
'   QuickWhatIf          - Custom what-if: user picks a driver and % change
'   RestoreBaseline      - Restore original Assumptions values saved before what-if
'
' DEPENDENCIES: modConfig, modPerformance, modLogger
' VERSION:  2.1.0
'===============================================================================

Private Const SH_WHATIF As String = "What-If Impact"
Private Const SH_BASELINE As String = "WhatIf_Baseline"

'===============================================================================
' RunWhatIfDemo - Preset scenario menu for live demo
'===============================================================================
Public Sub RunWhatIfDemo()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_ASSUMPTIONS) Then
        MsgBox "Assumptions sheet not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    Dim msg As String
    msg = "SELECT A WHAT-IF SCENARIO:" & vbCrLf & vbCrLf
    msg = msg & "1.  Revenue drops 15%" & vbCrLf
    msg = msg & "2.  Revenue increases 10%" & vbCrLf
    msg = msg & "3.  AWS costs increase 25%" & vbCrLf
    msg = msg & "4.  Headcount grows 20%" & vbCrLf
    msg = msg & "5.  All expenses cut 10%" & vbCrLf
    msg = msg & "6.  Best case: Revenue +15%, Expenses -5%" & vbCrLf
    msg = msg & "7.  Worst case: Revenue -20%, Expenses +15%" & vbCrLf & vbCrLf
    msg = msg & "8.  Custom (pick your own driver & %)" & vbCrLf
    msg = msg & "9.  Restore original values" & vbCrLf

    Dim choice As String
    choice = InputBox(msg, APP_NAME & " - What-If Scenario Demo")
    If choice = "" Then Exit Sub
    If Not IsNumeric(choice) Then Exit Sub

    Dim sel As Long: sel = CLng(choice)

    Select Case sel
        Case 1: ApplyPresetScenario "Revenue Drops 15%", "rev", -0.15
        Case 2: ApplyPresetScenario "Revenue Increases 10%", "rev", 0.1
        Case 3: ApplyPresetScenario "AWS Costs Increase 25%", "aws", 0.25
        Case 4: ApplyPresetScenario "Headcount Grows 20%", "head", 0.2
        Case 5: ApplyPresetScenario "All Expenses Cut 10%", "expense", -0.1
        Case 6: ApplyComboScenario "Best Case", "rev", 0.15, "expense", -0.05
        Case 7: ApplyComboScenario "Worst Case", "rev", -0.2, "expense", 0.15
        Case 8: QuickWhatIf
        Case 9: RestoreBaseline
        Case Else
            MsgBox "Invalid selection. Choose 1-9.", vbExclamation, APP_NAME
    End Select
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modWhatIf", "ERROR", Err.Description
    MsgBox "What-If error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' QuickWhatIf - Custom: user picks any Assumptions driver and a % change
'===============================================================================
Public Sub QuickWhatIf()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_ASSUMPTIONS) Then
        MsgBox "Assumptions sheet not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    Dim wsA As Worksheet: Set wsA = ThisWorkbook.Worksheets(SH_ASSUMPTIONS)
    Dim lastRow As Long: lastRow = modConfig.LastRow(wsA, 1)

    ' Build driver list
    Dim driverList As String: driverList = ""
    Dim driverCount As Long: driverCount = 0
    Dim r As Long
    For r = DATA_ROW_ASSUME To lastRow
        Dim dName As String: dName = Trim(CStr(wsA.Cells(r, 1).Value))
        If dName <> "" Then
            driverCount = driverCount + 1
            driverList = driverList & driverCount & ". " & dName & " = " & _
                         Format(wsA.Cells(r, 2).Value, "#,##0.00") & vbCrLf
        End If
    Next r

    Dim driverChoice As String
    driverChoice = InputBox("Select a driver to change:" & vbCrLf & vbCrLf & driverList, _
                            APP_NAME & " - Custom What-If")
    If driverChoice = "" Then Exit Sub
    If Not IsNumeric(driverChoice) Then Exit Sub

    Dim driverIdx As Long: driverIdx = CLng(driverChoice)
    If driverIdx < 1 Or driverIdx > driverCount Then
        MsgBox "Invalid selection.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim pctStr As String
    pctStr = InputBox("Enter percentage change:" & vbCrLf & vbCrLf & _
                      "Examples:" & vbCrLf & _
                      "  10  = increase by 10%" & vbCrLf & _
                      "  -15 = decrease by 15%" & vbCrLf & _
                      "  25  = increase by 25%", _
                      APP_NAME & " - Percentage Change")
    If pctStr = "" Then Exit Sub
    If Not IsNumeric(pctStr) Then
        MsgBox "Enter a number (e.g., 10 or -15).", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim pctChange As Double: pctChange = CDbl(pctStr) / 100

    ' Find the driver row
    Dim targetRow As Long: targetRow = 0
    Dim cnt As Long: cnt = 0
    For r = DATA_ROW_ASSUME To lastRow
        If Trim(CStr(wsA.Cells(r, 1).Value)) <> "" Then
            cnt = cnt + 1
            If cnt = driverIdx Then
                targetRow = r
                Exit For
            End If
        End If
    Next r

    If targetRow = 0 Then Exit Sub

    Dim driverName As String: driverName = Trim(CStr(wsA.Cells(targetRow, 1).Value))
    Dim scenarioName As String
    If pctChange >= 0 Then
        scenarioName = driverName & " +" & Format(pctChange, "0%")
    Else
        scenarioName = driverName & " " & Format(pctChange, "0%")
    End If

    ' Save baseline, apply change, build impact report
    SaveBaseline
    modPerformance.TurboOn
    modPerformance.UpdateStatus "Applying What-If: " & scenarioName, 0.3

    Dim origVal As Double: origVal = modConfig.SafeNum(wsA.Cells(targetRow, 2).Value)
    Dim newVal As Double: newVal = origVal * (1 + pctChange)
    wsA.Cells(targetRow, 2).Value = newVal

    Application.Calculate
    DoEvents

    BuildImpactReport scenarioName, Array(driverName), Array(origVal), Array(newVal)

    modPerformance.TurboOff
    modLogger.LogAction "modWhatIf", "QuickWhatIf", scenarioName & ": " & _
        Format(origVal, "#,##0.00") & " -> " & Format(newVal, "#,##0.00")
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modWhatIf", "ERROR", Err.Description
    MsgBox "Custom What-If error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' RestoreBaseline - Restore original Assumptions from saved baseline
'===============================================================================
Public Sub RestoreBaseline()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_BASELINE) Then
        MsgBox "No baseline saved. Run a What-If scenario first.", vbInformation, APP_NAME
        Exit Sub
    End If

    If MsgBox("Restore original Assumptions values?" & vbCrLf & _
              "This will undo the last What-If scenario.", _
              vbYesNo + vbQuestion, APP_NAME) = vbNo Then Exit Sub

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Restoring baseline...", 0.3

    Dim wsA As Worksheet: Set wsA = ThisWorkbook.Worksheets(SH_ASSUMPTIONS)
    Dim wsBL As Worksheet: Set wsBL = ThisWorkbook.Worksheets(SH_BASELINE)
    Dim lastRow As Long: lastRow = modConfig.LastRow(wsBL, 1)

    Dim restored As Long: restored = 0
    Dim r As Long
    For r = 1 To lastRow
        Dim blDriver As String: blDriver = Trim(CStr(wsBL.Cells(r, 1).Value))
        If blDriver <> "" Then
            ' Find matching row on Assumptions
            Dim ar As Long
            For ar = DATA_ROW_ASSUME To modConfig.LastRow(wsA, 1)
                If LCase(Trim(CStr(wsA.Cells(ar, 1).Value))) = LCase(blDriver) Then
                    wsA.Cells(ar, 2).Value = wsBL.Cells(r, 2).Value
                    restored = restored + 1
                    Exit For
                End If
            Next ar
        End If
    Next r

    ' Clean up
    modConfig.SafeDeleteSheet SH_BASELINE
    modConfig.SafeDeleteSheet SH_WHATIF

    Application.Calculate
    DoEvents

    modPerformance.TurboOff
    wsA.Activate

    modLogger.LogAction "modWhatIf", "RestoreBaseline", restored & " drivers restored"
    MsgBox "Baseline restored!" & vbCrLf & _
           restored & " driver values set back to original." & vbCrLf & vbCrLf & _
           "The model has been recalculated.", vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modWhatIf", "ERROR", Err.Description
    MsgBox "Restore error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' ApplyPresetScenario - Apply a single-category preset
'===============================================================================
Private Sub ApplyPresetScenario(ByVal scenarioName As String, _
                                 ByVal category As String, _
                                 ByVal pctChange As Double)
    On Error GoTo ErrHandler

    SaveBaseline
    modPerformance.TurboOn
    modPerformance.UpdateStatus "Applying: " & scenarioName, 0.2

    Dim wsA As Worksheet: Set wsA = ThisWorkbook.Worksheets(SH_ASSUMPTIONS)
    Dim lastRow As Long: lastRow = modConfig.LastRow(wsA, 1)

    Dim names() As String, origVals() As Double, newVals() As Double
    Dim changed As Long: changed = 0

    ' First pass: count matching drivers
    Dim matchCount As Long: matchCount = 0
    Dim r As Long
    For r = DATA_ROW_ASSUME To lastRow
        If DriverMatchesCategory(wsA.Cells(r, 1).Value, category) Then
            matchCount = matchCount + 1
        End If
    Next r

    If matchCount = 0 Then
        modPerformance.TurboOff
        MsgBox "No matching drivers found for category: " & category, vbExclamation, APP_NAME
        Exit Sub
    End If

    ReDim names(0 To matchCount - 1)
    ReDim origVals(0 To matchCount - 1)
    ReDim newVals(0 To matchCount - 1)

    ' Second pass: apply changes
    For r = DATA_ROW_ASSUME To lastRow
        Dim dName As String: dName = Trim(CStr(wsA.Cells(r, 1).Value))
        If dName <> "" And DriverMatchesCategory(dName, category) Then
            names(changed) = dName
            origVals(changed) = modConfig.SafeNum(wsA.Cells(r, 2).Value)
            newVals(changed) = origVals(changed) * (1 + pctChange)
            wsA.Cells(r, 2).Value = newVals(changed)
            changed = changed + 1
        End If
    Next r

    Application.Calculate
    DoEvents

    modPerformance.UpdateStatus "Building impact report...", 0.7
    BuildImpactReport scenarioName, names, origVals, newVals

    modPerformance.TurboOff
    modLogger.LogAction "modWhatIf", "ApplyPresetScenario", scenarioName & ": " & changed & " drivers changed"
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modWhatIf", "ERROR", Err.Description
    MsgBox "Scenario error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' ApplyComboScenario - Apply changes to two categories at once
'===============================================================================
Private Sub ApplyComboScenario(ByVal scenarioName As String, _
                                ByVal cat1 As String, ByVal pct1 As Double, _
                                ByVal cat2 As String, ByVal pct2 As Double)
    On Error GoTo ErrHandler

    SaveBaseline
    modPerformance.TurboOn
    modPerformance.UpdateStatus "Applying: " & scenarioName, 0.2

    Dim wsA As Worksheet: Set wsA = ThisWorkbook.Worksheets(SH_ASSUMPTIONS)
    Dim lastRow As Long: lastRow = modConfig.LastRow(wsA, 1)

    ' Collect all changes
    Dim allNames As String: allNames = ""
    Dim changed As Long: changed = 0
    Dim r As Long

    For r = DATA_ROW_ASSUME To lastRow
        Dim dName As String: dName = Trim(CStr(wsA.Cells(r, 1).Value))
        If dName = "" Then GoTo NextComboRow

        Dim origVal As Double: origVal = modConfig.SafeNum(wsA.Cells(r, 2).Value)

        If DriverMatchesCategory(dName, cat1) Then
            wsA.Cells(r, 2).Value = origVal * (1 + pct1)
            changed = changed + 1
            allNames = allNames & dName & " (" & Format(pct1, "+0%;-0%") & "), "
        ElseIf DriverMatchesCategory(dName, cat2) Then
            wsA.Cells(r, 2).Value = origVal * (1 + pct2)
            changed = changed + 1
            allNames = allNames & dName & " (" & Format(pct2, "+0%;-0%") & "), "
        End If
NextComboRow:
    Next r

    Application.Calculate
    DoEvents

    ' Build a simplified impact report for combo scenarios
    modConfig.SafeDeleteSheet SH_WHATIF

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = SH_WHATIF

    ' Title
    ws.Cells(1, 1).Value = "Keystone BenefitTech, Inc."
    ws.Cells(1, 1).Font.Size = 14: ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Color = CLR_NAVY

    ws.Cells(2, 1).Value = "What-If Scenario: " & scenarioName
    ws.Cells(2, 1).Font.Size = 12: ws.Cells(2, 1).Font.Bold = True

    ws.Cells(3, 1).Value = "Generated: " & Format(Now, "mmmm d, yyyy h:mm AM/PM")
    ws.Cells(3, 1).Font.Italic = True

    ws.Cells(5, 1).Value = "Drivers Changed: " & changed
    ws.Cells(5, 1).Font.Bold = True

    ws.Cells(7, 1).Value = "SCENARIO DETAILS"
    ws.Cells(7, 1).Font.Bold = True: ws.Cells(7, 1).Font.Size = 11
    ws.Cells(7, 1).Font.Color = CLR_NAVY

    ws.Cells(8, 1).Value = "Category 1: " & cat1 & " at " & Format(pct1, "+0%;-0%")
    ws.Cells(9, 1).Value = "Category 2: " & cat2 & " at " & Format(pct2, "+0%;-0%")

    ws.Cells(11, 1).Value = "Review the P&L Trend and Functional P&L sheets to see the full impact."
    ws.Cells(11, 1).Font.Italic = True

    ws.Cells(13, 1).Value = "To restore original values: Run 'Restore Baseline' from the Command Center"
    ws.Cells(13, 1).Font.Bold = True
    ws.Cells(13, 1).Font.Color = RGB(180, 0, 0)

    ws.Columns(1).ColumnWidth = 70
    ws.Activate

    modPerformance.TurboOff
    modLogger.LogAction "modWhatIf", "ApplyComboScenario", scenarioName & ": " & changed & " drivers changed"

    MsgBox "Scenario Applied: " & scenarioName & vbCrLf & vbCrLf & _
           changed & " drivers changed." & vbCrLf & _
           "The P&L has been recalculated." & vbCrLf & vbCrLf & _
           "Check the P&L Trend and Functional P&L sheets to see the impact." & vbCrLf & vbCrLf & _
           "Run 'Restore Baseline' when done to reset.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modWhatIf", "ERROR", Err.Description
    MsgBox "Combo scenario error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' DriverMatchesCategory - Check if driver name matches a scenario category
'===============================================================================
Private Function DriverMatchesCategory(ByVal driverName As Variant, _
                                        ByVal category As String) As Boolean
    Dim dn As String: dn = LCase(Trim(CStr(driverName)))
    DriverMatchesCategory = False

    Select Case LCase(category)
        Case "rev"
            DriverMatchesCategory = (InStr(dn, "revenue") > 0) Or _
                                    (InStr(dn, "rev share") > 0) Or _
                                    (InStr(dn, "sales") > 0) Or _
                                    (InStr(dn, "growth") > 0 And InStr(dn, "revenue") > 0)
        Case "aws"
            DriverMatchesCategory = (InStr(dn, "aws") > 0) Or _
                                    (InStr(dn, "cloud") > 0) Or _
                                    (InStr(dn, "compute") > 0)
        Case "head"
            DriverMatchesCategory = (InStr(dn, "headcount") > 0) Or _
                                    (InStr(dn, "salary") > 0) Or _
                                    (InStr(dn, "fte") > 0) Or _
                                    (InStr(dn, "compensation") > 0) Or _
                                    (InStr(dn, "payroll") > 0)
        Case "expense"
            ' Match anything that looks like an expense (not revenue)
            DriverMatchesCategory = (InStr(dn, "revenue") = 0) And _
                                    (InStr(dn, "rev share") = 0) And _
                                    (InStr(dn, "sales") = 0) And _
                                    (dn <> "")
        Case Else
            DriverMatchesCategory = False
    End Select
End Function

'===============================================================================
' SaveBaseline - Save current Assumptions to hidden baseline sheet
'===============================================================================
Private Sub SaveBaseline()
    ' Only save if no baseline exists yet (don't overwrite on consecutive runs)
    If modConfig.SheetExists(SH_BASELINE) Then Exit Sub

    Dim wsA As Worksheet: Set wsA = ThisWorkbook.Worksheets(SH_ASSUMPTIONS)
    Dim wsBL As Worksheet

    Set wsBL = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsBL.Name = SH_BASELINE

    Dim lastRow As Long: lastRow = modConfig.LastRow(wsA, 1)
    Dim outRow As Long: outRow = 1
    Dim r As Long

    For r = DATA_ROW_ASSUME To lastRow
        Dim dName As String: dName = Trim(CStr(wsA.Cells(r, 1).Value))
        If dName <> "" Then
            wsBL.Cells(outRow, 1).Value = dName
            wsBL.Cells(outRow, 2).Value = wsA.Cells(r, 2).Value
            outRow = outRow + 1
        End If
    Next r

    wsBL.Visible = xlSheetVeryHidden
End Sub

'===============================================================================
' BuildImpactReport - Create styled impact summary sheet
'===============================================================================
Private Sub BuildImpactReport(ByVal scenarioName As String, _
                               ByRef driverNames As Variant, _
                               ByRef origValues As Variant, _
                               ByRef newValues As Variant)
    modConfig.SafeDeleteSheet SH_WHATIF

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = SH_WHATIF

    ' Title block
    ws.Cells(1, 1).Value = "Keystone BenefitTech, Inc."
    ws.Cells(1, 1).Font.Size = 14: ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Color = CLR_NAVY

    ws.Cells(2, 1).Value = "What-If Impact Analysis: " & scenarioName
    ws.Cells(2, 1).Font.Size = 12: ws.Cells(2, 1).Font.Bold = True

    ws.Cells(3, 1).Value = "Generated: " & Format(Now, "mmmm d, yyyy h:mm AM/PM")
    ws.Cells(3, 1).Font.Italic = True

    ' Headers
    Dim hdrRow As Long: hdrRow = 5
    Dim headers As Variant
    headers = Array("Driver", "Original Value", "New Value", "Change", "Change %")
    modConfig.StyleHeader ws, hdrRow, headers

    ws.Columns(1).ColumnWidth = 35
    ws.Columns(2).ColumnWidth = 18
    ws.Columns(3).ColumnWidth = 18
    ws.Columns(4).ColumnWidth = 18
    ws.Columns(5).ColumnWidth = 12

    ' Data rows
    Dim r As Long: r = hdrRow + 1
    Dim i As Long
    For i = LBound(driverNames) To UBound(driverNames)
        ws.Cells(r, 1).Value = driverNames(i)
        ws.Cells(r, 2).Value = origValues(i)
        ws.Cells(r, 2).NumberFormat = "#,##0.00"
        ws.Cells(r, 3).Value = newValues(i)
        ws.Cells(r, 3).NumberFormat = "#,##0.00"

        Dim change As Double: change = newValues(i) - origValues(i)
        ws.Cells(r, 4).Value = change
        ws.Cells(r, 4).NumberFormat = "#,##0.00"

        If origValues(i) <> 0 Then
            ws.Cells(r, 5).Value = change / origValues(i)
        End If
        ws.Cells(r, 5).NumberFormat = "+0.0%;-0.0%"

        ' Color: green for favorable, red for unfavorable
        If change > 0 Then
            ws.Cells(r, 4).Font.Color = RGB(0, 128, 0)
            ws.Cells(r, 5).Font.Color = RGB(0, 128, 0)
        ElseIf change < 0 Then
            ws.Cells(r, 4).Font.Color = RGB(200, 0, 0)
            ws.Cells(r, 5).Font.Color = RGB(200, 0, 0)
        End If

        If r Mod 2 = 0 Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Interior.Color = CLR_ALT_ROW
        End If

        r = r + 1
    Next i

    ' Action note
    r = r + 2
    ws.Cells(r, 1).Value = "NEXT STEPS"
    ws.Cells(r, 1).Font.Bold = True: ws.Cells(r, 1).Font.Size = 11
    ws.Cells(r, 1).Font.Color = CLR_NAVY

    ws.Cells(r + 1, 1).Value = "1. Review the P&L Monthly Trend sheet to see the full financial impact"
    ws.Cells(r + 2, 1).Value = "2. Check Functional P&L Summary for department-level detail"
    ws.Cells(r + 3, 1).Value = "3. Run 'Restore Baseline' from the Command Center to reset"
    ws.Cells(r + 3, 1).Font.Bold = True
    ws.Cells(r + 3, 1).Font.Color = RGB(180, 0, 0)

    ws.Activate

    Dim changeCount As Long: changeCount = UBound(driverNames) - LBound(driverNames) + 1
    MsgBox "What-If Scenario Applied: " & scenarioName & vbCrLf & vbCrLf & _
           changeCount & " driver(s) changed." & vbCrLf & _
           "The P&L model has been recalculated." & vbCrLf & vbCrLf & _
           "Check P&L Trend and Functional P&L for full impact." & vbCrLf & vbCrLf & _
           "Run 'Restore Baseline' when done to reset values.", _
           vbInformation, APP_NAME
End Sub
