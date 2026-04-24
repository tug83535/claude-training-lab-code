Attribute VB_Name = "modDemo_VarianceNarrative"
Option Explicit

Public Sub GenerateDemoVarianceNarrative()
    Dim wsTrend As Worksheet
    Dim wsOut As Worksheet
    Dim headerRow As Long
    Dim firstMonthCol As Long
    Dim lastMonthCol As Long
    Dim outRow As Long
    Dim r As Long
    Dim lineName As String
    Dim firstVal As Double
    Dim lastVal As Double
    Dim deltaVal As Double
    Dim pctVal As Double
    Dim narrative As String

    On Error GoTo NarrativeFail

    DemoValidateWorkbookOrStop

    Set wsTrend = DemoGetSheet(DEMO_SHEET_PNL_TREND)
    Set wsOut = DemoGetOrCreateVarianceSheet()
    wsOut.Cells.Clear

    headerRow = FindHeaderRowForTrend(wsTrend)
    firstMonthCol = 2
    lastMonthCol = UTL_LastUsedColumn(wsTrend)

    wsOut.Range("A1:G1").Value = Array("Line Item", "First Period", "Latest Period", "Delta", "Delta %", "Status", "Narrative")
    wsOut.Rows(1).Font.Bold = True

    outRow = 2
    For r = headerRow + 1 To UTL_LastUsedRow(wsTrend)
        lineName = Trim$(CStr(wsTrend.Cells(r, 1).Value2))
        If Len(lineName) > 0 And IsNumeric(wsTrend.Cells(r, firstMonthCol).Value2) And IsNumeric(wsTrend.Cells(r, lastMonthCol).Value2) Then
            firstVal = CDbl(wsTrend.Cells(r, firstMonthCol).Value2)
            lastVal = CDbl(wsTrend.Cells(r, lastMonthCol).Value2)
            deltaVal = lastVal - firstVal
            pctVal = SafePct(deltaVal, firstVal)

            narrative = BuildVarianceNarrative(lineName, firstVal, lastVal, deltaVal, pctVal)

            wsOut.Cells(outRow, 1).Value = lineName
            wsOut.Cells(outRow, 2).Value = firstVal
            wsOut.Cells(outRow, 3).Value = lastVal
            wsOut.Cells(outRow, 4).Value = deltaVal
            wsOut.Cells(outRow, 5).Value = pctVal
            wsOut.Cells(outRow, 6).Value = LabelVariance(deltaVal, pctVal)
            wsOut.Cells(outRow, 7).Value = narrative
            outRow = outRow + 1
        End If
    Next r

    wsOut.Columns("B:D").NumberFormat = "$#,##0;($#,##0);""-"""
    wsOut.Columns("E:E").NumberFormat = "0.0%"
    wsOut.Columns("A:G").AutoFit

    DemoLog "GenerateDemoVarianceNarrative", "PASS", "Variance narrative refreshed"
    UTL_ShowCompletion "Variance Narrative", "Narrative output created on 'Exec_Variance_Narrative'."
    Exit Sub

NarrativeFail:
    DemoLog "GenerateDemoVarianceNarrative", "FAIL", Err.Description
    MsgBox "Variance narrative failed: " & Err.Description, vbExclamation, "Variance Narrative"
End Sub

Private Function DemoGetOrCreateVarianceSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Exec_Variance_Narrative")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "Exec_Variance_Narrative"
    End If

    Set DemoGetOrCreateVarianceSheet = ws
End Function

Private Function FindHeaderRowForTrend(ByVal ws As Worksheet) As Long
    FindHeaderRowForTrend = UTL_DetectHeaderRow(ws)
End Function

Private Function SafePct(ByVal deltaVal As Double, ByVal baseVal As Double) As Double
    If baseVal = 0 Then
        If deltaVal = 0 Then
            SafePct = 0
        Else
            SafePct = 1
        End If
    Else
        SafePct = deltaVal / baseVal
    End If
End Function

Private Function LabelVariance(ByVal deltaVal As Double, ByVal pctVal As Double) As String
    If Abs(deltaVal) >= DEMO_MATERIALITY_ABS And Abs(pctVal) >= DEMO_MATERIALITY_PCT Then
        If deltaVal >= 0 Then
            LabelVariance = "Material increase"
        Else
            LabelVariance = "Material decrease"
        End If
    ElseIf Abs(deltaVal) >= DEMO_MATERIALITY_ABS Or Abs(pctVal) >= DEMO_MATERIALITY_PCT Then
        LabelVariance = "Watch"
    Else
        LabelVariance = "Normal"
    End If
End Function

Private Function BuildVarianceNarrative(ByVal lineName As String, ByVal firstVal As Double, ByVal lastVal As Double, ByVal deltaVal As Double, ByVal pctVal As Double) As String
    Dim label As String

    label = LabelVariance(deltaVal, pctVal)

    Select Case label
        Case "Material increase"
            BuildVarianceNarrative = lineName & " moved up from " & Format$(firstVal, "$#,##0") & " to " & Format$(lastVal, "$#,##0") & ", a material increase of " & Format$(deltaVal, "$#,##0") & " (" & Format$(pctVal, "0.0%") & ")."
        Case "Material decrease"
            BuildVarianceNarrative = lineName & " moved down from " & Format$(firstVal, "$#,##0") & " to " & Format$(lastVal, "$#,##0") & ", a material decrease of " & Format$(deltaVal, "$#,##0") & " (" & Format$(pctVal, "0.0%") & ")."
        Case "Watch"
            BuildVarianceNarrative = lineName & " changed from " & Format$(firstVal, "$#,##0") & " to " & Format$(lastVal, "$#,##0") & ". Review this movement before finalizing close commentary."
        Case Else
            BuildVarianceNarrative = lineName & " remained within normal movement thresholds over the selected period."
    End Select
End Function
