Attribute VB_Name = "modDemo_WhatIfScenario"
Option Explicit

Private Type ScenarioResult
    ScenarioName As String
    RevenueDeltaPct As Double
    CostDeltaPct As Double
    BaseRevenue As Double
    BaseCost As Double
    NewRevenue As Double
    NewCost As Double
    NewMarginPct As Double
End Type

Public Sub RunDemoWhatIfScenarios()
    Dim wsOut As Worksheet
    Dim rowPtr As Long
    Dim result As ScenarioResult

    On Error GoTo ScenarioFail

    DemoValidateWorkbookOrStop
    Set wsOut = DemoGetOrCreateScenarioSheet()
    wsOut.Cells.Clear

    wsOut.Range("A1:H1").Value = Array("Scenario", "Revenue Delta %", "Cost Delta %", "Revenue", "Cost", "Margin %", "Variance Narrative", "Timestamp")
    wsOut.Rows(1).Font.Bold = True

    rowPtr = 2

    result = EvaluateScenario("Base Case", 0, 0)
    WriteScenarioRow wsOut, rowPtr, result
    rowPtr = rowPtr + 1

    result = EvaluateScenario("Growth Push", 0.08, 0.03)
    WriteScenarioRow wsOut, rowPtr, result
    rowPtr = rowPtr + 1

    result = EvaluateScenario("Margin Protection", 0.02, -0.04)
    WriteScenarioRow wsOut, rowPtr, result
    rowPtr = rowPtr + 1

    result = EvaluateScenario("Stress Case", -0.06, 0.05)
    WriteScenarioRow wsOut, rowPtr, result

    wsOut.Columns("A:H").AutoFit
    wsOut.Columns("D:E").NumberFormat = "$#,##0;($#,##0);""-"""
    wsOut.Columns("B:C,F").NumberFormat = "0.0%"

    DemoLog "RunDemoWhatIfScenarios", "PASS", "Scenario comparison refreshed"
    UTL_ShowCompletion "Demo What-If", "Scenario output created on 'Scenario_Compare'."
    Exit Sub

ScenarioFail:
    DemoLog "RunDemoWhatIfScenarios", "FAIL", Err.Description
    MsgBox "Scenario run failed: " & Err.Description, vbExclamation, "Demo What-If"
End Sub

Private Function EvaluateScenario(ByVal scenarioName As String, ByVal revenueDeltaPct As Double, ByVal costDeltaPct As Double) As ScenarioResult
    Dim wsTrend As Worksheet
    Dim revRow As Long
    Dim costRow As Long
    Dim baseRevenue As Double
    Dim baseCost As Double
    Dim result As ScenarioResult

    Set wsTrend = DemoGetSheet(DEMO_SHEET_PNL_TREND)

    revRow = FindRowByLabel(wsTrend, "Revenue")
    costRow = FindRowByLabel(wsTrend, "Cost of Revenue")

    baseRevenue = SumRowValues(wsTrend, revRow, 2)
    baseCost = SumRowValues(wsTrend, costRow, 2)

    result.ScenarioName = scenarioName
    result.RevenueDeltaPct = revenueDeltaPct
    result.CostDeltaPct = costDeltaPct
    result.BaseRevenue = baseRevenue
    result.BaseCost = baseCost
    result.NewRevenue = baseRevenue * (1 + revenueDeltaPct)
    result.NewCost = baseCost * (1 + costDeltaPct)
    result.NewMarginPct = SafeMargin(result.NewRevenue, result.NewCost)

    EvaluateScenario = result
End Function

Private Sub WriteScenarioRow(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal result As ScenarioResult)
    ws.Cells(rowNum, 1).Value = result.ScenarioName
    ws.Cells(rowNum, 2).Value = result.RevenueDeltaPct
    ws.Cells(rowNum, 3).Value = result.CostDeltaPct
    ws.Cells(rowNum, 4).Value = result.NewRevenue
    ws.Cells(rowNum, 5).Value = result.NewCost
    ws.Cells(rowNum, 6).Value = result.NewMarginPct
    ws.Cells(rowNum, 7).Value = BuildScenarioNarrative(result)
    ws.Cells(rowNum, 8).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

Private Function BuildScenarioNarrative(ByVal result As ScenarioResult) As String
    If result.NewMarginPct >= 0.6 Then
        BuildScenarioNarrative = result.ScenarioName & " keeps margin above 60%. This supports aggressive growth planning."
    ElseIf result.NewMarginPct >= 0.5 Then
        BuildScenarioNarrative = result.ScenarioName & " keeps margin in a controllable range. Monitor cost assumptions closely."
    Else
        BuildScenarioNarrative = result.ScenarioName & " compresses margin below target. Escalate this scenario for leadership review."
    End If
End Function

Private Function SumRowValues(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal startCol As Long) As Double
    Dim c As Long
    Dim lastCol As Long

    If rowNum = 0 Then Exit Function

    lastCol = UTL_LastUsedColumn(ws)
    For c = startCol To lastCol
        If IsNumeric(ws.Cells(rowNum, c).Value2) Then
            SumRowValues = SumRowValues + CDbl(ws.Cells(rowNum, c).Value2)
        End If
    Next c
End Function

Private Function FindRowByLabel(ByVal ws As Worksheet, ByVal labelText As String) As Long
    Dim r As Long
    For r = 1 To UTL_LastUsedRow(ws)
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value2)), labelText, vbTextCompare) = 0 Then
            FindRowByLabel = r
            Exit Function
        End If
    Next r
End Function

Private Function SafeMargin(ByVal revenue As Double, ByVal cost As Double) As Double
    If revenue = 0 Then
        SafeMargin = 0
    Else
        SafeMargin = (revenue - cost) / revenue
    End If
End Function

Private Function DemoGetOrCreateScenarioSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Scenario_Compare")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "Scenario_Compare"
    End If

    Set DemoGetOrCreateScenarioSheet = ws
End Function
