Attribute VB_Name = "modDemo_ExecBriefPack"
Option Explicit

Public Sub BuildDemoExecutiveBriefPack(Optional ByVal OutputFolder As String = "")
    Dim wsBrief As Worksheet
    Dim wsReport As Worksheet
    Dim wsChecks As Worksheet
    Dim wsTrend As Worksheet
    Dim pdfPath As String

    On Error GoTo BriefFail

    DemoValidateWorkbookOrStop

    Set wsReport = DemoGetSheet(DEMO_SHEET_REPORT)
    Set wsChecks = DemoGetSheet(DEMO_SHEET_CHECKS)
    Set wsTrend = DemoGetSheet(DEMO_SHEET_PNL_TREND)
    Set wsBrief = DemoGetOrCreateExecBriefSheet()

    wsBrief.Cells.Clear
    BuildBriefHeader wsBrief
    BuildKpiSection wsBrief, wsReport
    BuildCheckSection wsBrief, wsChecks
    BuildTrendSection wsBrief, wsTrend

    wsBrief.Columns("A:F").AutoFit

    If Len(OutputFolder) = 0 Then
        OutputFolder = ThisWorkbook.Path
    End If

    If Len(OutputFolder) > 0 Then
        pdfPath = OutputFolder & Application.PathSeparator & "Demo_Exec_Brief_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf"
        wsBrief.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False
    End If

    DemoLog "BuildDemoExecutiveBriefPack", "PASS", "Executive brief pack refreshed"
    UTL_ShowCompletion "Executive Brief Pack", "Executive brief created on 'Exec_Brief'" & IIf(Len(pdfPath) > 0, " and exported to PDF.", ".")
    Exit Sub

BriefFail:
    DemoLog "BuildDemoExecutiveBriefPack", "FAIL", Err.Description
    MsgBox "Executive brief build failed: " & Err.Description, vbExclamation, "Executive Brief Pack"
End Sub

Private Function DemoGetOrCreateExecBriefSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Exec_Brief")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "Exec_Brief"
    End If

    Set DemoGetOrCreateExecBriefSheet = ws
End Function

Private Sub BuildBriefHeader(ByVal ws As Worksheet)
    ws.Range("A1:F1").Merge
    ws.Range("A1").Value = "iPipeline"
    ws.Range("A1").Font.Name = "Arial"
    ws.Range("A1").Font.Size = 20
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = RGB(11, 71, 121)

    ws.Range("A2:F2").Merge
    ws.Range("A2").Value = "Finance & Accounting"
    ws.Range("A2").Font.Name = "Arial"
    ws.Range("A2").Font.Size = 10
    ws.Range("A2").Font.Color = RGB(17, 46, 81)

    ws.Range("A3:F3").Merge
    ws.Range("A3").Value = "Executive Brief"
    ws.Range("A3").Font.Name = "Arial"
    ws.Range("A3").Font.Size = 14
    ws.Range("A3").Font.Bold = True
    ws.Range("A3").Interior.Color = RGB(11, 71, 121)
    ws.Range("A3").Font.Color = RGB(249, 249, 249)

    ws.Range("A4").Value = "Run Timestamp"
    ws.Range("B4").Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

Private Sub BuildKpiSection(ByVal wsBrief As Worksheet, ByVal wsReport As Worksheet)
    wsBrief.Range("A6:B6").Value = Array("KPI", "Value")
    wsBrief.Range("A6:B6").Font.Bold = True

    wsBrief.Cells(7, 1).Value = "Full Year Revenue"
    wsBrief.Cells(7, 2).Value = wsReport.Range("B8").Value2

    wsBrief.Cells(8, 1).Value = "Contribution Margin %"
    wsBrief.Cells(8, 2).Value = wsReport.Range("D8").Value2

    wsBrief.Cells(9, 1).Value = "Top Product"
    wsBrief.Cells(9, 2).Value = wsReport.Range("F8").Value2

    wsBrief.Range("B7").NumberFormat = "$#,##0;($#,##0);""-"""
    wsBrief.Range("B8").NumberFormat = "0.0%"
End Sub

Private Sub BuildCheckSection(ByVal wsBrief As Worksheet, ByVal wsChecks As Worksheet)
    Dim rowPtr As Long
    Dim lastRow As Long
    Dim passCount As Long
    Dim failCount As Long

    rowPtr = 12
    wsBrief.Cells(rowPtr, 1).Value = "Control Summary"
    wsBrief.Cells(rowPtr, 1).Font.Bold = True

    lastRow = UTL_LastUsedRow(wsChecks)
    passCount = Application.WorksheetFunction.CountIf(wsChecks.Range("E5:E" & lastRow), "PASS")
    failCount = Application.WorksheetFunction.CountIf(wsChecks.Range("E5:E" & lastRow), "FAIL")

    wsBrief.Cells(rowPtr + 1, 1).Value = "PASS Checks"
    wsBrief.Cells(rowPtr + 1, 2).Value = passCount
    wsBrief.Cells(rowPtr + 2, 1).Value = "FAIL Checks"
    wsBrief.Cells(rowPtr + 2, 2).Value = failCount

    wsBrief.Cells(rowPtr + 4, 1).Value = "Action Note"
    wsBrief.Cells(rowPtr + 4, 2).Value = IIf(failCount > 0, "Review failed checks before final package release.", "No failed checks. Package ready for review.")
End Sub

Private Sub BuildTrendSection(ByVal wsBrief As Worksheet, ByVal wsTrend As Worksheet)
    Dim revenueRow As Long
    Dim startVal As Double
    Dim endVal As Double
    Dim pctVal As Double
    Dim lastCol As Long

    revenueRow = FindRowByLabel(wsTrend, "Revenue")
    If revenueRow = 0 Then Exit Sub

    lastCol = UTL_LastUsedColumn(wsTrend)
    startVal = CDbl(wsTrend.Cells(revenueRow, 2).Value2)
    endVal = CDbl(wsTrend.Cells(revenueRow, lastCol).Value2)
    pctVal = DemoSafePct(endVal - startVal, startVal)

    wsBrief.Range("D6:E6").Value = Array("Trend Metric", "Value")
    wsBrief.Range("D6:E6").Font.Bold = True

    wsBrief.Cells(7, 4).Value = "Revenue Start"
    wsBrief.Cells(7, 5).Value = startVal
    wsBrief.Cells(8, 4).Value = "Revenue Latest"
    wsBrief.Cells(8, 5).Value = endVal
    wsBrief.Cells(9, 4).Value = "Revenue Change %"
    wsBrief.Cells(9, 5).Value = pctVal

    wsBrief.Range("E7:E8").NumberFormat = "$#,##0;($#,##0);""-"""
    wsBrief.Range("E9").NumberFormat = "0.0%"
End Sub

Private Function FindRowByLabel(ByVal ws As Worksheet, ByVal labelText As String) As Long
    Dim r As Long
    For r = 1 To UTL_LastUsedRow(ws)
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value2)), labelText, vbTextCompare) = 0 Then
            FindRowByLabel = r
            Exit Function
        End If
    Next r
End Function

Private Function DemoSafePct(ByVal deltaVal As Double, ByVal baseVal As Double) As Double
    If baseVal = 0 Then
        If deltaVal = 0 Then
            DemoSafePct = 0
        Else
            DemoSafePct = 1
        End If
    Else
        DemoSafePct = deltaVal / baseVal
    End If
End Function
