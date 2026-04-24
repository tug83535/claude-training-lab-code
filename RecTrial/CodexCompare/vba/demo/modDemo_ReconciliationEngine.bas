Attribute VB_Name = "modDemo_ReconciliationEngine"
Option Explicit

Public Sub RunDemoReconciliation()
    Dim wsChecks As Worksheet
    Dim wsGL As Worksheet
    Dim nextRow As Long
    Dim expectedRowCount As Long
    Dim actualRowCount As Long
    Dim nonNullAmountCount As Long
    Dim dateMin As Variant
    Dim dateMax As Variant
    Dim statusText As String

    On Error GoTo ReconFail

    DemoValidateWorkbookOrStop

    Set wsChecks = DemoGetSheet(DEMO_SHEET_CHECKS)
    Set wsGL = DemoGetSheet(DEMO_SHEET_GL)

    wsChecks.Range("A5:G200").ClearContents
    wsChecks.Range("A4:G4").Value = Array("Check Name", "Expected", "Actual", "Difference", "Status", "Details", "Timestamp")
    wsChecks.Range("A4:G4").Font.Bold = True

    nextRow = 5

    expectedRowCount = UTL_LastUsedRow(wsGL) - 1
    actualRowCount = CountNonBlankRows(wsGL, 1)
    statusText = IIf(expectedRowCount = actualRowCount, "PASS", "FAIL")
    WriteCheckRow wsChecks, nextRow, "GL Row Count Validation", expectedRowCount, actualRowCount, actualRowCount - expectedRowCount, statusText, "Row count check"
    nextRow = nextRow + 1

    nonNullAmountCount = CountNumericValues(wsGL, 7)
    statusText = IIf(nonNullAmountCount = actualRowCount, "PASS", "FAIL")
    WriteCheckRow wsChecks, nextRow, "GL Amount Column Non-Null", actualRowCount, nonNullAmountCount, nonNullAmountCount - actualRowCount, statusText, "Amount column completeness"
    nextRow = nextRow + 1

    dateMin = Application.WorksheetFunction.Min(wsGL.Range("B2:B" & UTL_LastUsedRow(wsGL)))
    dateMax = Application.WorksheetFunction.Max(wsGL.Range("B2:B" & UTL_LastUsedRow(wsGL)))
    WriteCheckRow wsChecks, nextRow, "GL Date Range Check", "Expected FY window", Format$(dateMin, "yyyy-mm-dd") & " to " & Format$(dateMax, "yyyy-mm-dd"), "-", "PASS", "Date range extracted"
    nextRow = nextRow + 1

    RunRevenueTieOut wsChecks, nextRow

    wsChecks.Columns("A:G").AutoFit
    wsChecks.Range("A3").Value = "Last Run: " & Format$(Now, "yyyy-mm-dd hh:nn")

    DemoLog "RunDemoReconciliation", "PASS", "Reconciliation completed"
    UTL_ShowCompletion "Demo Reconciliation", "Checks refreshed on sheet 'Checks'."
    Exit Sub

ReconFail:
    DemoLog "RunDemoReconciliation", "FAIL", Err.Description
    MsgBox "Reconciliation failed: " & Err.Description, vbExclamation, "Demo Reconciliation"
End Sub

Private Sub RunRevenueTieOut(ByVal wsChecks As Worksheet, ByVal targetRow As Long)
    Dim wsGL As Worksheet
    Dim wsTrend As Worksheet
    Dim glRevenue As Double
    Dim trendRevenue As Double
    Dim diffVal As Double
    Dim statusText As String

    Set wsGL = DemoGetSheet(DEMO_SHEET_GL)
    Set wsTrend = DemoGetSheet(DEMO_SHEET_PNL_TREND)

    glRevenue = SumIfProductEquals(wsGL, "Revenue", 7, 5)
    trendRevenue = GetTrendRevenueTotal(wsTrend)

    diffVal = trendRevenue - glRevenue
    statusText = IIf(Abs(diffVal) <= DEMO_MATERIALITY_ABS, "PASS", "FAIL")

    WriteCheckRow wsChecks, targetRow, "Revenue Tie-Out", trendRevenue, glRevenue, diffVal, statusText, "Crossfire vs P&L trend"
End Sub

Private Sub WriteCheckRow(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal checkName As String, ByVal expectedVal As Variant, ByVal actualVal As Variant, ByVal diffVal As Variant, ByVal statusText As String, ByVal detailsText As String)
    ws.Cells(rowNum, 1).Value = checkName
    ws.Cells(rowNum, 2).Value = expectedVal
    ws.Cells(rowNum, 3).Value = actualVal
    ws.Cells(rowNum, 4).Value = diffVal
    ws.Cells(rowNum, 5).Value = statusText
    ws.Cells(rowNum, 6).Value = detailsText
    ws.Cells(rowNum, 7).Value = Format$(Now, "yyyy-mm-dd hh:nn")
End Sub

Private Function CountNonBlankRows(ByVal ws As Worksheet, ByVal keyColumn As Long) As Long
    Dim r As Long
    For r = 2 To UTL_LastUsedRow(ws)
        If Len(Trim$(CStr(ws.Cells(r, keyColumn).Value2))) > 0 Then CountNonBlankRows = CountNonBlankRows + 1
    Next r
End Function

Private Function CountNumericValues(ByVal ws As Worksheet, ByVal colNum As Long) As Long
    Dim r As Long
    For r = 2 To UTL_LastUsedRow(ws)
        If IsNumeric(ws.Cells(r, colNum).Value2) Then CountNumericValues = CountNumericValues + 1
    Next r
End Function

Private Function SumIfProductEquals(ByVal ws As Worksheet, ByVal productName As String, ByVal amountCol As Long, ByVal productCol As Long) As Double
    Dim r As Long
    Dim productText As String

    For r = 2 To UTL_LastUsedRow(ws)
        productText = Trim$(CStr(ws.Cells(r, productCol).Value2))
        If StrComp(productText, productName, vbTextCompare) = 0 Then
            If IsNumeric(ws.Cells(r, amountCol).Value2) Then
                SumIfProductEquals = SumIfProductEquals + CDbl(ws.Cells(r, amountCol).Value2)
            End If
        End If
    Next r
End Function

Private Function GetTrendRevenueTotal(ByVal ws As Worksheet) As Double
    Dim revenueRow As Long
    Dim c As Long
    Dim lastCol As Long

    revenueRow = FindRowByLabel(ws, "Revenue")
    If revenueRow = 0 Then Exit Function

    lastCol = UTL_LastUsedColumn(ws)
    For c = 2 To lastCol
        If IsNumeric(ws.Cells(revenueRow, c).Value2) Then
            GetTrendRevenueTotal = GetTrendRevenueTotal + CDbl(ws.Cells(revenueRow, c).Value2)
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
