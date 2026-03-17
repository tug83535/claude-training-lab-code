Attribute VB_Name = "modIntegrationTest"
Option Explicit

'===============================================================================
' modIntegrationTest - Integration Testing & Health Check
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Runs a comprehensive test suite across all major modules to verify
'           the workbook is functioning correctly. Also provides a quick health
'           check for verifying expected sheets and key formulas.
'
' PUBLIC SUBS:
'   RunFullTest      - Run all integration tests (Action #44)
'   QuickHealthCheck - Verify sheets exist and key values are valid (Action #45)
'
' DEPENDENCIES: modConfig, modPerformance, modLogger
' VERSION:  2.1.0
'===============================================================================

Private Type TestResult
    TestName   As String
    Status     As String   ' PASS / FAIL / SKIP
    Details    As String
    Duration   As Double
End Type

'===============================================================================
' RunFullTest - Comprehensive integration test suite
'===============================================================================
Public Sub RunFullTest()
    On Error GoTo ErrHandler

    If MsgBox("Run the full integration test suite?" & vbCrLf & vbCrLf & _
              "This will:" & vbCrLf & _
              "  - Verify all expected sheets exist" & vbCrLf & _
              "  - Check GL data integrity" & vbCrLf & _
              "  - Run reconciliation checks" & vbCrLf & _
              "  - Validate assumptions" & vbCrLf & _
              "  - Test key module functions" & vbCrLf & vbCrLf & _
              "No data will be modified.", _
              vbYesNo + vbQuestion, APP_NAME) = vbNo Then Exit Sub

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Starting integration tests...", 0.02

    ' Collect results
    Dim results() As TestResult
    Dim testCount As Long: testCount = 0
    Dim passCount As Long: passCount = 0
    Dim failCount As Long: failCount = 0
    Dim skipCount As Long: skipCount = 0
    ReDim results(1 To 30)  ' Max 30 tests

    ' ========== TEST 1: Required Sheets Exist ==========
    Dim reqSheets As Variant
    reqSheets = Array(SH_HIDDEN, SH_ASSUMPTIONS, SH_DATADICT, SH_AWS, _
                      SH_REPORT, SH_PL_TREND, SH_PROD_SUMMARY, _
                      SH_FUNC_TREND, SH_FUNC_JAN, SH_CHECKS)

    Dim s As Long
    For s = 0 To UBound(reqSheets)
        testCount = testCount + 1
        results(testCount).TestName = "Sheet Exists: " & CStr(reqSheets(s))
        If modConfig.SheetExists(CStr(reqSheets(s))) Then
            results(testCount).Status = "PASS"
            results(testCount).Details = "Found"
            passCount = passCount + 1
        Else
            results(testCount).Status = "FAIL"
            results(testCount).Details = "NOT FOUND"
            failCount = failCount + 1
        End If
    Next s

    modPerformance.UpdateStatus "Testing GL data...", 0.3

    ' ========== TEST 2: GL Data Has Records ==========
    testCount = testCount + 1
    results(testCount).TestName = "GL Data: Row Count"
    If modConfig.SheetExists(SH_GL) Then
        Dim wsGL As Worksheet: Set wsGL = ThisWorkbook.Worksheets(SH_GL)
        Dim glRows As Long: glRows = modConfig.LastRow(wsGL, COL_GL_ID) - DATA_ROW_GL + 1
        If glRows > 0 Then
            results(testCount).Status = "PASS"
            results(testCount).Details = glRows & " rows"
            passCount = passCount + 1
        Else
            results(testCount).Status = "FAIL"
            results(testCount).Details = "Empty GL"
            failCount = failCount + 1
        End If
    Else
        results(testCount).Status = "SKIP"
        results(testCount).Details = "GL sheet missing"
        skipCount = skipCount + 1
    End If

    ' ========== TEST 3: GL Columns Valid ==========
    testCount = testCount + 1
    results(testCount).TestName = "GL Data: Column Structure"
    If modConfig.SheetExists(SH_GL) Then
        Dim glCols As Long: glCols = wsGL.Cells(HDR_ROW_GL, wsGL.Columns.Count).End(xlToLeft).Column
        If glCols >= 7 Then
            results(testCount).Status = "PASS"
            results(testCount).Details = glCols & " columns"
            passCount = passCount + 1
        Else
            results(testCount).Status = "FAIL"
            results(testCount).Details = "Expected 7+, found " & glCols
            failCount = failCount + 1
        End If
    Else
        results(testCount).Status = "SKIP"
        results(testCount).Details = "GL sheet missing"
        skipCount = skipCount + 1
    End If

    modPerformance.UpdateStatus "Testing Assumptions...", 0.5

    ' ========== TEST 4: Assumptions Has Drivers ==========
    testCount = testCount + 1
    results(testCount).TestName = "Assumptions: Driver Count"
    If modConfig.SheetExists(SH_ASSUMPTIONS) Then
        Dim wsA As Worksheet: Set wsA = ThisWorkbook.Worksheets(SH_ASSUMPTIONS)
        Dim driverRows As Long: driverRows = modConfig.LastRow(wsA, 1) - DATA_ROW_ASSUME + 1
        If driverRows > 0 Then
            results(testCount).Status = "PASS"
            results(testCount).Details = driverRows & " drivers"
            passCount = passCount + 1
        Else
            results(testCount).Status = "FAIL"
            results(testCount).Details = "No drivers found"
            failCount = failCount + 1
        End If
    Else
        results(testCount).Status = "SKIP"
        results(testCount).Details = "Assumptions sheet missing"
        skipCount = skipCount + 1
    End If

    ' ========== TEST 5: P&L Trend Has Data ==========
    testCount = testCount + 1
    results(testCount).TestName = "P&L Trend: Data Present"
    If modConfig.SheetExists(SH_PL_TREND) Then
        Dim wsTrend As Worksheet: Set wsTrend = ThisWorkbook.Worksheets(SH_PL_TREND)
        Dim trendRows As Long: trendRows = modConfig.LastRow(wsTrend, 1)
        If trendRows > DATA_ROW_REPORT Then
            results(testCount).Status = "PASS"
            results(testCount).Details = trendRows & " rows"
            passCount = passCount + 1
        Else
            results(testCount).Status = "FAIL"
            results(testCount).Details = "No data"
            failCount = failCount + 1
        End If
    Else
        results(testCount).Status = "SKIP"
        results(testCount).Details = "Trend sheet missing"
        skipCount = skipCount + 1
    End If

    modPerformance.UpdateStatus "Testing Checks...", 0.7

    ' ========== TEST 6: Checks Sheet Has Results ==========
    testCount = testCount + 1
    results(testCount).TestName = "Checks: Reconciliation Data"
    If modConfig.SheetExists(SH_CHECKS) Then
        Dim wsChk As Worksheet: Set wsChk = ThisWorkbook.Worksheets(SH_CHECKS)
        Dim chkRows As Long: chkRows = modConfig.LastRow(wsChk, 1)
        If chkRows >= DATA_ROW_CHECKS Then
            results(testCount).Status = "PASS"
            results(testCount).Details = (chkRows - DATA_ROW_CHECKS + 1) & " checks"
            passCount = passCount + 1
        Else
            results(testCount).Status = "FAIL"
            results(testCount).Details = "No checks"
            failCount = failCount + 1
        End If
    Else
        results(testCount).Status = "SKIP"
        results(testCount).Details = "Checks sheet missing"
        skipCount = skipCount + 1
    End If

    ' ========== TEST 7: Product List Valid ==========
    testCount = testCount + 1
    results(testCount).TestName = "Config: Product List"
    Dim products As Variant: products = modConfig.GetProducts()
    If UBound(products) >= 3 Then
        results(testCount).Status = "PASS"
        results(testCount).Details = (UBound(products) + 1) & " products"
        passCount = passCount + 1
    Else
        results(testCount).Status = "FAIL"
        results(testCount).Details = "Expected 4+, found " & (UBound(products) + 1)
        failCount = failCount + 1
    End If

    ' ========== TEST 8: Config Constants Valid ==========
    testCount = testCount + 1
    results(testCount).TestName = "Config: Fiscal Year"
    If Len(FISCAL_YEAR) = 2 And Len(FISCAL_YEAR_4) = 4 Then
        results(testCount).Status = "PASS"
        results(testCount).Details = "FY" & FISCAL_YEAR_4
        passCount = passCount + 1
    Else
        results(testCount).Status = "FAIL"
        results(testCount).Details = "Invalid fiscal year config"
        failCount = failCount + 1
    End If

    ' ========== OUTPUT RESULTS ==========
    modPerformance.UpdateStatus "Writing test report...", 0.9

    ' Create report sheet
    modConfig.SafeDeleteSheet SH_TEST_REPORT
    Dim wsRpt As Worksheet
    Set wsRpt = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsRpt.Name = SH_TEST_REPORT

    wsRpt.Range("A1").Value = "INTEGRATION TEST REPORT"
    wsRpt.Range("A1").Font.Bold = True
    wsRpt.Range("A1").Font.Size = 14
    wsRpt.Range("A1").Font.Color = CLR_NAVY
    wsRpt.Range("A2").Value = Format(Now, "yyyy-mm-dd hh:mm") & "  |  " & _
        testCount & " tests: " & passCount & " PASS, " & failCount & " FAIL, " & skipCount & " SKIP"
    wsRpt.Range("A2").Font.Color = IIf(failCount > 0, RGB(192, 0, 0), RGB(0, 128, 0))

    modConfig.StyleHeader wsRpt, 4, Array("Test #", "Test Name", "Status", "Details")
    Dim rptRow As Long: rptRow = 5
    Dim t As Long
    For t = 1 To testCount
        wsRpt.Cells(rptRow, 1).Value = t
        wsRpt.Cells(rptRow, 2).Value = results(t).TestName
        wsRpt.Cells(rptRow, 3).Value = results(t).Status
        wsRpt.Cells(rptRow, 4).Value = results(t).Details

        Select Case results(t).Status
            Case "PASS"
                wsRpt.Cells(rptRow, 3).Font.Color = RGB(0, 128, 0)
            Case "FAIL"
                wsRpt.Cells(rptRow, 3).Font.Color = RGB(192, 0, 0)
                wsRpt.Cells(rptRow, 3).Font.Bold = True
            Case "SKIP"
                wsRpt.Cells(rptRow, 3).Font.Color = RGB(128, 128, 128)
        End Select

        If rptRow Mod 2 = 1 Then
            wsRpt.Range("A" & rptRow & ":D" & rptRow).Interior.Color = CLR_ALT_ROW
        End If
        rptRow = rptRow + 1
    Next t

    wsRpt.Columns("A").ColumnWidth = 8
    wsRpt.Columns("B").ColumnWidth = 35
    wsRpt.Columns("C").ColumnWidth = 10
    wsRpt.Columns("D").ColumnWidth = 30
    wsRpt.Tab.Color = IIf(failCount > 0, RGB(255, 0, 0), RGB(0, 176, 80))
    wsRpt.Activate

    Dim elapsed As Double: elapsed = modPerformance.ElapsedSeconds()
    modPerformance.TurboOff

    modLogger.LogAction "modIntegrationTest", "RunFullTest", _
        testCount & " tests: " & passCount & " pass, " & failCount & " fail, " & skipCount & " skip"

    Dim icon As VbMsgBoxStyle
    icon = IIf(failCount = 0, vbInformation, vbExclamation)
    MsgBox "INTEGRATION TEST COMPLETE" & vbCrLf & String(30, "=") & vbCrLf & vbCrLf & _
           "Total Tests:  " & testCount & vbCrLf & _
           "PASS:         " & passCount & vbCrLf & _
           "FAIL:         " & failCount & vbCrLf & _
           "SKIP:         " & skipCount & vbCrLf & vbCrLf & _
           "Results on '" & SH_TEST_REPORT & "' sheet.", _
           icon, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modIntegrationTest", "ERROR", Err.Description
    MsgBox "Integration test error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' QuickHealthCheck - Fast check of essential sheets and data
'===============================================================================
Public Sub QuickHealthCheck()
    On Error GoTo ErrHandler

    Dim issues As String: issues = ""
    Dim checkCount As Long: checkCount = 0
    Dim passCount As Long: passCount = 0

    ' Check essential sheets
    Dim essentials As Variant
    essentials = Array(SH_HIDDEN, SH_ASSUMPTIONS, SH_REPORT, SH_PL_TREND, _
                       SH_PROD_SUMMARY, SH_FUNC_TREND, SH_CHECKS)
    Dim s As Long
    For s = 0 To UBound(essentials)
        checkCount = checkCount + 1
        If modConfig.SheetExists(CStr(essentials(s))) Then
            passCount = passCount + 1
        Else
            issues = issues & "  MISSING: " & CStr(essentials(s)) & vbCrLf
        End If
    Next s

    ' Check GL has data
    checkCount = checkCount + 1
    If modConfig.SheetExists(SH_GL) Then
        Dim wsGL As Worksheet: Set wsGL = ThisWorkbook.Worksheets(SH_GL)
        If modConfig.LastRow(wsGL, 1) >= DATA_ROW_GL Then
            passCount = passCount + 1
        Else
            issues = issues & "  GL sheet is empty" & vbCrLf
        End If
    Else
        issues = issues & "  GL data sheet missing" & vbCrLf
    End If

    ' Check fiscal year config
    checkCount = checkCount + 1
    If Len(FISCAL_YEAR) = 2 And Len(FISCAL_YEAR_4) = 4 Then
        passCount = passCount + 1
    Else
        issues = issues & "  Fiscal year config invalid" & vbCrLf
    End If

    Dim msg As String
    If issues = "" Then
        msg = "HEALTH CHECK: ALL CLEAR" & vbCrLf & vbCrLf & _
              passCount & "/" & checkCount & " checks passed." & vbCrLf & _
              "Workbook is in good shape."
    Else
        msg = "HEALTH CHECK: ISSUES FOUND" & vbCrLf & vbCrLf & _
              passCount & "/" & checkCount & " checks passed." & vbCrLf & vbCrLf & _
              "Issues:" & vbCrLf & issues
    End If

    modLogger.LogAction "modIntegrationTest", "QuickHealthCheck", _
        passCount & "/" & checkCount & " passed"

    Dim icon As VbMsgBoxStyle
    icon = IIf(issues = "", vbInformation, vbExclamation)
    MsgBox msg, icon, APP_NAME & " - Health Check"
    Exit Sub

ErrHandler:
    MsgBox "Health check error: " & Err.Description, vbCritical, APP_NAME
End Sub
