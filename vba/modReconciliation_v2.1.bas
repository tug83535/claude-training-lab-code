Attribute VB_Name = "modReconciliation"
Option Explicit

'===============================================================================
' modReconciliation - Automated Reconciliation & Validation Runner
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Read the Checks sheet, force-recalc all formulas, evaluate every
'           PASS/FAIL status, generate a summary report, and log results.
'           Also computes cross-sheet validations from raw data.
'
' PUBLIC SUBS:
'   RunAllChecks        - Evaluate existing Checks sheet PASS/FAIL formulas
'   ExportCheckResults  - Write results to timestamped text file on Desktop
'   ValidateCrossSheet  - Compute 4 cross-sheet validations from raw data
'
' VERSION:  2.1.0
' CHANGES:  v2.0 -> v2.1:
'           + ISSUE-011: Added ValidateCrossSheet (4 computed validations)
'           + Added WriteValidationRow helper (DRY pattern for output rows)
'           + Added FindFYCol helper (FY Total column detection)
'
' PRIOR FIXES (v2.0):
'   BUG-006  Loop starts at DATA_ROW_CHECKS (5) not row 2
'            Uses COL_CHECK_STATUS constant instead of hardcoded 5
'===============================================================================

Private Type CheckResult
    CheckName   As String
    SheetAVal   As Double
    SheetBVal   As Double
    Difference  As Double
    Status      As String
    RowNum      As Long
End Type

'===============================================================================
' RunAllChecks - Main entry point (reads existing Checks sheet formulas)
'===============================================================================
Public Sub RunAllChecks()
    On Error GoTo ErrHandler
    
    If Not modConfig.SheetExists(SH_CHECKS) Then
        MsgBox "Checks sheet not found.", vbCritical, APP_NAME
        Exit Sub
    End If
    
    modPerformance.TurboOn
    modPerformance.UpdateStatus "Running reconciliation checks...", 0
    
    ' Force full recalculation first
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    DoEvents
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_CHECKS)
    Dim lastRow As Long: lastRow = modConfig.LastRow(ws, 1)
    
    If lastRow < DATA_ROW_CHECKS Then
        modPerformance.TurboOff
        MsgBox "No checks found on the Checks sheet.", vbInformation, APP_NAME
        Exit Sub
    End If
    
    ' Read all check results
    Dim results() As CheckResult
    Dim passCount As Long, failCount As Long, totalChecks As Long
    passCount = 0: failCount = 0: totalChecks = 0
    
    Dim r As Long
    Dim detailsLog As String: detailsLog = ""
    
    For r = DATA_ROW_CHECKS To lastRow
        ' Skip blank rows or section headers
        If Trim(CStr(ws.Cells(r, 1).Value)) <> "" And _
           Trim(CStr(ws.Cells(r, COL_CHECK_STATUS).Value)) <> "" Then
            
            totalChecks = totalChecks + 1
            Dim status As String: status = UCase(Trim(CStr(ws.Cells(r, COL_CHECK_STATUS).Value)))
            
            If status = "PASS" Then
                passCount = passCount + 1
                ' Ensure green formatting
                ws.Cells(r, COL_CHECK_STATUS).Interior.Color = RGB(198, 239, 206)
                ws.Cells(r, COL_CHECK_STATUS).Font.Color = RGB(0, 97, 0)
            ElseIf status = "FAIL" Then
                failCount = failCount + 1
                ' Ensure red formatting
                ws.Cells(r, COL_CHECK_STATUS).Interior.Color = RGB(255, 199, 206)
                ws.Cells(r, COL_CHECK_STATUS).Font.Color = RGB(156, 0, 6)
                
                ' Capture failure details
                detailsLog = detailsLog & "  FAIL: " & ws.Cells(r, 1).Value
                If IsNumeric(ws.Cells(r, 4).Value) Then
                    detailsLog = detailsLog & " (Diff: $" & Format(ws.Cells(r, 4).Value, "#,##0.00") & ")"
                End If
                detailsLog = detailsLog & vbCrLf
            End If
            
            modPerformance.UpdateStatus "Evaluating check " & totalChecks & "...", _
                                        totalChecks / (lastRow - DATA_ROW_CHECKS + 1)
        End If
    Next r
    
    modPerformance.TurboOff
    
    ' Build summary message
    Dim summary As String
    If totalChecks = 0 Then
        summary = "No valid checks found on the Checks sheet."
    Else
        summary = "RECONCILIATION RESULTS" & vbCrLf & _
                  String(35, "=") & vbCrLf & vbCrLf & _
                  "Total Checks:  " & totalChecks & vbCrLf & _
                  "Passed:        " & passCount & "  (" & Format(passCount / totalChecks, "0%") & ")" & vbCrLf & _
                  "Failed:        " & failCount & "  (" & Format(failCount / totalChecks, "0%") & ")" & vbCrLf & _
                  "Run Time:      " & Format(Now, "yyyy-mm-dd hh:mm:ss") & vbCrLf
        
        If failCount > 0 Then
            summary = summary & vbCrLf & "FAILURES:" & vbCrLf & detailsLog
        End If
    End If
    
    ' Log to audit trail
    modLogger.LogAction "modReconciliation", "RunAllChecks", _
                        passCount & "/" & totalChecks & " passed. " & _
                        IIf(failCount > 0, failCount & " failures.", "All clear.") & _
                        " (" & Format(modPerformance.ElapsedSeconds(), "0.00") & "s)"
    
    ' Navigate to Checks sheet and show results
    ws.Activate
    
    Dim icon As VbMsgBoxStyle
    icon = IIf(failCount = 0, vbInformation, vbExclamation)
    MsgBox summary, icon, APP_NAME & " - Reconciliation"
    Exit Sub
    
ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modReconciliation", "ERROR", Err.Description
    MsgBox "Reconciliation error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' ExportCheckResults - Write results to timestamped text file
'===============================================================================
Public Sub ExportCheckResults()
    On Error GoTo ErrHandler
    
    If Not modConfig.SheetExists(SH_CHECKS) Then
        MsgBox "Checks sheet not found.", vbCritical, APP_NAME
        Exit Sub
    End If
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_CHECKS)
    Dim lastRow As Long: lastRow = modConfig.LastRow(ws, 1)
    
    Dim fPath As String
    fPath = Environ("USERPROFILE") & "\Desktop\KBT_ReconciliationReport_" & _
            Format(Now, "yyyymmdd_hhmmss") & ".txt"
    
    Dim fNum As Integer: fNum = FreeFile
    Open fPath For Output As #fNum
    
    Print #fNum, "KEYSTONE BENEFITTECH, INC."
    Print #fNum, "Reconciliation Report"
    Print #fNum, "Generated: " & Format(Now, "mmmm d, yyyy h:mm AM/PM")
    Print #fNum, "User: " & Environ("USERNAME")
    Print #fNum, String(60, "=")
    Print #fNum, ""
    
    Dim r As Long
    ' Print header row
    Print #fNum, Format(ws.Cells(HDR_ROW_CHECKS, 1).Value, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & " | " & _
                 Format(ws.Cells(HDR_ROW_CHECKS, COL_CHECK_STATUS).Value, "!@@@@@@")
    
    ' Print data rows starting at DATA_ROW_CHECKS
    For r = DATA_ROW_CHECKS To lastRow
        If Trim(CStr(ws.Cells(r, 1).Value)) <> "" Then
            Print #fNum, Format(ws.Cells(r, 1).Value, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & " | " & _
                         Format(ws.Cells(r, COL_CHECK_STATUS).Value, "!@@@@@@") & " | Diff: " & _
                         Format(ws.Cells(r, 4).Value, "$#,##0.00")
        End If
    Next r
    
    Close #fNum
    
    modLogger.LogAction "modReconciliation", "ExportCheckResults", "Exported to: " & fPath
    MsgBox "Report exported to:" & vbCrLf & fPath, vbInformation, APP_NAME
    Exit Sub
    
ErrHandler:
    If fNum > 0 Then Close #fNum
    MsgBox "Export error: " & Err.Description, vbCritical, APP_NAME
End Sub


'===============================================================================
'
' ===  CROSS-SHEET VALIDATION (v2.1 — ISSUE-011)  =============================
'
'===============================================================================

'===============================================================================
' ValidateCrossSheet - Compute validations from raw data across all sheets
' Unlike RunAllChecks (which reads existing Checks sheet formulas), this
' COMPUTES validations by summing raw data and comparing across sheets:
'   Check 1: GL Total Amount vs P&L Trend Total Revenue (FY Total)
'   Check 2: GL by Department vs Functional P&L totals
'   Check 3: GL by Product vs Product Line Summary totals
'   Check 4: Mirror existing Checks sheet PASS/FAIL results
' Ported from legacy T2 #15 ValidateCrossSheetData.
'===============================================================================
Public Sub ValidateCrossSheet()
    On Error GoTo ErrHandler
    
    modPerformance.TurboOn
    modPerformance.UpdateStatus "Running cross-sheet validation...", 0.05
    
    ' Create output sheet
    Dim valName As String: valName = "Cross-Sheet Validation"
    modConfig.SafeDeleteSheet valName
    
    Dim wsVal As Worksheet
    Set wsVal = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsVal.Name = valName
    
    ' Write title
    wsVal.Range("A1").Value = "CROSS-SHEET DATA VALIDATION"
    wsVal.Range("A1").Font.Bold = True
    wsVal.Range("A1").Font.Size = 14
    wsVal.Range("A1").Font.Color = CLR_NAVY

    modConfig.StyleHeader wsVal, 4, _
        Array("Check #", "Description", "Sheet A", "Value A", _
              "Sheet B", "Value B", "Difference", "Status")
    
    Dim outRow As Long: outRow = 5
    Dim checkNum As Long: checkNum = 0
    Dim passCount As Long: passCount = 0
    Dim failCount As Long: failCount = 0
    
    '=======================================================================
    ' CHECK 1: GL Total Amount vs P&L Trend Total Revenue (FY Total)
    '=======================================================================
    If modConfig.SheetExists(SH_GL) And modConfig.SheetExists(SH_PL_TREND) Then
        modPerformance.UpdateStatus "Check 1: GL totals vs P&L Trend...", 0.15
        
        Dim wsGL As Worksheet: Set wsGL = ThisWorkbook.Worksheets(SH_GL)
        Dim glLastRow As Long: glLastRow = wsGL.Cells(wsGL.Rows.Count, 1).End(xlUp).Row
        
        ' Sum all GL amounts (column 7 = Amount)
        Dim glTotal As Double: glTotal = 0
        Dim gr As Long
        For gr = 2 To glLastRow
            glTotal = glTotal + modConfig.SafeNum(wsGL.Cells(gr, 7).Value)
        Next gr
        
        ' Find Total Revenue on P&L Trend
        Dim wsTrend As Worksheet: Set wsTrend = ThisWorkbook.Worksheets(SH_PL_TREND)
        Dim trendLastRow As Long: trendLastRow = wsTrend.Cells(wsTrend.Rows.Count, 1).End(xlUp).Row
        Dim trendLastCol As Long: trendLastCol = wsTrend.Cells(1, wsTrend.Columns.Count).End(xlToLeft).Column
        Dim fyCol As Long: fyCol = FindFYCol(wsTrend, trendLastCol)
        
        Dim trendRevRow As Long: trendRevRow = 0
        Dim tr As Long
        For tr = 2 To trendLastRow
            If InStr(1, LCase(Trim(CStr(wsTrend.Cells(tr, 1).Value))), "total revenue") > 0 Then
                trendRevRow = tr: Exit For
            End If
        Next tr
        
        Dim trendRevVal As Double: trendRevVal = 0
        If trendRevRow > 0 Then trendRevVal = modConfig.SafeNum(wsTrend.Cells(trendRevRow, fyCol).Value)
        
        checkNum = checkNum + 1
        WriteValidationRow wsVal, outRow, checkNum, _
            "GL Total Amount vs P&L Trend Total Revenue (FY Total)", _
            SH_GL, glTotal, SH_PL_TREND, trendRevVal, passCount, failCount
        outRow = outRow + 1
    End If
    
    '=======================================================================
    ' CHECK 2: GL by Department vs Functional P&L Totals
    '=======================================================================
    If modConfig.SheetExists(SH_GL) And modConfig.SheetExists(SH_FUNC_JAN) Then
        modPerformance.UpdateStatus "Check 2: GL by Dept vs Functional P&L...", 0.35
        
        ' Sum GL amounts for January (Month = 1 or "Jan")
        Dim glJanTotal As Double: glJanTotal = 0
        Dim dateCol As Long: dateCol = COL_GL_DATE    ' Column B = Date
        Dim amtCol As Long: amtCol = COL_GL_AMOUNT  ' Column G = Amount
        
        If wsGL Is Nothing Then Set wsGL = ThisWorkbook.Worksheets(SH_GL)
        If glLastRow = 0 Then glLastRow = wsGL.Cells(wsGL.Rows.Count, 1).End(xlUp).Row
        
        For gr = 2 To glLastRow
            Dim glDate As Variant: glDate = wsGL.Cells(gr, dateCol).Value
            If IsDate(glDate) Then
                If Month(CDate(glDate)) = 1 Then
                    glJanTotal = glJanTotal + modConfig.SafeNum(wsGL.Cells(gr, amtCol).Value)
                End If
            End If
        Next gr
        
        ' Find total on Functional P&L - Jan
        Dim wsFuncJan As Worksheet: Set wsFuncJan = ThisWorkbook.Worksheets(SH_FUNC_JAN)
        Dim funcLastRow As Long: funcLastRow = wsFuncJan.Cells(wsFuncJan.Rows.Count, 1).End(xlUp).Row
        Dim funcLastCol As Long: funcLastCol = modConfig.LastCol(wsFuncJan, HDR_ROW_FUNC)
        
        ' Find a total row on the Functional P&L
        Dim funcTotal As Double: funcTotal = 0
        Dim funcTotRow As Long: funcTotRow = 0
        For tr = funcLastRow To DATA_ROW_FUNC Step -1
            Dim funcLbl As String: funcLbl = LCase(Trim(CStr(wsFuncJan.Cells(tr, 1).Value)))
            If InStr(funcLbl, "total revenue") > 0 Or InStr(funcLbl, "net") > 0 Then
                funcTotRow = tr: Exit For
            End If
        Next tr
        
        If funcTotRow > 0 Then
            funcTotal = modConfig.SafeNum(wsFuncJan.Cells(funcTotRow, funcLastCol).Value)
        End If
        
        checkNum = checkNum + 1
        WriteValidationRow wsVal, outRow, checkNum, _
            "GL January Amount vs Functional P&L Jan Total", _
            SH_GL & " (Jan)", glJanTotal, SH_FUNC_JAN, funcTotal, passCount, failCount
        outRow = outRow + 1
    End If
    
    '=======================================================================
    ' CHECK 3: GL by Product vs Product Line Summary
    '=======================================================================
    If modConfig.SheetExists(SH_GL) And modConfig.SheetExists(SH_PROD_SUMMARY) Then
        modPerformance.UpdateStatus "Check 3: GL by Product vs Product Summary...", 0.55
        
        Dim products As Variant: products = modConfig.GetProducts()
        Dim prodCol As Long: prodCol = 4   ' Column D = Product
        
        If wsGL Is Nothing Then Set wsGL = ThisWorkbook.Worksheets(SH_GL)
        If glLastRow = 0 Then glLastRow = wsGL.Cells(wsGL.Rows.Count, 1).End(xlUp).Row
        
        ' Sum GL by product
        Dim p As Long
        For p = 0 To UBound(products)
            Dim pName As String: pName = CStr(products(p))
            Dim glProdTotal As Double: glProdTotal = 0
            
            For gr = 2 To glLastRow
                Dim glProd As String: glProd = Trim(CStr(wsGL.Cells(gr, prodCol).Value))
                If InStr(1, glProd, pName, vbTextCompare) > 0 Then
                    glProdTotal = glProdTotal + modConfig.SafeNum(wsGL.Cells(gr, amtCol).Value)
                End If
            Next gr
            
            ' Write as REVIEW item (Product Summary doesn't have a simple single total)
            checkNum = checkNum + 1
            wsVal.Cells(outRow, 1).Value = checkNum
            wsVal.Cells(outRow, 2).Value = "GL Total for " & pName & " vs Product Summary"
            wsVal.Cells(outRow, 3).Value = SH_GL
            wsVal.Cells(outRow, 4).Value = glProdTotal
            wsVal.Cells(outRow, 4).NumberFormat = "$#,##0.00"
            wsVal.Cells(outRow, 5).Value = SH_PROD_SUMMARY
            wsVal.Cells(outRow, 6).Value = "-"
            wsVal.Cells(outRow, 7).Value = "Manual"
            wsVal.Cells(outRow, 8).Value = "REVIEW"
            wsVal.Cells(outRow, 8).Font.Color = RGB(255, 165, 0)
            wsVal.Cells(outRow, 8).Font.Bold = True
            outRow = outRow + 1
        Next p
    End If
    
    '=======================================================================
    ' CHECK 4: Mirror existing Checks sheet PASS/FAIL results
    '=======================================================================
    If modConfig.SheetExists(SH_CHECKS) Then
        modPerformance.UpdateStatus "Check 4: Mirroring Checks sheet...", 0.75
        Dim wsChk As Worksheet: Set wsChk = ThisWorkbook.Worksheets(SH_CHECKS)
        Dim chkLastRow As Long: chkLastRow = wsChk.Cells(wsChk.Rows.Count, 1).End(xlUp).Row
        
        For gr = DATA_ROW_CHECKS To chkLastRow
            Dim chkName As String: chkName = Trim(CStr(wsChk.Cells(gr, 1).Value))
            If chkName <> "" Then
                checkNum = checkNum + 1
                wsVal.Cells(outRow, 1).Value = checkNum
                wsVal.Cells(outRow, 2).Value = "[Checks] " & chkName
                wsVal.Cells(outRow, 3).Value = CStr(wsChk.Cells(gr, 2).Value)
                wsVal.Cells(outRow, 4).Value = modConfig.SafeNum(wsChk.Cells(gr, 2).Value)
                wsVal.Cells(outRow, 4).NumberFormat = "$#,##0.00"
                wsVal.Cells(outRow, 5).Value = CStr(wsChk.Cells(gr, 3).Value)
                wsVal.Cells(outRow, 6).Value = modConfig.SafeNum(wsChk.Cells(gr, 3).Value)
                wsVal.Cells(outRow, 6).NumberFormat = "$#,##0.00"
                wsVal.Cells(outRow, 7).Value = modConfig.SafeNum(wsChk.Cells(gr, 4).Value)
                wsVal.Cells(outRow, 7).NumberFormat = "$#,##0.00;($#,##0.00)"
                
                Dim chkStatus As String: chkStatus = UCase(Trim(CStr(wsChk.Cells(gr, COL_CHECK_STATUS).Value)))
                wsVal.Cells(outRow, 8).Value = chkStatus
                If chkStatus = "PASS" Then
                    wsVal.Cells(outRow, 8).Font.Color = RGB(0, 128, 0)
                    passCount = passCount + 1
                Else
                    wsVal.Cells(outRow, 8).Font.Color = RGB(192, 0, 0)
                    wsVal.Cells(outRow, 8).Font.Bold = True
                    failCount = failCount + 1
                End If
                outRow = outRow + 1
            End If
        Next gr
    End If
    
    '=======================================================================
    ' Summary & Format
    '=======================================================================
    wsVal.Range("A3").Value = "Total: " & checkNum & "  |  PASS: " & passCount & _
                              "  |  FAIL: " & failCount
    wsVal.Range("A3").Font.Bold = True
    wsVal.Range("A3").Font.Color = IIf(failCount > 0, RGB(192, 0, 0), RGB(0, 128, 0))
    
    wsVal.Columns("A").ColumnWidth = 8
    wsVal.Columns("B").ColumnWidth = 45
    wsVal.Columns("C").ColumnWidth = 20
    wsVal.Columns("D").ColumnWidth = 14
    wsVal.Columns("E").ColumnWidth = 20
    wsVal.Columns("F").ColumnWidth = 14
    wsVal.Columns("G").ColumnWidth = 14
    wsVal.Columns("H").ColumnWidth = 12
    
    ' Add borders and autofilter
    If outRow > 5 Then
        With wsVal.Range("A4:H" & outRow - 1).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(191, 191, 191)
        End With
    End If
    
    wsVal.Range("A4:H4").AutoFilter
    wsVal.Tab.Color = RGB(0, 176, 80)
    wsVal.Activate
    
    Dim elapsed As Double: elapsed = modPerformance.ElapsedSeconds()
    modPerformance.TurboOff
    
    modLogger.LogAction "modReconciliation", "ValidateCrossSheet", _
        checkNum & " checks: " & passCount & " pass, " & failCount & " fail", elapsed
    
    MsgBox "Cross-Sheet Validation Complete!" & vbCrLf & vbCrLf & _
           "  Total Checks: " & checkNum & vbCrLf & _
           "  PASS: " & passCount & vbCrLf & _
           "  FAIL: " & failCount & vbCrLf & vbCrLf & _
           "  Results on '" & valName & "'", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modReconciliation", "ERROR-ValidateCrossSheet", Err.Description
    MsgBox "Cross-sheet validation error: " & Err.Description, vbCritical, APP_NAME
End Sub


'===============================================================================
'
' ===  PRIVATE HELPERS  ========================================================
'
'===============================================================================

'===============================================================================
' WriteValidationRow - Write one validation row and update pass/fail counts
' Compares valA vs valB; PASS if difference <= RECON_TOLERANCE ($1 default).
'===============================================================================
Private Sub WriteValidationRow(ByVal ws As Worksheet, ByVal r As Long, _
    ByVal chkNum As Long, ByVal desc As String, _
    ByVal sheetA As String, ByVal valA As Double, _
    ByVal sheetB As String, ByVal valB As Double, _
    ByRef passCount As Long, ByRef failCount As Long)
    
    Dim diff As Double: diff = valA - valB
    
    ws.Cells(r, 1).Value = chkNum
    ws.Cells(r, 2).Value = desc
    ws.Cells(r, 3).Value = sheetA
    ws.Cells(r, 4).Value = valA: ws.Cells(r, 4).NumberFormat = "$#,##0.00"
    ws.Cells(r, 5).Value = sheetB
    ws.Cells(r, 6).Value = valB: ws.Cells(r, 6).NumberFormat = "$#,##0.00"
    ws.Cells(r, 7).Value = diff: ws.Cells(r, 7).NumberFormat = "$#,##0.00;($#,##0.00)"
    
    If Abs(diff) <= RECON_TOLERANCE Then
        ws.Cells(r, 8).Value = "PASS"
        ws.Cells(r, 8).Font.Color = RGB(0, 128, 0)
        passCount = passCount + 1
    Else
        ws.Cells(r, 8).Value = "FAIL"
        ws.Cells(r, 8).Font.Color = RGB(192, 0, 0)
        ws.Cells(r, 8).Font.Bold = True
        failCount = failCount + 1
    End If
End Sub

'===============================================================================
' FindFYCol - Locate the FY Total column on a sheet
' Searches row 1 for "FY Total", "FY2025", or FISCAL_YEAR_4 pattern.
' Falls back to lastCol - 1.
'===============================================================================
Private Function FindFYCol(ByVal ws As Worksheet, ByVal lastCol As Long) As Long
    Dim c As Long
    For c = 2 To lastCol
        Dim hdr As String: hdr = LCase(Trim(CStr(ws.Cells(1, c).Value)))
        If InStr(hdr, "fy") > 0 And InStr(hdr, "total") > 0 Then
            FindFYCol = c: Exit Function
        End If
        If InStr(hdr, "fy" & FISCAL_YEAR_4) > 0 Then
            FindFYCol = c: Exit Function
        End If
        If InStr(hdr, FISCAL_YEAR_4 & " total") > 0 Then
            FindFYCol = c: Exit Function
        End If
    Next c
    FindFYCol = Application.Max(2, lastCol - 1)  ' Fallback
End Function
