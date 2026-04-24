Attribute VB_Name = "modExecBrief"
Option Explicit

'===============================================================================
' modExecBrief - One-Page Executive Brief Auto-Generator
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  One button scans the entire workbook and generates a plain English
'           executive summary. Covers revenue trends, expense highlights,
'           reconciliation status, data quality, and key variances.
'           Ready to paste into an email or print for leadership.
'
' PUBLIC SUBS:
'   GenerateExecBrief   - Build the executive brief sheet
'
' DEPENDENCIES: modConfig, modPerformance, modLogger
' VERSION:  2.1.0
'===============================================================================

Private Const SH_BRIEF As String = "Executive Brief"

'===============================================================================
' GenerateExecBrief - Scan workbook and produce plain English summary
'===============================================================================
Public Sub GenerateExecBrief()
    On Error GoTo ErrHandler

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Scanning workbook for executive brief...", 0.1

    ' Remove old brief
    modConfig.SafeDeleteSheet SH_BRIEF

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = SH_BRIEF

    ' --- Page Setup ---
    ws.Columns(1).ColumnWidth = 45
    ws.Columns(2).ColumnWidth = 18
    ws.Columns(3).ColumnWidth = 12
    ws.Cells.Interior.Color = RGB(255, 255, 255)

    ' --- Title Block ---
    Dim r As Long: r = 1

    ws.Cells(r, 1).Value = "EXECUTIVE BRIEF"
    ws.Cells(r, 1).Font.Size = 20: ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Color = RGB(255, 255, 255)
    ws.Cells(r, 1).Font.Name = "Arial"
    ws.Range("A1:C1").Interior.Color = CLR_NAVY
    ws.Range("A1:C1").Merge
    ws.Rows(1).RowHeight = 40
    ws.Cells(r, 1).VerticalAlignment = xlCenter
    r = r + 1

    ws.Cells(r, 1).Value = "Keystone BenefitTech, Inc. - P&L Model Summary"
    ws.Cells(r, 1).Font.Size = 11: ws.Cells(r, 1).Font.Color = RGB(255, 255, 255)
    ws.Cells(r, 1).Font.Name = "Arial"
    ws.Range("A2:C2").Interior.Color = RGB(11, 71, 121)
    ws.Range("A2:C2").Merge
    r = r + 1

    ws.Cells(r, 1).Value = "Generated: " & Format(Now, "mmmm d, yyyy h:mm AM/PM")
    ws.Cells(r, 1).Font.Size = 9: ws.Cells(r, 1).Font.Italic = True
    r = r + 2

    ' Divider
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 1)).Interior.Color = CLR_NAVY
    ws.Rows(r).RowHeight = 3
    r = r + 2

    modPerformance.UpdateStatus "Analyzing P&L Trend...", 0.2

    ' =========================================================================
    ' SECTION 1: Revenue & P&L Highlights
    ' =========================================================================
    WriteSectionHeader ws, r, "1. REVENUE & P&L HIGHLIGHTS", RGB(0, 128, 0)
    r = r + 1

    If modConfig.SheetExists(SH_PL_TREND) Then
        Dim wsTrend As Worksheet: Set wsTrend = ThisWorkbook.Worksheets(SH_PL_TREND)
        Dim revInsights As String: revInsights = AnalyzeRevenue(wsTrend)
        ws.Cells(r, 1).Value = revInsights
        ws.Cells(r, 1).WrapText = True
        ws.Cells(r, 1).Font.Name = "Arial"
        ws.Cells(r, 1).Font.Size = 10
        r = r + CountLines(revInsights) + 1
    Else
        ws.Cells(r, 1).Value = "P&L Trend sheet not found - unable to analyze revenue."
        ws.Cells(r, 1).Font.Italic = True
        r = r + 2
    End If

    modPerformance.UpdateStatus "Analyzing Reconciliation...", 0.4

    ' =========================================================================
    ' SECTION 2: Reconciliation Status
    ' =========================================================================
    WriteSectionHeader ws, r, "2. RECONCILIATION STATUS", RGB(255, 165, 0)
    r = r + 1

    If modConfig.SheetExists(SH_CHECKS) Then
        Dim wsChecks As Worksheet: Set wsChecks = ThisWorkbook.Worksheets(SH_CHECKS)
        Dim reconInsights As String: reconInsights = AnalyzeReconciliation(wsChecks)
        ws.Cells(r, 1).Value = reconInsights
        ws.Cells(r, 1).WrapText = True
        r = r + CountLines(reconInsights) + 1
    Else
        ws.Cells(r, 1).Value = "Checks sheet not found - unable to analyze reconciliation."
        ws.Cells(r, 1).Font.Italic = True
        r = r + 2
    End If

    modPerformance.UpdateStatus "Analyzing Assumptions...", 0.6

    ' =========================================================================
    ' SECTION 3: Key Assumptions
    ' =========================================================================
    WriteSectionHeader ws, r, "3. KEY ASSUMPTIONS & DRIVERS", RGB(0, 112, 192)
    r = r + 1

    If modConfig.SheetExists(SH_ASSUMPTIONS) Then
        Dim wsAssume As Worksheet: Set wsAssume = ThisWorkbook.Worksheets(SH_ASSUMPTIONS)
        Dim assumeInsights As String: assumeInsights = AnalyzeAssumptions(wsAssume)
        ws.Cells(r, 1).Value = assumeInsights
        ws.Cells(r, 1).WrapText = True
        r = r + CountLines(assumeInsights) + 1
    Else
        ws.Cells(r, 1).Value = "Assumptions sheet not found."
        ws.Cells(r, 1).Font.Italic = True
        r = r + 2
    End If

    modPerformance.UpdateStatus "Analyzing product mix...", 0.8

    ' =========================================================================
    ' SECTION 4: Product Line Summary
    ' =========================================================================
    WriteSectionHeader ws, r, "4. PRODUCT LINE OVERVIEW", RGB(112, 48, 160)
    r = r + 1

    If modConfig.SheetExists(SH_PROD_SUMMARY) Then
        Dim wsProd As Worksheet: Set wsProd = ThisWorkbook.Worksheets(SH_PROD_SUMMARY)
        Dim prodInsights As String: prodInsights = AnalyzeProducts(wsProd)
        ws.Cells(r, 1).Value = prodInsights
        ws.Cells(r, 1).WrapText = True
        r = r + CountLines(prodInsights) + 1
    Else
        ws.Cells(r, 1).Value = "Product Line Summary sheet not found."
        ws.Cells(r, 1).Font.Italic = True
        r = r + 2
    End If

    ' =========================================================================
    ' SECTION 5: Workbook Health
    ' =========================================================================
    WriteSectionHeader ws, r, "5. WORKBOOK HEALTH", RGB(0, 128, 128)
    r = r + 1

    Dim healthInsights As String: healthInsights = AnalyzeWorkbookHealth()
    ws.Cells(r, 1).Value = healthInsights
    ws.Cells(r, 1).WrapText = True
    r = r + CountLines(healthInsights) + 1

    ' =========================================================================
    ' Footer
    ' =========================================================================
    r = r + 1
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 1)).Interior.Color = CLR_NAVY
    ws.Rows(r).RowHeight = 3
    r = r + 2

    ws.Cells(r, 1).Value = "Generated by Keystone BenefitTech Automation Toolkit v" & APP_VERSION
    ws.Cells(r, 1).Font.Size = 8: ws.Cells(r, 1).Font.Italic = True
    ws.Cells(r, 1).Font.Color = RGB(150, 150, 150)
    r = r + 1
    ws.Cells(r, 1).Value = "This brief can be copied and pasted into an email or printed for leadership review."
    ws.Cells(r, 1).Font.Size = 8: ws.Cells(r, 1).Font.Italic = True
    ws.Cells(r, 1).Font.Color = RGB(150, 150, 150)

    ws.Activate
    ws.Range("A1").Select

    modPerformance.TurboOff

    modLogger.LogAction "modExecBrief", "GenerateExecBrief", "Executive brief generated with 5 sections"

    MsgBox "Executive Brief generated!" & vbCrLf & vbCrLf & _
           "5 sections: Revenue, Reconciliation, Assumptions, Products, Health" & vbCrLf & vbCrLf & _
           "Ready to copy/paste into email or print for leadership.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modExecBrief", "ERROR", Err.Description
    MsgBox "Executive Brief error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' WriteSectionHeader - Formatted section header with colored status bar
'===============================================================================
Private Sub WriteSectionHeader(ByRef ws As Worksheet, ByVal r As Long, _
                                ByVal title As String, ByVal barColor As Long)
    ws.Cells(r, 1).Value = title
    ws.Cells(r, 1).Font.Size = 13
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Color = RGB(255, 255, 255)
    ws.Cells(r, 1).Font.Name = "Arial"
    ws.Cells(r, 1).VerticalAlignment = xlCenter
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Interior.Color = barColor
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Merge
    ws.Rows(r).RowHeight = 28
End Sub

'===============================================================================
' AnalyzeRevenue - Scan P&L Trend for revenue and margin insights
'===============================================================================
Private Function AnalyzeRevenue(ByRef wsTrend As Worksheet) As String
    Dim s As String: s = ""

    ' Find key rows
    Dim revRow As Long: revRow = modConfig.FindRowByLabel(wsTrend, "Total Revenue", DATA_ROW_REPORT)
    Dim expRow As Long: expRow = modConfig.FindRowByLabel(wsTrend, "Total Expense", DATA_ROW_REPORT)
    If expRow = 0 Then expRow = modConfig.FindRowByLabel(wsTrend, "Total Operating", DATA_ROW_REPORT)
    Dim netRow As Long: netRow = modConfig.FindRowByLabel(wsTrend, "Net Income", DATA_ROW_REPORT)
    If netRow = 0 Then netRow = modConfig.FindRowByLabel(wsTrend, "Net Operating", DATA_ROW_REPORT)

    ' Find month columns (start at col 2 on trend sheet)
    Dim mths As Variant: mths = modConfig.GetMonths()
    Dim lastDataCol As Long: lastDataCol = 2  ' At least Jan

    Dim c As Long
    For c = 2 To 13
        If Not IsEmpty(wsTrend.Cells(HDR_ROW_REPORT, c).Value) Then
            lastDataCol = c
        End If
    Next c

    Dim monthCount As Long: monthCount = lastDataCol - 1
    If monthCount < 1 Then monthCount = 1

    ' Get latest month values
    If revRow > 0 Then
        Dim latestRev As Double: latestRev = modConfig.SafeNum(wsTrend.Cells(revRow, lastDataCol).Value)
        s = s & "- Total Revenue (latest month): " & Format(latestRev, "$#,##0")

        ' MoM trend if we have 2+ months
        If monthCount >= 2 Then
            Dim prevRev As Double: prevRev = modConfig.SafeNum(wsTrend.Cells(revRow, lastDataCol - 1).Value)
            If prevRev <> 0 Then
                Dim revChange As Double: revChange = (latestRev - prevRev) / Abs(prevRev)
                If revChange >= 0 Then
                    s = s & " (up " & Format(revChange, "0.0%") & " MoM)"
                Else
                    s = s & " (down " & Format(Abs(revChange), "0.0%") & " MoM)"
                End If
            End If
        End If
        s = s & vbCrLf
    End If

    If expRow > 0 Then
        Dim latestExp As Double: latestExp = modConfig.SafeNum(wsTrend.Cells(expRow, lastDataCol).Value)
        s = s & "- Total Expenses (latest month): " & Format(Abs(latestExp), "$#,##0") & vbCrLf
    End If

    If netRow > 0 Then
        Dim latestNet As Double: latestNet = modConfig.SafeNum(wsTrend.Cells(netRow, lastDataCol).Value)
        s = s & "- Net Income (latest month): " & Format(latestNet, "$#,##0")
        If latestRev <> 0 Then
            s = s & " (margin: " & Format(latestNet / latestRev, "0.0%") & ")"
        End If
        s = s & vbCrLf
    End If

    If monthCount >= 1 Then
        s = s & "- Data covers " & monthCount & " month(s) of FY" & FISCAL_YEAR & vbCrLf
    End If

    If s = "" Then s = "- Unable to locate revenue/expense rows on P&L Trend." & vbCrLf
    AnalyzeRevenue = s
End Function

'===============================================================================
' AnalyzeReconciliation - Scan Checks sheet for PASS/FAIL status
'===============================================================================
Private Function AnalyzeReconciliation(ByRef wsChecks As Worksheet) As String
    Dim s As String: s = ""
    Dim lastRow As Long: lastRow = modConfig.LastRow(wsChecks, 1)

    Dim totalChecks As Long: totalChecks = 0
    Dim passCount As Long: passCount = 0
    Dim failCount As Long: failCount = 0
    Dim failNames As String: failNames = ""

    Dim r As Long
    For r = DATA_ROW_CHECKS To lastRow
        Dim checkName As String: checkName = Trim(CStr(wsChecks.Cells(r, 1).Value))
        If checkName = "" Then GoTo NextCheckRow

        totalChecks = totalChecks + 1
        Dim status As String: status = UCase(Trim(CStr(wsChecks.Cells(r, COL_CHECK_STATUS).Value)))

        If status = "PASS" Then
            passCount = passCount + 1
        ElseIf status = "FAIL" Then
            failCount = failCount + 1
            failNames = failNames & "    - " & checkName & vbCrLf
        End If
NextCheckRow:
    Next r

    If totalChecks = 0 Then
        s = "- No reconciliation checks found." & vbCrLf
    Else
        s = s & "- " & totalChecks & " reconciliation checks evaluated" & vbCrLf
        s = s & "- " & passCount & " PASS, " & failCount & " FAIL" & vbCrLf

        If failCount = 0 Then
            s = s & "- All checks passing - model is balanced and reconciled." & vbCrLf
        Else
            s = s & "- ATTENTION: " & failCount & " check(s) failing:" & vbCrLf
            s = s & failNames
        End If
    End If

    AnalyzeReconciliation = s
End Function

'===============================================================================
' AnalyzeAssumptions - Summarize key drivers from Assumptions sheet
'===============================================================================
Private Function AnalyzeAssumptions(ByRef wsA As Worksheet) As String
    Dim s As String: s = ""
    Dim lastRow As Long: lastRow = modConfig.LastRow(wsA, 1)
    Dim driverCount As Long: driverCount = 0

    Dim r As Long
    For r = DATA_ROW_ASSUME To lastRow
        Dim dName As String: dName = Trim(CStr(wsA.Cells(r, 1).Value))
        If dName <> "" Then
            driverCount = driverCount + 1
            ' Show first 8 drivers as sample
            If driverCount <= 8 Then
                Dim dVal As Variant: dVal = wsA.Cells(r, 2).Value
                If IsNumeric(dVal) Then
                    If Abs(CDbl(dVal)) < 1 Then
                        s = s & "- " & dName & ": " & Format(dVal, "0.0%") & vbCrLf
                    Else
                        s = s & "- " & dName & ": " & Format(dVal, "#,##0.00") & vbCrLf
                    End If
                Else
                    s = s & "- " & dName & ": " & CStr(dVal) & vbCrLf
                End If
            End If
        End If
    Next r

    If driverCount > 8 Then
        s = s & "- ... and " & (driverCount - 8) & " more drivers" & vbCrLf
    End If

    s = s & "- Total: " & driverCount & " assumption drivers configured" & vbCrLf
    AnalyzeAssumptions = s
End Function

'===============================================================================
' AnalyzeProducts - Summarize product line data
'===============================================================================
Private Function AnalyzeProducts(ByRef wsProd As Worksheet) As String
    Dim s As String: s = ""
    Dim products As Variant: products = modConfig.GetProducts()

    s = s & "- " & (UBound(products) + 1) & " product lines: " & PRODUCTS_CSV & vbCrLf

    ' Try to find revenue data for each product
    Dim p As Long
    For p = 0 To UBound(products)
        Dim prodCol As Long
        prodCol = modConfig.FindColByHeader(wsProd, CStr(products(p)), HDR_ROW_REPORT)
        If prodCol > 0 Then
            ' Find revenue row
            Dim revRow As Long
            revRow = modConfig.FindRowByLabel(wsProd, "Total Revenue", DATA_ROW_REPORT)
            If revRow = 0 Then revRow = modConfig.FindRowByLabel(wsProd, "Revenue", DATA_ROW_REPORT)
            If revRow > 0 Then
                Dim prodRev As Double
                prodRev = modConfig.SafeNum(wsProd.Cells(revRow, prodCol).Value)
                s = s & "- " & CStr(products(p)) & " Revenue: " & Format(prodRev, "$#,##0") & vbCrLf
            End If
        End If
    Next p

    AnalyzeProducts = s
End Function

'===============================================================================
' AnalyzeWorkbookHealth - General workbook statistics
'===============================================================================
Private Function AnalyzeWorkbookHealth() As String
    Dim s As String: s = ""

    ' Count sheets
    Dim visibleCount As Long: visibleCount = 0
    Dim hiddenCount As Long: hiddenCount = 0
    Dim sh As Worksheet

    For Each sh In ThisWorkbook.Worksheets
        If sh.Visible = xlSheetVisible Then
            visibleCount = visibleCount + 1
        Else
            hiddenCount = hiddenCount + 1
        End If
    Next sh

    s = s & "- Total sheets: " & ThisWorkbook.Worksheets.Count & _
            " (" & visibleCount & " visible, " & hiddenCount & " hidden)" & vbCrLf

    ' File size
    Dim fileSize As String
    Dim fSize As Long
    On Error Resume Next
    fSize = FileLen(ThisWorkbook.FullName)
    On Error GoTo 0

    If fSize > 0 Then
        If fSize > 1048576 Then
            fileSize = Format(fSize / 1048576, "#,##0.0") & " MB"
        Else
            fileSize = Format(fSize / 1024, "#,##0") & " KB"
        End If
        s = s & "- File size: " & fileSize & vbCrLf
    End If

    ' VBA module count
    On Error Resume Next
    Dim moduleCount As Long
    moduleCount = ThisWorkbook.VBProject.VBComponents.Count
    If Err.Number = 0 Then
        s = s & "- VBA modules: " & moduleCount & vbCrLf
    End If
    On Error GoTo 0

    s = s & "- Toolkit version: " & APP_VERSION & " (Build " & APP_BUILD_DATE & ")" & vbCrLf
    s = s & "- Fiscal year: FY" & FISCAL_YEAR_4 & vbCrLf

    AnalyzeWorkbookHealth = s
End Function

'===============================================================================
' CountLines - Count vbCrLf occurrences in a string (for row spacing)
'===============================================================================
Private Function CountLines(ByVal s As String) As Long
    Dim count As Long: count = 1
    Dim pos As Long: pos = 1
    Do
        pos = InStr(pos, s, vbCrLf)
        If pos = 0 Then Exit Do
        count = count + 1
        pos = pos + 2
    Loop
    CountLines = count
End Function
