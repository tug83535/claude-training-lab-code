Attribute VB_Name = "modDashboardAdvanced"
Option Explicit

'===============================================================================
' modDashboardAdvanced - Executive Dashboards & Advanced Chart Tools
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Advanced dashboard and chart tools split from modDashboard (2026-03-05).
'           These are the CFO/CEO-level executive views and specialized chart tools.
'
' PUBLIC SUBS:
'   CreateExecutiveDashboard    - KPI cards + summary table on dedicated sheet
'   WaterfallChart              - Revenue-to-Net-Income waterfall bridge
'   ProductComparison           - Side-by-side product metrics + ranking
'   LinkDynamicChartTitles      - Sync chart titles to month selector cell
'   CreateSmallMultiplesGrid    - 2x2 grid of product revenue charts
'
' VERSION:  2.1.0
' SPLIT:    Extracted from modDashboard_v2.1.bas (2026-03-05)
'           modDashboard retains: BuildDashboard, RefreshDashboard, ReformatChartsAndVisuals
'
' PRIOR FIXES (carried forward):
'   T5.01a  CreateExecutiveDashboard reads row 4 for headers, not row 1
'   T5.01b  ChrW() for Unicode arrows (>255 codepoint)
'   T5.01c  Multi-pass row label search with fallbacks
'   T5.02   WaterfallChart multi-variant row label search
'===============================================================================

'===============================================================================
' CreateExecutiveDashboard - Full executive dashboard on dedicated sheet
' Creates KPI cards (Revenue, GM%, OpEx, Net Income), trend arrows,
' and Actual vs Budget summary table with variance coloring.
'===============================================================================
Public Sub CreateExecutiveDashboard()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_PL_TREND) Then
        MsgBox "'" & SH_PL_TREND & "' not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Building Executive Dashboard...", 0.1

    '--- Gather data from P&L Trend ---
    Dim wsTrend As Worksheet: Set wsTrend = ThisWorkbook.Worksheets(SH_PL_TREND)
    Dim tLastRow As Long: tLastRow = modConfig.LastRow(wsTrend, 1)
    Dim tLastCol As Long: tLastCol = modConfig.LastCol(wsTrend, HDR_ROW_REPORT)

    Debug.Print "[ExecDash] Trend sheet: " & SH_PL_TREND & _
                " | HDR_ROW_REPORT=" & HDR_ROW_REPORT & _
                " | tLastRow=" & tLastRow & " | tLastCol=" & tLastCol

    Dim c As Long
    Dim hdrDump As String: hdrDump = "[ExecDash] Row 4 headers:"
    For c = 1 To tLastCol
        hdrDump = hdrDump & " [" & c & "]=""" & wsTrend.Cells(HDR_ROW_REPORT, c).Value & """"
    Next c
    Debug.Print hdrDump

    ' --- Find FY Total column using multiple strategies ---
    Dim fyCol As Long: fyCol = FindFYTotalCol(wsTrend)
    If fyCol = tLastCol Then
        Dim tryCol As Long
        tryCol = modConfig.FindColByHeader(wsTrend, "2025 Total", HDR_ROW_REPORT)
        If tryCol = 0 Then tryCol = modConfig.FindColByHeader(wsTrend, "YTD", HDR_ROW_REPORT)
        If tryCol = 0 Then tryCol = modConfig.FindColByHeader(wsTrend, "FY25", HDR_ROW_REPORT)
        If tryCol = 0 Then tryCol = modConfig.FindColByHeader(wsTrend, "Full Year", HDR_ROW_REPORT)
        If tryCol > 0 Then fyCol = tryCol
    End If
    Debug.Print "[ExecDash] FY Total column: " & fyCol & _
                " (header = """ & wsTrend.Cells(HDR_ROW_REPORT, fyCol).Value & """)"

    ' --- Find budget column ---
    Dim budCol As Long
    budCol = modConfig.FindColByHeader(wsTrend, "budget", HDR_ROW_REPORT)
    If budCol = 0 Then budCol = modConfig.FindColByHeader(wsTrend, "plan", HDR_ROW_REPORT)
    If budCol = 0 Then budCol = modConfig.FindColByHeader(wsTrend, "target", HDR_ROW_REPORT)
    If budCol = 0 Then budCol = tLastCol
    Debug.Print "[ExecDash] Budget column: " & budCol & _
                " (header = """ & wsTrend.Cells(HDR_ROW_REPORT, budCol).Value & """)"

    ' --- Find key P&L summary rows ---
    Dim revRow As Long, gpRow As Long, opexRow As Long, niRow As Long

    revRow = modConfig.FindRowByLabel(wsTrend, "total revenue", DATA_ROW_REPORT)
    If revRow = 0 Then revRow = modConfig.FindRowByLabel(wsTrend, "consolidated revenue", DATA_ROW_REPORT)
    If revRow = 0 Then revRow = modConfig.FindRowByLabel(wsTrend, "net revenue", DATA_ROW_REPORT)

    gpRow = modConfig.FindRowByLabel(wsTrend, "gross profit", DATA_ROW_REPORT)
    If gpRow = 0 Then gpRow = modConfig.FindRowByLabel(wsTrend, "gross margin", DATA_ROW_REPORT)

    opexRow = modConfig.FindRowByLabel(wsTrend, "total operating expense", DATA_ROW_REPORT)
    If opexRow = 0 Then opexRow = modConfig.FindRowByLabel(wsTrend, "operating expense", DATA_ROW_REPORT)
    If opexRow = 0 Then opexRow = modConfig.FindRowByLabel(wsTrend, "total opex", DATA_ROW_REPORT)
    If opexRow = 0 Then opexRow = modConfig.FindRowByLabel(wsTrend, "total expenses", DATA_ROW_REPORT)

    niRow = modConfig.FindRowByLabel(wsTrend, "net income", DATA_ROW_REPORT)
    If niRow = 0 Then niRow = modConfig.FindRowByLabel(wsTrend, "net operating income", DATA_ROW_REPORT)
    If niRow = 0 Then niRow = modConfig.FindRowByLabel(wsTrend, "net operating profit", DATA_ROW_REPORT)
    If niRow = 0 Then niRow = modConfig.FindRowByLabel(wsTrend, "operating income", DATA_ROW_REPORT)

    Dim rowDump As String: rowDump = "[ExecDash] Col A labels (rows 5-" & Application.Min(tLastRow, 44) & "):"
    Dim r As Long
    For r = DATA_ROW_REPORT To Application.Min(tLastRow, 44)
        Dim lbl As String: lbl = Trim(CStr(wsTrend.Cells(r, 1).Value))
        If Len(lbl) > 0 Then rowDump = rowDump & " [" & r & "]=""" & lbl & """"
    Next r
    Debug.Print rowDump

    Debug.Print "[ExecDash] Row lookup — revRow=" & revRow & " gpRow=" & gpRow & _
                " opexRow=" & opexRow & " niRow=" & niRow
    If revRow = 0 Then
        Debug.Print "[ExecDash] WARNING: No revenue row found! Dashboard will show $0."
    End If

    Dim fyRev As Double: If revRow > 0 Then fyRev = modConfig.SafeNum(wsTrend.Cells(revRow, fyCol).Value)
    Dim budRev As Double: If revRow > 0 Then budRev = modConfig.SafeNum(wsTrend.Cells(revRow, budCol).Value)
    Dim fyGP As Double: If gpRow > 0 Then fyGP = modConfig.SafeNum(wsTrend.Cells(gpRow, fyCol).Value)
    Dim fyOpex As Double: If opexRow > 0 Then fyOpex = modConfig.SafeNum(wsTrend.Cells(opexRow, fyCol).Value)
    Dim fyNI As Double: If niRow > 0 Then fyNI = modConfig.SafeNum(wsTrend.Cells(niRow, fyCol).Value)
    Debug.Print "[ExecDash] Values — fyRev=" & fyRev & " budRev=" & budRev & _
                " fyGP=" & fyGP & " fyOpex=" & fyOpex & " fyNI=" & fyNI

    Dim gmPct As Double: If fyRev <> 0 Then gmPct = fyGP / fyRev
    Dim revVar As Double: revVar = fyRev - budRev
    Dim revVarPct As Double: If budRev <> 0 Then revVarPct = revVar / Abs(budRev)

    '--- Create dashboard sheet ---
    Dim dashName As String: dashName = SH_DASHBOARD
    modConfig.SafeDeleteSheet dashName

    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsDash.Name = dashName
    wsDash.Cells.Interior.Color = RGB(245, 246, 250)

    With wsDash.Range("A1")
        .Value = "EXECUTIVE DASHBOARD - FY" & FISCAL_YEAR_4
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(0, 51, 102)
    End With
    wsDash.Range("A2").Value = "Generated: " & Format(Now, "mmmm d, yyyy h:mm AM/PM")
    wsDash.Range("A2").Font.Italic = True
    wsDash.Range("A2").Font.Color = RGB(128, 128, 128)

    modPerformance.UpdateStatus "Writing KPI cards...", 0.4

    '--- KPI Cards ---
    Dim kpiRow As Long: kpiRow = 4
    Dim kpiLabels As Variant
    kpiLabels = Array("Total Revenue", "Gross Margin %", "Operating Expenses", "Net Income")
    Dim kpiValues As Variant
    kpiValues = Array(fyRev, gmPct, fyOpex, fyNI)
    Dim kpiFmts As Variant
    kpiFmts = Array("$#,##0", "0.0%", "$#,##0", "$#,##0")
    Dim kpiColors As Variant
    kpiColors = Array(RGB(0, 112, 192), RGB(0, 176, 80), RGB(192, 0, 0), RGB(112, 48, 160))

    Dim k As Long
    For k = 0 To 3
        Dim kCol As Long: kCol = (k * 3) + 1

        wsDash.Cells(kpiRow, kCol).Value = kpiLabels(k)
        wsDash.Cells(kpiRow, kCol).Font.Size = 9
        wsDash.Cells(kpiRow, kCol).Font.Color = RGB(100, 100, 100)

        wsDash.Cells(kpiRow + 1, kCol).Value = kpiValues(k)
        wsDash.Cells(kpiRow + 1, kCol).NumberFormat = kpiFmts(k)
        wsDash.Cells(kpiRow + 1, kCol).Font.Size = 22
        wsDash.Cells(kpiRow + 1, kCol).Font.Bold = True
        wsDash.Cells(kpiRow + 1, kCol).Font.Color = RGB(0, 51, 102)

        Dim arrow As String
        If k = 0 Then
            If revVar >= 0 Then
                arrow = ChrW(9650) & " +" & Format(revVarPct, "0.0%")
                wsDash.Cells(kpiRow + 2, kCol).Font.Color = RGB(0, 128, 0)
            Else
                arrow = ChrW(9660) & " " & Format(revVarPct, "0.0%")
                wsDash.Cells(kpiRow + 2, kCol).Font.Color = RGB(192, 0, 0)
            End If
        Else
            arrow = "vs Budget"
            wsDash.Cells(kpiRow + 2, kCol).Font.Color = RGB(128, 128, 128)
        End If
        wsDash.Cells(kpiRow + 2, kCol).Value = arrow
        wsDash.Cells(kpiRow + 2, kCol).Font.Size = 9

        wsDash.Cells(kpiRow - 1, kCol).Interior.Color = kpiColors(k)
        wsDash.Range(wsDash.Cells(kpiRow - 1, kCol), wsDash.Cells(kpiRow - 1, kCol + 1)).Merge
        wsDash.Range(wsDash.Cells(kpiRow - 1, kCol), wsDash.Cells(kpiRow - 1, kCol + 1)).RowHeight = 5
    Next k

    '--- Summary Table: Actual vs Budget ---
    modPerformance.UpdateStatus "Building summary table...", 0.7
    Dim tblRow As Long: tblRow = kpiRow + 5

    modConfig.StyleHeader wsDash, tblRow, _
        Array("Metric", "FY Actual", "Budget", "Variance $", "Variance %")

    Dim metrics As Variant: metrics = Array("Total Revenue", "Gross Profit", "Operating Expenses", "Net Income")
    Dim actVals As Variant: actVals = Array(fyRev, fyGP, fyOpex, fyNI)

    Dim metricRows As Variant
    metricRows = Array(revRow, gpRow, opexRow, niRow)

    Dim mr As Long
    For mr = 0 To 3
        Dim mRow As Long: mRow = tblRow + 1 + mr
        Dim mActual As Double: mActual = CDbl(actVals(mr))
        Dim mBudget As Double
        If CLng(metricRows(mr)) > 0 Then
            mBudget = modConfig.SafeNum(wsTrend.Cells(CLng(metricRows(mr)), budCol).Value)
        End If
        Dim mVar As Double: mVar = mActual - mBudget

        wsDash.Cells(mRow, 1).Value = metrics(mr)
        wsDash.Cells(mRow, 1).Font.Bold = True
        wsDash.Cells(mRow, 2).Value = mActual: wsDash.Cells(mRow, 2).NumberFormat = "$#,##0"
        wsDash.Cells(mRow, 3).Value = mBudget: wsDash.Cells(mRow, 3).NumberFormat = "$#,##0"
        wsDash.Cells(mRow, 4).Value = mVar: wsDash.Cells(mRow, 4).NumberFormat = "$#,##0;($#,##0)"
        If mBudget <> 0 Then
            wsDash.Cells(mRow, 5).Value = mVar / Abs(mBudget)
            wsDash.Cells(mRow, 5).NumberFormat = "0.0%"
        End If

        If mVar > 0 Then
            wsDash.Cells(mRow, 4).Font.Color = RGB(0, 128, 0)
            wsDash.Cells(mRow, 5).Font.Color = RGB(0, 128, 0)
        ElseIf mVar < 0 Then
            wsDash.Cells(mRow, 4).Font.Color = RGB(192, 0, 0)
            wsDash.Cells(mRow, 5).Font.Color = RGB(192, 0, 0)
        End If
    Next mr

    '--- PRODUCT REVENUE BREAKDOWN ---
    modPerformance.UpdateStatus "Building product breakdown...", 0.8

    Dim prodRow As Long: prodRow = tblRow + 6
    wsDash.Cells(prodRow, 1).Value = "PRODUCT REVENUE BREAKDOWN"
    wsDash.Cells(prodRow, 1).Font.Size = 12
    wsDash.Cells(prodRow, 1).Font.Bold = True
    wsDash.Cells(prodRow, 1).Font.Color = RGB(0, 51, 102)
    Dim pc As Long
    For pc = 1 To 6
        wsDash.Cells(prodRow, pc).Borders(xlEdgeBottom).Color = RGB(0, 176, 240)
    Next pc

    modConfig.StyleHeader wsDash, prodRow + 1, _
        Array("Product", "FY Revenue", "% of Total", "Status")

    Dim products As Variant: products = modConfig.GetProducts()
    Dim pIdx As Long
    Dim prodRevRow As Long
    For pIdx = 0 To UBound(products)
        Dim pName As String: pName = CStr(products(pIdx))
        Dim pRow As Long: pRow = prodRow + 2 + pIdx

        ' Find this product's revenue row on P&L Trend
        prodRevRow = 0
        Dim sr As Long
        Dim inBlock As Boolean: inBlock = False
        For sr = DATA_ROW_REPORT To tLastRow
            Dim sLabel As String: sLabel = Trim(CStr(wsTrend.Cells(sr, 1).Value))
            If InStr(1, sLabel, pName, vbTextCompare) > 0 Then inBlock = True
            If inBlock And LCase(sLabel) = "revenue" Then
                prodRevRow = sr
                Exit For
            End If
            If inBlock And sr > DATA_ROW_REPORT + 2 And sLabel = "" Then inBlock = False
        Next sr

        Dim pRevenue As Double: pRevenue = 0
        If prodRevRow > 0 Then pRevenue = modConfig.SafeNum(wsTrend.Cells(prodRevRow, fyCol).Value)

        wsDash.Cells(pRow, 1).Value = pName
        wsDash.Cells(pRow, 1).Font.Bold = True
        wsDash.Cells(pRow, 1).Font.Size = 11
        wsDash.Cells(pRow, 2).Value = pRevenue
        wsDash.Cells(pRow, 2).NumberFormat = "$#,##0"

        If fyRev <> 0 Then
            wsDash.Cells(pRow, 3).Value = pRevenue / fyRev
            wsDash.Cells(pRow, 3).NumberFormat = "0.0%"
        End If

        ' Status indicator based on share
        Dim pShare As Double: If fyRev <> 0 Then pShare = pRevenue / fyRev
        If pShare >= 0.3 Then
            wsDash.Cells(pRow, 4).Value = ChrW(9679) & " Strong"
            wsDash.Cells(pRow, 4).Font.Color = RGB(0, 128, 0)
        ElseIf pShare >= 0.1 Then
            wsDash.Cells(pRow, 4).Value = ChrW(9679) & " Stable"
            wsDash.Cells(pRow, 4).Font.Color = RGB(255, 165, 0)
        Else
            wsDash.Cells(pRow, 4).Value = ChrW(9679) & " Watch"
            wsDash.Cells(pRow, 4).Font.Color = RGB(192, 0, 0)
        End If
        wsDash.Cells(pRow, 4).Font.Bold = True

        ' Alternate row shading
        If pIdx Mod 2 = 0 Then
            Dim shadeC As Long
            For shadeC = 1 To 4
                wsDash.Cells(pRow, shadeC).Interior.Color = RGB(240, 245, 250)
            Next shadeC
        End If
    Next pIdx

    '--- MONTHLY REVENUE TREND ---
    modPerformance.UpdateStatus "Building monthly trend...", 0.9

    Dim trendRow As Long: trendRow = prodRow + 3 + UBound(products) + 2
    wsDash.Cells(trendRow, 1).Value = "MONTHLY REVENUE TREND"
    wsDash.Cells(trendRow, 1).Font.Size = 12
    wsDash.Cells(trendRow, 1).Font.Bold = True
    wsDash.Cells(trendRow, 1).Font.Color = RGB(0, 51, 102)
    For pc = 1 To 13
        wsDash.Cells(trendRow, pc).Borders(xlEdgeBottom).Color = RGB(0, 176, 240)
    Next pc

    ' Month headers
    Dim mths As Variant: mths = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", _
                                       "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    wsDash.Cells(trendRow + 1, 1).Value = "Month"
    wsDash.Cells(trendRow + 1, 1).Font.Bold = True
    Dim mi As Long
    For mi = 0 To 11
        wsDash.Cells(trendRow + 1, mi + 2).Value = mths(mi)
        wsDash.Cells(trendRow + 1, mi + 2).Font.Bold = True
        wsDash.Cells(trendRow + 1, mi + 2).Font.Size = 9
        wsDash.Cells(trendRow + 1, mi + 2).Font.Color = RGB(255, 255, 255)
        wsDash.Cells(trendRow + 1, mi + 2).Interior.Color = RGB(11, 71, 121)
        wsDash.Cells(trendRow + 1, mi + 2).HorizontalAlignment = xlCenter
    Next mi

    ' Revenue values per month (from consolidated revenue row)
    wsDash.Cells(trendRow + 2, 1).Value = "Revenue"
    wsDash.Cells(trendRow + 2, 1).Font.Bold = True
    If revRow > 0 Then
        Dim maxMonthRev As Double: maxMonthRev = 0
        Dim minMonthRev As Double: minMonthRev = 999999999
        For mi = 0 To 11
            Dim monthVal As Double
            monthVal = modConfig.SafeNum(wsTrend.Cells(revRow, mi + 2).Value)
            wsDash.Cells(trendRow + 2, mi + 2).Value = monthVal
            wsDash.Cells(trendRow + 2, mi + 2).NumberFormat = "$#,##0"
            wsDash.Cells(trendRow + 2, mi + 2).HorizontalAlignment = xlCenter
            If monthVal > maxMonthRev Then maxMonthRev = monthVal
            If monthVal < minMonthRev And monthVal > 0 Then minMonthRev = monthVal
        Next mi

        ' Color code: highest = dark green, lowest = red, others = gradient
        For mi = 0 To 11
            Dim mVal As Double
            mVal = modConfig.SafeNum(wsDash.Cells(trendRow + 2, mi + 2).Value)
            If mVal = maxMonthRev Then
                wsDash.Cells(trendRow + 2, mi + 2).Font.Color = RGB(0, 100, 0)
                wsDash.Cells(trendRow + 2, mi + 2).Font.Bold = True
                wsDash.Cells(trendRow + 2, mi + 2).Interior.Color = RGB(198, 239, 206)
            ElseIf mVal = minMonthRev Then
                wsDash.Cells(trendRow + 2, mi + 2).Font.Color = RGB(156, 0, 6)
                wsDash.Cells(trendRow + 2, mi + 2).Interior.Color = RGB(255, 235, 238)
            Else
                wsDash.Cells(trendRow + 2, mi + 2).Interior.Color = RGB(240, 245, 250)
            End If
        Next mi
    End If

    ' MoM Growth row
    wsDash.Cells(trendRow + 3, 1).Value = "MoM Growth"
    wsDash.Cells(trendRow + 3, 1).Font.Bold = True
    wsDash.Cells(trendRow + 3, 1).Font.Size = 9
    If revRow > 0 Then
        For mi = 1 To 11
            Dim prevMo As Double: prevMo = modConfig.SafeNum(wsTrend.Cells(revRow, mi + 1).Value)
            Dim currMo As Double: currMo = modConfig.SafeNum(wsTrend.Cells(revRow, mi + 2).Value)
            If prevMo <> 0 Then
                Dim momGrowth As Double: momGrowth = (currMo - prevMo) / Abs(prevMo)
                wsDash.Cells(trendRow + 3, mi + 2).Value = momGrowth
                wsDash.Cells(trendRow + 3, mi + 2).NumberFormat = "0.0%"
                wsDash.Cells(trendRow + 3, mi + 2).HorizontalAlignment = xlCenter
                If momGrowth >= 0 Then
                    wsDash.Cells(trendRow + 3, mi + 2).Font.Color = RGB(0, 128, 0)
                Else
                    wsDash.Cells(trendRow + 3, mi + 2).Font.Color = RGB(192, 0, 0)
                End If
            End If
        Next mi
    End If

    '--- KPI STATUS INDICATORS (add to existing KPI cards) ---
    ' Add green/red dot next to the +3.3% arrow on row 6
    If revVar > 0 Then
        wsDash.Cells(kpiRow + 2, 1).Value = ChrW(9679) & " +" & Format(revVarPct, "0.0%") & " vs Budget"
        wsDash.Cells(kpiRow + 2, 1).Font.Color = RGB(0, 128, 0)
    Else
        wsDash.Cells(kpiRow + 2, 1).Value = ChrW(9679) & " " & Format(revVarPct, "0.0%") & " vs Budget"
        wsDash.Cells(kpiRow + 2, 1).Font.Color = RGB(192, 0, 0)
    End If
    wsDash.Cells(kpiRow + 2, 1).Font.Bold = True
    wsDash.Cells(kpiRow + 2, 1).Font.Size = 10

    wsDash.Columns("A:N").AutoFit
    wsDash.Tab.Color = RGB(0, 51, 102)
    wsDash.Activate
    wsDash.Range("A1").Select

    Dim elapsed As Double: elapsed = modPerformance.ElapsedSeconds()
    modPerformance.TurboOff

    modLogger.LogAction "modDashboardAdvanced", "CreateExecutiveDashboard", _
        "4 KPIs + summary + products + trend (" & Format(elapsed, "0.0") & "s)"
    Debug.Print "[ExecDash] SUCCESS — Dashboard created on '" & dashName & "'"
    MsgBox "Executive Dashboard created on '" & dashName & "'.", vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    Dim errMsg As String
    errMsg = "Executive dashboard error in modDashboardAdvanced.CreateExecutiveDashboard" & vbCrLf & vbCrLf & _
             "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
             "Source: " & Err.Source & vbCrLf & vbCrLf & _
             "Diagnostic values at time of error:" & vbCrLf & _
             "  HDR_ROW_REPORT = " & HDR_ROW_REPORT & vbCrLf & _
             "  tLastCol = " & tLastCol & vbCrLf & _
             "  fyCol = " & fyCol & vbCrLf & _
             "  budCol = " & budCol & vbCrLf & _
             "  revRow = " & revRow & " | gpRow = " & gpRow & vbCrLf & _
             "  opexRow = " & opexRow & " | niRow = " & niRow & vbCrLf & _
             "  fyRev = " & fyRev & vbCrLf & vbCrLf & _
             "Check the Immediate Window (Ctrl+G) for full Debug.Print trace."
    modPerformance.TurboOff
    modLogger.LogAction "modDashboardAdvanced", "ERROR-ExecDash", _
        "Err " & Err.Number & ": " & Err.Description & _
        " | fyCol=" & fyCol & " budCol=" & budCol & " revRow=" & revRow
    Debug.Print "[ExecDash] ERROR — " & Err.Number & ": " & Err.Description
    MsgBox errMsg, vbCritical, APP_NAME
End Sub

'===============================================================================
' WaterfallChart - Revenue-to-Net-Income waterfall bridge
'===============================================================================
Public Sub WaterfallChart()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_PL_TREND) Then
        MsgBox "'" & SH_PL_TREND & "' not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Building waterfall chart...", 0.2

    Dim wsSrc As Worksheet: Set wsSrc = ThisWorkbook.Worksheets(SH_PL_TREND)
    Dim fyCol As Long: fyCol = FindFYTotalCol(wsSrc)
    Dim lastRow As Long: lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row

    Dim revRow As Long, cogsRow As Long, gpRow As Long, opexRow As Long, niRow As Long

    revRow = modConfig.FindRowByLabel(wsSrc, "total revenue", DATA_ROW_REPORT)
    If revRow = 0 Then revRow = modConfig.FindRowByLabel(wsSrc, "consolidated revenue", DATA_ROW_REPORT)
    If revRow = 0 Then revRow = modConfig.FindRowByLabel(wsSrc, "net revenue", DATA_ROW_REPORT)
    If revRow = 0 Then revRow = modConfig.FindRowByLabel(wsSrc, "revenue", DATA_ROW_REPORT)

    cogsRow = modConfig.FindRowByLabel(wsSrc, "cost of", DATA_ROW_REPORT)
    If cogsRow = 0 Then cogsRow = modConfig.FindRowByLabel(wsSrc, "cogs", DATA_ROW_REPORT)

    gpRow = modConfig.FindRowByLabel(wsSrc, "gross profit", DATA_ROW_REPORT)
    If gpRow = 0 Then gpRow = modConfig.FindRowByLabel(wsSrc, "gross margin", DATA_ROW_REPORT)

    opexRow = modConfig.FindRowByLabel(wsSrc, "total operating", DATA_ROW_REPORT)
    If opexRow = 0 Then opexRow = modConfig.FindRowByLabel(wsSrc, "operating expense", DATA_ROW_REPORT)
    If opexRow = 0 Then opexRow = modConfig.FindRowByLabel(wsSrc, "total opex", DATA_ROW_REPORT)
    If opexRow = 0 Then opexRow = modConfig.FindRowByLabel(wsSrc, "total expenses", DATA_ROW_REPORT)

    niRow = modConfig.FindRowByLabel(wsSrc, "net income", DATA_ROW_REPORT)

    If revRow = 0 Then
        modPerformance.TurboOff
        MsgBox "Could not find Revenue row on P&L Trend." & vbCrLf & _
               "Expected labels: Total Revenue, Net Revenue, Revenue, etc.", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim wfLabels As Variant
    wfLabels = Array("Revenue", "COGS", "Gross Profit", "OpEx", "Net Income")
    Dim wfValues(0 To 4) As Double
    wfValues(0) = modConfig.SafeNum(wsSrc.Cells(revRow, fyCol).Value)
    If cogsRow > 0 Then wfValues(1) = modConfig.SafeNum(wsSrc.Cells(cogsRow, fyCol).Value)
    If gpRow > 0 Then wfValues(2) = modConfig.SafeNum(wsSrc.Cells(gpRow, fyCol).Value)
    If opexRow > 0 Then wfValues(3) = modConfig.SafeNum(wsSrc.Cells(opexRow, fyCol).Value)
    If niRow > 0 Then wfValues(4) = modConfig.SafeNum(wsSrc.Cells(niRow, fyCol).Value)

    Dim wfName As String: wfName = "P&L Waterfall"
    modConfig.SafeDeleteSheet wfName

    Dim wsWF As Worksheet
    Set wsWF = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsWF.Name = wfName

    wsWF.Range("A1").Value = "P&L WATERFALL - FY" & FISCAL_YEAR_4
    wsWF.Range("A1").Font.Size = 14
    wsWF.Range("A1").Font.Bold = True

    wsWF.Cells(3, 1).Value = "Category"
    wsWF.Cells(3, 2).Value = "Base"
    wsWF.Cells(3, 3).Value = "Increase"
    wsWF.Cells(3, 4).Value = "Decrease"

    Dim wfBase(0 To 4) As Double
    Dim wfInc(0 To 4) As Double
    Dim wfDec(0 To 4) As Double

    wfBase(0) = 0: wfInc(0) = wfValues(0): wfDec(0) = 0
    wfBase(1) = wfValues(0) + wfValues(1): wfInc(1) = 0: wfDec(1) = Abs(wfValues(1))
    wfBase(2) = 0: wfInc(2) = wfValues(2): wfDec(2) = 0
    wfBase(3) = wfValues(2) + wfValues(3): wfInc(3) = 0: wfDec(3) = Abs(wfValues(3))
    wfBase(4) = 0: wfInc(4) = wfValues(4): wfDec(4) = 0

    Dim wi As Long
    For wi = 0 To 4
        wsWF.Cells(4 + wi, 1).Value = wfLabels(wi)
        wsWF.Cells(4 + wi, 2).Value = wfBase(wi)
        wsWF.Cells(4 + wi, 3).Value = wfInc(wi)
        wsWF.Cells(4 + wi, 4).Value = wfDec(wi)
    Next wi

    Dim chartTop As Long: chartTop = wsWF.Cells(10, 1).Top
    Dim co As ChartObject
    Set co = wsWF.ChartObjects.Add(20, chartTop, 500, 320)
    co.Name = "WaterfallChart"

    With co.Chart
        .ChartType = xlBarStacked
        .SetSourceData Source:=wsWF.Range("A3:D8")
        .HasTitle = True
        .ChartTitle.Text = "P&L Waterfall - FY" & FISCAL_YEAR_4

        On Error Resume Next
        .SeriesCollection(1).Format.Fill.Visible = msoFalse
        .SeriesCollection(1).Format.Line.Visible = msoFalse
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(0, 176, 80)
        .SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(192, 0, 0)
        On Error GoTo ErrHandler

        .HasLegend = False
    End With

    wsWF.Tab.Color = RGB(0, 112, 192)
    wsWF.Activate

    Dim elapsed As Double: elapsed = modPerformance.ElapsedSeconds()
    modPerformance.TurboOff

    modLogger.LogAction "modDashboardAdvanced", "WaterfallChart", "5 P&L stages (" & Format(elapsed, "0.0") & "s)"
    MsgBox "Waterfall chart created on '" & wfName & "'.", vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modDashboardAdvanced", "ERROR-Waterfall", Err.Description
    MsgBox "Waterfall error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' ProductComparison - Side-by-side product line comparison table + chart
'===============================================================================
Public Sub ProductComparison()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_PL_TREND) Then
        MsgBox "'" & SH_PL_TREND & "' not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Building product comparison...", 0.1

    Dim wsSrc As Worksheet: Set wsSrc = ThisWorkbook.Worksheets(SH_PL_TREND)
    Dim srcLastRow As Long: srcLastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    Dim products As Variant: products = modConfig.GetProducts()

    Dim pcName As String: pcName = "Product Comparison"
    modConfig.SafeDeleteSheet pcName

    Dim wsPC As Worksheet
    Set wsPC = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsPC.Name = pcName

    Dim hdrs() As String
    ReDim hdrs(0 To UBound(products) + 1)
    hdrs(0) = "Metric"
    Dim p As Long
    For p = 0 To UBound(products)
        hdrs(p + 1) = CStr(products(p))
    Next p

    modConfig.StyleHeader wsPC, 4, hdrs

    Dim metricNames As Variant
    metricNames = Array("Revenue", "Contribution Margin", "Operating Expense", "Net Income")

    Dim outRow As Long: outRow = 5
    Dim mi As Long
    For mi = 0 To UBound(metricNames)
        wsPC.Cells(outRow, 1).Value = metricNames(mi)
        wsPC.Cells(outRow, 1).Font.Bold = True

        For p = 0 To UBound(products)
            Dim metricVal As Double
            metricVal = FindProductMetric(wsSrc, CStr(products(p)), CStr(metricNames(mi)), srcLastRow)
            wsPC.Cells(outRow, p + 2).Value = metricVal
            wsPC.Cells(outRow, p + 2).NumberFormat = "$#,##0"
        Next p

        outRow = outRow + 1
    Next mi

    wsPC.Cells(outRow, 1).Value = "Margin %"
    wsPC.Cells(outRow, 1).Font.Bold = True
    For p = 0 To UBound(products)
        Dim pRev As Double: pRev = modConfig.SafeNum(wsPC.Cells(5, p + 2).Value)
        Dim pCM As Double: pCM = modConfig.SafeNum(wsPC.Cells(6, p + 2).Value)
        If pRev <> 0 Then
            wsPC.Cells(outRow, p + 2).Value = pCM / pRev
            wsPC.Cells(outRow, p + 2).NumberFormat = "0.0%"
        End If
    Next p
    outRow = outRow + 1

    modPerformance.UpdateStatus "Calculating rankings...", 0.5
    wsPC.Cells(outRow, 1).Value = "Revenue Rank"
    wsPC.Cells(outRow, 1).Font.Bold = True

    Dim revVals() As Double: ReDim revVals(0 To UBound(products))
    For p = 0 To UBound(products)
        revVals(p) = modConfig.SafeNum(wsPC.Cells(5, p + 2).Value)
    Next p
    For p = 0 To UBound(products)
        Dim rank As Long: rank = 1
        Dim q As Long
        For q = 0 To UBound(products)
            If revVals(q) > revVals(p) Then rank = rank + 1
        Next q
        wsPC.Cells(outRow, p + 2).Value = "#" & rank
        wsPC.Cells(outRow, p + 2).HorizontalAlignment = xlCenter
        wsPC.Cells(outRow, p + 2).Font.Bold = True
    Next p

    modPerformance.UpdateStatus "Creating comparison chart...", 0.7

    Dim chartTop As Long: chartTop = wsPC.Cells(outRow + 2, 1).Top
    Dim co As ChartObject
    Set co = wsPC.ChartObjects.Add(20, chartTop, 500, 280)
    co.Name = "ProductCompChart"

    On Error Resume Next
    With co.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=wsPC.Range(wsPC.Cells(4, 1), _
            wsPC.Cells(4 + UBound(metricNames), UBound(products) + 2))
        .HasTitle = True
        .ChartTitle.Text = "Product Comparison - Key Metrics"
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With
    On Error GoTo ErrHandler

    wsPC.Columns("A").ColumnWidth = 22
    Dim fc As Long
    For fc = 2 To UBound(products) + 2
        wsPC.Columns(fc).ColumnWidth = 16
    Next fc
    wsPC.Tab.Color = RGB(0, 176, 80)
    wsPC.Activate

    Dim elapsed As Double: elapsed = modPerformance.ElapsedSeconds()
    modPerformance.TurboOff

    modLogger.LogAction "modDashboardAdvanced", "ProductComparison", _
        (UBound(products) + 1) & " products compared (" & Format(elapsed, "0.0") & "s)"
    MsgBox "Product Comparison created on '" & pcName & "'.", vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modDashboardAdvanced", "ERROR-ProductComp", Err.Description
    MsgBox "Product comparison error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' LinkDynamicChartTitles - Link all chart titles to a selector cell (#44)
'===============================================================================
Public Sub LinkDynamicChartTitles()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_REPORT) Then
        MsgBox "Report--> sheet not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    Dim selectedMonth As String: selectedMonth = ""
    If modConfig.SheetExists("FPL Summary - Dynamic") Then
        selectedMonth = modConfig.SafeStr( _
            ThisWorkbook.Worksheets("FPL Summary - Dynamic").Range("B4").Value)
    End If
    If Len(selectedMonth) = 0 Then
        selectedMonth = Format(Date, "mmm")
    End If

    Dim wsReport As Worksheet: Set wsReport = ThisWorkbook.Worksheets(SH_REPORT)
    Dim updateCount As Long: updateCount = 0

    Dim co As ChartObject
    For Each co In wsReport.ChartObjects
        On Error Resume Next
        If co.Chart.HasTitle Then
            Dim oldTitle As String: oldTitle = co.Chart.ChartTitle.Text
            If InStr(oldTitle, selectedMonth) = 0 Then
                Dim newTitle As String
                Dim mths As Variant: mths = modConfig.GetMonths()
                newTitle = oldTitle
                Dim m As Long
                For m = 0 To UBound(mths)
                    newTitle = Replace(newTitle, " - " & CStr(mths(m)), "")
                Next m
                newTitle = newTitle & " - " & selectedMonth
                co.Chart.ChartTitle.Text = newTitle
                updateCount = updateCount + 1
            End If
        End If
        On Error GoTo ErrHandler
    Next co

    modLogger.LogAction "modDashboardAdvanced", "LinkDynamicChartTitles", _
        updateCount & " chart title(s) updated to " & selectedMonth
    MsgBox updateCount & " chart title(s) updated to show: " & selectedMonth, _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    MsgBox "LinkDynamicChartTitles error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' CreateSmallMultiplesGrid - Generate one small chart per product (#86)
'===============================================================================
Public Sub CreateSmallMultiplesGrid()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_PL_TREND) Then
        MsgBox "'" & SH_PL_TREND & "' not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Building small multiples grid...", 0.1

    Dim wsSrc As Worksheet: Set wsSrc = ThisWorkbook.Worksheets(SH_PL_TREND)
    Dim products As Variant: products = modConfig.GetProducts()
    Dim mths As Variant: mths = modConfig.GetMonths()

    Dim lastDataCol As Long: lastDataCol = 1
    Dim revRow As Long: revRow = modConfig.FindRowByLabel(wsSrc, "total revenue", DATA_ROW_REPORT)
    If revRow = 0 Then revRow = DATA_ROW_REPORT
    Dim c As Long
    For c = 13 To 2 Step -1
        If modConfig.SafeNum(wsSrc.Cells(revRow, c).Value) <> 0 Then
            lastDataCol = c: Exit For
        End If
    Next c
    If lastDataCol < 2 Then lastDataCol = 13

    Dim monthCount As Long: monthCount = lastDataCol - 1
    Dim monthLabels() As String
    ReDim monthLabels(0 To monthCount - 1)
    Dim mi As Long
    For mi = 0 To monthCount - 1
        monthLabels(mi) = CStr(mths(mi))
    Next mi

    Dim smName As String: smName = "Product Small Multiples"
    modConfig.SafeDeleteSheet smName
    Dim wsSM As Worksheet
    Set wsSM = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsSM.Name = smName
    wsSM.Cells.Interior.Color = RGB(245, 246, 250)

    With wsSM.Range("A1")
        .Value = "Product Revenue — Small Multiples (FY" & FISCAL_YEAR_4 & ")"
        .Font.Size = 14: .Font.Bold = True: .Font.Color = CLR_NAVY
    End With

    Dim prodColors As Variant
    prodColors = Array(RGB(31, 78, 121), RGB(68, 114, 196), _
                       RGB(112, 173, 71), RGB(237, 125, 49))

    Dim chartLeft  As Variant: chartLeft  = Array(20, 390, 20, 390)
    Dim chartTop   As Variant: chartTop   = Array(60, 60, 280, 280)
    Dim chartW     As Long: chartW = 355
    Dim chartH     As Long: chartH = 205

    modPerformance.UpdateStatus "Creating 4 product charts...", 0.4

    Dim p As Long
    For p = 0 To Application.Min(3, UBound(products))
        Dim productName As String: productName = CStr(products(p))

        Dim pRevRow As Long: pRevRow = FindProductRevenueRow(wsSrc, productName)
        If pRevRow = 0 Then GoTo NextProduct

        Dim co As ChartObject
        Set co = wsSM.ChartObjects.Add( _
            Left:=CLng(chartLeft(p)), Top:=CLng(chartTop(p)), _
            Width:=chartW, Height:=chartH)
        co.Name = "SM_" & productName

        With co.Chart
            .ChartType = xlLine
            .HasTitle = True
            .ChartTitle.Text = productName & " Revenue"
            .ChartTitle.Font.Size = 10
            .ChartTitle.Font.Bold = True
            .ChartTitle.Font.Color = CLng(prodColors(p))

            Dim ser As Series
            Set ser = .SeriesCollection.NewSeries
            ser.Name = productName
            ser.Values = wsSrc.Range(wsSrc.Cells(pRevRow, 2), wsSrc.Cells(pRevRow, lastDataCol))
            ser.XValues = monthLabels

            On Error Resume Next
            ser.Format.Line.ForeColor.RGB = CLng(prodColors(p))
            ser.Format.Line.Weight = 2
            On Error GoTo ErrHandler

            .HasLegend = False
            .PlotArea.Interior.Color = CLR_WHITE
            .PlotArea.Border.LineStyle = 0

            Dim ax As Axis
            Set ax = .Axes(xlValue)
            ax.TickLabels.NumberFormat = "$#,##0"
            ax.TickLabels.Font.Size = 7

            Set ax = .Axes(xlCategory)
            ax.TickLabels.Font.Size = 7
        End With
NextProduct:
    Next p

    wsSM.Tab.Color = CLR_NAVY
    wsSM.Activate

    Dim elapsed As Double: elapsed = modPerformance.ElapsedSeconds()
    modPerformance.TurboOff
    modLogger.LogAction "modDashboardAdvanced", "CreateSmallMultiplesGrid", _
        UBound(products) + 1 & " product charts | " & monthCount & " months (" & Format(elapsed, "0.0") & "s)"
    MsgBox "Small multiples grid created on '" & smName & "'." & vbCrLf & _
           "4 product revenue charts at the same scale for easy comparison.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    MsgBox "CreateSmallMultiplesGrid error: " & Err.Description, vbCritical, APP_NAME
End Sub


'===============================================================================
'
' ===  PRIVATE HELPERS — Row/Column Finders  ===================================
'
'===============================================================================

'===============================================================================
' FindProductRevenueRow - Locate Revenue row within a product block
'===============================================================================
Private Function FindProductRevenueRow(ByVal ws As Worksheet, ByVal product As String) As Long
    Dim lastRow As Long: lastRow = modConfig.LastRow(ws, 1)
    Dim inProductBlock As Boolean: inProductBlock = False
    Dim r As Long

    For r = 1 To lastRow
        Dim cellVal As String: cellVal = Trim(CStr(ws.Cells(r, 1).Value))

        If InStr(1, cellVal, product, vbTextCompare) > 0 And _
           InStr(1, cellVal, "Consolidated", vbTextCompare) = 0 Then
            inProductBlock = True
        End If

        If inProductBlock And cellVal = "Revenue" Then
            FindProductRevenueRow = r
            Exit Function
        End If

        If inProductBlock And r > 5 And cellVal = "" Then
            inProductBlock = False
        End If
    Next r

    FindProductRevenueRow = 0
End Function

'===============================================================================
' FindProductCMPctRow - Locate Contribution Margin % row within a product block
'===============================================================================
Private Function FindProductCMPctRow(ByVal ws As Worksheet, ByVal product As String) As Long
    Dim lastRow As Long: lastRow = modConfig.LastRow(ws, 1)
    Dim inProductBlock As Boolean: inProductBlock = False
    Dim r As Long

    For r = 1 To lastRow
        Dim cellVal As String: cellVal = Trim(CStr(ws.Cells(r, 1).Value))

        If InStr(1, cellVal, product, vbTextCompare) > 0 And _
           InStr(1, cellVal, "Consolidated", vbTextCompare) = 0 Then
            inProductBlock = True
        End If

        If inProductBlock And cellVal = "Contribution Margin %" Then
            FindProductCMPctRow = r
            Exit Function
        End If

        If inProductBlock And r > 5 And cellVal = "" Then
            inProductBlock = False
        End If
    Next r

    FindProductCMPctRow = 0
End Function

'===============================================================================
' FindFYTotalCol - Locate the FY Total column on P&L Trend
'===============================================================================
Private Function FindFYTotalCol(ByVal ws As Worksheet) As Long
    Dim hdrRow As Long: hdrRow = HDR_ROW_REPORT
    Dim lastCol As Long: lastCol = ws.Cells(hdrRow, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long

    Debug.Print "[FindFYTotalCol] Scanning row " & hdrRow & " across " & lastCol & " columns"

    For c = 2 To lastCol
        Dim hdr As String: hdr = LCase(Trim(CStr(ws.Cells(hdrRow, c).Value)))
        If InStr(hdr, "fy") > 0 And InStr(hdr, "total") > 0 Then
            Debug.Print "[FindFYTotalCol] MATCH pass 1 — col " & c & " = """ & ws.Cells(hdrRow, c).Value & """"
            FindFYTotalCol = c: Exit Function
        End If
    Next c

    For c = 2 To lastCol
        hdr = LCase(Trim(CStr(ws.Cells(hdrRow, c).Value)))
        If InStr(hdr, "fy" & FISCAL_YEAR_4) > 0 Then
            FindFYTotalCol = c: Exit Function
        End If
        If InStr(hdr, FISCAL_YEAR_4 & " total") > 0 Then
            FindFYTotalCol = c: Exit Function
        End If
    Next c

    For c = 2 To lastCol
        hdr = LCase(Trim(CStr(ws.Cells(hdrRow, c).Value)))
        If hdr = "total" Or hdr = "year total" Or hdr = "annual total" Then
            FindFYTotalCol = c: Exit Function
        End If
    Next c

    Debug.Print "[FindFYTotalCol] No header match — falling back to lastCol=" & lastCol
    FindFYTotalCol = lastCol
End Function

'===============================================================================
' FindProductMetric - Find a specific metric value for a product on P&L Trend
'===============================================================================
Private Function FindProductMetric(ByVal ws As Worksheet, _
                                     ByVal product As String, _
                                     ByVal metric As String, _
                                     ByVal srcLastRow As Long) As Double
    Dim inBlock As Boolean: inBlock = False
    Dim r As Long
    Dim products As Variant: products = modConfig.GetProducts()

    For r = 1 To srcLastRow
        Dim cellVal As String: cellVal = Trim(CStr(ws.Cells(r, 1).Value))

        If InStr(1, cellVal, product, vbTextCompare) > 0 And _
           InStr(1, cellVal, "Consolidated", vbTextCompare) = 0 Then
            inBlock = True
        End If

        If inBlock Then
            Dim p As Long
            For p = 0 To UBound(products)
                If CStr(products(p)) <> product Then
                    If InStr(1, cellVal, CStr(products(p)), vbTextCompare) > 0 Then
                        FindProductMetric = 0
                        Exit Function
                    End If
                End If
            Next p

            If InStr(1, cellVal, metric, vbTextCompare) > 0 Then
                Dim fyCol As Long: fyCol = FindFYTotalCol(ws)
                Dim fyVal As Double: fyVal = modConfig.SafeNum(ws.Cells(r, fyCol).Value)
                If fyVal <> 0 Then
                    FindProductMetric = fyVal
                    Exit Function
                End If
                Dim lc As Long: lc = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
                Dim c As Long
                For c = 2 To lc
                    If IsNumeric(ws.Cells(r, c).Value) And ws.Cells(r, c).Value <> 0 Then
                        FindProductMetric = modConfig.SafeNum(ws.Cells(r, c).Value)
                        Exit Function
                    End If
                Next c
            End If
        End If
    Next r

    FindProductMetric = 0
End Function
