Attribute VB_Name = "modDashboard"
Option Explicit

'===============================================================================
' modDashboard - Dynamic Dashboard & Chart Generation
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Build visual dashboards and charts from P&L data.
'           Includes both quick Report--> charts and advanced analysis views.
'
' PUBLIC SUBS:
'   BuildDashboard              - 3 charts on Report--> (revenue, margin, mix)
'   RefreshDashboard            - Recalculate and refresh existing charts
'   CreateExecutiveDashboard    - KPI cards + summary table on dedicated sheet
'   WaterfallChart              - Revenue-to-Net-Income waterfall bridge
'   ProductComparison           - Side-by-side product metrics + ranking
'
' VERSION:  2.1.0
' CHANGES:  v2.0 -> v2.1:
'           + ISSUE-009: Added CreateExecutiveDashboard (Action #37 equivalent)
'           + Added WaterfallChart with invisible-base stacked bar technique
'           + Added ProductComparison with metrics table, ranking, and chart
'           + Added helpers: FindFYTotalCol, FindProductMetric
'           + Existing BuildDashboard and its 3 chart helpers unchanged
'
' PRIOR FIXES (v2.0):
'   BUG-003  Pie chart lastCol uses HDR_ROW_REPORT (row 4) not row 1
'   BUG-010  Revenue/margin charts dynamically detect last month with data
'===============================================================================

'===============================================================================
' BuildDashboard - Create all dashboard charts on Report--> sheet
'===============================================================================
Public Sub BuildDashboard()
    On Error GoTo ErrHandler
    
    If Not modConfig.SheetExists(SH_PL_TREND) Then
        MsgBox "P&L Trend sheet not found. Cannot build dashboard.", vbCritical, APP_NAME
        Exit Sub
    End If
    
    If MsgBox("Build dashboard charts on the Report--> sheet?" & vbCrLf & _
              "This will add revenue trend, margin, and product mix charts.", _
              vbYesNo + vbQuestion, APP_NAME) = vbNo Then Exit Sub
    
    modPerformance.TurboOn
    modPerformance.UpdateStatus "Building dashboard...", 0
    
    Dim wsReport As Worksheet
    If modConfig.SheetExists(SH_REPORT) Then
        Set wsReport = ThisWorkbook.Worksheets(SH_REPORT)
    Else
        MsgBox "Report--> sheet not found.", vbCritical, APP_NAME
        modPerformance.TurboOff
        Exit Sub
    End If
    
    ' Clear existing charts on the report sheet
    Dim chtObj As ChartObject
    For Each chtObj In wsReport.ChartObjects
        chtObj.Delete
    Next chtObj
    
    Dim wsTrend As Worksheet: Set wsTrend = ThisWorkbook.Worksheets(SH_PL_TREND)
    
    ' Find data positions on the P&L Trend sheet
    Dim products As Variant: products = modConfig.GetProducts()
    Dim mths As Variant: mths = modConfig.GetMonths()
    
    ' Detect how many months actually have data (not hardcoded 12).
    Dim lastDataMonthCol As Long
    lastDataMonthCol = DetectLastMonthWithData(wsTrend, products)
    If lastDataMonthCol < 2 Then lastDataMonthCol = 13  ' Fallback to all 12
    
    ' Build month label array for only populated months
    Dim monthCount As Long: monthCount = lastDataMonthCol - 1
    Dim monthLabels() As String
    ReDim monthLabels(0 To monthCount - 1)
    Dim m As Long
    For m = 0 To monthCount - 1
        monthLabels(m) = CStr(mths(m))
    Next m
    
    ' Chart 1: Revenue Trend by Product (Line Chart)
    modPerformance.UpdateStatus "Creating revenue trend chart...", 0.3
    CreateRevenueTrendChart wsReport, wsTrend, products, monthLabels, lastDataMonthCol
    
    ' Chart 2: Contribution Margin Trend (Line Chart)
    modPerformance.UpdateStatus "Creating margin trend chart...", 0.6
    CreateMarginTrendChart wsReport, wsTrend, products, monthLabels, lastDataMonthCol
    
    ' Chart 3: Product Revenue Mix (Pie Chart) - using Year Total
    modPerformance.UpdateStatus "Creating product mix chart...", 0.9
    CreateProductMixChart wsReport, wsTrend, products
    
    modPerformance.TurboOff
    wsReport.Activate
    
    modLogger.LogAction "modDashboard", "BuildDashboard", _
                        "3 charts created (" & monthCount & " months of data)", _
                        modPerformance.ElapsedSeconds()
    
    MsgBox "Dashboard built with 3 charts on '" & SH_REPORT & "' sheet." & vbCrLf & _
           "Showing " & monthCount & " months of actual data.", _
           vbInformation, APP_NAME
    Exit Sub
    
ErrHandler:
    modPerformance.TurboOff
    MsgBox "Dashboard error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' RefreshDashboard - Update existing charts with current data
'===============================================================================
Public Sub RefreshDashboard()
    Application.Calculate
    MsgBox "Dashboard refreshed with latest data.", vbInformation, APP_NAME
End Sub

'===============================================================================
'
' ===  ADVANCED DASHBOARDS (v2.1 — ISSUE-009)  ===============================
'
'===============================================================================

'===============================================================================
' CreateExecutiveDashboard - Full executive dashboard on dedicated sheet
' Creates KPI cards (Revenue, GM%, OpEx, Net Income), trend arrows,
' and Actual vs Budget summary table with variance coloring.
' Ported from legacy T3B1 #24.
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
    Dim tLastRow As Long: tLastRow = wsTrend.Cells(wsTrend.Rows.Count, 1).End(xlUp).Row
    Dim tLastCol As Long: tLastCol = wsTrend.Cells(1, wsTrend.Columns.Count).End(xlToLeft).Column
    
    Dim fyCol As Long: fyCol = FindFYTotalCol(wsTrend)
    
    ' Find budget column
    Dim budCol As Long: budCol = 0
    Dim c As Long
    For c = 2 To tLastCol
        If InStr(LCase(CStr(wsTrend.Cells(1, c).Value)), "budget") > 0 Then budCol = c: Exit For
    Next c
    If budCol = 0 Then budCol = tLastCol
    
    ' Find key P&L summary rows
    Dim revRow As Long, gpRow As Long, opexRow As Long, niRow As Long
    Dim r As Long
    For r = 2 To tLastRow
        Dim lbl As String: lbl = LCase(Trim(CStr(wsTrend.Cells(r, 1).Value)))
        If revRow = 0 And InStr(lbl, "total revenue") > 0 Then revRow = r
        If gpRow = 0 And InStr(lbl, "gross profit") > 0 Then gpRow = r
        If opexRow = 0 And InStr(lbl, "total operating") > 0 And InStr(lbl, "expense") > 0 Then opexRow = r
        If niRow = 0 And (InStr(lbl, "net income") > 0 Or InStr(lbl, "net operating") > 0) Then niRow = r
    Next r
    If revRow = 0 Then revRow = 2   ' Fallback
    
    Dim fyRev As Double: fyRev = modConfig.SafeNum(wsTrend.Cells(revRow, fyCol).Value)
    Dim budRev As Double: budRev = modConfig.SafeNum(wsTrend.Cells(revRow, budCol).Value)
    Dim fyGP As Double: If gpRow > 0 Then fyGP = modConfig.SafeNum(wsTrend.Cells(gpRow, fyCol).Value)
    Dim fyOpex As Double: If opexRow > 0 Then fyOpex = modConfig.SafeNum(wsTrend.Cells(opexRow, fyCol).Value)
    Dim fyNI As Double: If niRow > 0 Then fyNI = modConfig.SafeNum(wsTrend.Cells(niRow, fyCol).Value)
    
    Dim gmPct As Double: If fyRev <> 0 Then gmPct = fyGP / fyRev
    Dim revVar As Double: revVar = fyRev - budRev
    Dim revVarPct As Double: If budRev <> 0 Then revVarPct = revVar / Abs(budRev)
    
    '--- Create dashboard sheet ---
    Dim dashName As String: dashName = "Executive Dashboard"
    modConfig.SafeDeleteSheet dashName
    
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsDash.Name = dashName
    wsDash.Cells.Interior.Color = RGB(245, 246, 250)
    
    ' Title
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
    
    '--- KPI Cards as cell-based layout ---
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
        
        ' Label
        wsDash.Cells(kpiRow, kCol).Value = kpiLabels(k)
        wsDash.Cells(kpiRow, kCol).Font.Size = 9
        wsDash.Cells(kpiRow, kCol).Font.Color = RGB(100, 100, 100)
        
        ' Value
        wsDash.Cells(kpiRow + 1, kCol).Value = kpiValues(k)
        wsDash.Cells(kpiRow + 1, kCol).NumberFormat = kpiFmts(k)
        wsDash.Cells(kpiRow + 1, kCol).Font.Size = 22
        wsDash.Cells(kpiRow + 1, kCol).Font.Bold = True
        wsDash.Cells(kpiRow + 1, kCol).Font.Color = RGB(0, 51, 102)
        
        ' Trend arrow / context
        Dim arrow As String
        If k = 0 Then
            arrow = IIf(revVar >= 0, Chr(9650) & " +" & Format(revVarPct, "0.0%"), _
                                     Chr(9660) & " " & Format(revVarPct, "0.0%"))
            wsDash.Cells(kpiRow + 2, kCol).Font.Color = IIf(revVar >= 0, RGB(0, 128, 0), RGB(192, 0, 0))
        Else
            arrow = "vs Budget"
            wsDash.Cells(kpiRow + 2, kCol).Font.Color = RGB(128, 128, 128)
        End If
        wsDash.Cells(kpiRow + 2, kCol).Value = arrow
        wsDash.Cells(kpiRow + 2, kCol).Font.Size = 9
        
        ' Top accent bar via merged cell color
        wsDash.Cells(kpiRow - 1, kCol).Interior.Color = kpiColors(k)
        wsDash.Range(wsDash.Cells(kpiRow - 1, kCol), wsDash.Cells(kpiRow - 1, kCol + 1)).Merge
        wsDash.Range(wsDash.Cells(kpiRow - 1, kCol), wsDash.Cells(kpiRow - 1, kCol + 1)).RowHeight = 5
    Next k
    
    '--- Summary Table: Actual vs Budget ---
    modPerformance.UpdateStatus "Building summary table...", 0.7
    Dim tblRow As Long: tblRow = kpiRow + 5
    
    modConfig.StyleHeader wsDash, "", tblRow, _
        Array("Metric", "FY Actual", "Budget", "Variance $", "Variance %")
    
    Dim metrics As Variant: metrics = Array("Total Revenue", "Gross Profit", "Operating Expenses", "Net Income")
    Dim actVals As Variant: actVals = Array(fyRev, fyGP, fyOpex, fyNI)
    
    ' Map metrics to their source rows for budget lookup
    Dim metricRows As Variant
    If gpRow > 0 And opexRow > 0 And niRow > 0 Then
        metricRows = Array(revRow, gpRow, opexRow, niRow)
    Else
        metricRows = Array(revRow, revRow, revRow, revRow)
    End If
    
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
        
        ' Color variance cells
        If mVar > 0 Then
            wsDash.Cells(mRow, 4).Font.Color = RGB(0, 128, 0)
            wsDash.Cells(mRow, 5).Font.Color = RGB(0, 128, 0)
        ElseIf mVar < 0 Then
            wsDash.Cells(mRow, 4).Font.Color = RGB(192, 0, 0)
            wsDash.Cells(mRow, 5).Font.Color = RGB(192, 0, 0)
        End If
    Next mr
    
    wsDash.Columns("A:L").AutoFit
    wsDash.Tab.Color = RGB(0, 51, 102)
    wsDash.Activate
    
    Dim elapsed As Double: elapsed = modPerformance.ElapsedSeconds()
    modPerformance.TurboOff
    
    modLogger.LogAction "modDashboard", "CreateExecutiveDashboard", _
        "4 KPIs + summary table", elapsed
    MsgBox "Executive Dashboard created on '" & dashName & "'.", vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modDashboard", "ERROR-ExecDash", Err.Description
    MsgBox "Executive dashboard error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' WaterfallChart - Revenue-to-Net-Income waterfall bridge
' Uses stacked bar with invisible base segment (classic Excel waterfall technique).
' Creates a dedicated "P&L Waterfall" sheet.
' Ported from legacy T3B1 #25.
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
    
    ' Find key P&L rows
    Dim revRow As Long, cogsRow As Long, gpRow As Long, opexRow As Long, niRow As Long
    Dim r2 As Long
    For r2 = 2 To lastRow
        Dim lb2 As String: lb2 = LCase(Trim(CStr(wsSrc.Cells(r2, 1).Value)))
        If revRow = 0 And InStr(lb2, "total revenue") > 0 Then revRow = r2
        If cogsRow = 0 And (InStr(lb2, "cost of") > 0 Or InStr(lb2, "cogs") > 0) Then cogsRow = r2
        If gpRow = 0 And InStr(lb2, "gross profit") > 0 Then gpRow = r2
        If opexRow = 0 And InStr(lb2, "total operating") > 0 Then opexRow = r2
        If niRow = 0 And InStr(lb2, "net income") > 0 Then niRow = r2
    Next r2
    
    If revRow = 0 Then
        modPerformance.TurboOff
        MsgBox "Could not find Revenue row on P&L Trend.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' Build waterfall data (5 stages: Revenue -> COGS -> GP -> OpEx -> NI)
    Dim wfLabels As Variant
    wfLabels = Array("Revenue", "COGS", "Gross Profit", "OpEx", "Net Income")
    Dim wfValues(0 To 4) As Double
    wfValues(0) = modConfig.SafeNum(wsSrc.Cells(revRow, fyCol).Value)
    If cogsRow > 0 Then wfValues(1) = modConfig.SafeNum(wsSrc.Cells(cogsRow, fyCol).Value)
    If gpRow > 0 Then wfValues(2) = modConfig.SafeNum(wsSrc.Cells(gpRow, fyCol).Value)
    If opexRow > 0 Then wfValues(3) = modConfig.SafeNum(wsSrc.Cells(opexRow, fyCol).Value)
    If niRow > 0 Then wfValues(4) = modConfig.SafeNum(wsSrc.Cells(niRow, fyCol).Value)
    
    ' Create waterfall on a dedicated sheet
    Dim wfName As String: wfName = "P&L Waterfall"
    modConfig.SafeDeleteSheet wfName
    
    Dim wsWF As Worksheet
    Set wsWF = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsWF.Name = wfName
    
    ' Write data table for the stacked bar chart
    wsWF.Range("A1").Value = "P&L WATERFALL - FY" & FISCAL_YEAR_4
    wsWF.Range("A1").Font.Size = 14
    wsWF.Range("A1").Font.Bold = True
    
    ' Invisible base + visible bar columns (stacked bar technique)
    wsWF.Cells(3, 1).Value = "Category"
    wsWF.Cells(3, 2).Value = "Base"
    wsWF.Cells(3, 3).Value = "Increase"
    wsWF.Cells(3, 4).Value = "Decrease"
    
    ' Calculate base/increase/decrease for each waterfall stage:
    '   Revenue:    Full positive bar (no base)
    '   COGS:       Decrease from Revenue (base = Revenue - |COGS|)
    '   Gross Profit: Full positive bar (total = GP)
    '   OpEx:       Decrease from GP (base = GP - |OpEx|)
    '   Net Income: Full positive bar (total = NI)
    Dim wfBase(0 To 4) As Double
    Dim wfInc(0 To 4) As Double
    Dim wfDec(0 To 4) As Double
    
    wfBase(0) = 0: wfInc(0) = wfValues(0): wfDec(0) = 0                            ' Revenue = full bar
    wfBase(1) = wfValues(0) + wfValues(1): wfInc(1) = 0: wfDec(1) = Abs(wfValues(1)) ' COGS = decrease
    wfBase(2) = 0: wfInc(2) = wfValues(2): wfDec(2) = 0                            ' GP = total bar
    wfBase(3) = wfValues(2) + wfValues(3): wfInc(3) = 0: wfDec(3) = Abs(wfValues(3)) ' OpEx = decrease
    wfBase(4) = 0: wfInc(4) = wfValues(4): wfDec(4) = 0                            ' NI = total bar
    
    Dim wi As Long
    For wi = 0 To 4
        wsWF.Cells(4 + wi, 1).Value = wfLabels(wi)
        wsWF.Cells(4 + wi, 2).Value = wfBase(wi)
        wsWF.Cells(4 + wi, 3).Value = wfInc(wi)
        wsWF.Cells(4 + wi, 4).Value = wfDec(wi)
    Next wi
    
    ' Create stacked bar chart
    Dim co As ChartObject
    Set co = wsWF.ChartObjects.Add(20, 120, 500, 320)
    co.Name = "WaterfallChart"
    
    With co.Chart
        .ChartType = xlBarStacked
        .SetSourceData Source:=wsWF.Range("A3:D8")
        .HasTitle = True
        .ChartTitle.Text = "P&L Waterfall - FY" & FISCAL_YEAR_4
        
        ' Make base series invisible (key waterfall trick)
        On Error Resume Next
        .SeriesCollection(1).Format.Fill.Visible = msoFalse
        .SeriesCollection(1).Format.Line.Visible = msoFalse
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(0, 176, 80)   ' Increases = green
        .SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(192, 0, 0)    ' Decreases = red
        On Error GoTo ErrHandler
        
        .HasLegend = False
    End With
    
    wsWF.Tab.Color = RGB(0, 112, 192)
    wsWF.Activate
    
    Dim elapsed As Double: elapsed = modPerformance.ElapsedSeconds()
    modPerformance.TurboOff
    
    modLogger.LogAction "modDashboard", "WaterfallChart", "5 P&L stages", elapsed
    MsgBox "Waterfall chart created on '" & wfName & "'.", vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modDashboard", "ERROR-Waterfall", Err.Description
    MsgBox "Waterfall error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' ProductComparison - Side-by-side product line comparison table + chart
' Creates a dedicated "Product Comparison" sheet with metrics, margin %,
' revenue ranking, and clustered column chart.
' Ported from legacy T3B1 #26.
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
    
    ' Create comparison sheet
    Dim pcName As String: pcName = "Product Comparison"
    modConfig.SafeDeleteSheet pcName
    
    Dim wsPC As Worksheet
    Set wsPC = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsPC.Name = pcName
    
    ' Build header row with product names
    Dim hdrs() As String
    ReDim hdrs(0 To UBound(products) + 1)
    hdrs(0) = "Metric"
    Dim p As Long
    For p = 0 To UBound(products)
        hdrs(p + 1) = CStr(products(p))
    Next p
    
    modConfig.StyleHeader wsPC, "PRODUCT LINE COMPARISON - FY" & FISCAL_YEAR_4, 4, hdrs
    
    ' Metrics to compare (row labels)
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
    
    ' Add margin % row
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
    
    ' Revenue ranking row
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
    
    ' Create comparison bar chart
    modPerformance.UpdateStatus "Creating comparison chart...", 0.7
    
    Dim co As ChartObject
    Set co = wsPC.ChartObjects.Add(20, (outRow + 2) * 15, 500, 280)
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
    
    ' Format columns
    wsPC.Columns("A").ColumnWidth = 22
    Dim fc As Long
    For fc = 2 To UBound(products) + 2
        wsPC.Columns(fc).ColumnWidth = 16
    Next fc
    wsPC.Tab.Color = RGB(0, 176, 80)
    wsPC.Activate
    
    Dim elapsed As Double: elapsed = modPerformance.ElapsedSeconds()
    modPerformance.TurboOff
    
    modLogger.LogAction "modDashboard", "ProductComparison", _
        (UBound(products) + 1) & " products compared", elapsed
    MsgBox "Product Comparison created on '" & pcName & "'.", vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modDashboard", "ERROR-ProductComp", Err.Description
    MsgBox "Product comparison error: " & Err.Description, vbCritical, APP_NAME
End Sub


'===============================================================================
'
' ===  PRIVATE HELPERS — BuildDashboard Charts  ================================
'
'===============================================================================

'===============================================================================
' DetectLastMonthWithData - Find rightmost month column (B-M) with non-zero data
'===============================================================================
Private Function DetectLastMonthWithData(ByVal wsSrc As Worksheet, _
                                          ByVal products As Variant) As Long
    Dim maxCol As Long: maxCol = 2  ' At minimum, Jan
    Dim p As Long, c As Long
    
    For p = 0 To UBound(products)
        Dim revRow As Long
        revRow = FindProductRevenueRow(wsSrc, CStr(products(p)))
        If revRow > 0 Then
            For c = 13 To 2 Step -1
                If modConfig.SafeNum(wsSrc.Cells(revRow, c).Value) <> 0 Then
                    If c > maxCol Then maxCol = c
                    Exit For
                End If
            Next c
        End If
    Next p
    
    DetectLastMonthWithData = maxCol
End Function

'===============================================================================
' CreateRevenueTrendChart - Monthly revenue by product (dynamic month range)
'===============================================================================
Private Sub CreateRevenueTrendChart(ByVal wsTarget As Worksheet, _
                                     ByVal wsSrc As Worksheet, _
                                     ByVal products As Variant, _
                                     ByRef monthLabels() As String, _
                                     ByVal lastMonthCol As Long)
    Dim cht As ChartObject
    Set cht = wsTarget.ChartObjects.Add(Left:=400, Top:=20, Width:=520, Height:=300)
    
    With cht.Chart
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Monthly Revenue by Product - FY" & FISCAL_YEAR_4
        .ChartTitle.Font.Size = 11
        .ChartTitle.Font.Name = "Calibri"
        
        Dim p As Long
        For p = 0 To UBound(products)
            Dim revRow As Long
            revRow = FindProductRevenueRow(wsSrc, CStr(products(p)))
            
            If revRow > 0 Then
                Dim ser As Series
                Set ser = .SeriesCollection.NewSeries
                ser.Name = CStr(products(p))
                ser.Values = wsSrc.Range(wsSrc.Cells(revRow, 2), wsSrc.Cells(revRow, lastMonthCol))
                ser.XValues = monthLabels
            End If
        Next p
        
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 9
        
        Dim ax As Axis
        Set ax = .Axes(xlValue)
        ax.HasTitle = True
        ax.AxisTitle.Text = "Revenue ($)"
        ax.AxisTitle.Font.Size = 9
        ax.TickLabels.NumberFormat = "$#,##0"
        
        .PlotArea.Interior.Color = CLR_WHITE
    End With
End Sub

'===============================================================================
' CreateMarginTrendChart - CM% trend by product (dynamic month range)
'===============================================================================
Private Sub CreateMarginTrendChart(ByVal wsTarget As Worksheet, _
                                    ByVal wsSrc As Worksheet, _
                                    ByVal products As Variant, _
                                    ByRef monthLabels() As String, _
                                    ByVal lastMonthCol As Long)
    Dim cht As ChartObject
    Set cht = wsTarget.ChartObjects.Add(Left:=400, Top:=340, Width:=520, Height:=300)
    
    With cht.Chart
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Contribution Margin % Trend - FY" & FISCAL_YEAR_4
        .ChartTitle.Font.Size = 11
        .ChartTitle.Font.Name = "Calibri"
        
        Dim p As Long
        For p = 0 To UBound(products)
            Dim cmRow As Long
            cmRow = FindProductCMPctRow(wsSrc, CStr(products(p)))
            
            If cmRow > 0 Then
                Dim ser As Series
                Set ser = .SeriesCollection.NewSeries
                ser.Name = CStr(products(p))
                ser.Values = wsSrc.Range(wsSrc.Cells(cmRow, 2), wsSrc.Cells(cmRow, lastMonthCol))
                ser.XValues = monthLabels
            End If
        Next p
        
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        
        Dim ax As Axis
        Set ax = .Axes(xlValue)
        ax.HasTitle = True
        ax.AxisTitle.Text = "CM %"
        ax.AxisTitle.Font.Size = 9
        ax.TickLabels.NumberFormat = "0%"
        
        .PlotArea.Interior.Color = CLR_WHITE
    End With
End Sub

'===============================================================================
' CreateProductMixChart - Revenue share pie chart (using Year Total column)
'===============================================================================
Private Sub CreateProductMixChart(ByVal wsTarget As Worksheet, _
                                   ByVal wsSrc As Worksheet, _
                                   ByVal products As Variant)
    Dim cht As ChartObject
    Set cht = wsTarget.ChartObjects.Add(Left:=940, Top:=20, Width:=350, Height:=300)
    
    With cht.Chart
        .ChartType = xlPie
        .HasTitle = True
        .ChartTitle.Text = "FY" & FISCAL_YEAR_4 & " Revenue Mix"
        .ChartTitle.Font.Size = 11
        .ChartTitle.Font.Name = "Calibri"
        
        ' Use HDR_ROW_REPORT (row 4) for last column detection (BUG-003 fix)
        Dim lastCol As Long: lastCol = modConfig.LastCol(wsSrc, HDR_ROW_REPORT)
        
        Dim revValues(0 To 3) As Double
        Dim prodNames(0 To 3) As String
        Dim p As Long
        For p = 0 To UBound(products)
            prodNames(p) = CStr(products(p))
            Dim revRow As Long
            revRow = FindProductRevenueRow(wsSrc, prodNames(p))
            If revRow > 0 Then
                revValues(p) = modConfig.SafeNum(wsSrc.Cells(revRow, lastCol).Value)
            End If
        Next p
        
        Dim ser As Series
        Set ser = .SeriesCollection.NewSeries
        ser.Values = revValues
        ser.XValues = prodNames
        
        ' Data labels
        ser.HasDataLabels = True
        ser.DataLabels.ShowPercentage = True
        ser.DataLabels.ShowValue = False
        ser.DataLabels.Font.Size = 10
        
        ' Custom product colors
        If ser.Points.Count >= 4 Then
            ser.Points(1).Interior.Color = RGB(31, 78, 121)   ' Navy - iGO
            ser.Points(2).Interior.Color = RGB(68, 114, 196)  ' Blue - Affirm
            ser.Points(3).Interior.Color = RGB(112, 173, 71)  ' Green - InsureSight
            ser.Points(4).Interior.Color = RGB(237, 125, 49)  ' Orange - DocFast
        End If
        
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With
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
' Searches row 1 for headers containing "FY" + "Total" or "FY2025" etc.
' Falls back to last column.
'===============================================================================
Private Function FindFYTotalCol(ByVal ws As Worksheet) As Long
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    
    ' First pass: look for "FY Total" or "FY2025 Total" style headers
    For c = 2 To lastCol
        Dim hdr As String: hdr = LCase(Trim(CStr(ws.Cells(1, c).Value)))
        If InStr(hdr, "fy") > 0 And InStr(hdr, "total") > 0 Then
            FindFYTotalCol = c: Exit Function
        End If
    Next c
    
    ' Second pass: look for "FY2025" or FISCAL_YEAR_4 pattern
    For c = 2 To lastCol
        hdr = LCase(Trim(CStr(ws.Cells(1, c).Value)))
        If InStr(hdr, "fy" & FISCAL_YEAR_4) > 0 Then
            FindFYTotalCol = c: Exit Function
        End If
        If InStr(hdr, FISCAL_YEAR_4 & " total") > 0 Then
            FindFYTotalCol = c: Exit Function
        End If
    Next c
    
    ' Fallback: last column
    FindFYTotalCol = lastCol
End Function

'===============================================================================
' FindProductMetric - Find a specific metric value for a product on P&L Trend
' Searches for the product block, then the metric label within that block.
' Returns the first non-zero numeric value found in the row.
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
        
        ' Detect product block start
        If InStr(1, cellVal, product, vbTextCompare) > 0 And _
           InStr(1, cellVal, "Consolidated", vbTextCompare) = 0 Then
            inBlock = True
        End If
        
        ' Detect product block end (another product header or empty row after data)
        If inBlock Then
            Dim p As Long
            For p = 0 To UBound(products)
                If CStr(products(p)) <> product Then
                    If InStr(1, cellVal, CStr(products(p)), vbTextCompare) > 0 Then
                        ' Entered a different product block — stop
                        FindProductMetric = 0
                        Exit Function
                    End If
                End If
            Next p
            
            ' Check if this row has our metric
            If InStr(1, cellVal, metric, vbTextCompare) > 0 Then
                ' Return FY Total column value, or first non-zero numeric cell
                Dim fyCol As Long: fyCol = FindFYTotalCol(ws)
                Dim fyVal As Double: fyVal = modConfig.SafeNum(ws.Cells(r, fyCol).Value)
                If fyVal <> 0 Then
                    FindProductMetric = fyVal
                    Exit Function
                End If
                ' Fallback: scan for first non-zero value
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


'===============================================================================
'
' ===  NEW DASHBOARD TOOLS (v2.1 — from NewTesting ideas #44, #86)  ============
'
'===============================================================================

'===============================================================================
' LinkDynamicChartTitles - Link all chart titles to a selector cell (#44)
' Loops through every ChartObject on the Report--> sheet and rewrites the
' chart title to include the current month value from the FPL Dynamic sheet
' selector (cell B4). Run this after changing the month dropdown so all
' chart titles update instantly without manual editing.
'===============================================================================
Public Sub LinkDynamicChartTitles()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_REPORT) Then
        MsgBox "Report--> sheet not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    ' Read the selected month from FPL Summary - Dynamic B4 if it exists
    Dim selectedMonth As String: selectedMonth = ""
    If modConfig.SheetExists("FPL Summary - Dynamic") Then
        selectedMonth = modConfig.SafeStr( _
            ThisWorkbook.Worksheets("FPL Summary - Dynamic").Range("B4").Value)
    End If
    If Len(selectedMonth) = 0 Then
        selectedMonth = Format(Date, "mmm")  ' Fall back to current month abbreviation
    End If

    Dim wsReport As Worksheet: Set wsReport = ThisWorkbook.Worksheets(SH_REPORT)
    Dim updateCount As Long: updateCount = 0

    Dim co As ChartObject
    For Each co In wsReport.ChartObjects
        On Error Resume Next
        If co.Chart.HasTitle Then
            Dim oldTitle As String: oldTitle = co.Chart.ChartTitle.Text
            ' Append month to title if not already present
            If InStr(oldTitle, selectedMonth) = 0 Then
                ' Replace existing month abbreviation or append
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

    modLogger.LogAction "modDashboard", "LinkDynamicChartTitles", _
        updateCount & " chart title(s) updated to " & selectedMonth
    MsgBox updateCount & " chart title(s) updated to show: " & selectedMonth, _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    MsgBox "LinkDynamicChartTitles error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' CreateSmallMultiplesGrid - Generate one small chart per product (#86)
' Creates a dedicated "Product Small Multiples" sheet with 4 small line charts
' arranged in a 2x2 grid, one per product. Each chart shows that product's
' monthly revenue across all populated months. A CFO-level visual that lets
' you compare all 4 products at the same scale in one glance.
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

    ' Detect last populated month
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

    ' Create output sheet
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

    ' Product colors for visual variety
    Dim prodColors As Variant
    prodColors = Array(RGB(31, 78, 121), RGB(68, 114, 196), _
                       RGB(112, 173, 71), RGB(237, 125, 49))

    ' Grid layout: 2 columns x 2 rows of charts
    ' Each chart: width=350, height=200, margins: left=20/400, top=60/280
    Dim chartLeft  As Variant: chartLeft  = Array(20, 390, 20, 390)
    Dim chartTop   As Variant: chartTop   = Array(60, 60, 280, 280)
    Dim chartW     As Long: chartW = 355
    Dim chartH     As Long: chartH = 205

    modPerformance.UpdateStatus "Creating 4 product charts...", 0.4

    Dim p As Long
    For p = 0 To Application.Min(3, UBound(products))
        Dim productName As String: productName = CStr(products(p))

        ' Find revenue row for this product
        Dim pRevRow As Long: pRevRow = FindProductRevenueRow(wsSrc, productName)
        If pRevRow = 0 Then GoTo NextProduct

        ' Create chart
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
    modLogger.LogAction "modDashboard", "CreateSmallMultiplesGrid", _
        UBound(products) + 1 & " product charts | " & monthCount & " months", elapsed
    MsgBox "Small multiples grid created on '" & smName & "'." & vbCrLf & _
           "4 product revenue charts at the same scale for easy comparison.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    MsgBox "CreateSmallMultiplesGrid error: " & Err.Description, vbCritical, APP_NAME
End Sub
