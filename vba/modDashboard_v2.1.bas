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
'   ReformatChartsAndVisuals    - Reflow Charts & Visuals sheet into clean grid
'
' VERSION:  2.1.0
' SPLIT:    2026-03-05 — Advanced subs moved to modDashboardAdvanced_v2.1.bas:
'           CreateExecutiveDashboard, WaterfallChart, ProductComparison,
'           LinkDynamicChartTitles, CreateSmallMultiplesGrid + their helpers
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
    
    ' Place charts below all existing content on Report--> to avoid overlap
    Dim chartAnchorRow As Long
    Dim anchorCell As Range
    Set anchorCell = wsReport.Cells.Find(What:="*", SearchOrder:=xlByRows, _
                                          SearchDirection:=xlPrevious)
    If anchorCell Is Nothing Then
        chartAnchorRow = 1
    Else
        chartAnchorRow = anchorCell.row
    End If
    Dim chartTopStart As Long: chartTopStart = wsReport.Cells(chartAnchorRow + 2, 1).Top

    ' Chart 1: Revenue Trend by Product (Line Chart)
    modPerformance.UpdateStatus "Creating revenue trend chart...", 0.3
    CreateRevenueTrendChart wsReport, wsTrend, products, monthLabels, lastDataMonthCol, chartTopStart

    ' Chart 2: Contribution Margin Trend (Line Chart)
    modPerformance.UpdateStatus "Creating margin trend chart...", 0.6
    CreateMarginTrendChart wsReport, wsTrend, products, monthLabels, lastDataMonthCol, chartTopStart

    ' Chart 3: Product Revenue Mix (Pie Chart) - using Year Total
    modPerformance.UpdateStatus "Creating product mix chart...", 0.9
    CreateProductMixChart wsReport, wsTrend, products, chartTopStart
    
    modPerformance.TurboOff
    wsReport.Activate
    
    modLogger.LogAction "modDashboard", "BuildDashboard", _
                        "3 charts created (" & monthCount & " months of data) (" & Format(modPerformance.ElapsedSeconds(), "0.0") & "s)"
    
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

'
'===============================================================================


'===============================================================================
'
' ===  PRIVATE HELPERS — BuildDashboard Charts  ================================
'
'===============================================================================

' NOTE: CreateExecutiveDashboard, WaterfallChart, ProductComparison,
'       LinkDynamicChartTitles, CreateSmallMultiplesGrid and their helpers
'       (FindFYTotalCol, FindProductMetric) moved to modDashboardAdvanced_v2.1.bas
'       as of 2026-03-05.


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
                                     ByVal lastMonthCol As Long, _
                                     ByVal baseTop As Long)
    ' Row 1: Revenue chart (left) — 520x300
    Dim cht As ChartObject
    Set cht = wsTarget.ChartObjects.Add(Left:=20, Top:=baseTop, Width:=520, Height:=300)
    
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
                                    ByVal lastMonthCol As Long, _
                                    ByVal baseTop As Long)
    ' Row 1: Margin chart (right of Revenue) — 520x300
    Dim cht As ChartObject
    Set cht = wsTarget.ChartObjects.Add(Left:=560, Top:=baseTop, Width:=520, Height:=300)
    
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
                                   ByVal products As Variant, _
                                   ByVal baseTop As Long)
    ' Row 2: Pie chart centered below the two line charts — 400x300
    Dim cht As ChartObject
    Set cht = wsTarget.ChartObjects.Add(Left:=280, Top:=baseTop + 320, Width:=400, Height:=300)
    
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
'
' ===  CHARTS & VISUALS REFORMATTER  ==========================================
'
'===============================================================================

'===============================================================================
' ReformatChartsAndVisuals - Reflow all charts on "Charts & Visuals" into a
' clean, non-overlapping 2-column grid with consistent sizing, proper spacing,
' and visible chart titles/labels. Also cleans up any overlapping text boxes.
'
' Grid layout:
'   Row 1: [Chart 1]  [Chart 2]      (top = row 4)
'   Row 2: [Chart 3]  [Chart 4]      (top = row 4 + chartH + gap)
'   ...etc
'
' Each chart: 480w x 300h, 20px margin between columns, 30px between rows.
' Title row preserved at top with sheet heading and generated timestamp.
'===============================================================================
Public Sub ReformatChartsAndVisuals()
    On Error GoTo ErrHandler

    ' Find the Charts & Visuals sheet (try common name variants)
    Dim wsCV As Worksheet
    Dim shName As String: shName = ""
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, "Charts", vbTextCompare) > 0 And _
           InStr(1, ws.Name, "Visual", vbTextCompare) > 0 Then
            shName = ws.Name
            Exit For
        End If
    Next ws

    If Len(shName) = 0 Then
        MsgBox "No 'Charts & Visuals' sheet found in this workbook.", _
               vbExclamation, APP_NAME
        Exit Sub
    End If
    Set wsCV = ThisWorkbook.Worksheets(shName)

    Dim chartCount As Long: chartCount = wsCV.ChartObjects.Count
    If chartCount = 0 Then
        MsgBox "'" & shName & "' has no charts to reformat.", vbExclamation, APP_NAME
        Exit Sub
    End If

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Reformatting Charts & Visuals...", 0.1

    ' --- Grid constants ---
    Dim chartW As Long: chartW = 480       ' chart width
    Dim chartH As Long: chartH = 300       ' chart height
    Dim colGap As Long: colGap = 20        ' horizontal gap between charts
    Dim rowGap As Long: rowGap = 30        ' vertical gap between chart rows
    Dim gridCols As Long: gridCols = 2     ' 2 charts per row
    Dim leftMargin As Long: leftMargin = 10
    Dim topStart As Long: topStart = wsCV.Cells(4, 1).Top  ' below title rows

    ' --- Clean up title area ---
    wsCV.Range("A1").Value = "Charts & Visuals - FY" & FISCAL_YEAR_4
    wsCV.Range("A1").Font.Size = 16
    wsCV.Range("A1").Font.Bold = True
    wsCV.Range("A1").Font.Color = CLR_NAVY
    wsCV.Range("A2").Value = "Reformatted: " & Format(Now, "mmmm d, yyyy h:mm AM/PM")
    wsCV.Range("A2").Font.Italic = True
    wsCV.Range("A2").Font.Color = RGB(128, 128, 128)
    wsCV.Range("A2").Font.Size = 10

    ' --- Remove overlapping text boxes (Shapes that are not charts) ---
    Dim shp As Shape
    Dim shpIdx As Long
    For shpIdx = wsCV.Shapes.Count To 1 Step -1
        Set shp = wsCV.Shapes(shpIdx)
        If shp.Type = msoTextBox Or shp.Type = msoAutoShape Then
            ' Check if it overlaps any chart area — if so, delete it
            Dim co2 As ChartObject
            For Each co2 In wsCV.ChartObjects
                If shp.Top < (co2.Top + co2.Height) And _
                   (shp.Top + shp.Height) > co2.Top And _
                   shp.Left < (co2.Left + co2.Width) And _
                   (shp.Left + shp.Width) > co2.Left Then
                    shp.Delete
                    Exit For
                End If
            Next co2
        End If
    Next shpIdx

    ' --- Reflow charts into 2-column grid ---
    modPerformance.UpdateStatus "Repositioning " & chartCount & " charts...", 0.4
    Dim co As ChartObject
    Dim i As Long: i = 0
    For Each co In wsCV.ChartObjects
        Dim gridRow As Long: gridRow = i \ gridCols
        Dim gridCol As Long: gridCol = i Mod gridCols

        co.Left = leftMargin + gridCol * (chartW + colGap)
        co.Top = topStart + gridRow * (chartH + rowGap)
        co.Width = chartW
        co.Height = chartH

        ' Ensure chart title is visible and properly formatted
        On Error Resume Next
        If co.Chart.HasTitle Then
            co.Chart.ChartTitle.Font.Size = 11
            co.Chart.ChartTitle.Font.Name = "Calibri"
            co.Chart.ChartTitle.Font.Bold = True
        End If
        ' Ensure legend doesn't overlap plot area
        If co.Chart.HasLegend Then
            co.Chart.Legend.Position = xlLegendPositionBottom
            co.Chart.Legend.Font.Size = 9
        End If
        ' Clean up plot area
        co.Chart.PlotArea.Interior.Color = CLR_WHITE
        On Error GoTo ErrHandler

        i = i + 1
    Next co

    wsCV.Columns("A:L").AutoFit
    wsCV.Activate
    wsCV.Range("A1").Select

    Dim elapsed As Double: elapsed = modPerformance.ElapsedSeconds()
    modPerformance.TurboOff

    modLogger.LogAction "modDashboard", "ReformatChartsAndVisuals", _
        chartCount & " charts reflowed into 2-col grid (" & Format(elapsed, "0.0") & "s)"
    MsgBox chartCount & " charts on '" & shName & "' reformatted." & vbCrLf & vbCrLf & _
           "Layout: 2-column grid, " & chartW & "x" & chartH & "px each." & vbCrLf & _
           "All overlapping text removed. Titles and legends cleaned up.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    MsgBox "ReformatChartsAndVisuals error: " & Err.Description, vbCritical, APP_NAME
End Sub
