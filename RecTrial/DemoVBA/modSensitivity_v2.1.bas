Attribute VB_Name = "modSensitivity"
Option Explicit

'===============================================================================
' modSensitivity - What-If Sensitivity Analysis
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Reads driver values from the Assumptions sheet, varies each one
'           by +/-10% and +/-20%, and shows the impact on Total Revenue and
'           Contribution Margin. Outputs a formatted "Sensitivity Analysis"
'           sheet with a tornado-style summary table.
'
' PUBLIC SUBS:
'   RunSensitivityAnalysis - Main entry (Action #5)
'
' DEPENDENCIES: modConfig, modPerformance, modLogger
' VERSION:  2.1.0
'===============================================================================

'===============================================================================
' RunSensitivityAnalysis - Main entry point
'===============================================================================
Public Sub RunSensitivityAnalysis()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_ASSUMPTIONS) Then
        MsgBox "Assumptions sheet not found.", vbCritical, APP_NAME
        Exit Sub
    End If
    If Not modConfig.SheetExists(SH_PL_TREND) Then
        MsgBox "P&L - Monthly Trend sheet not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Running sensitivity analysis...", 0.05

    ' Read drivers from Assumptions
    Dim wsA As Worksheet: Set wsA = ThisWorkbook.Worksheets(SH_ASSUMPTIONS)
    Dim lastRow As Long: lastRow = modConfig.LastRow(wsA, 1)

    If lastRow < DATA_ROW_ASSUME Then
        modPerformance.TurboOff
        MsgBox "No drivers found on Assumptions sheet.", vbInformation, APP_NAME
        Exit Sub
    End If

    ' Collect driver names and values
    Dim driverCount As Long: driverCount = 0
    Dim driverNames() As String
    Dim driverValues() As Double
    Dim driverRows() As Long
    ReDim driverNames(1 To lastRow)
    ReDim driverValues(1 To lastRow)
    ReDim driverRows(1 To lastRow)

    Dim r As Long
    For r = DATA_ROW_ASSUME To lastRow
        Dim dName As String: dName = Trim(CStr(wsA.Cells(r, 1).Value))
        Dim dVal As Double: dVal = modConfig.SafeNum(wsA.Cells(r, 2).Value)
        If dName <> "" And dVal <> 0 Then
            driverCount = driverCount + 1
            driverNames(driverCount) = dName
            driverValues(driverCount) = dVal
            driverRows(driverCount) = r
        End If
    Next r

    If driverCount = 0 Then
        modPerformance.TurboOff
        MsgBox "No valid drivers with non-zero values found.", vbInformation, APP_NAME
        Exit Sub
    End If

    ' Read baseline Total Revenue from P&L Trend (FY Total column)
    Dim wsTrend As Worksheet: Set wsTrend = ThisWorkbook.Worksheets(SH_PL_TREND)
    Dim trendLastCol As Long: trendLastCol = modConfig.LastCol(wsTrend, HDR_ROW_REPORT)
    Dim fyCol As Long: fyCol = FindFYTotalCol(wsTrend, trendLastCol)

    Dim revRow As Long: revRow = modConfig.FindRowByLabel(wsTrend, "Total Revenue", DATA_ROW_REPORT)
    Dim cmRow As Long: cmRow = modConfig.FindRowByLabel(wsTrend, "Contribution Margin", DATA_ROW_REPORT)
    If cmRow = 0 Then cmRow = modConfig.FindRowByLabel(wsTrend, "Gross Margin", DATA_ROW_REPORT)

    Dim baseRevenue As Double: baseRevenue = 0
    Dim baseCM As Double: baseCM = 0
    If revRow > 0 Then baseRevenue = modConfig.SafeNum(wsTrend.Cells(revRow, fyCol).Value)
    If cmRow > 0 Then baseCM = modConfig.SafeNum(wsTrend.Cells(cmRow, fyCol).Value)

    ' Create output sheet
    modConfig.SafeDeleteSheet SH_SENSITIVITY
    Dim wsOut As Worksheet
    Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsOut.Name = SH_SENSITIVITY

    ' Title
    wsOut.Range("A1").Value = "SENSITIVITY ANALYSIS"
    wsOut.Range("A1").Font.Bold = True
    wsOut.Range("A1").Font.Size = 14
    wsOut.Range("A1").Font.Color = CLR_NAVY

    wsOut.Range("A2").Value = "Baseline Total Revenue: " & Format(baseRevenue, "$#,##0")
    If baseCM <> 0 Then
        wsOut.Range("A2").Value = wsOut.Range("A2").Value & "  |  Baseline CM: " & Format(baseCM, "$#,##0")
    End If
    wsOut.Range("A2").Font.Italic = True

    ' Header row
    modConfig.StyleHeader wsOut, 4, _
        Array("Driver", "Base Value", "-20%", "-10%", "+10%", "+20%", _
              "Range ($)", "Impact Rating")

    ' Calculate sensitivity for each driver
    Dim outRow As Long: outRow = 5
    Dim scenarios As Variant: scenarios = Array(-0.2, -0.1, 0.1, 0.2)
    Dim d As Long

    For d = 1 To driverCount
        modPerformance.UpdateStatus "Analyzing driver " & d & " of " & driverCount & "...", d / driverCount

        wsOut.Cells(outRow, 1).Value = driverNames(d)
        wsOut.Cells(outRow, 2).Value = driverValues(d)
        wsOut.Cells(outRow, 2).NumberFormat = "#,##0.00"

        ' For each scenario, estimate revenue impact proportionally
        ' If the driver is a % or rate (value < 10), scale impact against baseline revenue
        ' If the driver is a dollar amount (value >= 10), use direct proportional impact
        Dim driverShare As Double
        If Abs(driverValues(d)) < 10 Then
            ' Percentage/rate driver — impact = baseline revenue * driver value * scenario change
            driverShare = driverValues(d)
        Else
            ' Dollar driver — impact = driver value * scenario change * 12 (annualize)
            driverShare = (driverValues(d) * 12) / Application.Max(baseRevenue, 1)
        End If

        Dim minImpact As Double: minImpact = 0
        Dim maxImpact As Double: maxImpact = 0
        Dim sc As Long
        For sc = 0 To 3
            Dim pctChange As Double: pctChange = scenarios(sc)
            Dim impactEst As Double
            impactEst = baseRevenue * driverShare * pctChange
            wsOut.Cells(outRow, 3 + sc).Value = impactEst
            wsOut.Cells(outRow, 3 + sc).NumberFormat = "$#,##0;($#,##0)"

            ' Color code: positive = green, negative = red
            If impactEst >= 0 Then
                wsOut.Cells(outRow, 3 + sc).Font.Color = RGB(0, 128, 0)
            Else
                wsOut.Cells(outRow, 3 + sc).Font.Color = RGB(192, 0, 0)
            End If

            If impactEst < minImpact Then minImpact = impactEst
            If impactEst > maxImpact Then maxImpact = impactEst
        Next sc

        ' Range column
        Dim rangeVal As Double: rangeVal = maxImpact - minImpact
        wsOut.Cells(outRow, 7).Value = rangeVal
        wsOut.Cells(outRow, 7).NumberFormat = "$#,##0"

        ' Impact rating
        If rangeVal > baseRevenue * 0.05 Then
            wsOut.Cells(outRow, 8).Value = "HIGH"
            wsOut.Cells(outRow, 8).Font.Color = RGB(192, 0, 0)
            wsOut.Cells(outRow, 8).Font.Bold = True
        ElseIf rangeVal > baseRevenue * 0.01 Then
            wsOut.Cells(outRow, 8).Value = "MEDIUM"
            wsOut.Cells(outRow, 8).Font.Color = RGB(255, 165, 0)
        Else
            wsOut.Cells(outRow, 8).Value = "LOW"
            wsOut.Cells(outRow, 8).Font.Color = RGB(0, 128, 0)
        End If

        ' Alternate row shading
        If outRow Mod 2 = 1 Then
            wsOut.Range("A" & outRow & ":H" & outRow).Interior.Color = CLR_ALT_ROW
        End If

        outRow = outRow + 1
    Next d

    ' Format
    wsOut.Columns("A").ColumnWidth = 30
    wsOut.Columns("B").ColumnWidth = 14
    wsOut.Columns("C:F").ColumnWidth = 14
    wsOut.Columns("G").ColumnWidth = 14
    wsOut.Columns("H").ColumnWidth = 14
    wsOut.Tab.Color = RGB(255, 192, 0)
    wsOut.Activate

    Dim elapsed As Double: elapsed = modPerformance.ElapsedSeconds()
    modPerformance.TurboOff

    modLogger.LogAction "modSensitivity", "RunSensitivityAnalysis", _
        driverCount & " drivers analyzed in " & Format(elapsed, "0.0") & "s"

    MsgBox "Sensitivity Analysis Complete!" & vbCrLf & vbCrLf & _
           driverCount & " drivers analyzed." & vbCrLf & _
           "Results on '" & SH_SENSITIVITY & "' sheet.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modSensitivity", "ERROR", Err.Description
    MsgBox "Sensitivity analysis error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' FindFYTotalCol - Locate FY Total column on a trend sheet
'===============================================================================
Private Function FindFYTotalCol(ByVal ws As Worksheet, ByVal lastCol As Long) As Long
    Dim c As Long
    For c = 2 To lastCol
        Dim hdr As String: hdr = LCase(Trim(CStr(ws.Cells(HDR_ROW_REPORT, c).Value)))
        If InStr(hdr, "total") > 0 Then
            FindFYTotalCol = c: Exit Function
        End If
        If InStr(hdr, FISCAL_YEAR_4) > 0 Then
            FindFYTotalCol = c: Exit Function
        End If
    Next c
    FindFYTotalCol = lastCol  ' Fallback to last column
End Function
