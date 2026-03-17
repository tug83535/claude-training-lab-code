Attribute VB_Name = "modForecast"
Option Explicit

'===============================================================================
' modForecast - Rolling Forecast & Trend Append
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Reads actual monthly P&L data from the trend sheet, calculates
'           a 3-month rolling average forecast for future months, and can
'           append a completed month's summary into the trend.
'
' PUBLIC SUBS:
'   RollingForecast  - Generate forecast for remaining months (Action #18)
'   AppendToTrend    - Copy a monthly summary into the P&L trend (Action #19)
'
' DEPENDENCIES: modConfig, modPerformance, modLogger
' VERSION:  2.1.0
'===============================================================================

Private Const FORECAST_WINDOW As Long = 3  ' 3-month rolling average

'===============================================================================
' RollingForecast - Calculate rolling average forecast for upcoming months
'===============================================================================
Public Sub RollingForecast()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_PL_TREND) Then
        MsgBox "P&L - Monthly Trend sheet not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Building rolling forecast...", 0.1

    Dim wsTrend As Worksheet: Set wsTrend = ThisWorkbook.Worksheets(SH_PL_TREND)
    Dim lastRow As Long: lastRow = modConfig.LastRow(wsTrend, 1)
    Dim lastCol As Long: lastCol = modConfig.LastCol(wsTrend, HDR_ROW_REPORT)

    ' Identify month columns (skip col A = labels, look for month names in header)
    Dim months As Variant: months = modConfig.GetMonths()
    Dim monthCols() As Long
    ReDim monthCols(0 To 11)
    Dim actualMonths As Long: actualMonths = 0

    Dim m As Long, mc As Long
    For m = 0 To 11
        mc = modConfig.FindColByHeader(wsTrend, CStr(months(m)), HDR_ROW_REPORT)
        monthCols(m) = mc
        ' Count months with actual data
        If mc > 0 Then
            Dim testVal As Double: testVal = modConfig.SafeNum(wsTrend.Cells(DATA_ROW_REPORT, mc).Value)
            If testVal <> 0 Then actualMonths = actualMonths + 1
        End If
    Next m

    If actualMonths < FORECAST_WINDOW Then
        modPerformance.TurboOff
        MsgBox "Need at least " & FORECAST_WINDOW & " months of actual data." & vbCrLf & _
               "Found " & actualMonths & " months.", vbExclamation, APP_NAME
        Exit Sub
    End If

    modPerformance.UpdateStatus "Calculating forecasts...", 0.4

    ' For each row (line item), calculate rolling average forecast
    Dim forecastCount As Long: forecastCount = 0
    Dim r As Long

    For r = DATA_ROW_REPORT To lastRow
        Dim label As String: label = Trim(CStr(wsTrend.Cells(r, 1).Value))
        If label = "" Then GoTo NextRow  ' Skip blank rows
        ' Skip header/section rows (check if they have numeric data)
        Dim hasData As Boolean: hasData = False
        For m = 0 To actualMonths - 1
            If monthCols(m) > 0 Then
                If modConfig.SafeNum(wsTrend.Cells(r, monthCols(m)).Value) <> 0 Then
                    hasData = True: Exit For
                End If
            End If
        Next m
        If Not hasData Then GoTo NextRow

        ' Calculate 3-month rolling average from last 3 actual months
        Dim rollingSum As Double: rollingSum = 0
        Dim rollingN As Long: rollingN = 0
        For m = actualMonths - 1 To Application.Max(0, actualMonths - FORECAST_WINDOW) Step -1
            If monthCols(m) > 0 Then
                rollingSum = rollingSum + modConfig.SafeNum(wsTrend.Cells(r, monthCols(m)).Value)
                rollingN = rollingN + 1
            End If
        Next m

        If rollingN = 0 Then GoTo NextRow
        Dim forecast As Double: forecast = rollingSum / rollingN

        ' Write forecast to remaining months
        For m = actualMonths To 11
            If monthCols(m) > 0 Then
                wsTrend.Cells(r, monthCols(m)).Value = forecast
                wsTrend.Cells(r, monthCols(m)).Font.Italic = True
                wsTrend.Cells(r, monthCols(m)).Font.Color = RGB(0, 0, 192)
                wsTrend.Cells(r, monthCols(m)).NumberFormat = "#,##0"
                forecastCount = forecastCount + 1
            End If
        Next m
NextRow:
    Next r

    wsTrend.Activate

    Dim elapsed As Double: elapsed = modPerformance.ElapsedSeconds()
    modPerformance.TurboOff

    modLogger.LogAction "modForecast", "RollingForecast", _
        forecastCount & " cells forecast from " & actualMonths & " months of actuals"

    MsgBox "ROLLING FORECAST COMPLETE" & vbCrLf & String(30, "=") & vbCrLf & vbCrLf & _
           "Actual Months:    " & actualMonths & vbCrLf & _
           "Forecast Method:  " & FORECAST_WINDOW & "-month rolling average" & vbCrLf & _
           "Cells Forecast:   " & forecastCount & vbCrLf & vbCrLf & _
           "Forecast values shown in blue italic.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modForecast", "ERROR", Err.Description
    MsgBox "Forecast error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' AppendToTrend - Copy a monthly summary's totals into the P&L trend
'===============================================================================
Public Sub AppendToTrend()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_PL_TREND) Then
        MsgBox "P&L - Monthly Trend sheet not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    ' Ask which month to append
    Dim months As Variant: months = modConfig.GetMonths()
    Dim monthList As String: monthList = ""
    Dim m As Long
    For m = 0 To 11
        monthList = monthList & (m + 1) & ". " & CStr(months(m)) & vbCrLf
    Next m

    Dim choice As String
    choice = InputBox("Which month to append to the trend?" & vbCrLf & vbCrLf & _
                      monthList, APP_NAME & " - Append to Trend")
    If choice = "" Then Exit Sub
    If Not IsNumeric(choice) Then Exit Sub

    Dim monthIdx As Long: monthIdx = CLng(choice) - 1
    If monthIdx < 0 Or monthIdx > 11 Then
        MsgBox "Invalid month number.", vbExclamation, APP_NAME
        Exit Sub
    End If

    ' Find the source monthly summary sheet
    Dim srcName As String
    srcName = "Functional P&L Summary - " & CStr(months(monthIdx)) & " " & FISCAL_YEAR

    If Not modConfig.SheetExists(srcName) Then
        MsgBox "Sheet '" & srcName & "' not found." & vbCrLf & _
               "Generate monthly tabs first (Action #1).", vbExclamation, APP_NAME
        Exit Sub
    End If

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Appending " & CStr(months(monthIdx)) & " to trend...", 0.2

    Dim wsSrc As Worksheet: Set wsSrc = ThisWorkbook.Worksheets(srcName)
    Dim wsTrend As Worksheet: Set wsTrend = ThisWorkbook.Worksheets(SH_PL_TREND)

    ' Find the target column on the trend sheet
    Dim tgtCol As Long
    tgtCol = modConfig.FindColByHeader(wsTrend, CStr(months(monthIdx)), HDR_ROW_REPORT)

    If tgtCol = 0 Then
        modPerformance.TurboOff
        MsgBox "Cannot find column for " & CStr(months(monthIdx)) & " on the trend sheet.", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    ' Find the US Total column on the source summary (typically last column)
    Dim srcLastCol As Long: srcLastCol = modConfig.LastCol(wsSrc, HDR_ROW_FUNC)
    Dim usCol As Long: usCol = modConfig.FindColByHeader(wsSrc, "US", HDR_ROW_FUNC)
    If usCol = 0 Then usCol = srcLastCol

    ' Match line items by label and copy US Total values
    Dim trendLastRow As Long: trendLastRow = modConfig.LastRow(wsTrend, 1)
    Dim srcLastRow As Long: srcLastRow = modConfig.LastRow(wsSrc, 1)
    Dim copyCount As Long: copyCount = 0

    Dim r As Long
    For r = DATA_ROW_REPORT To trendLastRow
        Dim trendLabel As String: trendLabel = LCase(Trim(CStr(wsTrend.Cells(r, 1).Value)))
        If trendLabel = "" Then GoTo NextTrendRow

        ' Find matching label on source sheet
        Dim sr As Long
        For sr = DATA_ROW_FUNC To srcLastRow
            Dim srcLabel As String: srcLabel = LCase(Trim(CStr(wsSrc.Cells(sr, 1).Value)))
            If srcLabel = trendLabel Then
                Dim val As Double: val = modConfig.SafeNum(wsSrc.Cells(sr, usCol).Value)
                wsTrend.Cells(r, tgtCol).Value = val
                wsTrend.Cells(r, tgtCol).NumberFormat = "#,##0"
                copyCount = copyCount + 1
                Exit For
            End If
        Next sr
NextTrendRow:
    Next r

    wsTrend.Activate

    Dim elapsed As Double: elapsed = modPerformance.ElapsedSeconds()
    modPerformance.TurboOff

    modLogger.LogAction "modForecast", "AppendToTrend", _
        CStr(months(monthIdx)) & ": " & copyCount & " line items copied to trend"

    MsgBox CStr(months(monthIdx)) & " data appended to P&L Trend." & vbCrLf & _
           copyCount & " line items copied.", vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modForecast", "ERROR", Err.Description
    MsgBox "Append error: " & Err.Description, vbCritical, APP_NAME
End Sub
