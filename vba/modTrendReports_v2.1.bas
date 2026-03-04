Attribute VB_Name = "modTrendReports"
Option Explicit

'===============================================================================
' modTrendReports - Trend Views & Historical Archiving
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Builds rolling time views and preserves historical reconciliation
'           results so month-over-month trends are always available.
'
' PUBLIC SUBS:
'   CreateRolling12MonthView          - Build a dynamic rolling 12-month P&L (#77)
'   CreateReconciliationTrendChart    - Chart PASS/FAIL counts over time (#156)
'   ArchiveReconciliationResults      - Save a dated snapshot of the Checks tab (#163)
'
' VERSION:  2.1.0 (New module — 2026-03-01)
' SOURCE:   Ideas from NewTesting/VBA Examples (200) — items #77, #156, #163
'===============================================================================

' Sheet name constants used by this module
Private Const SH_ROLLING_12    As String = "Rolling 12-Month P&L"
Private Const SH_RECON_TREND   As String = "Recon Trend Chart"
Private Const SH_RECON_ARCHIVE As String = "Recon Archive"

'===============================================================================
' CreateRolling12MonthView - Build a dynamic rolling 12-month P&L (#77)
' Reads the P&L Monthly Trend sheet and creates a new sheet that always
' shows the most recent 12 months available, regardless of which month
' it is run in. Column headers auto-detect the last populated month and
' count back 12. A chart of Total Revenue across those 12 months is added.
'===============================================================================
Public Sub CreateRolling12MonthView()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_PL_TREND) Then
        MsgBox "'" & SH_PL_TREND & "' not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Building rolling 12-month view...", 0.1

    Dim wsSrc As Worksheet: Set wsSrc = ThisWorkbook.Worksheets(SH_PL_TREND)
    Dim srcLastRow As Long: srcLastRow = modConfig.LastRow(wsSrc, 1)

    ' Find how many months have data (columns B onward = Jan–Dec, col 2–13)
    ' Detect the last column that has any non-zero value in the revenue row
    Dim revRow As Long: revRow = modConfig.FindRowByLabel(wsSrc, "total revenue", DATA_ROW_REPORT)
    If revRow = 0 Then revRow = DATA_ROW_REPORT

    Dim lastDataCol As Long: lastDataCol = 1
    Dim c As Long
    For c = 13 To 2 Step -1
        If modConfig.SafeNum(wsSrc.Cells(revRow, c).Value) <> 0 Then
            lastDataCol = c
            Exit For
        End If
    Next c
    If lastDataCol < 2 Then lastDataCol = 13  ' Default to all 12 if nothing found

    ' Calculate rolling window: up to 12 months ending at lastDataCol
    Dim windowSize As Long: windowSize = Application.Min(12, lastDataCol - 1)
    Dim startCol   As Long: startCol   = lastDataCol - windowSize + 1

    ' Build month label array for the window
    Dim mths As Variant: mths = modConfig.GetMonths()
    Dim windowLabels() As String
    ReDim windowLabels(0 To windowSize - 1)
    Dim mi As Long
    For mi = 0 To windowSize - 1
        Dim mthIdx As Long: mthIdx = (startCol - 2 + mi)
        If mthIdx >= 0 And mthIdx <= 11 Then
            windowLabels(mi) = CStr(mths(mthIdx)) & " " & FISCAL_YEAR_4
        Else
            windowLabels(mi) = "Month " & (mi + 1)
        End If
    Next mi

    ' Create output sheet
    modConfig.SafeDeleteSheet SH_ROLLING_12
    Dim wsOut As Worksheet
    Set wsOut = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsOut.Name = SH_ROLLING_12

    ' Title
    wsOut.Range("A1").Value = "Keystone BenefitTech " & ChrW(8212) & " Rolling 12-Month P&L"
    wsOut.Range("A1").Font.Size = 14: wsOut.Range("A1").Font.Bold = True
    wsOut.Range("A1").Font.Color = CLR_NAVY
    wsOut.Range("A2").Value = "Generated: " & Format(Now, "mmmm d, yyyy")
    wsOut.Range("A2").Font.Italic = True
    wsOut.Range("A2").Font.Color = RGB(120, 120, 120)

    modPerformance.UpdateStatus "Copying P&L data...", 0.4

    ' Write column headers (row 4)
    wsOut.Cells(HDR_ROW_REPORT, 1).Value = "Line Item"
    Dim hdrCol As Long
    For hdrCol = 0 To windowSize - 1
        wsOut.Cells(HDR_ROW_REPORT, hdrCol + 2).Value = windowLabels(hdrCol)
    Next hdrCol
    wsOut.Range(wsOut.Cells(HDR_ROW_REPORT, 1), _
                wsOut.Cells(HDR_ROW_REPORT, windowSize + 1)).Font.Bold = True
    wsOut.Range(wsOut.Cells(HDR_ROW_REPORT, 1), _
                wsOut.Cells(HDR_ROW_REPORT, windowSize + 1)).Interior.Color = CLR_NAVY
    wsOut.Range(wsOut.Cells(HDR_ROW_REPORT, 1), _
                wsOut.Cells(HDR_ROW_REPORT, windowSize + 1)).Font.Color = CLR_WHITE

    ' Copy row labels and data values for the rolling window
    Dim outR As Long: outR = DATA_ROW_REPORT
    Dim r As Long
    For r = DATA_ROW_REPORT To srcLastRow
        Dim rowLbl As String: rowLbl = modConfig.SafeStr(wsSrc.Cells(r, 1).Value)
        wsOut.Cells(outR, 1).Value = rowLbl
        If Len(rowLbl) > 0 And wsSrc.Cells(r, 1).Font.Bold Then
            wsOut.Cells(outR, 1).Font.Bold = True
        End If
        Dim dc As Long
        For dc = 0 To windowSize - 1
            Dim srcCol As Long: srcCol = startCol + dc
            Dim cellVal As Double: cellVal = modConfig.SafeNum(wsSrc.Cells(r, srcCol).Value)
            wsOut.Cells(outR, dc + 2).Value = cellVal
            If cellVal <> 0 Then
                ' Detect percentage rows (e.g., "Contribution Margin %", "GM%") — use % format
                If InStr(1, rowLbl, "%", vbTextCompare) > 0 Then
                    wsOut.Cells(outR, dc + 2).NumberFormat = "0.0%"
                Else
                    wsOut.Cells(outR, dc + 2).NumberFormat = "$#,##0"
                End If
            End If
        Next dc
        outR = outR + 1
    Next r

    ' Revenue trend chart for the rolling window — placed below all data rows
    modPerformance.UpdateStatus "Adding trend chart...", 0.75
    Dim chartTop As Long: chartTop = wsOut.Cells(outR + 1, 1).Top
    Dim co As ChartObject
    Set co = wsOut.ChartObjects.Add(Left:=20, Top:=chartTop, Width:=500, Height:=280)
    co.Name = "Rolling12Chart"

    ' Find Total Revenue row on the OUTPUT sheet (revRow is from wsSrc, not wsOut)
    Dim outRevRow As Long: outRevRow = 0
    Dim sr As Long
    For sr = DATA_ROW_REPORT To outR - 1
        Dim srLbl As String: srLbl = LCase(modConfig.SafeStr(wsOut.Cells(sr, 1).Value))
        If srLbl = "total revenue" Or srLbl = "revenue" Or srLbl = "net revenue" Then
            outRevRow = sr
            Exit For
        End If
    Next sr
    If outRevRow = 0 Then outRevRow = DATA_ROW_REPORT  ' Fallback to first data row

    With co.Chart
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Revenue " & ChrW(8212) & " Rolling " & windowSize & " Months"
        Dim ser As Series
        Set ser = .SeriesCollection.NewSeries
        ser.Name = "Total Revenue"
        ser.Values = wsOut.Range(wsOut.Cells(outRevRow, 2), wsOut.Cells(outRevRow, windowSize + 1))
        ser.XValues = windowLabels
        .Axes(xlValue).TickLabels.NumberFormat = "$#,##0"
        .HasLegend = False
        .PlotArea.Interior.Color = CLR_WHITE
    End With

    wsOut.Columns("A:M").AutoFit
    wsOut.Tab.Color = CLR_NAVY
    wsOut.Activate

    Dim elapsed As Double: elapsed = modPerformance.ElapsedSeconds()
    modPerformance.TurboOff
    modLogger.LogAction "modTrendReports", "CreateRolling12MonthView", _
        windowSize & "-month window (cols " & startCol & " to " & lastDataCol & ") (" & Format(elapsed, "0.0") & "s)"
    MsgBox "Rolling " & windowSize & "-month P&L view created on '" & SH_ROLLING_12 & "'.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    MsgBox "CreateRolling12MonthView error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' CreateReconciliationTrendChart - Chart PASS/FAIL counts over time (#156)
' Reads the Recon Archive sheet (created by ArchiveReconciliationResults) and
' charts how many checks passed or failed across each archived run.
' If no archive exists yet, prompts the user to run the archiver first.
'===============================================================================
Public Sub CreateReconciliationTrendChart()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_RECON_ARCHIVE) Then
        MsgBox "No reconciliation archive found." & vbCrLf & vbCrLf & _
               "Run ArchiveReconciliationResults after each month's checks to" & vbCrLf & _
               "build up the history needed for this trend chart.", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim wsArch As Worksheet: Set wsArch = ThisWorkbook.Worksheets(SH_RECON_ARCHIVE)
    Dim lastRow As Long: lastRow = modConfig.LastRow(wsArch, 1)

    If lastRow < 2 Then
        MsgBox "Archive sheet exists but has no data rows yet.", vbExclamation, APP_NAME
        Exit Sub
    End If

    modPerformance.TurboOn

    ' Expected archive columns: RunDate | RunLabel | PassCount | FailCount
    ' Collect all run summary rows (distinct RunDate values)
    Dim runDates()  As String
    Dim passArr()   As Long
    Dim failArr()   As Long
    Dim runCount    As Long: runCount = 0
    ReDim runDates(1 To lastRow)
    ReDim passArr(1 To lastRow)
    ReDim failArr(1 To lastRow)

    Dim r As Long
    For r = 2 To lastRow
        Dim cellDate As String: cellDate = modConfig.SafeStr(wsArch.Cells(r, 1).Value)
        Dim cellType As String: cellType = modConfig.SafeStr(wsArch.Cells(r, 2).Value)
        ' Summary rows are tagged "SUMMARY" in column 2
        If cellType = "SUMMARY" Then
            runCount = runCount + 1
            runDates(runCount) = cellDate
            passArr(runCount)  = CLng(modConfig.SafeNum(wsArch.Cells(r, 3).Value))
            failArr(runCount)  = CLng(modConfig.SafeNum(wsArch.Cells(r, 4).Value))
        End If
    Next r

    If runCount < 1 Then
        modPerformance.TurboOff
        MsgBox "No SUMMARY rows found in the archive. Archive may be from an older format.", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    ' Create trend chart sheet
    modConfig.SafeDeleteSheet SH_RECON_TREND
    Dim wsChart As Worksheet
    Set wsChart = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsChart.Name = SH_RECON_TREND

    ' Write data table for the chart
    wsChart.Range("A1").Value = "Reconciliation Trend — PASS / FAIL by Run"
    wsChart.Range("A1").Font.Bold = True
    wsChart.Range("A1").Font.Size = 13
    wsChart.Cells(3, 1).Value = "Run Date": wsChart.Cells(3, 2).Value = "PASS": wsChart.Cells(3, 3).Value = "FAIL"
    wsChart.Range("A3:C3").Font.Bold = True

    Dim i As Long
    For i = 1 To runCount
        wsChart.Cells(3 + i, 1).Value = runDates(i)
        wsChart.Cells(3 + i, 2).Value = passArr(i)
        wsChart.Cells(3 + i, 3).Value = failArr(i)
    Next i

    ' Add chart
    Dim co As ChartObject
    Set co = wsChart.ChartObjects.Add(Left:=20, Top:=120, Width:=500, Height:=300)
    co.Name = "ReconTrendChart"
    With co.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=wsChart.Range(wsChart.Cells(3, 1), wsChart.Cells(3 + runCount, 3))
        .HasTitle = True
        .ChartTitle.Text = "Reconciliation Pass/Fail Trend"
        On Error Resume Next
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(0, 176, 80)   ' PASS = green
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(192, 0, 0)    ' FAIL = red
        On Error GoTo ErrHandler
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Check Count"
    End With

    wsChart.Columns("A:C").AutoFit
    wsChart.Tab.Color = RGB(0, 176, 80)
    wsChart.Activate

    modPerformance.TurboOff
    modLogger.LogAction "modTrendReports", "CreateReconciliationTrendChart", _
        runCount & " run(s) charted"
    MsgBox "Reconciliation trend chart created (" & runCount & " run(s) charted).", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    MsgBox "CreateReconciliationTrendChart error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' ArchiveReconciliationResults - Save a dated snapshot of the Checks tab (#163)
' Appends every check row from the current Checks tab into a permanent
' "Recon Archive" sheet, tagged with today's date and a PASS/FAIL summary row.
' Run this once per month after reviewing and approving the checks.
' The archive is the data source for CreateReconciliationTrendChart.
'===============================================================================
Public Sub ArchiveReconciliationResults()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_CHECKS) Then
        MsgBox "Checks sheet '" & SH_CHECKS & "' not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    Dim wsChk As Worksheet: Set wsChk = ThisWorkbook.Worksheets(SH_CHECKS)
    Dim chkLastRow As Long: chkLastRow = modConfig.LastRow(wsChk, 1)

    If chkLastRow < DATA_ROW_CHECKS Then
        MsgBox "No check data found on the Checks sheet.", vbExclamation, APP_NAME
        Exit Sub
    End If

    ' Confirm before appending
    Dim runLabel As String
    runLabel = Format(Date, "yyyy-mm-dd") & " — " & _
               Format(Date, "mmmm yyyy") & " close"

    If MsgBox("Archive current check results?" & vbCrLf & vbCrLf & _
              "Run label: " & runLabel & vbCrLf & _
              "Checks rows: " & (chkLastRow - DATA_ROW_CHECKS + 1), _
              vbYesNo + vbQuestion, APP_NAME) = vbNo Then Exit Sub

    ' Create archive sheet if it does not exist
    If Not modConfig.SheetExists(SH_RECON_ARCHIVE) Then
        Dim wsNew As Worksheet
        Set wsNew = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsNew.Name = SH_RECON_ARCHIVE
        modConfig.StyleHeader wsNew, 1, _
            Array("Archive Date", "Type / Check Name", "Pass Count", "Fail Count", "Check Status", "Difference")
        wsNew.Columns("A:F").AutoFit
    End If

    Dim wsArch As Worksheet: Set wsArch = ThisWorkbook.Worksheets(SH_RECON_ARCHIVE)
    Dim archNextRow As Long: archNextRow = modConfig.LastRow(wsArch, 1) + 1

    ' Append each check row
    Dim passCount As Long: passCount = 0
    Dim failCount As Long: failCount = 0
    Dim r As Long
    Dim runDate As String: runDate = Format(Now, "yyyy-mm-dd hh:nn:ss")

    For r = DATA_ROW_CHECKS To chkLastRow
        Dim checkName As String: checkName = modConfig.SafeStr(wsChk.Cells(r, 1).Value)
        If Len(checkName) = 0 Then GoTo NextChk

        Dim chkStatus As String: chkStatus = UCase(modConfig.SafeStr(wsChk.Cells(r, COL_CHECK_STATUS).Value))
        Dim chkDiff   As Double: chkDiff   = modConfig.SafeNum(wsChk.Cells(r, 4).Value)

        wsArch.Cells(archNextRow, 1).Value = runDate
        wsArch.Cells(archNextRow, 2).Value = checkName
        wsArch.Cells(archNextRow, 3).Value = IIf(chkStatus = "PASS", 1, 0)
        wsArch.Cells(archNextRow, 4).Value = IIf(chkStatus = "FAIL", 1, 0)
        wsArch.Cells(archNextRow, 5).Value = chkStatus
        wsArch.Cells(archNextRow, 6).Value = chkDiff
        wsArch.Cells(archNextRow, 6).NumberFormat = "$#,##0.00"

        If chkStatus = "PASS" Then
            wsArch.Cells(archNextRow, 5).Interior.Color = RGB(200, 255, 200)
            passCount = passCount + 1
        ElseIf chkStatus = "FAIL" Then
            wsArch.Cells(archNextRow, 5).Interior.Color = RGB(255, 200, 200)
            failCount = failCount + 1
        End If
        archNextRow = archNextRow + 1
NextChk:
    Next r

    ' Append SUMMARY row for trend chart
    wsArch.Cells(archNextRow, 1).Value = runDate
    wsArch.Cells(archNextRow, 2).Value = "SUMMARY"
    wsArch.Cells(archNextRow, 3).Value = passCount
    wsArch.Cells(archNextRow, 4).Value = failCount
    wsArch.Cells(archNextRow, 5).Value = runLabel
    wsArch.Range(wsArch.Cells(archNextRow, 1), wsArch.Cells(archNextRow, 6)).Font.Bold = True
    wsArch.Range(wsArch.Cells(archNextRow, 1), wsArch.Cells(archNextRow, 6)).Interior.Color = CLR_LIGHT_GRAY

    wsArch.Columns("A:F").AutoFit

    modLogger.LogAction "modTrendReports", "ArchiveReconciliationResults", _
        runLabel & " | " & passCount & " PASS / " & failCount & " FAIL"
    MsgBox "Reconciliation results archived." & vbCrLf & _
           runLabel & vbCrLf & vbCrLf & _
           passCount & " PASS  |  " & failCount & " FAIL" & vbCrLf & vbCrLf & _
           "Run CreateReconciliationTrendChart to see trend over time.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    MsgBox "ArchiveReconciliationResults error: " & Err.Description, vbCritical, APP_NAME
End Sub
