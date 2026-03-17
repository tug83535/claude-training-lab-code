Attribute VB_Name = "modTimeSaved"
Option Explicit

'===============================================================================
' modTimeSaved - Time Saved Calculator
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Calculates and displays how long each Command Center action would
'           take manually vs. running the macro. Builds a styled summary sheet
'           with per-action savings and a grand total.
'           Perfect talking point for CFO/CEO: "This saves X hours per month."
'
' PUBLIC SUBS:
'   ShowTimeSavedReport   - Build the Time Saved report sheet
'
' DEPENDENCIES: modConfig, modPerformance, modLogger
' VERSION:  2.1.0
'===============================================================================

Private Const SH_TIME_SAVED As String = "Time Saved Analysis"

'===============================================================================
' ShowTimeSavedReport - Build styled report showing manual vs automated time
'===============================================================================
Public Sub ShowTimeSavedReport()
    On Error GoTo ErrHandler

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Building Time Saved report...", 0.1

    ' Remove old sheet if it exists
    modConfig.SafeDeleteSheet SH_TIME_SAVED

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = SH_TIME_SAVED

    ' --- Title Block ---
    ws.Cells(1, 1).Value = "Keystone BenefitTech, Inc."
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Color = CLR_NAVY

    ws.Cells(2, 1).Value = "Time Saved Analysis - Automation ROI"
    ws.Cells(2, 1).Font.Size = 11
    ws.Cells(2, 1).Font.Italic = True
    ws.Cells(2, 1).Font.Color = CLR_NAVY

    ws.Cells(3, 1).Value = "Generated: " & Format(Now, "mmmm d, yyyy h:mm AM/PM")
    ws.Cells(3, 1).Font.Size = 9
    ws.Cells(3, 1).Font.Italic = True

    ' --- Headers ---
    Dim hdrRow As Long: hdrRow = 5
    Dim headers As Variant
    headers = Array("#", "Action Name", "Category", _
                    "Manual Time (min)", "Automated Time (min)", _
                    "Time Saved (min)", "Savings %")
    modConfig.StyleHeader ws, hdrRow, headers

    ' Column widths
    ws.Columns(1).ColumnWidth = 5
    ws.Columns(2).ColumnWidth = 42
    ws.Columns(3).ColumnWidth = 22
    ws.Columns(4).ColumnWidth = 18
    ws.Columns(5).ColumnWidth = 20
    ws.Columns(6).ColumnWidth = 18
    ws.Columns(7).ColumnWidth = 12

    modPerformance.UpdateStatus "Building Time Saved report...", 0.3

    ' --- Populate Data ---
    Dim r As Long: r = hdrRow + 1
    Dim totalManual As Double: totalManual = 0
    Dim totalAuto As Double: totalAuto = 0

    ' Action data: num, name, category, manual minutes, auto minutes
    Dim actions As Variant
    actions = GetActionTimeData()

    Dim i As Long
    For i = 0 To UBound(actions)
        Dim parts As Variant: parts = actions(i)
        ws.Cells(r, 1).Value = parts(0)       ' Action #
        ws.Cells(r, 2).Value = parts(1)       ' Name
        ws.Cells(r, 3).Value = parts(2)       ' Category
        ws.Cells(r, 4).Value = parts(3)       ' Manual min
        ws.Cells(r, 4).NumberFormat = "#,##0.0"
        ws.Cells(r, 5).Value = parts(4)       ' Auto min
        ws.Cells(r, 5).NumberFormat = "#,##0.0"

        Dim saved As Double: saved = parts(3) - parts(4)
        ws.Cells(r, 6).Value = saved
        ws.Cells(r, 6).NumberFormat = "#,##0.0"

        If parts(3) > 0 Then
            ws.Cells(r, 7).Value = saved / parts(3)
        Else
            ws.Cells(r, 7).Value = 0
        End If
        ws.Cells(r, 7).NumberFormat = "0%"

        ' Alternating row color
        If r Mod 2 = 0 Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Interior.Color = CLR_ALT_ROW
        End If

        totalManual = totalManual + parts(3)
        totalAuto = totalAuto + parts(4)
        r = r + 1
    Next i

    modPerformance.UpdateStatus "Building Time Saved report...", 0.8

    ' --- Totals Row ---
    Dim totRow As Long: totRow = r + 1
    ws.Cells(totRow, 2).Value = "TOTAL (per monthly close)"
    ws.Cells(totRow, 4).Value = totalManual
    ws.Cells(totRow, 4).NumberFormat = "#,##0.0"
    ws.Cells(totRow, 5).Value = totalAuto
    ws.Cells(totRow, 5).NumberFormat = "#,##0.0"
    ws.Cells(totRow, 6).Value = totalManual - totalAuto
    ws.Cells(totRow, 6).NumberFormat = "#,##0.0"
    If totalManual > 0 Then
        ws.Cells(totRow, 7).Value = (totalManual - totalAuto) / totalManual
    End If
    ws.Cells(totRow, 7).NumberFormat = "0%"

    With ws.Range(ws.Cells(totRow, 1), ws.Cells(totRow, 7))
        .Font.Bold = True
        .Interior.Color = CLR_NAVY
        .Font.Color = CLR_WHITE
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlMedium
    End With

    ' --- Executive Summary Box ---
    Dim sumRow As Long: sumRow = totRow + 3
    Dim manualHrs As Double: manualHrs = Round(totalManual / 60, 1)
    Dim autoHrs As Double: autoHrs = Round(totalAuto / 60, 1)
    Dim savedHrs As Double: savedHrs = Round((totalManual - totalAuto) / 60, 1)
    Dim annualSaved As Double: annualSaved = Round(savedHrs * 12, 0)

    ws.Cells(sumRow, 1).Value = "EXECUTIVE SUMMARY"
    ws.Cells(sumRow, 1).Font.Size = 13
    ws.Cells(sumRow, 1).Font.Bold = True
    ws.Cells(sumRow, 1).Font.Color = CLR_NAVY

    ws.Cells(sumRow + 2, 1).Value = "Manual Process:"
    ws.Cells(sumRow + 2, 2).Value = manualHrs & " hours per monthly close"
    ws.Cells(sumRow + 2, 1).Font.Bold = True

    ws.Cells(sumRow + 3, 1).Value = "Automated:"
    ws.Cells(sumRow + 3, 2).Value = autoHrs & " hours per monthly close"
    ws.Cells(sumRow + 3, 1).Font.Bold = True

    ws.Cells(sumRow + 4, 1).Value = "Time Saved:"
    ws.Cells(sumRow + 4, 2).Value = savedHrs & " hours per monthly close"
    ws.Cells(sumRow + 4, 1).Font.Bold = True
    ws.Cells(sumRow + 4, 2).Font.Bold = True
    ws.Cells(sumRow + 4, 2).Font.Color = RGB(0, 128, 0)

    ws.Cells(sumRow + 5, 1).Value = "Annual Savings:"
    ws.Cells(sumRow + 5, 2).Value = annualSaved & " hours per year (~" & Round(annualSaved / 2080 * 100, 0) & "% of one FTE)"
    ws.Cells(sumRow + 5, 1).Font.Bold = True
    ws.Cells(sumRow + 5, 2).Font.Bold = True
    ws.Cells(sumRow + 5, 2).Font.Color = RGB(0, 128, 0)
    ws.Cells(sumRow + 5, 2).Font.Size = 12

    ' Border around the summary box
    Dim sumRange As Range
    Set sumRange = ws.Range(ws.Cells(sumRow, 1), ws.Cells(sumRow + 6, 4))
    With sumRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = CLR_NAVY
    End With

    ' Freeze panes at data start
    ws.Activate
    ws.Cells(hdrRow + 1, 1).Select
    ActiveWindow.FreezePanes = True

    modPerformance.TurboOff

    modLogger.LogAction "modTimeSaved", "ShowTimeSavedReport", _
        "Report built: Manual " & manualHrs & "h, Auto " & autoHrs & "h, Saved " & savedHrs & "h/month"

    MsgBox "Time Saved Report Complete!" & vbCrLf & vbCrLf & _
           "Manual:  " & manualHrs & " hours/month" & vbCrLf & _
           "Automated:  " & autoHrs & " hours/month" & vbCrLf & _
           "SAVED:  " & savedHrs & " hours/month" & vbCrLf & _
           "Annual:  ~" & annualSaved & " hours/year", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modTimeSaved", "ERROR", Err.Description
    MsgBox "Time Saved error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' GetActionTimeData - Returns array of (num, name, category, manualMin, autoMin)
' Manual times are realistic estimates for a Finance analyst doing each task
' by hand in Excel. Automated times are typical macro runtimes.
'===============================================================================
Private Function GetActionTimeData() As Variant
    Dim d(0 To 61) As Variant

    ' --- Monthly Operations ---
    d(0) = Array(1, "Generate Monthly Tabs (Apr-Dec)", "Monthly Ops", 45, 0.5)
    d(1) = Array(2, "Delete Generated Tabs", "Monthly Ops", 10, 0.2)
    d(2) = Array(3, "Run All Reconciliation Checks", "Monthly Ops", 60, 0.3)
    d(3) = Array(4, "Export Check Results", "Monthly Ops", 15, 0.2)

    ' --- Analysis ---
    d(4) = Array(5, "Sensitivity Analysis", "Analysis", 90, 0.5)
    d(5) = Array(6, "Variance Analysis (MoM)", "Analysis", 45, 0.3)

    ' --- Data Quality ---
    d(6) = Array(7, "Scan All Data Quality Issues", "Data Quality", 30, 0.3)
    d(7) = Array(8, "Fix Text-Stored Numbers", "Data Quality", 20, 0.2)
    d(8) = Array(9, "Fix Duplicate Entries", "Data Quality", 25, 0.2)

    ' --- Reporting ---
    d(9) = Array(10, "Export Full Report Package (PDF)", "Reporting", 30, 0.5)
    d(10) = Array(11, "Export Single Sheet (PDF)", "Reporting", 5, 0.1)
    d(11) = Array(12, "Build Dashboard Charts", "Reporting", 60, 0.5)

    ' --- Utilities ---
    d(12) = Array(13, "Refresh Table of Contents", "Utilities", 15, 0.1)
    d(13) = Array(14, "Validate & Recalc AWS Allocation", "Utilities", 30, 0.3)
    d(14) = Array(15, "Quick Jump to Sheet", "Utilities", 0.5, 0.1)
    d(15) = Array(16, "Go Home", "Utilities", 0.2, 0.1)

    ' --- Data Import ---
    d(16) = Array(17, "Import Data Pipeline (CSV/Excel)", "Data Import", 20, 0.3)

    ' --- Forecasting ---
    d(17) = Array(18, "Rolling Forecast", "Forecasting", 120, 0.5)
    d(18) = Array(19, "Append Forecast to Trend", "Forecasting", 15, 0.2)

    ' --- Scenarios ---
    d(19) = Array(20, "Save Scenario", "Scenarios", 10, 0.2)
    d(20) = Array(21, "Load Scenario", "Scenarios", 10, 0.2)
    d(21) = Array(22, "Compare Scenarios", "Scenarios", 30, 0.2)
    d(22) = Array(23, "Delete Scenario", "Scenarios", 5, 0.1)

    ' --- Allocation ---
    d(23) = Array(24, "Run Allocation Engine", "Allocation", 45, 0.4)
    d(24) = Array(25, "Allocation Preview", "Allocation", 20, 0.2)

    ' --- Consolidation ---
    d(25) = Array(26, "Consolidation Menu", "Consolidation", 5, 0.1)
    d(26) = Array(27, "Add Entity", "Consolidation", 15, 0.2)
    d(27) = Array(28, "Generate Consolidated P&L", "Consolidation", 90, 0.5)
    d(28) = Array(29, "List Entities", "Consolidation", 5, 0.1)
    d(29) = Array(30, "Add IC Elimination", "Consolidation", 20, 0.2)

    ' --- Version Control ---
    d(30) = Array(31, "Version Menu", "Version Control", 5, 0.1)
    d(31) = Array(32, "Save Version Snapshot", "Version Control", 15, 0.2)
    d(32) = Array(33, "Compare Versions", "Version Control", 30, 0.3)
    d(33) = Array(34, "Restore Version", "Version Control", 20, 0.2)
    d(34) = Array(35, "List All Versions", "Version Control", 10, 0.1)

    ' --- Governance ---
    d(35) = Array(36, "Generate Tech Documentation", "Governance", 120, 0.5)
    d(36) = Array(37, "Change Management Menu", "Governance", 5, 0.1)
    d(37) = Array(38, "Add Change Request", "Governance", 10, 0.2)
    d(38) = Array(39, "Update Change Status", "Governance", 5, 0.1)
    d(39) = Array(40, "Change Management Summary", "Governance", 15, 0.2)

    ' --- Admin & Testing ---
    d(40) = Array(41, "View Audit Log", "Admin", 10, 0.1)
    d(41) = Array(42, "Export Audit Log", "Admin", 10, 0.2)
    d(42) = Array(43, "Clear Audit Log", "Admin", 5, 0.1)
    d(43) = Array(44, "Run Full Integration Test", "Admin", 120, 0.8)
    d(44) = Array(45, "Quick Health Check", "Admin", 30, 0.3)

    ' --- Advanced ---
    d(45) = Array(46, "Generate Variance Commentary", "Advanced", 45, 0.3)
    d(46) = Array(47, "Cross-Sheet Validation", "Advanced", 40, 0.3)
    d(47) = Array(48, "Toggle Executive Mode", "Advanced", 5, 0.1)
    d(48) = Array(49, "Force Recalculate", "Advanced", 2, 0.1)
    d(49) = Array(50, "About / Version Info", "Advanced", 0, 0.1)

    ' --- Sheet Tools ---
    d(50) = Array(51, "Delete Blank Rows", "Sheet Tools", 15, 0.2)
    d(51) = Array(52, "Unhide All Sheets", "Sheet Tools", 5, 0.1)
    d(52) = Array(53, "Sort Sheets Alphabetically", "Sheet Tools", 10, 0.2)
    d(53) = Array(54, "Toggle Freeze Panes", "Sheet Tools", 2, 0.1)
    d(54) = Array(55, "Convert Formulas to Values", "Sheet Tools", 10, 0.2)
    d(55) = Array(56, "AutoFit All Columns", "Sheet Tools", 5, 0.1)
    d(56) = Array(57, "Protect All Sheets", "Sheet Tools", 15, 0.2)
    d(57) = Array(58, "Unprotect All Sheets", "Sheet Tools", 10, 0.2)
    d(58) = Array(59, "Find & Replace All Sheets", "Sheet Tools", 15, 0.2)
    d(59) = Array(60, "Highlight Hardcoded Numbers", "Sheet Tools", 20, 0.3)
    d(60) = Array(61, "Toggle Presentation Mode", "Sheet Tools", 3, 0.1)
    d(61) = Array(62, "Unmerge & Fill Down", "Sheet Tools", 10, 0.2)

    GetActionTimeData = d
End Function
