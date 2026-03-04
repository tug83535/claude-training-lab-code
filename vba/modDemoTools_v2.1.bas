Attribute VB_Name = "modDemoTools"
Option Explicit

'===============================================================================
' modDemoTools - Demo Presentation & Print Tools
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Macros that improve the live demo experience and produce
'           print-ready output for the CFO/CEO presentation.
'
' PUBLIC SUBS:
'   AddControlSheetButtons       - Add clickable macro buttons to Report--> (#17)
'   SetParameterizedPrintArea    - Set print area by selected month/product (#63)
'   CreatePrintableExecSummary   - Build one-page print layout for CFO (#64)
'
' VERSION:  2.1.0 (New module — 2026-03-01)
' SOURCE:   Ideas from NewTesting/VBA Examples (200) — items #17, #63, #64
'===============================================================================

'===============================================================================
' AddControlSheetButtons - Add labeled macro buttons to Report--> sheet (#17)
' Creates 5 buttons for the most common demo actions so the sheet looks
' like a polished app, not a raw spreadsheet. Old DemoBtn* buttons are
' removed and replaced each time this runs.
'===============================================================================
Public Sub AddControlSheetButtons()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_REPORT) Then
        MsgBox "Report--> sheet not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_REPORT)

    ' Remove any existing demo buttons (tagged by name prefix "DemoBtn")
    Dim btn As Object
    For Each btn In ws.Buttons
        If Left(btn.Name, 7) = "DemoBtn" Then btn.Delete
    Next btn

    ' Place buttons below all existing content with a 2-row gap
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells.Find(What:="*", SearchOrder:=xlByRows, _
                                SearchDirection:=xlPrevious).row
    If lastUsedRow < 1 Then lastUsedRow = 1
    Dim startRow As Long: startRow = lastUsedRow + 2
    Dim topPos As Long: topPos = ws.Cells(startRow, 1).Top

    ' Section label above buttons
    ws.Cells(startRow, 1).Value = "Quick Actions"
    ws.Cells(startRow, 1).Font.Bold = True
    ws.Cells(startRow, 1).Font.Size = 11
    ws.Cells(startRow, 1).Font.Color = CLR_NAVY
    topPos = ws.Cells(startRow + 1, 1).Top

    ' Define: Caption, MacroToCall, Width, Height
    Dim btnDefs As Variant
    btnDefs = Array( _
        Array("Run Reconciliation",  "modReconciliation.RunAllChecks",   170, 28), _
        Array("Build Dashboard",     "modDashboard.BuildDashboard",      170, 28), _
        Array("Data Quality Check",  "modDataQuality.ScanAll",           170, 28), _
        Array("Export PDF",          "modPDFExport.ExportReportPackage",  170, 28), _
        Array("Validate Assumptions","modDataGuards.ValidateAssumptionsPresence", 170, 28))

    Dim btnSpacing As Long: btnSpacing = 34  ' vertical gap between buttons
    Dim i As Long
    For i = 0 To 4
        Dim def As Variant: def = btnDefs(i)
        Dim newBtn As Object
        Set newBtn = ws.Buttons.Add(20, topPos + (i * btnSpacing), CLng(def(2)), CLng(def(3)))
        newBtn.Caption  = CStr(def(0))
        newBtn.OnAction = CStr(def(1))
        newBtn.Name     = "DemoBtn" & i
        With newBtn.Font
            .Name = "Calibri"
            .Size = 9
            .Bold = True
        End With
    Next i

    modLogger.LogAction "modDemoTools", "AddControlSheetButtons", _
        "5 buttons added to " & SH_REPORT
    MsgBox "5 control buttons added to '" & SH_REPORT & "'." & vbCrLf & _
           "Click any button during the demo to run its action instantly.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    MsgBox "AddControlSheetButtons error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' SetParameterizedPrintArea - Set print area by month/sheet selection (#63)
' Targets the FPL Summary - Dynamic sheet if it exists, otherwise falls back
' to the first visible Functional P&L Summary sheet found.
' Configures full page setup so it prints clean in one click.
'===============================================================================
Public Sub SetParameterizedPrintArea()
    On Error GoTo ErrHandler

    ' Prefer the dynamic sheet; fall back to any static functional summary
    Dim targetName As String: targetName = "FPL Summary - Dynamic"
    If Not modConfig.SheetExists(targetName) Then
        Dim ws2 As Worksheet
        For Each ws2 In ThisWorkbook.Worksheets
            If InStr(ws2.Name, "Functional P&L Summary") > 0 And _
               ws2.Visible = xlSheetVisible Then
                targetName = ws2.Name
                Exit For
            End If
        Next ws2
    End If

    If Not modConfig.SheetExists(targetName) Then
        MsgBox "No Functional P&L Summary sheet found." & vbCrLf & _
               "Run CreateDynamicSummary first.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(targetName)
    Dim lastRow As Long: lastRow = modConfig.LastRow(ws, 1)
    Dim lastCol As Long: lastCol = modConfig.LastCol(ws, HDR_ROW_FUNC)
    If lastRow < 5  Then lastRow = 40
    If lastCol < 2  Then lastCol = 5

    ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Address

    With ws.PageSetup
        .Orientation       = xlPortrait
        .FitToPagesWide    = 1
        .FitToPagesTall    = 1
        .Zoom              = False
        .LeftMargin        = Application.InchesToPoints(0.5)
        .RightMargin       = Application.InchesToPoints(0.5)
        .TopMargin         = Application.InchesToPoints(0.75)
        .BottomMargin      = Application.InchesToPoints(0.75)
        .CenterHorizontally = True
        .PrintTitleRows    = ws.Rows("1:4").Address
        .LeftHeader        = "Keystone BenefitTech, Inc."
        .CenterHeader      = "&B" & ws.Name
        .RightHeader       = "FY" & FISCAL_YEAR_4
        .CenterFooter      = "Page &P of &N"
    End With

    ws.Activate
    modLogger.LogAction "modDemoTools", "SetParameterizedPrintArea", _
        "Print area set on " & targetName & " (" & lastRow & "r x " & lastCol & "c)"
    MsgBox "Print area set on '" & targetName & "'." & vbCrLf & _
           "Fits to 1 page portrait. Use File > Print to preview.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    MsgBox "SetParameterizedPrintArea error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' CreatePrintableExecSummary - One-page CFO handout sheet (#64)
' Assembles 4 KPI cells, a divider, and a 4-product breakdown table on a
' dedicated "Exec Summary - Print" sheet with full print page setup applied.
' Hand this to the CFO/CEO at the meeting or export it to PDF.
'===============================================================================
Public Sub CreatePrintableExecSummary()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_PL_TREND) Then
        MsgBox "'" & SH_PL_TREND & "' not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Building printable executive summary...", 0.1

    Dim wsSrc As Worksheet: Set wsSrc = ThisWorkbook.Worksheets(SH_PL_TREND)

    ' Find FY Total column using multiple strategies (same as CreateExecutiveDashboard)
    Dim tLastCol As Long: tLastCol = modConfig.LastCol(wsSrc, HDR_ROW_REPORT)
    Dim fyCol As Long: fyCol = 0
    Dim fc As Long
    ' Pass 1: "FY Total" or "FY2025 Total"
    For fc = 2 To tLastCol
        Dim hdr As String: hdr = LCase(Trim(CStr(wsSrc.Cells(HDR_ROW_REPORT, fc).Value)))
        If InStr(hdr, "fy") > 0 And InStr(hdr, "total") > 0 Then fyCol = fc: Exit For
    Next fc
    ' Pass 2: "FY2025" or FISCAL_YEAR_4 pattern
    If fyCol = 0 Then
        For fc = 2 To tLastCol
            hdr = LCase(Trim(CStr(wsSrc.Cells(HDR_ROW_REPORT, fc).Value)))
            If InStr(hdr, "fy" & FISCAL_YEAR_4) > 0 Then fyCol = fc: Exit For
            If InStr(hdr, FISCAL_YEAR_4 & " total") > 0 Then fyCol = fc: Exit For
        Next fc
    End If
    ' Pass 3: standalone "total" / "year total" (skip column A to avoid label matches)
    If fyCol = 0 Then
        For fc = 2 To tLastCol
            hdr = LCase(Trim(CStr(wsSrc.Cells(HDR_ROW_REPORT, fc).Value)))
            If hdr = "total" Or hdr = "year total" Or hdr = "annual total" Then fyCol = fc: Exit For
        Next fc
    End If
    ' Pass 4: try modConfig.FindColByHeader with specific keywords
    If fyCol = 0 Then fyCol = modConfig.FindColByHeader(wsSrc, "2025 Total", HDR_ROW_REPORT)
    If fyCol = 0 Then fyCol = modConfig.FindColByHeader(wsSrc, "YTD", HDR_ROW_REPORT)
    If fyCol = 0 Then fyCol = modConfig.FindColByHeader(wsSrc, "Full Year", HDR_ROW_REPORT)
    ' Last resort: rightmost column
    If fyCol = 0 Then fyCol = tLastCol

    ' Pull key metrics — try multiple label variants (same as CreateExecutiveDashboard)
    Dim revRow  As Long
    revRow = modConfig.FindRowByLabel(wsSrc, "total revenue", DATA_ROW_REPORT)
    If revRow = 0 Then revRow = modConfig.FindRowByLabel(wsSrc, "consolidated revenue", DATA_ROW_REPORT)
    If revRow = 0 Then revRow = modConfig.FindRowByLabel(wsSrc, "net revenue", DATA_ROW_REPORT)

    Dim gpRow As Long
    gpRow = modConfig.FindRowByLabel(wsSrc, "gross profit", DATA_ROW_REPORT)
    If gpRow = 0 Then gpRow = modConfig.FindRowByLabel(wsSrc, "gross margin", DATA_ROW_REPORT)

    Dim opexRow As Long
    opexRow = modConfig.FindRowByLabel(wsSrc, "total operating expense", DATA_ROW_REPORT)
    If opexRow = 0 Then opexRow = modConfig.FindRowByLabel(wsSrc, "operating expense", DATA_ROW_REPORT)
    If opexRow = 0 Then opexRow = modConfig.FindRowByLabel(wsSrc, "total opex", DATA_ROW_REPORT)
    If opexRow = 0 Then opexRow = modConfig.FindRowByLabel(wsSrc, "total expenses", DATA_ROW_REPORT)

    Dim niRow As Long
    niRow = modConfig.FindRowByLabel(wsSrc, "net income", DATA_ROW_REPORT)
    If niRow = 0 Then niRow = modConfig.FindRowByLabel(wsSrc, "net operating income", DATA_ROW_REPORT)
    If niRow = 0 Then niRow = modConfig.FindRowByLabel(wsSrc, "operating income", DATA_ROW_REPORT)

    Dim fyRev  As Double: fyRev  = modConfig.SafeNum(wsSrc.Cells(revRow,  fyCol).Value)
    Dim fyGP   As Double: If gpRow   > 0 Then fyGP   = modConfig.SafeNum(wsSrc.Cells(gpRow,   fyCol).Value)
    Dim fyOpex As Double: If opexRow > 0 Then fyOpex = modConfig.SafeNum(wsSrc.Cells(opexRow, fyCol).Value)
    Dim fyNI   As Double: If niRow   > 0 Then fyNI   = modConfig.SafeNum(wsSrc.Cells(niRow,   fyCol).Value)
    Dim gmPct  As Double: If fyRev <> 0 Then gmPct = fyGP / fyRev

    ' Create / reset the print sheet
    Dim shName As String: shName = "Exec Summary - Print"
    modConfig.SafeDeleteSheet shName
    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsP.Name = shName
    wsP.Cells.Interior.Color = RGB(255, 255, 255)

    '── Title block ──────────────────────────────────────────────────────
    With wsP.Range("A1")
        .Value = "Keystone BenefitTech, Inc."
        .Font.Size = 16: .Font.Bold = True: .Font.Color = CLR_NAVY
    End With
    With wsP.Range("A2")
        .Value = "FY" & FISCAL_YEAR_4 & " Financial Summary"
        .Font.Size = 12: .Font.Color = RGB(80, 80, 80)
    End With
    wsP.Range("A3").Value = "Prepared: " & Format(Date, "mmmm d, yyyy")
    wsP.Range("A3").Font.Italic = True
    wsP.Range("A3").Font.Color = RGB(130, 130, 130)

    '── KPI row ──────────────────────────────────────────────────────────
    modPerformance.UpdateStatus "Writing KPI cells...", 0.4
    Dim kpiLabels As Variant: kpiLabels = Array("Total Revenue", "Gross Margin", "Operating Expenses", "Net Income")
    Dim kpiVals   As Variant: kpiVals   = Array(fyRev, gmPct, fyOpex, fyNI)
    Dim kpiFmts   As Variant: kpiFmts   = Array("$#,##0", "0.0%", "$#,##0", "$#,##0")
    Dim kpiRow    As Long:    kpiRow = 5

    Dim k As Long
    For k = 0 To 3
        Dim kCol As Long: kCol = (k * 2) + 1
        wsP.Cells(kpiRow - 1, kCol).Interior.Color = CLR_NAVY
        wsP.Cells(kpiRow - 1, kCol).RowHeight = 4
        wsP.Cells(kpiRow, kCol).Value = kpiLabels(k)
        wsP.Cells(kpiRow, kCol).Font.Size = 8
        wsP.Cells(kpiRow, kCol).Font.Color = RGB(100, 100, 100)
        wsP.Cells(kpiRow + 1, kCol).Value = kpiVals(k)
        wsP.Cells(kpiRow + 1, kCol).NumberFormat = kpiFmts(k)
        wsP.Cells(kpiRow + 1, kCol).Font.Size = 18
        wsP.Cells(kpiRow + 1, kCol).Font.Bold = True
        wsP.Cells(kpiRow + 1, kCol).Font.Color = CLR_NAVY
    Next k

    '── Divider ──────────────────────────────────────────────────────────
    Dim divRow As Long: divRow = kpiRow + 3
    wsP.Range(wsP.Cells(divRow, 1), wsP.Cells(divRow, 8)).Interior.Color = RGB(200, 200, 200)
    wsP.Range(wsP.Cells(divRow, 1), wsP.Cells(divRow, 8)).RowHeight = 2

    '── Product breakdown table ──────────────────────────────────────────
    modPerformance.UpdateStatus "Building product breakdown...", 0.7
    Dim tblRow As Long: tblRow = divRow + 2
    modConfig.StyleHeader wsP, tblRow, Array("Product", "FY Revenue", "Gross Margin %", "Net Income")

    Dim products As Variant: products = modConfig.GetProducts()
    Dim pr As Long
    For pr = 0 To UBound(products)
        Dim pRevRow As Long
        pRevRow = modConfig.FindRowByLabel(wsSrc, CStr(products(pr)), DATA_ROW_REPORT)
        Dim pRev As Double: pRev = 0
        If pRevRow > 0 Then pRev = modConfig.SafeNum(wsSrc.Cells(pRevRow, fyCol).Value)

        wsP.Cells(tblRow + 1 + pr, 1).Value = CStr(products(pr))
        wsP.Cells(tblRow + 1 + pr, 1).Font.Bold = True
        wsP.Cells(tblRow + 1 + pr, 2).Value = pRev
        wsP.Cells(tblRow + 1 + pr, 2).NumberFormat = "$#,##0"
        ' Avoid IIf — VBA evaluates both branches, causing Overflow when fyRev=0
        If fyRev <> 0 Then
            wsP.Cells(tblRow + 1 + pr, 3).Value = gmPct
            wsP.Cells(tblRow + 1 + pr, 4).Value = fyNI * (pRev / fyRev)
        Else
            wsP.Cells(tblRow + 1 + pr, 3).Value = 0
            wsP.Cells(tblRow + 1 + pr, 4).Value = 0
        End If
        wsP.Cells(tblRow + 1 + pr, 3).NumberFormat = "0.0%"
        wsP.Cells(tblRow + 1 + pr, 4).NumberFormat = "$#,##0"
    Next pr

    '── Page setup ───────────────────────────────────────────────────────
    With wsP.PageSetup
        .Orientation        = xlPortrait
        .FitToPagesWide     = 1
        .FitToPagesTall     = 1
        .Zoom               = False
        .PrintArea          = wsP.UsedRange.Address
        .CenterHorizontally = True
        .LeftHeader         = "Keystone BenefitTech, Inc."
        .RightHeader        = "CONFIDENTIAL"
        .CenterFooter       = "Page &P"
    End With

    wsP.Columns("A:H").AutoFit
    wsP.Tab.Color = CLR_NAVY
    wsP.Activate

    Dim elapsed As Double: elapsed = modPerformance.ElapsedSeconds()
    modPerformance.TurboOff
    modLogger.LogAction "modDemoTools", "CreatePrintableExecSummary", _
        "Print sheet created: " & shName & " (" & Format(elapsed, "0.0") & "s)"
    MsgBox "'" & shName & "' is ready." & vbCrLf & _
           "Use File > Print or Ctrl+P to print or save as PDF.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    MsgBox "CreatePrintableExecSummary error: " & Err.Description, vbCritical, APP_NAME
End Sub
