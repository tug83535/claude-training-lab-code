Attribute VB_Name = "modDrillDown"
Option Explicit

'===============================================================================
' modDrillDown - Reconciliation Drill & Comparison Tools
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Makes the Checks tab interactive and self-diagnosing.
'           Lets users jump from a failing check to the source transactions.
'
' PUBLIC SUBS:
'   AddReconciliationDrillLinks       - Hyperlinks from Checks rows to GL data (#18)
'   AutoPopulateReconciliationChecks  - Recalculate + verify named ranges (#55)
'   ApplyReconciliationHeatmap        - Color Checks tab by variance size (#56)
'   RunGoldenFileCompare              - Compare current P&L to saved baseline (#90)
'
' VERSION:  2.1.0 (New module — 2026-03-01)
' SOURCE:   Ideas from NewTesting/VBA Examples (200) — items #18, #55, #56, #90
'===============================================================================

'===============================================================================
' AddReconciliationDrillLinks - Hyperlinks from each Checks row to GL data (#18)
' Adds a "View Data" hyperlink in column F of every check row that jumps
' directly to the GL sheet (made visible). Column F gets a header on first run.
'===============================================================================
Public Sub AddReconciliationDrillLinks()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_CHECKS) Then
        MsgBox "Checks sheet '" & SH_CHECKS & "' not found.", vbCritical, APP_NAME
        Exit Sub
    End If
    If Not modConfig.SheetExists(SH_HIDDEN) Then
        MsgBox "GL sheet '" & SH_HIDDEN & "' not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    ' GL must be visible for internal hyperlinks to function
    modConfig.ShowGLSheet

    Dim wsChk As Worksheet: Set wsChk = ThisWorkbook.Worksheets(SH_CHECKS)
    Dim lastRow As Long: lastRow = modConfig.LastRow(wsChk, 1)

    ' Write column F header if blank
    If wsChk.Cells(HDR_ROW_CHECKS, 6).Value = "" Then
        wsChk.Cells(HDR_ROW_CHECKS, 6).Value = "Drill To Data"
        wsChk.Cells(HDR_ROW_CHECKS, 6).Font.Bold = True
        wsChk.Cells(HDR_ROW_CHECKS, 6).Interior.Color = CLR_NAVY
        wsChk.Cells(HDR_ROW_CHECKS, 6).Font.Color = CLR_WHITE
    End If

    Dim linkCount As Long: linkCount = 0
    Dim r As Long
    For r = DATA_ROW_CHECKS To lastRow
        Dim checkName As String: checkName = modConfig.SafeStr(wsChk.Cells(r, 1).Value)
        If Len(checkName) > 0 Then
            wsChk.Cells(r, 6).Hyperlinks.Delete
            wsChk.Hyperlinks.Add _
                Anchor:=wsChk.Cells(r, 6), _
                Address:="", _
                SubAddress:="'" & SH_HIDDEN & "'!A1", _
                TextToDisplay:="View Data"
            wsChk.Cells(r, 6).Font.Color = RGB(0, 70, 180)
            linkCount = linkCount + 1
        End If
    Next r

    ' Set GL to regular hidden (not very-hidden) so drill hyperlinks can navigate to it
    If modConfig.SheetExists(SH_HIDDEN) Then
        ThisWorkbook.Worksheets(SH_HIDDEN).Visible = xlSheetHidden
    End If
    modLogger.LogAction "modDrillDown", "AddReconciliationDrillLinks", _
        linkCount & " drill links added to " & SH_CHECKS
    MsgBox linkCount & " 'View Data' hyperlinks added to the Checks tab (column F)." & vbCrLf & _
           "Click any link to jump directly to the GL transaction data.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modConfig.HideGLSheet
    MsgBox "AddReconciliationDrillLinks error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' AutoPopulateReconciliationChecks - Force recalc + verify named ranges (#55)
' Triggers Application.CalculateFull to refresh all Checks tab formulas,
' timestamps the sheet, and reports any named ranges that are missing.
'===============================================================================
Public Sub AutoPopulateReconciliationChecks()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_CHECKS) Then
        MsgBox "Checks sheet '" & SH_CHECKS & "' not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    ' Ensure revenue share named ranges exist before recalculating
    modConfig.AddNamedRanges

    Application.CalculateFull

    Dim wsChk As Worksheet: Set wsChk = ThisWorkbook.Worksheets(SH_CHECKS)
    wsChk.Range("E1").Value = "Last Refreshed: " & Format(Now, "yyyy-mm-dd hh:nn:ss")

    ' Verify key named ranges that drive check formulas
    Dim requiredNames As Variant
    requiredNames = Array("drv_iGO_RevShare", "drv_Affirm_RevShare", _
                          "drv_InsureSight_RevShare", "drv_DocFast_RevShare")
    Dim missingNames As String
    Dim nm As Variant
    For Each nm In requiredNames
        On Error Resume Next
        Dim testRef As String: testRef = ThisWorkbook.Names(CStr(nm)).RefersTo
        If Err.Number <> 0 Then
            missingNames = missingNames & vbCrLf & "  " & CStr(nm)
        End If
        Err.Clear
        On Error GoTo ErrHandler
    Next nm

    modLogger.LogAction "modDrillDown", "AutoPopulateReconciliationChecks", _
        "Checks recalculated | Missing: " & IIf(Len(missingNames) = 0, "none", missingNames)

    If Len(missingNames) > 0 Then
        MsgBox "Checks tab refreshed." & vbCrLf & vbCrLf & _
               "WARNING: These named ranges could not be created:" & _
               missingNames & vbCrLf & vbCrLf & _
               "Label rows on Assumptions with product name + 'Rev Share' in column A.", _
               vbExclamation, APP_NAME
    Else
        MsgBox "Checks tab refreshed. All named ranges verified.", vbInformation, APP_NAME
    End If
    Exit Sub

ErrHandler:
    MsgBox "AutoPopulateReconciliationChecks error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' ApplyReconciliationHeatmap - Color Checks tab by variance size (#56)
' Reads column D (Difference) and applies a 3-color scale:
'   Green  = |difference| < $1     (OK)
'   Yellow = $1 to $100            (review)
'   Red    = > $100                (action required)
' Also re-colors the Status column (col E) for PASS / FAIL.
'===============================================================================
Public Sub ApplyReconciliationHeatmap()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_CHECKS) Then
        MsgBox "Checks sheet '" & SH_CHECKS & "' not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    Dim wsChk As Worksheet: Set wsChk = ThisWorkbook.Worksheets(SH_CHECKS)
    Dim lastRow As Long: lastRow = modConfig.LastRow(wsChk, 1)

    If lastRow < DATA_ROW_CHECKS Then
        MsgBox "No check data found on the Checks sheet.", vbInformation, APP_NAME
        Exit Sub
    End If

    Dim r As Long
    For r = DATA_ROW_CHECKS To lastRow
        If Len(modConfig.SafeStr(wsChk.Cells(r, 1).Value)) = 0 Then GoTo NextRow

        ' Column D — absolute difference
        Dim diffVal As Double: diffVal = Abs(modConfig.SafeNum(wsChk.Cells(r, 4).Value))
        Dim diffColor As Long
        If diffVal < 1 Then
            diffColor = RGB(200, 255, 200)   ' Green — below $1
        ElseIf diffVal < 100 Then
            diffColor = RGB(255, 255, 180)   ' Yellow — $1 to $100
        Else
            diffColor = RGB(255, 200, 200)   ' Red — over $100
        End If
        wsChk.Cells(r, 4).Interior.Color = diffColor

        ' Column E — PASS / FAIL
        Dim statusText As String
        statusText = UCase(modConfig.SafeStr(wsChk.Cells(r, COL_CHECK_STATUS).Value))
        If statusText = "PASS" Then
            wsChk.Cells(r, COL_CHECK_STATUS).Interior.Color = RGB(200, 255, 200)
            wsChk.Cells(r, COL_CHECK_STATUS).Font.Color     = RGB(0, 100, 0)
        ElseIf statusText = "FAIL" Then
            wsChk.Cells(r, COL_CHECK_STATUS).Interior.Color = RGB(255, 200, 200)
            wsChk.Cells(r, COL_CHECK_STATUS).Font.Color     = RGB(150, 0, 0)
        End If

NextRow:
    Next r

    modLogger.LogAction "modDrillDown", "ApplyReconciliationHeatmap", _
        "Heatmap applied (" & (lastRow - DATA_ROW_CHECKS + 1) & " rows)"
    MsgBox "Reconciliation heatmap applied." & vbCrLf & _
           "Green = OK  |  Yellow = Review  |  Red = Action Required", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    MsgBox "ApplyReconciliationHeatmap error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' RunGoldenFileCompare - Compare current P&L to saved baseline (#90)
' FIRST RUN:  Saves current FY Total column values on a very-hidden sheet
'             called "GoldenBaseline". Shows confirmation before saving.
' LATER RUNS: Compares current values to the saved baseline and writes a
'             "Golden Compare Report" sheet showing every change >= $1.
'===============================================================================
Public Sub RunGoldenFileCompare()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_PL_TREND) Then
        MsgBox "'" & SH_PL_TREND & "' not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    Dim goldenName As String: goldenName = "GoldenBaseline"
    Dim wsSrc As Worksheet: Set wsSrc = ThisWorkbook.Worksheets(SH_PL_TREND)
    Dim fyCol As Long: fyCol = modConfig.FindColByHeader(wsSrc, "total", HDR_ROW_REPORT)
    If fyCol = 0 Then fyCol = modConfig.LastCol(wsSrc, HDR_ROW_REPORT)
    Dim lastRow As Long: lastRow = modConfig.LastRow(wsSrc, 1)

    '── First run: save baseline ─────────────────────────────────────────
    If Not modConfig.SheetExists(goldenName) Then
        If MsgBox("No golden baseline exists yet." & vbCrLf & vbCrLf & _
                  "Save current FY Total values as the baseline now?", _
                  vbYesNo + vbQuestion, APP_NAME) = vbNo Then Exit Sub

        Dim wsGold As Worksheet
        Set wsGold = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsGold.Name    = goldenName
        wsGold.Visible = xlSheetVeryHidden
        wsGold.Cells(1, 1).Value = "SavedOn"
        wsGold.Cells(1, 2).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
        wsGold.Cells(2, 1).Value = "Label"
        wsGold.Cells(2, 2).Value = "GoldenValue"

        Dim r As Long
        For r = DATA_ROW_REPORT To lastRow
            Dim lbl As String: lbl = modConfig.SafeStr(wsSrc.Cells(r, 1).Value)
            If Len(lbl) > 0 Then
                wsGold.Cells(r + 1, 1).Value = lbl
                wsGold.Cells(r + 1, 2).Value = modConfig.SafeNum(wsSrc.Cells(r, fyCol).Value)
            End If
        Next r

        modLogger.LogAction "modDrillDown", "RunGoldenFileCompare", "Golden baseline saved"
        MsgBox "Golden baseline saved (" & (lastRow - DATA_ROW_REPORT + 1) & " rows)." & vbCrLf & _
               "Run this macro again after making changes to compare.", _
               vbInformation, APP_NAME
        Exit Sub
    End If

    '── Later runs: compare to baseline ─────────────────────────────────
    Dim wsGold2 As Worksheet: Set wsGold2 = ThisWorkbook.Worksheets(goldenName)
    Dim goldLastRow As Long: goldLastRow = modConfig.LastRow(wsGold2, 1)

    Dim rptName As String: rptName = "Golden Compare Report"
    modConfig.SafeDeleteSheet rptName
    Dim wsRpt As Worksheet
    Set wsRpt = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsRpt.Name = rptName

    modConfig.StyleHeader wsRpt, 1, _
        Array("Row Label", "Golden Value", "Current Value", "Difference", "Status")

    Dim outRow    As Long: outRow    = 2
    Dim diffCount As Long: diffCount = 0
    Dim g As Long
    For g = 3 To goldLastRow
        Dim goldenLbl As String: goldenLbl = modConfig.SafeStr(wsGold2.Cells(g, 1).Value)
        If Len(goldenLbl) = 0 Then GoTo NextG

        Dim goldenVal As Double: goldenVal = modConfig.SafeNum(wsGold2.Cells(g, 2).Value)
        Dim curRow As Long: curRow = modConfig.FindRowByLabel(wsSrc, goldenLbl, DATA_ROW_REPORT)
        Dim curVal As Double: curVal = 0
        If curRow > 0 Then curVal = modConfig.SafeNum(wsSrc.Cells(curRow, fyCol).Value)

        Dim diff   As Double: diff = curVal - goldenVal
        Dim stat   As String: stat = IIf(Abs(diff) < 1, "MATCH", "CHANGED")
        If stat = "CHANGED" Then diffCount = diffCount + 1

        wsRpt.Cells(outRow, 1).Value = goldenLbl
        wsRpt.Cells(outRow, 2).Value = goldenVal:  wsRpt.Cells(outRow, 2).NumberFormat = "$#,##0"
        wsRpt.Cells(outRow, 3).Value = curVal:     wsRpt.Cells(outRow, 3).NumberFormat = "$#,##0"
        wsRpt.Cells(outRow, 4).Value = diff:       wsRpt.Cells(outRow, 4).NumberFormat = "$#,##0;($#,##0)"
        wsRpt.Cells(outRow, 5).Value = stat
        If stat = "CHANGED" Then
            wsRpt.Cells(outRow, 5).Interior.Color = RGB(255, 200, 200)
            wsRpt.Cells(outRow, 4).Font.Color     = RGB(150, 0, 0)
        Else
            wsRpt.Cells(outRow, 5).Interior.Color = RGB(200, 255, 200)
        End If
        outRow = outRow + 1
NextG:
    Next g

    wsRpt.Columns("A:E").AutoFit
    wsRpt.Tab.Color = IIf(diffCount > 0, RGB(192, 0, 0), RGB(0, 176, 80))
    wsRpt.Activate

    modLogger.LogAction "modDrillDown", "RunGoldenFileCompare", _
        diffCount & " change(s) detected vs golden baseline"
    MsgBox "Golden file comparison complete." & vbCrLf & _
           diffCount & " line(s) changed since the baseline was saved." & vbCrLf & _
           "Full results on '" & rptName & "'.", _
           IIf(diffCount > 0, vbExclamation, vbInformation), APP_NAME
    Exit Sub

ErrHandler:
    MsgBox "RunGoldenFileCompare error: " & Err.Description, vbCritical, APP_NAME
End Sub
