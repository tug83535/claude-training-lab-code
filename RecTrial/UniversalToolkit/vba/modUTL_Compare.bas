Attribute VB_Name = "modUTL_Compare"
'==============================================================================
' modUTL_Compare — Sheet & Range Comparison Tool
'==============================================================================
' PURPOSE:  Compare two sheets or two ranges cell-by-cell. Highlights
'           differences and builds a styled summary report.
'
' PUBLIC SUBS:
'   CompareSheets      — Pick two sheets, compare cell-by-cell
'   CompareRanges      — Compare two user-selected ranges
'   ClearCompareHighlights — Remove comparison highlighting
'
' DEPENDENCIES: None (standalone). Works in any Excel workbook.
' VERSION:  1.0.0 | DATE: 2026-03-12
'==============================================================================
Option Explicit

Private Const REPORT_SHEET As String = "UTL_CompareReport"
Private Const CLR_DIFF As Long = 255        ' Red highlight for differences
Private Const CLR_MATCH As Long = 5287936   ' Green for matches (RGB 80,200,0)
Private Const CLR_HDR As Long = 7948043     ' RGB(11,71,121) iPipeline Blue

'==============================================================================
' PUBLIC: CompareSheets
' User picks two sheets from a list. Compares cell-by-cell within UsedRange.
'==============================================================================
Public Sub CompareSheets()
    On Error GoTo ErrHandler

    If ThisWorkbook.Sheets.Count < 2 Then
        MsgBox "You need at least 2 sheets to compare.", vbExclamation, "Compare Sheets"
        Exit Sub
    End If

    '--- Build sheet list for user ---
    Dim sheetList As String
    sheetList = "Available sheets:" & vbCrLf & String(35, "-") & vbCrLf

    Dim i As Long
    For i = 1 To ThisWorkbook.Sheets.Count
        sheetList = sheetList & "  " & i & ". " & ThisWorkbook.Sheets(i).Name & vbCrLf
    Next i

    '--- Get first sheet ---
    sheetList = sheetList & vbCrLf & "Enter the NUMBER of the FIRST sheet to compare:"
    Dim choice1 As String
    choice1 = InputBox(sheetList, "Compare Sheets - Step 1 of 3")
    If Len(Trim(choice1)) = 0 Then Exit Sub
    If Not IsNumeric(choice1) Then
        MsgBox "Please enter a number.", vbExclamation, "Compare Sheets"
        Exit Sub
    End If
    Dim idx1 As Long
    idx1 = CLng(choice1)
    If idx1 < 1 Or idx1 > ThisWorkbook.Sheets.Count Then
        MsgBox "Invalid sheet number.", vbExclamation, "Compare Sheets"
        Exit Sub
    End If

    '--- Get second sheet ---
    Dim choice2 As String
    choice2 = InputBox("Enter the NUMBER of the SECOND sheet to compare:" & vbCrLf & vbCrLf & _
                        "(Comparing against: " & ThisWorkbook.Sheets(idx1).Name & ")", _
                        "Compare Sheets - Step 2 of 3")
    If Len(Trim(choice2)) = 0 Then Exit Sub
    If Not IsNumeric(choice2) Then
        MsgBox "Please enter a number.", vbExclamation, "Compare Sheets"
        Exit Sub
    End If
    Dim idx2 As Long
    idx2 = CLng(choice2)
    If idx2 < 1 Or idx2 > ThisWorkbook.Sheets.Count Then
        MsgBox "Invalid sheet number.", vbExclamation, "Compare Sheets"
        Exit Sub
    End If
    If idx1 = idx2 Then
        MsgBox "Please pick two different sheets.", vbExclamation, "Compare Sheets"
        Exit Sub
    End If

    '--- Ask about highlighting ---
    Dim highlightChoice As VbMsgBoxResult
    highlightChoice = MsgBox("Do you want to highlight differences directly on the sheets?" & vbCrLf & vbCrLf & _
                              "YES = Highlight differences in red on both sheets" & vbCrLf & _
                              "NO = Only create the summary report (no changes to your sheets)", _
                              vbYesNoCancel + vbQuestion, "Compare Sheets - Step 3 of 3")
    If highlightChoice = vbCancel Then Exit Sub

    Dim ws1 As Worksheet, ws2 As Worksheet
    Set ws1 = ThisWorkbook.Sheets(idx1)
    Set ws2 = ThisWorkbook.Sheets(idx2)

    Application.ScreenUpdating = False
    Application.StatusBar = "Comparing sheets..."

    '--- Determine comparison range ---
    Dim maxRow As Long, maxCol As Long
    Dim lastRow1 As Long, lastRow2 As Long
    Dim lastCol1 As Long, lastCol2 As Long

    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    lastCol1 = ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column
    lastCol2 = ws2.Cells(1, ws2.Columns.Count).End(xlToLeft).Column

    ' Use UsedRange as fallback for better detection
    If ws1.UsedRange.Rows.Count + ws1.UsedRange.Row - 1 > lastRow1 Then
        lastRow1 = ws1.UsedRange.Rows.Count + ws1.UsedRange.Row - 1
    End If
    If ws2.UsedRange.Rows.Count + ws2.UsedRange.Row - 1 > lastRow2 Then
        lastRow2 = ws2.UsedRange.Rows.Count + ws2.UsedRange.Row - 1
    End If
    If ws1.UsedRange.Columns.Count + ws1.UsedRange.Column - 1 > lastCol1 Then
        lastCol1 = ws1.UsedRange.Columns.Count + ws1.UsedRange.Column - 1
    End If
    If ws2.UsedRange.Columns.Count + ws2.UsedRange.Column - 1 > lastCol2 Then
        lastCol2 = ws2.UsedRange.Columns.Count + ws2.UsedRange.Column - 1
    End If

    If lastRow1 > lastRow2 Then maxRow = lastRow1 Else maxRow = lastRow2
    If lastCol1 > lastCol2 Then maxCol = lastCol1 Else maxCol = lastCol2

    ' Safety cap
    If maxRow > 10000 Then maxRow = 10000
    If maxCol > 256 Then maxCol = 256

    '--- Compare cell by cell ---
    Dim diffCount As Long, matchCount As Long, totalCells As Long
    diffCount = 0
    matchCount = 0
    totalCells = maxRow * maxCol

    ' Store differences for report (cap at 5000)
    Dim diffCells() As String, diffVal1() As String, diffVal2() As String
    ReDim diffCells(1 To 5000)
    ReDim diffVal1(1 To 5000)
    ReDim diffVal2(1 To 5000)

    Dim r As Long, c As Long
    For r = 1 To maxRow
        If r Mod 500 = 0 Then
            Application.StatusBar = "Comparing row " & r & " of " & maxRow & "..."
        End If
        For c = 1 To maxCol
            Dim v1 As Variant, v2 As Variant
            v1 = ws1.Cells(r, c).Value
            v2 = ws2.Cells(r, c).Value

            If Not CellsMatch(v1, v2) Then
                diffCount = diffCount + 1

                If diffCount <= 5000 Then
                    diffCells(diffCount) = ws1.Cells(r, c).Address(False, False)
                    diffVal1(diffCount) = Left(CStr(Nz(v1, "")), 100)
                    diffVal2(diffCount) = Left(CStr(Nz(v2, "")), 100)
                End If

                If highlightChoice = vbYes Then
                    ws1.Cells(r, c).Interior.Color = CLR_DIFF
                    ws2.Cells(r, c).Interior.Color = CLR_DIFF
                End If
            Else
                matchCount = matchCount + 1
            End If
        Next c
    Next r

    '--- Build report ---
    BuildCompareReport ws1.Name, ws2.Name, diffCount, matchCount, totalCells, _
                       diffCells, diffVal1, diffVal2

    Application.StatusBar = False
    Application.ScreenUpdating = True

    Dim pctMatch As Double
    If totalCells > 0 Then pctMatch = Round((matchCount / totalCells) * 100, 1) Else pctMatch = 100

    MsgBox "Comparison complete!" & vbCrLf & vbCrLf & _
           "Sheet 1: " & ws1.Name & vbCrLf & _
           "Sheet 2: " & ws2.Name & vbCrLf & _
           String(30, "-") & vbCrLf & _
           "Total cells compared: " & Format(totalCells, "#,##0") & vbCrLf & _
           "Matches: " & Format(matchCount, "#,##0") & vbCrLf & _
           "Differences: " & Format(diffCount, "#,##0") & vbCrLf & _
           "Match rate: " & pctMatch & "%" & vbCrLf & vbCrLf & _
           "See '" & REPORT_SHEET & "' for details.", _
           vbInformation, "Compare Sheets"

    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Compare Sheets"
End Sub

'==============================================================================
' PUBLIC: CompareRanges
' User selects two ranges (can be on different sheets). Compares cell-by-cell.
'==============================================================================
Public Sub CompareRanges()
    On Error GoTo ErrHandler

    MsgBox "You will be asked to select TWO ranges to compare." & vbCrLf & vbCrLf & _
           "The ranges should be the same size." & vbCrLf & _
           "They can be on different sheets." & vbCrLf & vbCrLf & _
           "Click OK, then select the FIRST range.", _
           vbInformation, "Compare Ranges"

    Dim rng1 As Range
    On Error Resume Next
    Set rng1 = Application.InputBox("Select the FIRST range to compare:", _
                                     "Compare Ranges - Range 1", Type:=8)
    On Error GoTo ErrHandler
    If rng1 Is Nothing Then Exit Sub

    MsgBox "First range selected: " & rng1.Address(External:=True) & vbCrLf & _
           "(" & rng1.Rows.Count & " rows x " & rng1.Columns.Count & " columns)" & vbCrLf & vbCrLf & _
           "Click OK, then select the SECOND range.", _
           vbInformation, "Compare Ranges"

    Dim rng2 As Range
    On Error Resume Next
    Set rng2 = Application.InputBox("Select the SECOND range to compare:", _
                                     "Compare Ranges - Range 2", Type:=8)
    On Error GoTo ErrHandler
    If rng2 Is Nothing Then Exit Sub

    '--- Validate sizes match ---
    If rng1.Rows.Count <> rng2.Rows.Count Or rng1.Columns.Count <> rng2.Columns.Count Then
        Dim sizeMsg As VbMsgBoxResult
        sizeMsg = MsgBox("Range sizes don't match:" & vbCrLf & _
                          "Range 1: " & rng1.Rows.Count & " rows x " & rng1.Columns.Count & " cols" & vbCrLf & _
                          "Range 2: " & rng2.Rows.Count & " rows x " & rng2.Columns.Count & " cols" & vbCrLf & vbCrLf & _
                          "Compare anyway? (will use the smaller dimensions)", _
                          vbYesNo + vbQuestion, "Compare Ranges")
        If sizeMsg = vbNo Then Exit Sub
    End If

    '--- Ask about highlighting ---
    Dim doHighlight As VbMsgBoxResult
    doHighlight = MsgBox("Highlight differences in red on both ranges?" & vbCrLf & vbCrLf & _
                          "YES = Highlight differences" & vbCrLf & _
                          "NO = Report only (no changes)", _
                          vbYesNoCancel + vbQuestion, "Compare Ranges")
    If doHighlight = vbCancel Then Exit Sub

    Application.ScreenUpdating = False
    Application.StatusBar = "Comparing ranges..."

    Dim maxRow As Long, maxCol As Long
    If rng1.Rows.Count < rng2.Rows.Count Then maxRow = rng1.Rows.Count Else maxRow = rng2.Rows.Count
    If rng1.Columns.Count < rng2.Columns.Count Then maxCol = rng1.Columns.Count Else maxCol = rng2.Columns.Count

    Dim diffCount As Long, matchCount As Long, totalCells As Long
    diffCount = 0
    matchCount = 0
    totalCells = maxRow * maxCol

    Dim diffCells() As String, diffV1() As String, diffV2() As String
    ReDim diffCells(1 To 5000)
    ReDim diffV1(1 To 5000)
    ReDim diffV2(1 To 5000)

    Dim r As Long, c As Long
    For r = 1 To maxRow
        For c = 1 To maxCol
            Dim val1 As Variant, val2 As Variant
            val1 = rng1.Cells(r, c).Value
            val2 = rng2.Cells(r, c).Value

            If Not CellsMatch(val1, val2) Then
                diffCount = diffCount + 1
                If diffCount <= 5000 Then
                    diffCells(diffCount) = rng1.Cells(r, c).Address(False, False)
                    diffV1(diffCount) = Left(CStr(Nz(val1, "")), 100)
                    diffV2(diffCount) = Left(CStr(Nz(val2, "")), 100)
                End If
                If doHighlight = vbYes Then
                    rng1.Cells(r, c).Interior.Color = CLR_DIFF
                    rng2.Cells(r, c).Interior.Color = CLR_DIFF
                End If
            Else
                matchCount = matchCount + 1
            End If
        Next c
    Next r

    Dim src1 As String, src2 As String
    src1 = rng1.Parent.Name & "!" & rng1.Address(False, False)
    src2 = rng2.Parent.Name & "!" & rng2.Address(False, False)

    BuildCompareReport src1, src2, diffCount, matchCount, totalCells, _
                       diffCells, diffV1, diffV2

    Application.StatusBar = False
    Application.ScreenUpdating = True

    Dim pct As Double
    If totalCells > 0 Then pct = Round((matchCount / totalCells) * 100, 1) Else pct = 100

    MsgBox "Comparison complete!" & vbCrLf & vbCrLf & _
           "Differences: " & Format(diffCount, "#,##0") & " of " & Format(totalCells, "#,##0") & " cells" & vbCrLf & _
           "Match rate: " & pct & "%" & vbCrLf & vbCrLf & _
           "See '" & REPORT_SHEET & "' for details.", _
           vbInformation, "Compare Ranges"

    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Compare Ranges"
End Sub

'==============================================================================
' PUBLIC: ClearCompareHighlights
' Removes red highlighting from comparison on a user-selected sheet.
'==============================================================================
Public Sub ClearCompareHighlights()
    On Error GoTo ErrHandler

    Dim choice As VbMsgBoxResult
    choice = MsgBox("Clear comparison highlighting from:" & vbCrLf & vbCrLf & _
                     "YES = Active sheet only" & vbCrLf & _
                     "NO = ALL sheets in workbook", _
                     vbYesNoCancel + vbQuestion, "Clear Highlights")
    If choice = vbCancel Then Exit Sub

    Application.ScreenUpdating = False

    Dim ws As Worksheet
    If choice = vbYes Then
        ClearRedHighlights ActiveSheet
        MsgBox "Highlights cleared from " & ActiveSheet.Name & ".", vbInformation, "Clear Highlights"
    Else
        For Each ws In ThisWorkbook.Worksheets
            ClearRedHighlights ws
        Next ws
        MsgBox "Highlights cleared from all sheets.", vbInformation, "Clear Highlights"
    End If

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Clear Highlights"
End Sub

'==============================================================================
' PRIVATE: CellsMatch — Compare two values with type awareness
'==============================================================================
Private Function CellsMatch(ByVal v1 As Variant, ByVal v2 As Variant) As Boolean
    ' Both empty = match
    If IsEmpty(v1) And IsEmpty(v2) Then CellsMatch = True: Exit Function
    If IsEmpty(v1) Or IsEmpty(v2) Then CellsMatch = False: Exit Function

    ' Both errors
    If IsError(v1) And IsError(v2) Then
        CellsMatch = (CStr(v1) = CStr(v2))
        Exit Function
    End If
    If IsError(v1) Or IsError(v2) Then CellsMatch = False: Exit Function

    ' Numeric comparison with tolerance for floating point
    If IsNumeric(v1) And IsNumeric(v2) Then
        If Abs(CDbl(v1) - CDbl(v2)) < 0.0001 Then
            CellsMatch = True
        Else
            CellsMatch = False
        End If
        Exit Function
    End If

    ' String comparison (case-insensitive)
    CellsMatch = (CStr(v1) = CStr(v2))
End Function

'==============================================================================
' PRIVATE: Nz — Null/Empty to string
'==============================================================================
Private Function Nz(ByVal v As Variant, ByVal defaultVal As String) As String
    If IsEmpty(v) Then
        Nz = defaultVal
    ElseIf IsError(v) Then
        Nz = CStr(v)
    ElseIf IsNull(v) Then
        Nz = defaultVal
    Else
        Nz = CStr(v)
    End If
End Function

'==============================================================================
' PRIVATE: ClearRedHighlights — Remove red interior color from used range
'==============================================================================
Private Sub ClearRedHighlights(ByVal ws As Worksheet)
    Dim cell As Range
    Dim rng As Range

    Set rng = Nothing
    On Error Resume Next
    Set rng = ws.UsedRange.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0

    ' Also check cells with formulas
    Dim rngF As Range
    Set rngF = Nothing
    On Error Resume Next
    Set rngF = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    ' Combine ranges
    Dim combined As Range
    If Not rng Is Nothing And Not rngF Is Nothing Then
        Set combined = Union(rng, rngF)
    ElseIf Not rng Is Nothing Then
        Set combined = rng
    ElseIf Not rngF Is Nothing Then
        Set combined = rngF
    Else
        Exit Sub
    End If

    For Each cell In combined.Cells
        If cell.Interior.Color = CLR_DIFF Then
            cell.Interior.ColorIndex = xlNone
        End If
    Next cell
End Sub

'==============================================================================
' PRIVATE: BuildCompareReport — Create styled summary sheet
'==============================================================================
Private Sub BuildCompareReport(ByVal name1 As String, ByVal name2 As String, _
                                ByVal diffCount As Long, ByVal matchCount As Long, _
                                ByVal totalCells As Long, _
                                ByRef diffCells() As String, _
                                ByRef diffV1() As String, _
                                ByRef diffV2() As String)

    '--- Create or clear report sheet ---
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(REPORT_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = REPORT_SHEET
    Else
        ws.Cells.Clear
    End If

    '--- Header ---
    ws.Range("A1").Value = "Sheet Comparison Report"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    ws.Range("A2").Value = "Generated: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Range("A2").Font.Italic = True

    '--- Summary ---
    ws.Range("A4").Value = "Source 1:"
    ws.Range("B4").Value = name1
    ws.Range("A5").Value = "Source 2:"
    ws.Range("B5").Value = name2
    ws.Range("A4:A5").Font.Bold = True

    ws.Range("A7").Value = "Total Cells:"
    ws.Range("B7").Value = totalCells
    ws.Range("B7").NumberFormat = "#,##0"
    ws.Range("A8").Value = "Matches:"
    ws.Range("B8").Value = matchCount
    ws.Range("B8").NumberFormat = "#,##0"
    ws.Range("A9").Value = "Differences:"
    ws.Range("B9").Value = diffCount
    ws.Range("B9").NumberFormat = "#,##0"
    ws.Range("B9").Font.Color = vbRed
    ws.Range("B9").Font.Bold = True
    ws.Range("A10").Value = "Match Rate:"
    If totalCells > 0 Then
        ws.Range("B10").Value = matchCount / totalCells
    Else
        ws.Range("B10").Value = 1
    End If
    ws.Range("B10").NumberFormat = "0.0%"
    ws.Range("A7:A10").Font.Bold = True

    '--- Difference detail ---
    If diffCount > 0 Then
        Dim hdr As Long
        hdr = 12
        ws.Cells(hdr, 1).Value = "Cell"
        ws.Cells(hdr, 2).Value = "Value in " & name1
        ws.Cells(hdr, 3).Value = "Value in " & name2

        Dim hdrRng As Range
        Set hdrRng = ws.Range(ws.Cells(hdr, 1), ws.Cells(hdr, 3))
        hdrRng.Font.Bold = True
        hdrRng.Font.Color = RGB(255, 255, 255)
        hdrRng.Interior.Color = CLR_HDR

        Dim showCount As Long
        If diffCount > 5000 Then showCount = 5000 Else showCount = diffCount

        Dim d As Long
        For d = 1 To showCount
            ws.Cells(hdr + d, 1).Value = diffCells(d)
            ws.Cells(hdr + d, 2).Value = diffV1(d)
            ws.Cells(hdr + d, 3).Value = diffV2(d)

            If d Mod 2 = 0 Then
                ws.Range(ws.Cells(hdr + d, 1), ws.Cells(hdr + d, 3)).Interior.Color = RGB(235, 241, 250)
            End If
        Next d

        If diffCount > 5000 Then
            ws.Cells(hdr + showCount + 2, 1).Value = "... and " & Format(diffCount - 5000, "#,##0") & " more differences (showing first 5,000)"
            ws.Cells(hdr + showCount + 2, 1).Font.Italic = True
        End If
    End If

    ws.Columns("A:C").AutoFit
    ws.Activate
    ws.Range("A1").Select
End Sub

'==============================================================================
' DIRECTOR WRAPPERS — Silent subs for video automation (no dialogs)
'==============================================================================

'==============================================================================
' DirectorCompareSheets
' Compares two sheets by name, creates diff report with color-coded differences.
' No InputBox/MsgBox. Highlights differences in red on both sheets.
'==============================================================================
Public Sub DirectorCompareSheets(sheet1Name As String, sheet2Name As String)
    On Error Resume Next

    If Len(sheet1Name) = 0 Or Len(sheet2Name) = 0 Then
        Debug.Print "[Director] CompareSheets: Both sheet names required."
        Exit Sub
    End If
    If sheet1Name = sheet2Name Then
        Debug.Print "[Director] CompareSheets: Cannot compare a sheet to itself."
        Exit Sub
    End If

    Dim ws1 As Worksheet, ws2 As Worksheet
    Set ws1 = ThisWorkbook.Sheets(sheet1Name)
    Set ws2 = ThisWorkbook.Sheets(sheet2Name)

    If ws1 Is Nothing Then
        Debug.Print "[Director] CompareSheets: Sheet '" & sheet1Name & "' not found."
        Exit Sub
    End If
    If ws2 Is Nothing Then
        Debug.Print "[Director] CompareSheets: Sheet '" & sheet2Name & "' not found."
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "Comparing sheets..."

    ' Determine comparison range
    Dim maxRow As Long, maxCol As Long
    Dim lastRow1 As Long, lastRow2 As Long
    Dim lastCol1 As Long, lastCol2 As Long

    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    lastCol1 = ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column
    lastCol2 = ws2.Cells(1, ws2.Columns.Count).End(xlToLeft).Column

    If ws1.UsedRange.Rows.Count + ws1.UsedRange.Row - 1 > lastRow1 Then
        lastRow1 = ws1.UsedRange.Rows.Count + ws1.UsedRange.Row - 1
    End If
    If ws2.UsedRange.Rows.Count + ws2.UsedRange.Row - 1 > lastRow2 Then
        lastRow2 = ws2.UsedRange.Rows.Count + ws2.UsedRange.Row - 1
    End If
    If ws1.UsedRange.Columns.Count + ws1.UsedRange.Column - 1 > lastCol1 Then
        lastCol1 = ws1.UsedRange.Columns.Count + ws1.UsedRange.Column - 1
    End If
    If ws2.UsedRange.Columns.Count + ws2.UsedRange.Column - 1 > lastCol2 Then
        lastCol2 = ws2.UsedRange.Columns.Count + ws2.UsedRange.Column - 1
    End If

    If lastRow1 > lastRow2 Then maxRow = lastRow1 Else maxRow = lastRow2
    If lastCol1 > lastCol2 Then maxCol = lastCol1 Else maxCol = lastCol2

    ' Safety cap
    If maxRow > 10000 Then maxRow = 10000
    If maxCol > 256 Then maxCol = 256

    ' Compare cell by cell
    Dim diffCount As Long, matchCount As Long, totalCells As Long
    diffCount = 0
    matchCount = 0
    totalCells = maxRow * maxCol

    Dim diffCells() As String, diffVal1() As String, diffVal2() As String
    ReDim diffCells(1 To 5000)
    ReDim diffVal1(1 To 5000)
    ReDim diffVal2(1 To 5000)

    Dim r As Long, c As Long
    For r = 1 To maxRow
        If r Mod 500 = 0 Then
            Application.StatusBar = "Comparing row " & r & " of " & maxRow & "..."
        End If
        For c = 1 To maxCol
            Dim v1 As Variant, v2 As Variant
            v1 = ws1.Cells(r, c).Value
            v2 = ws2.Cells(r, c).Value

            If Not CellsMatch(v1, v2) Then
                diffCount = diffCount + 1
                If diffCount <= 5000 Then
                    diffCells(diffCount) = ws1.Cells(r, c).Address(False, False)
                    diffVal1(diffCount) = Left(CStr(Nz(v1, "")), 100)
                    diffVal2(diffCount) = Left(CStr(Nz(v2, "")), 100)
                End If
                ' Highlight differences in red on both sheets
                ws1.Cells(r, c).Interior.Color = CLR_DIFF
                ws2.Cells(r, c).Interior.Color = CLR_DIFF
            Else
                matchCount = matchCount + 1
            End If
        Next c
    Next r

    ' Build report
    BuildCompareReport ws1.Name, ws2.Name, diffCount, matchCount, totalCells, _
                       diffCells, diffVal1, diffVal2

    Application.StatusBar = False
    Application.ScreenUpdating = True

    Debug.Print "[Director] CompareSheets: '" & sheet1Name & "' vs '" & sheet2Name & "' — " & _
                diffCount & " difference(s) found out of " & totalCells & " cells."
End Sub


' ============================================================
' UTL_QuickRowCompareCount — Fast row-level "are these close?" check
' Returns the count of rows on sheet1 that have an exact match on sheet2
' (by pipe-delimited row signature). Much faster than a full cell-by-cell
' compare. Use as a pre-check before running the full Compare.
' Cherry-picked from Codex comparison (Batch 2, 2026-04-20).
'
' Example:
'   matches = UTL_QuickRowCompareCount("Q1 Revenue", "Q1 Revenue v2")
'   If matches = 0 Then ' sheets are totally different, skip cell-by-cell
' ============================================================
Public Function UTL_QuickRowCompareCount(ByVal sheet1Name As String, _
                                          ByVal sheet2Name As String) As Long
    UTL_QuickRowCompareCount = 0
    On Error Resume Next

    Dim ws1 As Worksheet, ws2 As Worksheet
    Set ws1 = ActiveWorkbook.Worksheets(sheet1Name)
    Set ws2 = ActiveWorkbook.Worksheets(sheet2Name)
    If ws1 Is Nothing Or ws2 Is Nothing Then
        Debug.Print "[UTL Compare] QuickRowCompare: one or both sheets not found."
        Exit Function
    End If

    Dim hdr1 As Long, hdr2 As Long
    hdr1 = UTL_DetectHeaderRow(ws1)
    hdr2 = UTL_DetectHeaderRow(ws2)

    Dim map2 As Object
    Set map2 = BuildRowHashMap(ws2, hdr2)
    If map2 Is Nothing Then Exit Function

    Dim lastRow1 As Long, lastCol1 As Long
    lastRow1 = UTL_LastRow(ws1, 1)
    lastCol1 = UTL_LastCol(ws1, hdr1)
    If lastRow1 <= hdr1 Or lastCol1 < 1 Then Exit Function

    Dim matches As Long
    Dim r As Long, c As Long
    Dim rowKey As String
    For r = hdr1 + 1 To lastRow1
        rowKey = ""
        For c = 1 To lastCol1
            rowKey = rowKey & "|" & Trim$(CStr(ws1.Cells(r, c).Value2))
        Next c
        If map2.Exists(rowKey) Then matches = matches + 1
    Next r

    Debug.Print "[UTL Compare] QuickRowCompare: '" & sheet1Name & "' vs '" & sheet2Name & _
                "' — " & matches & " matching row(s) out of " & (lastRow1 - hdr1)
    UTL_QuickRowCompareCount = matches
End Function

' Private helper — builds a Scripting.Dictionary of pipe-delimited row
' signatures for every data row on a sheet. Used by UTL_QuickRowCompareCount.
Private Function BuildRowHashMap(ByVal ws As Worksheet, ByVal headerRow As Long) As Object
    Set BuildRowHashMap = Nothing
    If ws Is Nothing Then Exit Function

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long, lastCol As Long
    lastRow = UTL_LastRow(ws, 1)
    lastCol = UTL_LastCol(ws, headerRow)
    If lastRow <= headerRow Or lastCol < 1 Then
        Set BuildRowHashMap = dict
        Exit Function
    End If

    Dim r As Long, c As Long
    Dim rowKey As String
    For r = headerRow + 1 To lastRow
        rowKey = ""
        For c = 1 To lastCol
            rowKey = rowKey & "|" & Trim$(CStr(ws.Cells(r, c).Value2))
        Next c
        If Not dict.Exists(rowKey) Then dict.Add rowKey, True
    Next r

    Set BuildRowHashMap = dict
End Function
