Attribute VB_Name = "modUTL_AuditPlus"
Option Explicit

' ============================================================
' KBT Universal Tools — Audit Plus Module
' Works on ANY Excel file — no project-specific setup required
' Tools: 4 | All Small-Medium effort
' Date: 2026-03-05
' ============================================================
' Tool 05 — Data Boundary Detector
' Tool 06 — Header Validator (Fuzzy Matching)
' Tool 07 — Formula Error Finder
' Tool 08 — Formula Consistency Checker
' ============================================================

Private Sub UTL_TurboOn()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

Private Sub UTL_TurboOff()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

' ============================================================
' TOOL 05 — Data Boundary Detector                    [SMALL]
' Scans active sheet to find the actual data rectangle
' (first row, last row, first col, last col) and reports
' any unexpected gaps, blank rows/cols inside the data area.
' ============================================================
Sub DataBoundaryDetector()
    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim usedRng As Range
    Set usedRng = ws.UsedRange
    If usedRng Is Nothing Then
        UTL_TurboOff
        MsgBox "Sheet '" & ws.Name & "' appears empty.", vbInformation
        Exit Sub
    End If

    Dim firstRow As Long: firstRow = usedRng.Row
    Dim firstCol As Long: firstCol = usedRng.Column
    Dim lastRow  As Long: lastRow = firstRow + usedRng.Rows.Count - 1
    Dim lastCol  As Long: lastCol = firstCol + usedRng.Columns.Count - 1

    ' Count blank rows inside data area
    Dim blankRows As Long: blankRows = 0
    Dim r As Long
    For r = firstRow To lastRow
        If Application.CountA(ws.Range(ws.Cells(r, firstCol), ws.Cells(r, lastCol))) = 0 Then
            blankRows = blankRows + 1
        End If
    Next r

    ' Count blank columns inside data area
    Dim blankCols As Long: blankCols = 0
    Dim c As Long
    For c = firstCol To lastCol
        If Application.CountA(ws.Range(ws.Cells(firstRow, c), ws.Cells(lastRow, c))) = 0 Then
            blankCols = blankCols + 1
        End If
    Next c

    Dim totalRows As Long: totalRows = lastRow - firstRow + 1
    Dim totalCols As Long: totalCols = lastCol - firstCol + 1

    UTL_TurboOff

    Dim msg As String
    msg = "Data Boundary Report for '" & ws.Name & "'" & vbCrLf & vbCrLf & _
          "Data rectangle: " & ws.Cells(firstRow, firstCol).Address(False, False) & _
          " to " & ws.Cells(lastRow, lastCol).Address(False, False) & vbCrLf & _
          "Rows: " & totalRows & "  |  Columns: " & totalCols & vbCrLf & vbCrLf

    If blankRows > 0 Or blankCols > 0 Then
        msg = msg & "GAPS DETECTED:" & vbCrLf
        If blankRows > 0 Then msg = msg & "  " & blankRows & " entirely blank row(s)" & vbCrLf
        If blankCols > 0 Then msg = msg & "  " & blankCols & " entirely blank column(s)" & vbCrLf
        msg = msg & vbCrLf & "Blank rows/columns inside data can break PivotTables, filters, and formulas."
        MsgBox msg, vbExclamation, "Data Boundary Detector"
    Else
        msg = msg & "No gaps found — data area is contiguous."
        MsgBox msg, vbInformation, "Data Boundary Detector"
    End If
    Exit Sub

ErrHandler:
    UTL_TurboOff
    MsgBox "Data Boundary Detector error: " & Err.Description, vbCritical
End Sub

' ============================================================
' TOOL 06 — Header Validator (Fuzzy Matching)          [MEDIUM]
' User provides a comma-separated list of expected headers.
' Scans row 1 (or user-specified row) and reports:
'   - Exact matches, fuzzy matches (Levenshtein-like), missing
' ============================================================
Sub HeaderValidator()
    On Error GoTo ErrHandler

    Dim headerRow As String
    headerRow = InputBox("Which row contains headers?" & vbCrLf & _
                         "(Enter row number, e.g. 1)", _
                         "Header Validator", "1")
    If headerRow = "" Then Exit Sub
    Dim hRow As Long: hRow = CLng(headerRow)

    Dim expectedStr As String
    expectedStr = InputBox("Enter expected header names, separated by commas:" & vbCrLf & _
                           "Example: Date, Amount, Description, Status", _
                           "Header Validator — Expected Headers")
    If expectedStr = "" Then Exit Sub

    UTL_TurboOn

    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastCol As Long: lastCol = ws.Cells(hRow, ws.Columns.Count).End(xlToLeft).Column

    ' Parse expected headers
    Dim expected() As String: expected = Split(expectedStr, ",")
    Dim i As Long
    For i = LBound(expected) To UBound(expected)
        expected(i) = Trim(expected(i))
    Next i

    ' Read actual headers
    Dim actualCount As Long: actualCount = lastCol
    Dim actual() As String
    ReDim actual(1 To actualCount)
    Dim c As Long
    For c = 1 To actualCount
        actual(c) = Trim(CStr(ws.Cells(hRow, c).Value))
    Next c

    ' Match each expected header
    Dim report As String: report = ""
    Dim exactCount As Long: exactCount = 0
    Dim fuzzyCount As Long: fuzzyCount = 0
    Dim missingCount As Long: missingCount = 0

    For i = LBound(expected) To UBound(expected)
        If Len(expected(i)) = 0 Then GoTo NextExpected
        Dim found As Boolean: found = False
        Dim bestMatch As String: bestMatch = ""
        Dim bestScore As Long: bestScore = 999

        For c = 1 To actualCount
            If UCase(actual(c)) = UCase(expected(i)) Then
                found = True
                exactCount = exactCount + 1
                report = report & "  EXACT: '" & expected(i) & "' found in column " & c & vbCrLf
                GoTo NextExpected
            End If
            ' Simple fuzzy: check if one contains the other
            Dim score As Long
            score = FuzzyDistance(UCase(expected(i)), UCase(actual(c)))
            If score < bestScore Then
                bestScore = score
                bestMatch = actual(c)
            End If
        Next c

        If Not found Then
            ' If best fuzzy score is <= 3 edits, suggest it
            If bestScore <= 3 And Len(bestMatch) > 0 Then
                fuzzyCount = fuzzyCount + 1
                report = report & "  FUZZY: '" & expected(i) & "' not found — did you mean '" & bestMatch & "'?" & vbCrLf
            Else
                missingCount = missingCount + 1
                report = report & "  MISSING: '" & expected(i) & "' not found" & vbCrLf
            End If
        End If
NextExpected:
    Next i

    UTL_TurboOff

    Dim summary As String
    summary = "Header Validation Report — '" & ws.Name & "' Row " & hRow & vbCrLf & vbCrLf & _
              "Exact matches: " & exactCount & vbCrLf & _
              "Fuzzy matches: " & fuzzyCount & vbCrLf & _
              "Missing: " & missingCount & vbCrLf & vbCrLf & _
              report
    MsgBox summary, IIf(missingCount > 0, vbExclamation, vbInformation), "Header Validator"
    Exit Sub

ErrHandler:
    UTL_TurboOff
    MsgBox "Header Validator error: " & Err.Description, vbCritical
End Sub

' Simple edit-distance approximation (Levenshtein)
Private Function FuzzyDistance(ByVal s1 As String, ByVal s2 As String) As Long
    Dim len1 As Long: len1 = Len(s1)
    Dim len2 As Long: len2 = Len(s2)
    If len1 = 0 Then FuzzyDistance = len2: Exit Function
    If len2 = 0 Then FuzzyDistance = len1: Exit Function

    ' Use two-row approach to save memory
    Dim prev() As Long, curr() As Long
    ReDim prev(0 To len2)
    ReDim curr(0 To len2)

    Dim i As Long, j As Long
    For j = 0 To len2: prev(j) = j: Next j

    For i = 1 To len1
        curr(0) = i
        For j = 1 To len2
            Dim cost As Long
            If Mid(s1, i, 1) = Mid(s2, j, 1) Then cost = 0 Else cost = 1
            Dim ins As Long: ins = curr(j - 1) + 1
            Dim del As Long: del = prev(j) + 1
            Dim sub1 As Long: sub1 = prev(j - 1) + cost
            curr(j) = ins
            If del < curr(j) Then curr(j) = del
            If sub1 < curr(j) Then curr(j) = sub1
        Next j
        Dim tmp() As Long: tmp = prev: prev = curr: curr = tmp
    Next i
    FuzzyDistance = prev(len2)
End Function

' ============================================================
' TOOL 07 — Formula Error Finder                       [SMALL]
' Scans all sheets for cells containing Excel errors
' (#REF!, #VALUE!, #N/A, #DIV/0!, #NAME?, #NULL!, #NUM!)
' Reports results on a new sheet.
' ============================================================
Sub FormulaErrorFinder()
    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim rptName As String: rptName = "UTL_ErrorReport"

    ' Delete old report if exists
    Dim wsOld As Worksheet
    On Error Resume Next
    Set wsOld = wb.Worksheets(rptName)
    On Error GoTo ErrHandler
    If Not wsOld Is Nothing Then
        Application.DisplayAlerts = False
        wsOld.Delete
        Application.DisplayAlerts = True
    End If

    Dim wsRpt As Worksheet
    Set wsRpt = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsRpt.Name = rptName

    ' Headers
    wsRpt.Cells(1, 1).Value = "Sheet"
    wsRpt.Cells(1, 2).Value = "Cell"
    wsRpt.Cells(1, 3).Value = "Error Type"
    wsRpt.Cells(1, 4).Value = "Formula"
    Dim hdrCol As Long
    For hdrCol = 1 To 4
        wsRpt.Cells(1, hdrCol).Font.Bold = True
        wsRpt.Cells(1, hdrCol).Interior.Color = RGB(11, 71, 121)
        wsRpt.Cells(1, hdrCol).Font.Color = RGB(255, 255, 255)
    Next hdrCol

    Dim outRow As Long: outRow = 2
    Dim totalErrors As Long: totalErrors = 0
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If ws.Name = rptName Then GoTo NextSheet

        Dim errRng As Range: Set errRng = Nothing
        On Error Resume Next
        Set errRng = ws.UsedRange.SpecialCells(xlCellTypeFormulas, xlErrors)
        On Error GoTo ErrHandler
        If errRng Is Nothing Then GoTo NextSheet

        Dim cell As Range
        For Each cell In errRng
            wsRpt.Cells(outRow, 1).Value = ws.Name
            wsRpt.Cells(outRow, 2).Value = cell.Address(False, False)
            wsRpt.Cells(outRow, 3).Value = CStr(cell.Value)
            On Error Resume Next
            wsRpt.Cells(outRow, 4).Value = "'" & cell.Formula
            On Error GoTo ErrHandler
            totalErrors = totalErrors + 1
            outRow = outRow + 1

            If totalErrors >= 5000 Then
                wsRpt.Cells(outRow, 1).Value = "--- LIMIT REACHED (5,000 errors) ---"
                GoTo DoneScanning
            End If
        Next cell
NextSheet:
    Next ws

DoneScanning:
    wsRpt.Columns("A:D").AutoFit
    wsRpt.Activate

    UTL_TurboOff

    If totalErrors = 0 Then
        MsgBox "No formula errors found in any sheet.", vbInformation, "Formula Error Finder"
        Application.DisplayAlerts = False
        wsRpt.Delete
        Application.DisplayAlerts = True
    Else
        MsgBox totalErrors & " formula error(s) found across all sheets." & vbCrLf & _
               "Results on '" & rptName & "'.", vbExclamation, "Formula Error Finder"
    End If
    Exit Sub

ErrHandler:
    UTL_TurboOff
    MsgBox "Formula Error Finder error: " & Err.Description, vbCritical
End Sub

' ============================================================
' TOOL 08 — Formula Consistency Checker                [MEDIUM]
' Picks a column (user input) and checks if all formulas in
' that column follow the same pattern. Flags any row where
' the formula structure differs from the majority.
' ============================================================
Sub FormulaConsistencyChecker()
    On Error GoTo ErrHandler

    Dim colInput As String
    colInput = InputBox("Enter the column letter to check for formula consistency:" & vbCrLf & _
                        "Example: D", "Formula Consistency Checker", "D")
    If colInput = "" Then Exit Sub

    UTL_TurboOn

    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim checkCol As Long: checkCol = Range(colInput & "1").Column
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, checkCol).End(xlUp).Row

    If lastRow < 2 Then
        UTL_TurboOff
        MsgBox "No data found in column " & colInput & ".", vbInformation
        Exit Sub
    End If

    ' Collect formula patterns (replace row numbers with #)
    Dim patterns As Object: Set patterns = CreateObject("Scripting.Dictionary")
    Dim r As Long
    Dim formulaCount As Long: formulaCount = 0

    For r = 2 To lastRow
        If ws.Cells(r, checkCol).HasFormula Then
            Dim rawFormula As String: rawFormula = ws.Cells(r, checkCol).Formula
            Dim pattern As String: pattern = NormalizeFormula(rawFormula, r)
            formulaCount = formulaCount + 1
            If patterns.Exists(pattern) Then
                patterns(pattern) = patterns(pattern) + 1
            Else
                patterns.Add pattern, 1
            End If
        End If
    Next r

    If formulaCount = 0 Then
        UTL_TurboOff
        MsgBox "No formulas found in column " & colInput & ".", vbInformation
        Exit Sub
    End If

    ' Find the dominant pattern
    Dim dominantPattern As String: dominantPattern = ""
    Dim maxCount As Long: maxCount = 0
    Dim key As Variant
    For Each key In patterns.Keys
        If patterns(key) > maxCount Then
            maxCount = patterns(key)
            dominantPattern = CStr(key)
        End If
    Next key

    ' Find inconsistent rows
    Dim inconsistent As String: inconsistent = ""
    Dim inconsistentCount As Long: inconsistentCount = 0

    For r = 2 To lastRow
        If ws.Cells(r, checkCol).HasFormula Then
            Dim thisPattern As String
            thisPattern = NormalizeFormula(ws.Cells(r, checkCol).Formula, r)
            If thisPattern <> dominantPattern Then
                inconsistentCount = inconsistentCount + 1
                If inconsistentCount <= 20 Then
                    inconsistent = inconsistent & "  Row " & r & ": " & ws.Cells(r, checkCol).Formula & vbCrLf
                End If
            End If
        End If
    Next r

    UTL_TurboOff

    Dim msg As String
    msg = "Formula Consistency Report — Column " & colInput & vbCrLf & vbCrLf & _
          "Total formulas: " & formulaCount & vbCrLf & _
          "Dominant pattern count: " & maxCount & vbCrLf & _
          "Inconsistent: " & inconsistentCount & vbCrLf

    If inconsistentCount > 0 Then
        msg = msg & vbCrLf & "Inconsistent formulas:" & vbCrLf & inconsistent
        If inconsistentCount > 20 Then
            msg = msg & "  ... and " & (inconsistentCount - 20) & " more"
        End If
        MsgBox msg, vbExclamation, "Formula Consistency Checker"
    Else
        msg = msg & vbCrLf & "All formulas follow the same pattern."
        MsgBox msg, vbInformation, "Formula Consistency Checker"
    End If
    Exit Sub

ErrHandler:
    UTL_TurboOff
    MsgBox "Formula Consistency Checker error: " & Err.Description, vbCritical
End Sub

' Normalize a formula by replacing its own row number with #
' so =SUM(B5:G5) on row 5 becomes =SUM(B#:G#)
Private Function NormalizeFormula(ByVal formula As String, ByVal rowNum As Long) As String
    NormalizeFormula = Replace(formula, CStr(rowNum), "#")
End Function
