Attribute VB_Name = "modUTL_Audit"
Option Explicit

' ============================================================
' KBT Universal Tools — Audit & Compliance Module
' Works on ANY Excel file — no project-specific setup required
' Install in Personal.xlsb to use across all Excel sessions
' Tools: 8 | Tier 1: 3 | Tier 2: 5
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
' TOOL 1 — External Link Finder                      [TIER 1]
' Lists all cells that reference external workbooks
' Creates a report sheet with file paths and cell addresses
' ============================================================
Sub ExternalLinkFinder()
    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim reportWS As Worksheet
    Dim wsName As String
    wsName = "UTL_ExternalLinks"

    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Sheets(wsName).Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    Set reportWS = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    reportWS.Name = wsName

    With reportWS
        .Range("A1").Value = "External Link Report — " & ActiveWorkbook.Name
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A3").Value = "Sheet"
        .Range("B3").Value = "Cell"
        .Range("C3").Value = "Formula / Linked File"
        .Range("A3:C3").Font.Bold = True
        .Range("A3:C3").Interior.Color = RGB(31, 73, 125)
        .Range("A3:C3").Font.Color = RGB(255, 255, 255)
    End With

    Dim rowNum As Long
    rowNum = 4
    Dim found As Long
    Dim ws As Worksheet
    Dim c As Range

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> wsName Then
            Dim fRng As Range: Set fRng = Nothing
            On Error Resume Next
            Set fRng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
            On Error GoTo ErrHandler
            If Not fRng Is Nothing Then
                For Each c In fRng
                    If InStr(c.Formula, "[") > 0 Then
                        reportWS.Cells(rowNum, 1).Value = ws.Name
                        reportWS.Cells(rowNum, 2).Value = c.Address
                        reportWS.Cells(rowNum, 3).Value = c.Formula
                        rowNum = rowNum + 1
                        found = found + 1
                    End If
                Next c
            End If
        End If
    Next ws

    reportWS.Columns("A:C").AutoFit
    UTL_TurboOff

    If found = 0 Then
        Application.DisplayAlerts = False
        reportWS.Delete
        Application.DisplayAlerts = True
        MsgBox "No external links found. This workbook is self-contained.", _
               vbInformation, "UTL Audit"
    Else
        reportWS.Activate
        MsgBox "Found " & found & " external link(s). Report created on sheet '" & wsName & "'." & Chr(10) & _
               "Review each link to confirm it should be there.", _
               vbExclamation, "UTL Audit"
    End If
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Audit"
End Sub

' ============================================================
' TOOL 2 — Circular Reference Detector               [TIER 1]
' Finds and reports all circular references in the workbook
' ============================================================
Sub CircularReferenceDetector()
    On Error GoTo ErrHandler

    Dim circCount As Long
    circCount = 0

    Dim ws As Worksheet
    Dim circ As Range

    Dim report As String
    report = "CIRCULAR REFERENCE REPORT" & Chr(10) & String(35, "-") & Chr(10)

    For Each ws In ActiveWorkbook.Worksheets
        Dim circRange As Range
        On Error Resume Next
        Set circRange = ws.CircularReference
        On Error GoTo ErrHandler
        If Not circRange Is Nothing Then
            For Each circ In circRange
                report = report & "Sheet: " & ws.Name & " | Cell: " & circ.Address & Chr(10)
                circCount = circCount + 1
            Next circ
        End If
        Set circRange = Nothing
    Next ws

    If circCount = 0 Then
        MsgBox "No circular references found. Workbook is clean.", vbInformation, "UTL Audit"
    Else
        report = report & Chr(10) & "Total Found: " & circCount & Chr(10) & _
                 Chr(10) & "Each of these cells references itself (directly or indirectly)." & Chr(10) & _
                 "Circular references can cause incorrect calculations."
        MsgBox report, vbExclamation, "UTL Audit — Circular References"
    End If
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Audit"
End Sub

' ============================================================
' TOOL 3 — Workbook Error Scanner                    [TIER 1]
' Lists every cell with an error value across all sheets
' Creates a report sheet with exact locations
' ============================================================
Sub WorkbookErrorScanner()
    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim reportWS As Worksheet
    Dim wsName As String
    wsName = "UTL_ErrorReport"

    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Sheets(wsName).Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    Set reportWS = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    reportWS.Name = wsName

    With reportWS
        .Range("A1").Value = "Error Cell Report — " & ActiveWorkbook.Name
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A3").Value = "Sheet"
        .Range("B3").Value = "Cell"
        .Range("C3").Value = "Error Type"
        .Range("D3").Value = "Formula"
        .Range("A3:D3").Font.Bold = True
        .Range("A3:D3").Interior.Color = RGB(192, 0, 0)
        .Range("A3:D3").Font.Color = RGB(255, 255, 255)
    End With

    Dim rowNum As Long
    rowNum = 4
    Dim totalErrors As Long
    Dim ws As Worksheet
    Dim c As Range

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> wsName Then
            Dim errRng As Range: Set errRng = Nothing
            On Error Resume Next
            Set errRng = ws.UsedRange.SpecialCells(xlCellTypeFormulas, xlErrors)
            On Error GoTo ErrHandler
            If Not errRng Is Nothing Then
                For Each c In errRng
                    reportWS.Cells(rowNum, 1).Value = ws.Name
                    reportWS.Cells(rowNum, 2).Value = c.Address
                    Select Case CStr(c.Value)
                        Case "Error 2007": reportWS.Cells(rowNum, 3).Value = "#DIV/0!"
                        Case "Error 2015": reportWS.Cells(rowNum, 3).Value = "#VALUE!"
                        Case "Error 2023": reportWS.Cells(rowNum, 3).Value = "#REF!"
                        Case "Error 2029": reportWS.Cells(rowNum, 3).Value = "#NAME?"
                        Case "Error 2042": reportWS.Cells(rowNum, 3).Value = "#N/A"
                        Case "Error 2036": reportWS.Cells(rowNum, 3).Value = "#NUM!"
                        Case "Error 2000": reportWS.Cells(rowNum, 3).Value = "#NULL!"
                        Case Else: reportWS.Cells(rowNum, 3).Value = "Error"
                    End Select
                    If c.HasFormula Then reportWS.Cells(rowNum, 4).Value = c.Formula
                    rowNum = rowNum + 1
                    totalErrors = totalErrors + 1
                Next c
            End If
        End If
    Next ws

    reportWS.Columns("A:D").AutoFit
    UTL_TurboOff

    If totalErrors = 0 Then
        Application.DisplayAlerts = False
        reportWS.Delete
        Application.DisplayAlerts = True
        MsgBox "No error cells found. Workbook is clean.", vbInformation, "UTL Audit"
    Else
        reportWS.Activate
        MsgBox "Found " & totalErrors & " error cell(s). Full report on sheet '" & wsName & "'.", _
               vbExclamation, "UTL Audit"
    End If
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Audit"
End Sub

' ============================================================
' TOOL 4 — Data Quality Scorecard                    [TIER 2]
' Generates a column-by-column summary of data quality issues
' Covers: blanks, errors, duplicates, data types per column
' ============================================================
Sub DataQualityScorecard()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim scoreWS As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Sheets("UTL Data Quality").Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    Set scoreWS = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    scoreWS.Name = "UTL Data Quality"

    With scoreWS
        .Range("A1").Value = "Data Quality Scorecard — " & ws.Name
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A2").Value = "Source sheet: " & ws.Name & " | Rows analyzed: " & (lastRow - 1) & " | Date: " & Format(Now, "MM/DD/YYYY")
        .Range("A4").Value = "Column"
        .Range("B4").Value = "Header"
        .Range("C4").Value = "Total Rows"
        .Range("D4").Value = "Blanks"
        .Range("E4").Value = "Errors"
        .Range("F4").Value = "Duplicates"
        .Range("G4").Value = "Numeric"
        .Range("H4").Value = "Text"
        .Range("I4").Value = "Dates"
        .Range("A4:I4").Font.Bold = True
        .Range("A4:I4").Interior.Color = RGB(31, 73, 125)
        .Range("A4:I4").Font.Color = RGB(255, 255, 255)
    End With

    Dim colNum As Long
    Dim scoreRow As Long
    scoreRow = 5

    For colNum = 1 To lastCol
        Dim colHeader As String
        colHeader = CStr(ws.Cells(1, colNum).Value)

        Dim blankCount As Long
        Dim errorCount As Long
        Dim numCount As Long
        Dim textCount As Long
        Dim dateCount As Long
        blankCount = 0: errorCount = 0: numCount = 0: textCount = 0: dateCount = 0

        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
        Dim dupCount As Long
        dupCount = 0

        Dim r As Long
        For r = 2 To lastRow
            Dim cellVal As Variant
            cellVal = ws.Cells(r, colNum).Value
            If IsEmpty(cellVal) Or cellVal = "" Then
                blankCount = blankCount + 1
            ElseIf IsError(cellVal) Then
                errorCount = errorCount + 1
            ElseIf IsDate(cellVal) Then
                dateCount = dateCount + 1
            ElseIf IsNumeric(cellVal) Then
                numCount = numCount + 1
            Else
                textCount = textCount + 1
            End If

            Dim dictKey As String
            dictKey = CStr(cellVal)
            If Not IsEmpty(cellVal) And cellVal <> "" Then
                If dict.exists(dictKey) Then
                    dupCount = dupCount + 1
                Else
                    dict.Add dictKey, True
                End If
            End If
        Next r

        With scoreWS
            .Cells(scoreRow, 1).Value = colNum
            .Cells(scoreRow, 2).Value = colHeader
            .Cells(scoreRow, 3).Value = lastRow - 1
            .Cells(scoreRow, 4).Value = blankCount
            .Cells(scoreRow, 5).Value = errorCount
            .Cells(scoreRow, 6).Value = dupCount
            .Cells(scoreRow, 7).Value = numCount
            .Cells(scoreRow, 8).Value = textCount
            .Cells(scoreRow, 9).Value = dateCount

            ' Highlight issues
            If blankCount > 0 Then .Cells(scoreRow, 4).Interior.Color = RGB(255, 235, 59)
            If errorCount > 0 Then .Cells(scoreRow, 5).Interior.Color = RGB(255, 100, 100)
            If dupCount > 0 Then .Cells(scoreRow, 6).Interior.Color = RGB(255, 200, 100)
        End With
        scoreRow = scoreRow + 1
    Next colNum

    scoreWS.Columns("A:I").AutoFit
    UTL_TurboOff
    scoreWS.Activate
    MsgBox "Data Quality Scorecard complete!" & Chr(10) & _
           "Yellow = blanks | Red = errors | Orange = duplicates", _
           vbInformation, "UTL Audit"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Audit"
End Sub

' ============================================================
' TOOL 5 — Named Range Auditor                       [TIER 2]
' Reports all Named Ranges and flags broken references
' ============================================================
Sub NamedRangeAuditor()
    On Error GoTo ErrHandler

    If ActiveWorkbook.Names.Count = 0 Then
        MsgBox "No named ranges found in this workbook.", vbInformation, "UTL Audit"
        Exit Sub
    End If

    Dim reportWS As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Sheets("UTL_NamedRanges").Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    Set reportWS = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    reportWS.Name = "UTL_NamedRanges"

    With reportWS
        .Range("A1").Value = "Named Range Audit — " & ActiveWorkbook.Name
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A3").Value = "Name"
        .Range("B3").Value = "Refers To"
        .Range("C3").Value = "Scope"
        .Range("D3").Value = "Status"
        .Range("A3:D3").Font.Bold = True
        .Range("A3:D3").Interior.Color = RGB(31, 73, 125)
        .Range("A3:D3").Font.Color = RGB(255, 255, 255)
    End With

    Dim row As Long
    row = 4
    Dim brokenCount As Long
    Dim nm As Name

    For Each nm In ActiveWorkbook.Names
        reportWS.Cells(row, 1).Value = nm.Name
        reportWS.Cells(row, 2).Value = nm.RefersTo
        reportWS.Cells(row, 3).Value = IIf(InStr(nm.Name, "!") > 0, "Sheet", "Workbook")

        Dim status As String
        status = "OK"
        On Error Resume Next
        Dim testRange As Range
        Set testRange = Nothing
        Set testRange = nm.RefersToRange
        If testRange Is Nothing Then
            status = "BROKEN — #REF! or invalid"
            reportWS.Cells(row, 4).Interior.Color = RGB(255, 100, 100)
            brokenCount = brokenCount + 1
        End If
        On Error GoTo ErrHandler
        reportWS.Cells(row, 4).Value = status
        row = row + 1
    Next nm

    reportWS.Columns("A:D").AutoFit
    reportWS.Activate
    MsgBox "Named Range Audit complete." & Chr(10) & _
           ActiveWorkbook.Names.Count & " ranges found." & Chr(10) & _
           brokenCount & " broken reference(s) flagged red.", _
           IIf(brokenCount > 0, vbExclamation, vbInformation), "UTL Audit"
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Audit"
End Sub

' ============================================================
' TOOL 6 — Data Validation Checker                   [TIER 2]
' Scans for dropdown cells with broken source ranges
' ============================================================
Sub DataValidationChecker()
    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim reportWS As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Sheets("UTL_Validation").Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    Set reportWS = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    reportWS.Name = "UTL_Validation"

    With reportWS
        .Range("A1").Value = "Data Validation Audit — " & ActiveWorkbook.Name
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A3").Value = "Sheet"
        .Range("B3").Value = "Cell"
        .Range("C3").Value = "Validation Type"
        .Range("D3").Value = "Formula / Source"
        .Range("E3").Value = "Status"
        .Range("A3:E3").Font.Bold = True
        .Range("A3:E3").Interior.Color = RGB(31, 73, 125)
        .Range("A3:E3").Font.Color = RGB(255, 255, 255)
    End With

    Dim row As Long
    row = 4
    Dim totalDV As Long
    Dim brokenDV As Long
    Dim ws As Worksheet
    Dim c As Range

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> "UTL_Validation" Then
            For Each c In ws.UsedRange
                On Error Resume Next
                Dim dvType As Long
                dvType = c.Validation.Type
                If Err.Number = 0 Then
                    totalDV = totalDV + 1
                    reportWS.Cells(row, 1).Value = ws.Name
                    reportWS.Cells(row, 2).Value = c.Address

                    Dim dvTypeName As String
                    Select Case dvType
                        Case 3: dvTypeName = "List (Dropdown)"
                        Case 1: dvTypeName = "Whole Number"
                        Case 2: dvTypeName = "Decimal"
                        Case 4: dvTypeName = "Date"
                        Case Else: dvTypeName = "Other (" & dvType & ")"
                    End Select
                    reportWS.Cells(row, 3).Value = dvTypeName

                    Dim dvFormula As String
                    dvFormula = c.Validation.Formula1
                    reportWS.Cells(row, 4).Value = dvFormula

                    ' Check if source range is valid (for list dropdowns)
                    Dim dvStatus As String
                    dvStatus = "OK"
                    If dvType = 3 And Left(dvFormula, 1) = "=" Then
                        Dim sourceRng As Range
                        Set sourceRng = Nothing
                        Set sourceRng = ws.Range(Mid(dvFormula, 2))
                        If sourceRng Is Nothing Then
                            dvStatus = "BROKEN SOURCE"
                            reportWS.Cells(row, 5).Interior.Color = RGB(255, 100, 100)
                            brokenDV = brokenDV + 1
                        End If
                    End If
                    reportWS.Cells(row, 5).Value = dvStatus
                    row = row + 1
                End If
                Err.Clear
                On Error GoTo ErrHandler
            Next c
        End If
    Next ws

    reportWS.Columns("A:E").AutoFit
    UTL_TurboOff

    If totalDV = 0 Then
        Application.DisplayAlerts = False
        reportWS.Delete
        Application.DisplayAlerts = True
        MsgBox "No data validation rules found in this workbook.", vbInformation, "UTL Audit"
    Else
        reportWS.Activate
        MsgBox totalDV & " validation rule(s) found." & Chr(10) & _
               brokenDV & " broken source range(s) flagged red.", _
               IIf(brokenDV > 0, vbExclamation, vbInformation), "UTL Audit"
    End If
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Audit"
End Sub

' ============================================================
' TOOL 7 — Inconsistent Formulas Auditor             [TIER 2]
' Flags cells in a column where the formula differs from majority
' Catches accidental hardcodes hiding inside formula columns
' ============================================================
Sub InconsistentFormulasAuditor()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select the column range to audit.", vbExclamation, "UTL Audit"
        Exit Sub
    End If

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim rng As Range
    Set rng = Selection.Columns(1)

    ' Build formula frequency dictionary
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim c As Range

    For Each c In rng
        If c.HasFormula Then
            ' Normalize formula using R1C1 notation so row numbers don't cause false positives
            Dim normFormula As String
            normFormula = c.FormulaR1C1
            Dim key As String
            key = normFormula
            If dict.exists(key) Then
                dict(key) = dict(key) + 1
            Else
                dict.Add key, 1
            End If
        End If
    Next c

    ' Find the most common formula
    Dim maxCount As Long
    Dim mostCommon As String
    Dim k As Variant
    For Each k In dict.Keys
        If dict(k) > maxCount Then
            maxCount = dict(k)
            mostCommon = k
        End If
    Next k

    ' Flag outliers
    Dim flagged As Long
    For Each c In rng
        If c.HasFormula Then
            If c.FormulaR1C1 <> mostCommon Then
                c.Interior.Color = RGB(255, 200, 100)
                flagged = flagged + 1
            End If
        ElseIf Not IsEmpty(c) Then
            ' Hardcoded value in a formula column — flag red
            c.Interior.Color = RGB(255, 100, 100)
            flagged = flagged + 1
        End If
    Next c

    UTL_TurboOff
    If flagged = 0 Then
        MsgBox "All formulas in selection are consistent. No issues found.", vbInformation, "UTL Audit"
    Else
        MsgBox flagged & " inconsistency(ies) flagged." & Chr(10) & _
               "Orange = different formula | Red = hardcoded value in formula column." & Chr(10) & Chr(10) & _
               "Most common formula: " & Chr(10) & mostCommon, _
               vbExclamation, "UTL Audit"
    End If
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Audit"
End Sub

' ============================================================
' TOOL 8 — External Link Severance Protocol          [TIER 2]
' Replaces external file references with their current values
' Logs original formulas in a comment on each affected cell
' ============================================================
Sub ExternalLinkSeveranceProtocol()
    If MsgBox("This will replace all external link formulas with their CURRENT VALUES." & Chr(10) & _
              "A backup of each sheet will be created first." & Chr(10) & _
              "The original formula will be saved as a comment on each cell." & Chr(10) & Chr(10) & _
              "Continue?", vbExclamation + vbYesNo, "UTL Audit") = vbNo Then Exit Sub

    On Error GoTo ErrHandler
    UTL_TurboOn

    ' Create backup of all sheets before severance
    Dim bkWs As Worksheet
    For Each bkWs In ActiveWorkbook.Worksheets
        modUTL_Core.UTL_BackupSheet bkWs
    Next bkWs

    Dim severed As Long
    Dim ws As Worksheet
    Dim c As Range

    For Each ws In ActiveWorkbook.Worksheets
        Dim sevRng As Range: Set sevRng = Nothing
        On Error Resume Next
        Set sevRng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo ErrHandler
        If Not sevRng Is Nothing Then
            For Each c In sevRng
                If InStr(c.Formula, "[") > 0 Then
                    Dim originalFormula As String
                    originalFormula = c.Formula
                    Dim currentValue As Variant
                    currentValue = c.Value

                    ' Add comment with original formula
                    On Error Resume Next
                    c.Comment.Delete
                    On Error GoTo ErrHandler
                    c.AddComment "ORIGINAL FORMULA (severed " & Format(Now, "MM/DD/YYYY") & "):" & Chr(10) & originalFormula
                    c.Comment.Shape.TextFrame.AutoSize = True

                    ' Replace with value
                    c.Value = currentValue
                    c.Interior.Color = RGB(255, 235, 59)
                    severed = severed + 1
                End If
            Next c
        End If
    Next ws

    UTL_TurboOff
    If severed = 0 Then
        MsgBox "No external links found to sever.", vbInformation, "UTL Audit"
    Else
        MsgBox "Done! " & severed & " external link(s) severed." & Chr(10) & _
               "Affected cells highlighted yellow." & Chr(10) & _
               "Original formulas saved as cell comments.", _
               vbInformation, "UTL Audit"
    End If
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Audit"
End Sub


' ============================================================
' CreateRunReceiptSheet — Compliance artifact for any toolkit run
' Drops a 6-row UTL_RunReceipt sheet with:
'   Timestamp, User, Workbook, Feature, Sheets Touched, Cells Changed
' Any toolkit tool can call this after a meaningful run.
' Brand-styled header (iPipeline Blue / Arctic White).
' No dialogs. No MsgBox. Always overwrites prior receipt.
' Cherry-picked from Codex comparison (Batch 1, 2026-04-20).
' ============================================================
Public Sub CreateRunReceiptSheet(ByVal featureName As String, _
                                 ByVal sheetsTouched As String, _
                                 ByVal cellsChanged As Long)
    On Error Resume Next
    Dim sheetName As String: sheetName = "UTL_RunReceipt"

    ' Remove any prior receipt quietly
    Application.DisplayAlerts = False
    Dim oldWs As Worksheet
    Set oldWs = Nothing
    Set oldWs = ActiveWorkbook.Sheets(sheetName)
    If Not oldWs Is Nothing Then oldWs.Delete
    Application.DisplayAlerts = True

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    ws.Name = sheetName

    ' Title
    ws.Range("A1").Value = "Run Receipt — iPipeline Universal Toolkit"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Name = "Arial"
    ws.Range("A1").Font.Color = RGB(17, 46, 81)        ' Navy
    ws.Range("A1:B1").Merge

    ' Header row (row 3)
    ws.Range("A3").Value = "Field"
    ws.Range("B3").Value = "Value"
    ws.Range("A3:B3").Font.Bold = True
    ws.Range("A3:B3").Font.Name = "Arial"
    ws.Range("A3:B3").Font.Color = RGB(249, 249, 249)  ' Arctic White
    ws.Range("A3:B3").Interior.Color = RGB(11, 71, 121) ' iPipeline Blue

    ' Receipt rows
    ws.Range("A4").Value = "Timestamp"
    ws.Range("B4").Value = Format(Now, "yyyy-mm-dd hh:nn:ss")

    ws.Range("A5").Value = "User"
    ws.Range("B5").Value = Environ("USERNAME")

    ws.Range("A6").Value = "Workbook"
    ws.Range("B6").Value = ActiveWorkbook.Name

    ws.Range("A7").Value = "Feature"
    ws.Range("B7").Value = featureName

    ws.Range("A8").Value = "Sheets Touched"
    ws.Range("B8").Value = sheetsTouched

    ws.Range("A9").Value = "Cells Changed"
    ws.Range("B9").Value = cellsChanged

    ' Style field column
    ws.Range("A4:A9").Font.Bold = True
    ws.Range("A4:A9").Font.Name = "Arial"
    ws.Range("A4:A9").Font.Color = RGB(22, 22, 22)     ' Charcoal

    ws.Range("B4:B9").Font.Name = "Arial"
    ws.Range("B4:B9").Font.Color = RGB(22, 22, 22)

    ' Light alternating fill on value cells for readability
    ws.Range("A5:B5").Interior.Color = RGB(240, 240, 238)
    ws.Range("A7:B7").Interior.Color = RGB(240, 240, 238)
    ws.Range("A9:B9").Interior.Color = RGB(240, 240, 238)

    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B").ColumnWidth = 60

    Debug.Print "[UTL] RunReceipt: " & featureName & " | " & cellsChanged & " cells | " & sheetsTouched
End Sub
