Attribute VB_Name = "modUTL_Consolidate"
'==============================================================================
' modUTL_Consolidate — Sheet Consolidation Tool
'==============================================================================
' PURPOSE:  Combine data from multiple identically-structured sheets into one
'           master sheet. User picks which sheets to include. Adds a "Source"
'           column so you know where each row came from.
'
' PUBLIC SUBS:
'   ConsolidateSheets    — Pick sheets from a list, combine into one
'   ConsolidateByPattern — Combine sheets matching a name pattern
'
' DEPENDENCIES: None (standalone). Works in any Excel workbook.
' VERSION:  1.0.0 | DATE: 2026-03-12
'==============================================================================
Option Explicit

Private Const OUTPUT_SHEET As String = "UTL_Consolidated"
Private Const CLR_HDR As Long = 7948043   ' RGB(11,71,121)

'==============================================================================
' PUBLIC: ConsolidateSheets
' User picks sheets from a numbered list. Data is combined into one sheet.
'==============================================================================
Public Sub ConsolidateSheets()
    On Error GoTo ErrHandler

    If ThisWorkbook.Sheets.Count < 2 Then
        MsgBox "You need at least 2 sheets to consolidate.", vbExclamation, "Consolidate Sheets"
        Exit Sub
    End If

    '--- Build sheet list ---
    Dim sheetList As String
    sheetList = "Select sheets to consolidate:" & vbCrLf & String(40, "-") & vbCrLf

    Dim i As Long
    For i = 1 To ThisWorkbook.Sheets.Count
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = Nothing
        Set ws = ThisWorkbook.Worksheets(i)
        On Error GoTo ErrHandler
        If Not ws Is Nothing Then
            sheetList = sheetList & "  " & i & ". " & ws.Name & vbCrLf
        End If
    Next i

    sheetList = sheetList & vbCrLf & "Enter sheet numbers separated by commas:" & vbCrLf & _
                "Example: 1,3,5,6"

    Dim choice As String
    choice = InputBox(sheetList, "Consolidate Sheets - Step 1 of 3")
    If Len(Trim(choice)) = 0 Then Exit Sub

    '--- Parse selections ---
    Dim parts() As String
    parts = Split(choice, ",")

    Dim selectedSheets() As String
    Dim selectedCount As Long
    selectedCount = 0
    ReDim selectedSheets(1 To ThisWorkbook.Sheets.Count)

    Dim p As Long
    For p = LBound(parts) To UBound(parts)
        Dim num As String
        num = Trim(parts(p))
        If IsNumeric(num) Then
            Dim idx As Long
            idx = CLng(num)
            If idx >= 1 And idx <= ThisWorkbook.Sheets.Count Then
                selectedCount = selectedCount + 1
                selectedSheets(selectedCount) = ThisWorkbook.Sheets(idx).Name
            End If
        End If
    Next p

    If selectedCount < 2 Then
        MsgBox "Please select at least 2 sheets to consolidate.", vbExclamation, "Consolidate Sheets"
        Exit Sub
    End If

    '--- Ask about headers ---
    Dim headerChoice As VbMsgBoxResult
    headerChoice = MsgBox("Does the first row of each sheet contain column headers?" & vbCrLf & vbCrLf & _
                           "YES = Skip the header row on sheets 2+ (avoid duplicate headers)" & vbCrLf & _
                           "NO = Copy all rows from every sheet", _
                           vbYesNoCancel + vbQuestion, "Consolidate Sheets - Step 2 of 3")
    If headerChoice = vbCancel Then Exit Sub
    Dim hasHeaders As Boolean
    hasHeaders = (headerChoice = vbYes)

    '--- Ask about source column ---
    Dim sourceChoice As VbMsgBoxResult
    sourceChoice = MsgBox("Add a 'Source Sheet' column at the end?" & vbCrLf & vbCrLf & _
                           "This adds a column showing which sheet each row came from." & vbCrLf & _
                           "Recommended for tracking.", _
                           vbYesNo + vbQuestion, "Consolidate Sheets - Step 3 of 3")
    Dim addSource As Boolean
    addSource = (sourceChoice = vbYes)

    '--- Create output sheet ---
    Application.ScreenUpdating = False
    Application.StatusBar = "Consolidating sheets..."

    Dim wsOut As Worksheet
    On Error Resume Next
    Set wsOut = ThisWorkbook.Sheets(OUTPUT_SHEET)
    On Error GoTo ErrHandler

    If Not wsOut Is Nothing Then
        Application.DisplayAlerts = False
        wsOut.Delete
        Application.DisplayAlerts = True
    End If

    Set wsOut = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsOut.Name = OUTPUT_SHEET

    '--- Pre-scan all sheets to find max column width (for consistent source column) ---
    Dim maxColWidth As Long
    maxColWidth = 0
    If addSource Then
        Dim sc As Long
        For sc = 1 To selectedCount
            Dim wsPre As Worksheet
            Set wsPre = ThisWorkbook.Sheets(selectedSheets(sc))
            Dim preCol As Long
            preCol = wsPre.Cells(1, wsPre.Columns.Count).End(xlToLeft).Column
            If wsPre.UsedRange.Columns.Count + wsPre.UsedRange.Column - 1 > preCol Then
                preCol = wsPre.UsedRange.Columns.Count + wsPre.UsedRange.Column - 1
            End If
            If preCol > maxColWidth Then maxColWidth = preCol
        Next sc
    End If

    '--- Consolidate data ---
    Dim outRow As Long
    outRow = 1
    Dim totalRows As Long
    totalRows = 0

    Dim s As Long
    For s = 1 To selectedCount
        Application.StatusBar = "Consolidating sheet " & s & " of " & selectedCount & ": " & selectedSheets(s) & "..."

        Dim wsSrc As Worksheet
        Set wsSrc = ThisWorkbook.Sheets(selectedSheets(s))

        Dim srcLastRow As Long, srcLastCol As Long
        srcLastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
        srcLastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

        ' Better detection via UsedRange
        If wsSrc.UsedRange.Columns.Count + wsSrc.UsedRange.Column - 1 > srcLastCol Then
            srcLastCol = wsSrc.UsedRange.Columns.Count + wsSrc.UsedRange.Column - 1
        End If

        If srcLastRow < 1 Then GoTo NextSheet
        If srcLastCol < 1 Then srcLastCol = 1

        Dim startRow As Long
        If s = 1 Then
            startRow = 1   ' Always copy first sheet fully (including headers)
        Else
            If hasHeaders Then startRow = 2 Else startRow = 1
        End If

        If startRow > srcLastRow Then GoTo NextSheet

        '--- Copy data ---
        Dim srcRange As Range
        Set srcRange = wsSrc.Range(wsSrc.Cells(startRow, 1), wsSrc.Cells(srcLastRow, srcLastCol))
        srcRange.Copy wsOut.Cells(outRow, 1)

        '--- Add source column ---
        If addSource Then
            Dim srcCol As Long
            srcCol = maxColWidth + 1

            ' Header on first sheet
            If s = 1 And hasHeaders Then
                wsOut.Cells(1, srcCol).Value = "Source Sheet"
                Dim dataStart As Long
                If hasHeaders Then dataStart = 2 Else dataStart = 1
                Dim dr As Long
                For dr = dataStart To outRow + (srcLastRow - startRow)
                    wsOut.Cells(dr, srcCol).Value = selectedSheets(s)
                Next dr
            Else
                Dim dr2 As Long
                For dr2 = outRow To outRow + (srcLastRow - startRow)
                    wsOut.Cells(dr2, srcCol).Value = selectedSheets(s)
                Next dr2
            End If
        End If

        totalRows = totalRows + (srcLastRow - startRow + 1)
        outRow = outRow + (srcLastRow - startRow + 1)

NextSheet:
    Next s

    '--- Style the header row ---
    If hasHeaders And outRow > 1 Then
        Dim hdrLastCol As Long
        hdrLastCol = wsOut.Cells(1, wsOut.Columns.Count).End(xlToLeft).Column
        Dim hdrRng As Range
        Set hdrRng = wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, hdrLastCol))
        hdrRng.Font.Bold = True
        hdrRng.Font.Color = RGB(255, 255, 255)
        hdrRng.Interior.Color = CLR_HDR
    End If

    wsOut.Columns.AutoFit
    wsOut.Activate
    wsOut.Range("A1").Select

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.CutCopyMode = False

    MsgBox "Consolidation complete!" & vbCrLf & vbCrLf & _
           "Sheets combined: " & selectedCount & vbCrLf & _
           "Total rows: " & Format(totalRows, "#,##0") & vbCrLf & _
           "Output sheet: " & OUTPUT_SHEET, _
           vbInformation, "Consolidate Sheets"

    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Consolidate Sheets"
End Sub

'==============================================================================
' PUBLIC: ConsolidateByPattern
' Combine sheets whose names match a keyword/pattern the user types.
'==============================================================================
Public Sub ConsolidateByPattern()
    On Error GoTo ErrHandler

    Dim keyword As String
    keyword = InputBox("Enter a keyword to match sheet names:" & vbCrLf & vbCrLf & _
                        "Examples:" & vbCrLf & _
                        "  'Q1' = all sheets with Q1 in the name" & vbCrLf & _
                        "  '2025' = all sheets with 2025 in the name" & vbCrLf & _
                        "  'Jan' = all sheets with Jan in the name", _
                        "Consolidate by Pattern - Step 1 of 3")
    If Len(Trim(keyword)) = 0 Then Exit Sub
    keyword = Trim(keyword)

    '--- Find matching sheets ---
    Dim matchNames() As String
    Dim matchCount As Long
    matchCount = 0
    ReDim matchNames(1 To ThisWorkbook.Sheets.Count)

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, keyword, vbTextCompare) > 0 Then
            matchCount = matchCount + 1
            matchNames(matchCount) = ws.Name
        End If
    Next ws

    If matchCount < 2 Then
        MsgBox "Found " & matchCount & " sheet(s) matching '" & keyword & "'." & vbCrLf & _
               "Need at least 2 sheets to consolidate.", vbExclamation, "Consolidate by Pattern"
        Exit Sub
    End If

    '--- Confirm matches ---
    Dim confirmMsg As String
    confirmMsg = "Found " & matchCount & " sheets matching '" & keyword & "':" & vbCrLf & vbCrLf
    Dim i As Long
    For i = 1 To matchCount
        confirmMsg = confirmMsg & "  " & i & ". " & matchNames(i) & vbCrLf
    Next i
    confirmMsg = confirmMsg & vbCrLf & "Consolidate these sheets?"

    Dim confirm As VbMsgBoxResult
    confirm = MsgBox(confirmMsg, vbYesNo + vbQuestion, "Consolidate by Pattern - Step 2 of 3")
    If confirm = vbNo Then Exit Sub

    '--- Ask about headers ---
    Dim headerChoice As VbMsgBoxResult
    headerChoice = MsgBox("Does the first row of each sheet contain column headers?" & vbCrLf & vbCrLf & _
                           "YES = Skip headers on sheets 2+" & vbCrLf & _
                           "NO = Copy all rows", _
                           vbYesNoCancel + vbQuestion, "Consolidate by Pattern - Step 3 of 3")
    If headerChoice = vbCancel Then Exit Sub
    Dim hasHeaders As Boolean
    hasHeaders = (headerChoice = vbYes)

    '--- Create output ---
    Application.ScreenUpdating = False
    Application.StatusBar = "Consolidating..."

    Dim wsOut As Worksheet
    On Error Resume Next
    Set wsOut = ThisWorkbook.Sheets(OUTPUT_SHEET)
    On Error GoTo ErrHandler

    If Not wsOut Is Nothing Then
        Application.DisplayAlerts = False
        wsOut.Delete
        Application.DisplayAlerts = True
    End If

    Set wsOut = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsOut.Name = OUTPUT_SHEET

    Dim outRow As Long
    outRow = 1
    Dim totalRows As Long
    totalRows = 0

    Dim s As Long
    For s = 1 To matchCount
        Application.StatusBar = "Consolidating " & s & " of " & matchCount & "..."

        Dim wsSrc As Worksheet
        Set wsSrc = ThisWorkbook.Sheets(matchNames(s))

        Dim srcLastRow As Long, srcLastCol As Long
        srcLastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
        srcLastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
        If wsSrc.UsedRange.Columns.Count + wsSrc.UsedRange.Column - 1 > srcLastCol Then
            srcLastCol = wsSrc.UsedRange.Columns.Count + wsSrc.UsedRange.Column - 1
        End If

        If srcLastRow < 1 Then GoTo NextPatternSheet

        Dim startRow As Long
        If s = 1 Then startRow = 1 Else If hasHeaders Then startRow = 2 Else startRow = 1

        If startRow > srcLastRow Then GoTo NextPatternSheet

        Dim srcRange As Range
        Set srcRange = wsSrc.Range(wsSrc.Cells(startRow, 1), wsSrc.Cells(srcLastRow, srcLastCol))
        srcRange.Copy wsOut.Cells(outRow, 1)

        ' Source column
        Dim srcCol As Long
        srcCol = srcLastCol + 1
        If s = 1 And hasHeaders Then wsOut.Cells(1, srcCol).Value = "Source Sheet"
        Dim fillStart As Long
        If s = 1 And hasHeaders Then fillStart = 2 Else fillStart = outRow
        Dim fr As Long
        For fr = fillStart To outRow + (srcLastRow - startRow)
            wsOut.Cells(fr, srcCol).Value = matchNames(s)
        Next fr

        totalRows = totalRows + (srcLastRow - startRow + 1)
        outRow = outRow + (srcLastRow - startRow + 1)

NextPatternSheet:
    Next s

    '--- Style header ---
    If hasHeaders And outRow > 1 Then
        Dim hdrLastCol As Long
        hdrLastCol = wsOut.Cells(1, wsOut.Columns.Count).End(xlToLeft).Column
        Dim hdrRng As Range
        Set hdrRng = wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, hdrLastCol))
        hdrRng.Font.Bold = True
        hdrRng.Font.Color = RGB(255, 255, 255)
        hdrRng.Interior.Color = CLR_HDR
    End If

    wsOut.Columns.AutoFit
    wsOut.Activate
    wsOut.Range("A1").Select

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.CutCopyMode = False

    MsgBox "Consolidation complete!" & vbCrLf & vbCrLf & _
           "Pattern: '" & keyword & "'" & vbCrLf & _
           "Sheets combined: " & matchCount & vbCrLf & _
           "Total rows: " & Format(totalRows, "#,##0") & vbCrLf & _
           "Output sheet: " & OUTPUT_SHEET, _
           vbInformation, "Consolidate by Pattern"

    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Consolidate by Pattern"
End Sub

'==============================================================================
' DIRECTOR WRAPPERS — Silent subs for video automation (no dialogs)
'==============================================================================

'==============================================================================
' DirectorConsolidateSheets
' Consolidates the named sheets into one output sheet with source tracking.
' Takes an array of sheet names (String). Assumes headers in row 1.
' No InputBox/MsgBox.
'
' Usage: DirectorConsolidateSheets Array("Sheet1","Sheet2","Sheet3")
'==============================================================================
Public Sub DirectorConsolidateSheets(sheetNames As Variant)
    On Error Resume Next

    If Not IsArray(sheetNames) Then
        Debug.Print "[Director] ConsolidateSheets: sheetNames must be an array."
        Exit Sub
    End If

    Dim selectedCount As Long
    selectedCount = 0

    ' Validate all sheet names exist
    Dim validNames() As String
    ReDim validNames(1 To UBound(sheetNames) - LBound(sheetNames) + 1)

    Dim n As Variant
    Dim ws As Worksheet
    For Each n In sheetNames
        Set ws = Nothing
        Set ws = ThisWorkbook.Sheets(CStr(n))
        If Not ws Is Nothing Then
            selectedCount = selectedCount + 1
            validNames(selectedCount) = CStr(n)
        End If
    Next n

    If selectedCount < 2 Then
        Debug.Print "[Director] ConsolidateSheets: Need at least 2 valid sheets. Found " & selectedCount & "."
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "Consolidating sheets..."

    ' Create or replace output sheet
    Dim wsOut As Worksheet
    Set wsOut = Nothing
    Set wsOut = ThisWorkbook.Sheets(OUTPUT_SHEET)

    If Not wsOut Is Nothing Then
        Application.DisplayAlerts = False
        wsOut.Delete
        Application.DisplayAlerts = True
    End If

    Set wsOut = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsOut.Name = OUTPUT_SHEET

    ' Pre-scan for max column width (for consistent source column)
    Dim maxColWidth As Long
    maxColWidth = 0
    Dim sc As Long
    For sc = 1 To selectedCount
        Dim wsPre As Worksheet
        Set wsPre = ThisWorkbook.Sheets(validNames(sc))
        Dim preCol As Long
        preCol = wsPre.Cells(1, wsPre.Columns.Count).End(xlToLeft).Column
        If wsPre.UsedRange.Columns.Count + wsPre.UsedRange.Column - 1 > preCol Then
            preCol = wsPre.UsedRange.Columns.Count + wsPre.UsedRange.Column - 1
        End If
        If preCol > maxColWidth Then maxColWidth = preCol
    Next sc

    ' Consolidate data (assumes headers in row 1)
    Dim outRow As Long
    outRow = 1
    Dim totalRows As Long
    totalRows = 0

    Dim s As Long
    For s = 1 To selectedCount
        Application.StatusBar = "Consolidating sheet " & s & " of " & selectedCount & ": " & validNames(s) & "..."

        Dim wsSrc As Worksheet
        Set wsSrc = ThisWorkbook.Sheets(validNames(s))

        Dim srcLastRow As Long, srcLastCol As Long
        srcLastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
        srcLastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
        If wsSrc.UsedRange.Columns.Count + wsSrc.UsedRange.Column - 1 > srcLastCol Then
            srcLastCol = wsSrc.UsedRange.Columns.Count + wsSrc.UsedRange.Column - 1
        End If

        If srcLastRow < 1 Then GoTo DirNextSheet
        If srcLastCol < 1 Then srcLastCol = 1

        Dim startRow As Long
        If s = 1 Then startRow = 1 Else startRow = 2  ' Skip header on sheets 2+

        If startRow > srcLastRow Then GoTo DirNextSheet

        ' Copy data
        Dim srcRange As Range
        Set srcRange = wsSrc.Range(wsSrc.Cells(startRow, 1), wsSrc.Cells(srcLastRow, srcLastCol))
        srcRange.Copy wsOut.Cells(outRow, 1)

        ' Add source column
        Dim srcCol As Long
        srcCol = maxColWidth + 1

        If s = 1 Then
            wsOut.Cells(1, srcCol).Value = "Source Sheet"
            Dim dr As Long
            For dr = 2 To outRow + (srcLastRow - startRow)
                wsOut.Cells(dr, srcCol).Value = validNames(s)
            Next dr
        Else
            Dim dr2 As Long
            For dr2 = outRow To outRow + (srcLastRow - startRow)
                wsOut.Cells(dr2, srcCol).Value = validNames(s)
            Next dr2
        End If

        totalRows = totalRows + (srcLastRow - startRow + 1)
        outRow = outRow + (srcLastRow - startRow + 1)

DirNextSheet:
    Next s

    ' Style the header row
    If outRow > 1 Then
        Dim hdrLastCol As Long
        hdrLastCol = wsOut.Cells(1, wsOut.Columns.Count).End(xlToLeft).Column
        Dim hdrRng As Range
        Set hdrRng = wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, hdrLastCol))
        hdrRng.Font.Bold = True
        hdrRng.Font.Color = RGB(255, 255, 255)
        hdrRng.Interior.Color = CLR_HDR
    End If

    wsOut.Columns.AutoFit
    wsOut.Activate
    wsOut.Range("A1").Select

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.CutCopyMode = False

    Debug.Print "[Director] ConsolidateSheets: " & selectedCount & " sheets combined, " & totalRows & " total rows. Output: " & OUTPUT_SHEET
End Sub
