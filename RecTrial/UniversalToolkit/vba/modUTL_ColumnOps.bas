Attribute VB_Name = "modUTL_ColumnOps"
'==============================================================================
' modUTL_ColumnOps — Column Split, Combine & Extract Tools
'==============================================================================
' PURPOSE:  Split one column into multiple, combine columns, or extract
'           patterns from text. User selects ranges and chooses options.
'
' PUBLIC SUBS:
'   SplitColumn        — Split a column by a delimiter into multiple columns
'   CombineColumns     — Merge multiple columns into one with a separator
'   ExtractPattern     — Extract numbers, emails, or custom patterns from text
'   SwapColumns        — Swap the contents of two columns
'
' DEPENDENCIES: None (standalone). Works in any Excel workbook.
' VERSION:  1.0.0 | DATE: 2026-03-12
'==============================================================================
Option Explicit

'==============================================================================
' PUBLIC: SplitColumn
' Splits a selected column by a delimiter. User picks the delimiter.
' New columns are inserted to the right — existing data is NOT overwritten.
'==============================================================================
Public Sub SplitColumn()
    On Error GoTo ErrHandler

    Dim rng As Range

    ' If a multi-cell range is already selected, use it (Director-friendly)
    If Not TypeOf Selection Is Range Then GoTo AskSplitRange
    If Selection.Cells.Count > 1 And Selection.Columns.Count = 1 Then
        Set rng = Selection
        GoTo HaveSplitRange
    End If

AskSplitRange:
    MsgBox "Select the column range you want to split." & vbCrLf & vbCrLf & _
           "Select ONLY the data cells (not the header)." & vbCrLf & _
           "New columns will be inserted to the right.", _
           vbInformation, "Split Column"

    On Error Resume Next
    Set rng = Application.InputBox("Select the cells to split:", _
                                    "Split Column - Step 1 of 2", Type:=8)
    On Error GoTo ErrHandler
    If rng Is Nothing Then Exit Sub

HaveSplitRange:

    If rng.Columns.Count > 1 Then
        MsgBox "Please select cells from a single column only.", vbExclamation, "Split Column"
        Exit Sub
    End If

    '--- Ask for delimiter ---
    Dim delimChoice As String
    delimChoice = InputBox("Choose a delimiter:" & vbCrLf & vbCrLf & _
                           "  1. Comma  ( , )" & vbCrLf & _
                           "  2. Semicolon  ( ; )" & vbCrLf & _
                           "  3. Space" & vbCrLf & _
                           "  4. Dash  ( - )" & vbCrLf & _
                           "  5. Pipe  ( | )" & vbCrLf & _
                           "  6. Custom (you type it)" & vbCrLf & vbCrLf & _
                           "Enter number or type your custom delimiter:", _
                           "Split Column - Step 2 of 2")
    If Len(delimChoice) = 0 Then Exit Sub

    Dim delim As String
    Select Case Trim(delimChoice)
        Case "1": delim = ","
        Case "2": delim = ";"
        Case "3": delim = " "
        Case "4": delim = "-"
        Case "5": delim = "|"
        Case "6"
            delim = InputBox("Type your custom delimiter:", "Custom Delimiter")
            If Len(delim) = 0 Then Exit Sub
        Case Else
            delim = delimChoice  ' User typed the actual delimiter
    End Select

    Application.ScreenUpdating = False

    '--- Find max number of parts ---
    Dim maxParts As Long
    maxParts = 1
    Dim cell As Range
    For Each cell In rng.Cells
        If Not IsEmpty(cell.Value) Then
            Dim testParts() As String
            testParts = Split(CStr(cell.Value), delim)
            If UBound(testParts) + 1 > maxParts Then maxParts = UBound(testParts) + 1
        End If
    Next cell

    If maxParts <= 1 Then
        Application.ScreenUpdating = True
        MsgBox "No cells contain the delimiter '" & delim & "'." & vbCrLf & _
               "Nothing to split.", vbInformation, "Split Column"
        Exit Sub
    End If

    '--- Insert new columns to the right ---
    Dim insertCol As Long
    insertCol = rng.Column + 1

    Dim colsNeeded As Long
    colsNeeded = maxParts - 1  ' First part stays, rest go in new columns

    Dim c As Long
    For c = 1 To colsNeeded
        rng.Parent.Columns(insertCol).Insert Shift:=xlToRight
    Next c

    '--- Split the data ---
    Dim splitCount As Long
    splitCount = 0

    For Each cell In rng.Cells
        If Not IsEmpty(cell.Value) Then
            Dim cellParts() As String
            cellParts = Split(CStr(cell.Value), delim)

            If UBound(cellParts) >= 1 Then
                ' Keep first part in original cell
                cell.Value = Trim(cellParts(0))

                ' Put remaining parts in new columns
                Dim pi As Long
                For pi = 1 To UBound(cellParts)
                    cell.Offset(0, pi).Value = Trim(cellParts(pi))
                Next pi

                splitCount = splitCount + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Split complete!" & vbCrLf & vbCrLf & _
           "Cells split: " & splitCount & vbCrLf & _
           "New columns added: " & colsNeeded & vbCrLf & _
           "Delimiter used: '" & delim & "'", _
           vbInformation, "Split Column"

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Split Column"
End Sub

'==============================================================================
' PUBLIC: CombineColumns
' User selects multiple columns. Combines them into a new column with a separator.
'==============================================================================
Public Sub CombineColumns()
    On Error GoTo ErrHandler

    Dim rng As Range

    ' If a multi-column range is already selected, use it (Director-friendly)
    If Not TypeOf Selection Is Range Then GoTo AskCombineRange
    If Selection.Columns.Count >= 2 Then
        Set rng = Selection
        GoTo HaveCombineRange
    End If

AskCombineRange:
    MsgBox "Select the columns you want to combine." & vbCrLf & vbCrLf & _
           "Select the full data range (all columns to merge)." & vbCrLf & _
           "The combined result will be placed in a new column to the right.", _
           vbInformation, "Combine Columns"

    On Error Resume Next
    Set rng = Application.InputBox("Select the columns to combine:", _
                                    "Combine Columns - Step 1 of 2", Type:=8)
    On Error GoTo ErrHandler
    If rng Is Nothing Then Exit Sub

    If rng.Columns.Count < 2 Then
        MsgBox "Please select at least 2 columns to combine.", vbExclamation, "Combine Columns"
        Exit Sub
    End If

HaveCombineRange:

    '--- Ask for separator ---
    Dim sepChoice As String
    sepChoice = InputBox("Choose a separator for the combined values:" & vbCrLf & vbCrLf & _
                          "  1. Comma + Space  ( , )" & vbCrLf & _
                          "  2. Space" & vbCrLf & _
                          "  3. Dash  ( - )" & vbCrLf & _
                          "  4. Pipe  ( | )" & vbCrLf & _
                          "  5. No separator (just concatenate)" & vbCrLf & _
                          "  6. Custom (you type it)" & vbCrLf & vbCrLf & _
                          "Enter number:", _
                          "Combine Columns - Step 2 of 2")
    If Len(sepChoice) = 0 Then Exit Sub

    Dim sep As String
    Select Case Trim(sepChoice)
        Case "1": sep = ", "
        Case "2": sep = " "
        Case "3": sep = " - "
        Case "4": sep = " | "
        Case "5": sep = ""
        Case "6"
            sep = InputBox("Type your custom separator:", "Custom Separator")
        Case Else
            sep = sepChoice
    End Select

    Application.ScreenUpdating = False

    '--- Insert result column ---
    Dim resultCol As Long
    resultCol = rng.Column + rng.Columns.Count
    rng.Parent.Columns(resultCol).Insert Shift:=xlToRight

    '--- Combine values ---
    Dim r As Long
    Dim combined As String
    Dim rowCount As Long
    rowCount = 0

    For r = 1 To rng.Rows.Count
        combined = ""
        Dim ci As Long
        For ci = 1 To rng.Columns.Count
            Dim cellVal As String
            cellVal = CStr(Nz(rng.Cells(r, ci).Value))
            If Len(cellVal) > 0 Then
                If Len(combined) > 0 And Len(sep) > 0 Then
                    combined = combined & sep
                End If
                combined = combined & cellVal
            End If
        Next ci

        rng.Parent.Cells(rng.Row + r - 1, resultCol).Value = combined
        rowCount = rowCount + 1
    Next r

    Application.ScreenUpdating = True

    MsgBox "Combine complete!" & vbCrLf & vbCrLf & _
           "Rows combined: " & rowCount & vbCrLf & _
           "Result column: " & Split(rng.Parent.Cells(1, resultCol).Address, "$")(1) & vbCrLf & _
           "Separator: '" & sep & "'", _
           vbInformation, "Combine Columns"

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Combine Columns"
End Sub

'==============================================================================
' PUBLIC: ExtractPattern
' Extract numbers, emails, or dates from text cells into a new column.
'==============================================================================
Public Sub ExtractPattern()
    On Error GoTo ErrHandler

    MsgBox "Select the cells containing text to extract from." & vbCrLf & vbCrLf & _
           "The extracted values will be placed in a new column to the right.", _
           vbInformation, "Extract Pattern"

    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox("Select the cells to extract from:", _
                                    "Extract Pattern - Step 1 of 2", Type:=8)
    On Error GoTo ErrHandler
    If rng Is Nothing Then Exit Sub

    If rng.Columns.Count > 1 Then
        MsgBox "Please select a single column.", vbExclamation, "Extract Pattern"
        Exit Sub
    End If

    '--- Ask what to extract ---
    Dim extractChoice As String
    extractChoice = InputBox("What do you want to extract?" & vbCrLf & vbCrLf & _
                              "  1. Numbers only (first number found)" & vbCrLf & _
                              "  2. All numbers (concatenated)" & vbCrLf & _
                              "  3. Text before a delimiter" & vbCrLf & _
                              "  4. Text after a delimiter" & vbCrLf & _
                              "  5. First N characters" & vbCrLf & _
                              "  6. Last N characters" & vbCrLf & vbCrLf & _
                              "Enter number:", _
                              "Extract Pattern - Step 2 of 2")
    If Len(Trim(extractChoice)) = 0 Then Exit Sub

    Application.ScreenUpdating = False

    '--- Insert result column ---
    Dim resultCol As Long
    resultCol = rng.Column + 1
    rng.Parent.Columns(resultCol).Insert Shift:=xlToRight

    Dim cell As Range
    Dim extracted As String
    Dim extractCount As Long
    extractCount = 0

    Select Case Trim(extractChoice)
        Case "1"  ' First number
            For Each cell In rng.Cells
                extracted = ExtractFirstNumber(CStr(Nz(cell.Value)))
                cell.Offset(0, 1).Value = extracted
                If Len(extracted) > 0 Then extractCount = extractCount + 1
            Next cell

        Case "2"  ' All numbers
            For Each cell In rng.Cells
                extracted = ExtractAllNumbers(CStr(Nz(cell.Value)))
                cell.Offset(0, 1).Value = extracted
                If Len(extracted) > 0 Then extractCount = extractCount + 1
            Next cell

        Case "3"  ' Before delimiter
            Dim delimBefore As String
            Application.ScreenUpdating = True
            delimBefore = InputBox("Type the delimiter:" & vbCrLf & _
                                   "Text BEFORE this delimiter will be extracted.", _
                                   "Extract Before Delimiter")
            Application.ScreenUpdating = False
            If Len(delimBefore) = 0 Then GoTo CleanUp

            For Each cell In rng.Cells
                Dim txt3 As String
                txt3 = CStr(Nz(cell.Value))
                Dim pos3 As Long
                pos3 = InStr(1, txt3, delimBefore)
                If pos3 > 0 Then
                    extracted = Trim(Left(txt3, pos3 - 1))
                    extractCount = extractCount + 1
                Else
                    extracted = ""
                End If
                cell.Offset(0, 1).Value = extracted
            Next cell

        Case "4"  ' After delimiter
            Dim delimAfter As String
            Application.ScreenUpdating = True
            delimAfter = InputBox("Type the delimiter:" & vbCrLf & _
                                  "Text AFTER this delimiter will be extracted.", _
                                  "Extract After Delimiter")
            Application.ScreenUpdating = False
            If Len(delimAfter) = 0 Then GoTo CleanUp

            For Each cell In rng.Cells
                Dim txt4 As String
                txt4 = CStr(Nz(cell.Value))
                Dim pos4 As Long
                pos4 = InStr(1, txt4, delimAfter)
                If pos4 > 0 Then
                    extracted = Trim(Mid(txt4, pos4 + Len(delimAfter)))
                    extractCount = extractCount + 1
                Else
                    extracted = ""
                End If
                cell.Offset(0, 1).Value = extracted
            Next cell

        Case "5"  ' First N chars
            Dim firstN As String
            Application.ScreenUpdating = True
            firstN = InputBox("How many characters from the start?", "First N Characters")
            Application.ScreenUpdating = False
            If Not IsNumeric(firstN) Then GoTo CleanUp
            Dim n5 As Long
            n5 = CLng(firstN)

            For Each cell In rng.Cells
                Dim txt5 As String
                txt5 = CStr(Nz(cell.Value))
                cell.Offset(0, 1).Value = Left(txt5, n5)
                If Len(txt5) > 0 Then extractCount = extractCount + 1
            Next cell

        Case "6"  ' Last N chars
            Dim lastN As String
            Application.ScreenUpdating = True
            lastN = InputBox("How many characters from the end?", "Last N Characters")
            Application.ScreenUpdating = False
            If Not IsNumeric(lastN) Then GoTo CleanUp
            Dim n6 As Long
            n6 = CLng(lastN)

            For Each cell In rng.Cells
                Dim txt6 As String
                txt6 = CStr(Nz(cell.Value))
                cell.Offset(0, 1).Value = Right(txt6, n6)
                If Len(txt6) > 0 Then extractCount = extractCount + 1
            Next cell

        Case Else
            Application.ScreenUpdating = True
            MsgBox "Invalid choice.", vbExclamation, "Extract Pattern"
            Exit Sub
    End Select

CleanUp:
    Application.ScreenUpdating = True

    MsgBox "Extraction complete!" & vbCrLf & vbCrLf & _
           "Values extracted: " & extractCount, _
           vbInformation, "Extract Pattern"

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Extract Pattern"
End Sub

'==============================================================================
' PUBLIC: SwapColumns
' Swap the contents of two user-selected columns.
'==============================================================================
Public Sub SwapColumns()
    On Error GoTo ErrHandler

    MsgBox "You will select two single-column ranges to swap." & vbCrLf & _
           "Both ranges must have the same number of rows.", _
           vbInformation, "Swap Columns"

    Dim rng1 As Range, rng2 As Range

    On Error Resume Next
    Set rng1 = Application.InputBox("Select the FIRST column:", "Swap Columns - Column 1", Type:=8)
    On Error GoTo ErrHandler
    If rng1 Is Nothing Then Exit Sub

    On Error Resume Next
    Set rng2 = Application.InputBox("Select the SECOND column:", "Swap Columns - Column 2", Type:=8)
    On Error GoTo ErrHandler
    If rng2 Is Nothing Then Exit Sub

    If rng1.Columns.Count > 1 Or rng2.Columns.Count > 1 Then
        MsgBox "Please select single columns only.", vbExclamation, "Swap Columns"
        Exit Sub
    End If

    If rng1.Rows.Count <> rng2.Rows.Count Then
        MsgBox "Both ranges must have the same number of rows.", vbExclamation, "Swap Columns"
        Exit Sub
    End If

    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("Swap contents of:" & vbCrLf & _
                      "  Column 1: " & rng1.Address(External:=True) & vbCrLf & _
                      "  Column 2: " & rng2.Address(External:=True) & vbCrLf & vbCrLf & _
                      "Proceed?", vbYesNo + vbQuestion, "Swap Columns")
    If confirm = vbNo Then Exit Sub

    Application.ScreenUpdating = False

    ' Store column 1 values
    Dim temp() As Variant
    ReDim temp(1 To rng1.Rows.Count)

    Dim r As Long
    For r = 1 To rng1.Rows.Count
        temp(r) = rng1.Cells(r, 1).Value
    Next r

    ' Copy column 2 to column 1
    For r = 1 To rng1.Rows.Count
        rng1.Cells(r, 1).Value = rng2.Cells(r, 1).Value
    Next r

    ' Copy saved column 1 to column 2
    For r = 1 To rng2.Rows.Count
        rng2.Cells(r, 1).Value = temp(r)
    Next r

    Application.ScreenUpdating = True

    MsgBox "Columns swapped successfully! (" & rng1.Rows.Count & " rows)", _
           vbInformation, "Swap Columns"

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Swap Columns"
End Sub

'==============================================================================
' PRIVATE Helpers
'==============================================================================
Private Function ExtractFirstNumber(ByVal txt As String) As String
    Dim i As Long
    Dim result As String
    Dim inNumber As Boolean
    result = ""
    inNumber = False

    For i = 1 To Len(txt)
        Dim ch As String
        ch = Mid(txt, i, 1)
        If ch >= "0" And ch <= "9" Then
            result = result & ch
            inNumber = True
        ElseIf ch = "." And inNumber And InStr(result, ".") = 0 Then
            result = result & ch
        ElseIf inNumber Then
            Exit For
        End If
    Next i

    ExtractFirstNumber = result
End Function

Private Function ExtractAllNumbers(ByVal txt As String) As String
    Dim i As Long
    Dim result As String
    result = ""

    For i = 1 To Len(txt)
        Dim ch As String
        ch = Mid(txt, i, 1)
        If ch >= "0" And ch <= "9" Or ch = "." Then
            result = result & ch
        End If
    Next i

    ExtractAllNumbers = result
End Function

Private Function Nz(ByVal v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then Nz = "" Else Nz = CStr(v)
End Function

'==============================================================================
' DIRECTOR WRAPPERS — Silent subs for video automation (no dialogs)
'==============================================================================

'==============================================================================
' DirectorSplitColumn
' Splits the given single-column range by delimiter. Inserts new columns to
' the right. No InputBox/MsgBox/Application.InputBox.
'==============================================================================
Public Sub DirectorSplitColumn(ByVal sheetName As String, ByVal rangeAddress As String, ByVal delimiter As String)
    ' Robust string-parameter version. Avoids Range marshaling issues through
    ' Application.Run by resolving the sheet + address on arrival.
    On Error Resume Next

    If Len(sheetName) = 0 Or Len(rangeAddress) = 0 Or Len(delimiter) = 0 Then
        Debug.Print "[Director] SplitColumn: Missing required parameter."
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = Nothing
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    If ws Is Nothing Then
        Debug.Print "[Director] SplitColumn: Sheet '" & sheetName & "' not found."
        Exit Sub
    End If

    Dim rng As Range
    Set rng = Nothing
    Set rng = ws.Range(rangeAddress)
    If rng Is Nothing Then
        Debug.Print "[Director] SplitColumn: Range '" & rangeAddress & "' not valid."
        Exit Sub
    End If
    If rng.Columns.Count > 1 Then
        Debug.Print "[Director] SplitColumn: Range must be a single column."
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Read all values into a 1-D array first so we're not iterating a Range
    ' object across structural changes (column inserts).
    Dim rowCount As Long
    rowCount = rng.Rows.Count
    Dim startRowIdx As Long
    startRowIdx = rng.Row
    Dim srcCol As Long
    srcCol = rng.Column

    Dim vals() As String
    ReDim vals(1 To rowCount)
    Dim i As Long
    For i = 1 To rowCount
        vals(i) = CStr(ws.Cells(startRowIdx + i - 1, srcCol).Value)
    Next i

    ' Determine max parts in any cell
    Dim maxParts As Long
    maxParts = 1
    For i = 1 To rowCount
        If Len(vals(i)) > 0 Then
            Dim testParts() As String
            testParts = Split(vals(i), delimiter)
            If UBound(testParts) + 1 > maxParts Then maxParts = UBound(testParts) + 1
        End If
    Next i

    If maxParts <= 1 Then
        Application.ScreenUpdating = True
        Debug.Print "[Director] SplitColumn: No cells contain delimiter '" & delimiter & "'."
        Exit Sub
    End If

    ' Insert empty columns immediately to the right of the source column
    Dim colsNeeded As Long
    colsNeeded = maxParts - 1
    Dim c As Long
    For c = 1 To colsNeeded
        ws.Columns(srcCol + 1).Insert Shift:=xlToRight
    Next c

    ' Write split parts back using explicit row/column indices (no Range
    ' enumeration — safe across the earlier insert).
    Dim splitCount As Long
    splitCount = 0
    For i = 1 To rowCount
        If Len(vals(i)) > 0 Then
            Dim cellParts() As String
            cellParts = Split(vals(i), delimiter)
            If UBound(cellParts) >= 1 Then
                ws.Cells(startRowIdx + i - 1, srcCol).Value = Trim(cellParts(0))
                Dim pi As Long
                For pi = 1 To UBound(cellParts)
                    ws.Cells(startRowIdx + i - 1, srcCol + pi).Value = Trim(cellParts(pi))
                Next pi
                splitCount = splitCount + 1
            End If
        End If
    Next i

    Application.ScreenUpdating = True

    Debug.Print "[Director] SplitColumn: '" & sheetName & "' " & rangeAddress & _
                " — " & splitCount & " cell(s) split by '" & delimiter & "', " & _
                colsNeeded & " new column(s) added."
End Sub

'==============================================================================
' DirectorCombineColumns
' Combines columns in the given multi-column range with separator. Result
' goes in a new column to the right. No InputBox/MsgBox/Application.InputBox.
'==============================================================================
Public Sub DirectorCombineColumns(ByVal sheetName As String, ByVal rangeAddress As String, ByVal separator As String)
    ' Robust string-parameter version. Avoids Range marshaling via Application.Run.
    On Error Resume Next

    If Len(sheetName) = 0 Or Len(rangeAddress) = 0 Then
        Debug.Print "[Director] CombineColumns: Missing required parameter."
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = Nothing
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    If ws Is Nothing Then
        Debug.Print "[Director] CombineColumns: Sheet '" & sheetName & "' not found."
        Exit Sub
    End If

    Dim rng As Range
    Set rng = Nothing
    Set rng = ws.Range(rangeAddress)
    If rng Is Nothing Then
        Debug.Print "[Director] CombineColumns: Range '" & rangeAddress & "' not valid."
        Exit Sub
    End If
    If rng.Columns.Count < 2 Then
        Debug.Print "[Director] CombineColumns: Need at least 2 columns."
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Snapshot all source values into a 2-D array first so we read before the
    ' structural change (column insert) rearranges references.
    Dim rowCount As Long, colCount As Long
    rowCount = rng.Rows.Count
    colCount = rng.Columns.Count
    Dim startRowIdx As Long, startColIdx As Long
    startRowIdx = rng.Row
    startColIdx = rng.Column

    Dim snap() As String
    ReDim snap(1 To rowCount, 1 To colCount)
    Dim r As Long, c As Long
    For r = 1 To rowCount
        For c = 1 To colCount
            Dim v As Variant
            v = ws.Cells(startRowIdx + r - 1, startColIdx + c - 1).Value
            If IsEmpty(v) Or IsNull(v) Then
                snap(r, c) = ""
            Else
                snap(r, c) = CStr(v)
            End If
        Next c
    Next r

    ' Insert the result column immediately after the source range
    Dim resultCol As Long
    resultCol = startColIdx + colCount
    ws.Columns(resultCol).Insert Shift:=xlToRight

    ' Write combined values from the array snapshot
    Dim combined As String
    For r = 1 To rowCount
        combined = ""
        For c = 1 To colCount
            If Len(snap(r, c)) > 0 Then
                If Len(combined) > 0 And Len(separator) > 0 Then
                    combined = combined & separator
                End If
                combined = combined & snap(r, c)
            End If
        Next c
        ws.Cells(startRowIdx + r - 1, resultCol).Value = combined
    Next r

    Application.ScreenUpdating = True

    Debug.Print "[Director] CombineColumns: '" & sheetName & "' " & rangeAddress & _
                " — " & rowCount & " row(s) combined with separator '" & separator & "'."
End Sub
