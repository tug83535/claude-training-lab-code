Attribute VB_Name = "modUTL_DataCleaning"
Option Explicit

' ============================================================
' KBT Universal Tools — Data Cleaning Module
' Works on ANY Excel file — no project-specific setup required
' Install in Personal.xlsb to use across all Excel sessions
' Tools: 12 | Tier 1: 9 | Tier 2: 3
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
' TOOL 1 — Unmerge Cells & Fill Down                [TIER 1]
' Unmerges every merged cell in selection, fills value down
' Run: select the range first, then run this macro
' ============================================================
Sub UnmergeAndFillDown()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first, then run this tool.", vbExclamation, "UTL Data Cleaning"
        Exit Sub
    End If

    Dim rng As Range
    Set rng = Selection

    On Error GoTo ErrHandler
    UTL_TurboOn

    rng.UnMerge

    Dim c As Range
    For Each c In rng
        If c.Row > rng.Row Then
            If IsEmpty(c) Or c.Value = "" Then
                c.Value = c.Offset(-1, 0).Value
            End If
        End If
    Next c

    UTL_TurboOff
    MsgBox "Done! All merged cells unmerged and filled down.", vbInformation, "UTL Data Cleaning"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Data Cleaning"
End Sub

' ============================================================
' TOOL 2 — Fill Blanks Down                          [TIER 1]
' Fills every blank cell with the value from the cell above
' Run: select the column(s) first, then run this macro
' ============================================================
Sub FillBlanksDown()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first, then run this tool.", vbExclamation, "UTL Data Cleaning"
        Exit Sub
    End If

    Dim rng As Range
    Set rng = Selection

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim count As Long
    Dim c As Range
    For Each c In rng
        If c.Row > rng.Row Then
            If IsEmpty(c) Or c.Value = "" Then
                c.Value = c.Offset(-1, 0).Value
                count = count + 1
            End If
        End If
    Next c

    UTL_TurboOff
    MsgBox "Done! " & count & " blank cells filled down.", vbInformation, "UTL Data Cleaning"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Data Cleaning"
End Sub

' ============================================================
' TOOL 3 — Convert Text to Numbers                   [TIER 1]
' Fixes cells that store numbers as text so they sum correctly
' Run: select the range, then run this macro
' ============================================================
Sub ConvertTextToNumbers()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first, then run this tool.", vbExclamation, "UTL Data Cleaning"
        Exit Sub
    End If

    Dim rng As Range
    Set rng = Selection

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim count As Long
    Dim c As Range
    For Each c In rng
        If Not IsEmpty(c) Then
            If c.HasFormula = False Then
                If IsNumeric(c.Value) And c.NumberFormat = "@" Then
                    c.Value = CDbl(c.Value)
                    c.NumberFormat = "#,##0.00"
                    count = count + 1
                ElseIf VarType(c.Value) = vbString And IsNumeric(Trim(c.Value)) Then
                    c.Value = CDbl(Trim(c.Value))
                    count = count + 1
                End If
            End If
        End If
    Next c

    UTL_TurboOff
    MsgBox "Done! " & count & " text-stored numbers converted.", vbInformation, "UTL Data Cleaning"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Data Cleaning"
End Sub

' ============================================================
' TOOL 4 — Remove Leading/Trailing Spaces            [TIER 1]
' Trims invisible spaces from all text cells in selection
' Run: select the range, then run this macro
' ============================================================
Sub RemoveLeadingTrailingSpaces()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first, then run this tool.", vbExclamation, "UTL Data Cleaning"
        Exit Sub
    End If

    Dim rng As Range
    Set rng = Selection

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim count As Long
    Dim c As Range
    For Each c In rng
        If Not IsEmpty(c) And VarType(c.Value) = vbString Then
            Dim cleaned As String
            cleaned = WorksheetFunction.Trim(c.Value)
            If cleaned <> c.Value Then
                c.Value = cleaned
                count = count + 1
            End If
        End If
    Next c

    UTL_TurboOff
    MsgBox "Done! Spaces cleaned in " & count & " cells.", vbInformation, "UTL Data Cleaning"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Data Cleaning"
End Sub

' ============================================================
' TOOL 5 — Delete Blank Rows                         [TIER 1]
' Removes completely empty rows from the active sheet
' Run: no selection needed — works on entire used range
' ============================================================
Sub DeleteBlankRows()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    If MsgBox("Delete all blank rows on sheet '" & ws.Name & "'?", _
              vbQuestion + vbYesNo, "UTL Data Cleaning") = vbNo Then Exit Sub

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim count As Long
    Dim i As Long
    For i = lastRow To 1 Step -1
        If WorksheetFunction.CountA(ws.Rows(i)) = 0 Then
            ws.Rows(i).Delete
            count = count + 1
        End If
    Next i

    UTL_TurboOff
    MsgBox "Done! " & count & " blank rows deleted.", vbInformation, "UTL Data Cleaning"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Data Cleaning"
End Sub

' ============================================================
' TOOL 6 — Replace Error Values                      [TIER 1]
' Replaces #N/A, #REF!, #VALUE!, #DIV/0!, #NAME? with blank
' Run: select the range, then run this macro
' ============================================================
Sub ReplaceErrorValues()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first, then run this tool.", vbExclamation, "UTL Data Cleaning"
        Exit Sub
    End If

    Dim replacement As String
    replacement = InputBox("Replace errors with: (leave blank to clear them)", _
                           "UTL Data Cleaning", "")

    Dim rng As Range
    Set rng = Selection

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim count As Long
    Dim c As Range
    For Each c In rng
        If IsError(c.Value) Then
            c.Value = replacement
            count = count + 1
        End If
    Next c

    UTL_TurboOff
    MsgBox "Done! " & count & " error cells replaced.", vbInformation, "UTL Data Cleaning"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Data Cleaning"
End Sub

' ============================================================
' TOOL 7 — Highlight Duplicate Rows                  [TIER 1]
' Colors duplicate rows yellow — does NOT delete anything
' Run: select the key column to check for duplicates
' ============================================================
Sub HighlightDuplicateRows()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select the column to check for duplicates.", vbExclamation, "UTL Data Cleaning"
        Exit Sub
    End If

    Dim keyCol As Range
    Set keyCol = Selection.Columns(1)

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim count As Long
    Dim c As Range
    For Each c In keyCol
        If Not IsEmpty(c) Then
            Dim key As String
            key = CStr(c.Value)
            If dict.exists(key) Then
                c.EntireRow.Interior.Color = RGB(255, 235, 59)
                dict(key).EntireRow.Interior.Color = RGB(255, 235, 59)
                count = count + 1
            Else
                dict.Add key, c
            End If
        End If
    Next c

    UTL_TurboOff
    If count = 0 Then
        MsgBox "No duplicates found in the selected column.", vbInformation, "UTL Data Cleaning"
    Else
        MsgBox "Found " & count & " duplicate values — rows highlighted yellow." & Chr(10) & _
               "Review before deleting. Use Remove Duplicate Rows when ready.", _
               vbInformation, "UTL Data Cleaning"
    End If
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Data Cleaning"
End Sub

' ============================================================
' TOOL 8 — Remove Duplicate Rows                     [TIER 1]
' Permanently deletes duplicate rows based on selected column
' Tip: run Highlight Duplicates first to review before deleting
' ============================================================
Sub RemoveDuplicateRows()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select the key column to check for duplicates.", vbExclamation, "UTL Data Cleaning"
        Exit Sub
    End If

    If MsgBox("This will PERMANENTLY DELETE duplicate rows." & Chr(10) & _
              "Tip: Run 'Highlight Duplicate Rows' first to review." & Chr(10) & Chr(10) & _
              "Continue?", vbExclamation + vbYesNo, "UTL Data Cleaning") = vbNo Then Exit Sub

    Dim keyCol As Range
    Set keyCol = Selection.Columns(1)

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim count As Long
    Dim i As Long
    For i = keyCol.Cells.Count To 1 Step -1
        Dim c As Range
        Set c = keyCol.Cells(i)
        If Not IsEmpty(c) Then
            Dim key As String
            key = CStr(c.Value)
            If dict.exists(key) Then
                c.EntireRow.Delete
                count = count + 1
            Else
                dict.Add key, True
            End If
        End If
    Next i

    UTL_TurboOff
    MsgBox "Done! " & count & " duplicate rows deleted.", vbInformation, "UTL Data Cleaning"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Data Cleaning"
End Sub

' ============================================================
' TOOL 9 — Multi-Replace Data Cleaner                [TIER 1]
' Batch find-and-replace based on a two-column mapping table
' Setup: create a sheet with Find in column A, Replace in column B
' Perfect for standardizing account names, cost center codes, etc.
' ============================================================
Sub MultiReplaceDataCleaner()
    Dim mapSheetName As String
    mapSheetName = InputBox("Enter the name of your mapping sheet:" & Chr(10) & _
                            "(Sheet must have: Find value in column A, Replace value in column B)", _
                            "UTL Data Cleaning", "Mapping")

    If mapSheetName = "" Then Exit Sub

    Dim mapWS As Worksheet
    On Error Resume Next
    Set mapWS = ThisWorkbook.Sheets(mapSheetName)
    On Error GoTo ErrHandler

    If mapWS Is Nothing Then
        MsgBox "Sheet '" & mapSheetName & "' not found. Please check the name and try again.", _
               vbExclamation, "UTL Data Cleaning"
        Exit Sub
    End If

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select the range to apply replacements to.", vbExclamation, "UTL Data Cleaning"
        Exit Sub
    End If

    Dim targetRng As Range
    Set targetRng = Selection

    UTL_TurboOn

    Dim lastRow As Long
    lastRow = mapWS.Cells(mapWS.Rows.Count, 1).End(xlUp).Row

    Dim replaced As Long
    Dim i As Long
    For i = 1 To lastRow
        Dim findVal As String
        Dim replaceVal As String
        findVal = CStr(mapWS.Cells(i, 1).Value)
        replaceVal = CStr(mapWS.Cells(i, 2).Value)
        If findVal <> "" Then
            Dim c As Range
            For Each c In targetRng
                If Not IsEmpty(c) And VarType(c.Value) = vbString Then
                    If InStr(1, c.Value, findVal, vbTextCompare) > 0 Then
                        c.Value = Replace(c.Value, findVal, replaceVal, 1, -1, vbTextCompare)
                        replaced = replaced + 1
                    End If
                End If
            Next c
        End If
    Next i

    UTL_TurboOff
    MsgBox "Done! " & replaced & " replacements made using " & (lastRow) & " mapping rules.", _
           vbInformation, "UTL Data Cleaning"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Data Cleaning"
End Sub

' ============================================================
' TOOL 10 — Formula to Value Hardcoder               [TIER 2]
' Converts all formulas in selection to static values
' Useful before sharing a file or archiving a period
' ============================================================
Sub FormulaToValueHardcoder()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first, then run this tool.", vbExclamation, "UTL Data Cleaning"
        Exit Sub
    End If

    If MsgBox("Convert all FORMULAS in selection to static VALUES?" & Chr(10) & _
              "This cannot be undone after saving.", _
              vbExclamation + vbYesNo, "UTL Data Cleaning") = vbNo Then Exit Sub

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim count As Long
    Dim c As Range
    For Each c In Selection
        If c.HasFormula Then
            c.Value = c.Value
            count = count + 1
        End If
    Next c

    UTL_TurboOff
    MsgBox "Done! " & count & " formulas converted to static values.", vbInformation, "UTL Data Cleaning"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Data Cleaning"
End Sub

' ============================================================
' TOOL 11 — Phantom Hyperlink Purger                 [TIER 2]
' Removes all embedded hyperlinks from the active sheet
' Great for cleaning files before distribution
' ============================================================
Sub PhantomHyperlinkPurger()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim linkCount As Long
    linkCount = ws.Hyperlinks.Count

    If linkCount = 0 Then
        MsgBox "No hyperlinks found on this sheet.", vbInformation, "UTL Data Cleaning"
        Exit Sub
    End If

    If MsgBox("Remove all " & linkCount & " hyperlinks from sheet '" & ws.Name & "'?", _
              vbQuestion + vbYesNo, "UTL Data Cleaning") = vbNo Then Exit Sub

    On Error GoTo ErrHandler
    ws.Hyperlinks.Delete
    MsgBox "Done! " & linkCount & " hyperlinks removed.", vbInformation, "UTL Data Cleaning"
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Data Cleaning"
End Sub

' ============================================================
' TOOL 12 — Convert Numbers to Words                 [TIER 2]
' Translates numeric values to written text for formal documents
' Example: 1250.00 → "One Thousand Two Hundred Fifty Dollars"
' Run: select the cells containing numbers, then run this macro
' ============================================================
Sub ConvertNumbersToWords()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select the cells containing numbers.", vbExclamation, "UTL Data Cleaning"
        Exit Sub
    End If

    Dim writeAdj As String
    writeAdj = InputBox("Write converted text to which column offset?" & Chr(10) & _
                        "0 = overwrite, 1 = next column to the right, etc.", _
                        "UTL Data Cleaning", "1")
    If writeAdj = "" Then Exit Sub
    If Not IsNumeric(writeAdj) Then
        MsgBox "Please enter a number.", vbExclamation, "UTL Data Cleaning"
        Exit Sub
    End If

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim count As Long
    Dim c As Range
    For Each c In Selection
        If IsNumeric(c.Value) And Not IsEmpty(c) Then
            Dim targetCell As Range
            Set targetCell = c.Offset(0, CLng(writeAdj))
            targetCell.Value = UTL_NumberToWords(CDbl(c.Value))
            count = count + 1
        End If
    Next c

    UTL_TurboOff
    MsgBox "Done! " & count & " numbers converted to words.", vbInformation, "UTL Data Cleaning"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Data Cleaning"
End Sub

Private Function UTL_NumberToWords(ByVal num As Double) As String
    Dim ones()  As String
    Dim tens()  As String
    ones  = Array("", "One", "Two", "Three", "Four", "Five", "Six", "Seven", _
                  "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", _
                  "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen")
    tens  = Array("", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety")

    If num = 0 Then
        UTL_NumberToWords = "Zero Dollars"
        Exit Function
    End If

    Dim negative As Boolean
    negative = (num < 0)
    num = Abs(num)

    Dim dollars As Long
    Dim cents As Long
    dollars = CLng(Int(num))
    cents = CLng(Round((num - Int(num)) * 100))

    Dim result As String
    result = UTL_ConvertGroup(dollars, ones, tens)

    If dollars = 1 Then
        result = result & " Dollar"
    Else
        result = result & " Dollars"
    End If

    If cents > 0 Then
        result = result & " and " & UTL_ConvertGroup(cents, ones, tens)
        If cents = 1 Then
            result = result & " Cent"
        Else
            result = result & " Cents"
        End If
    End If

    If negative Then result = "Negative " & result
    UTL_NumberToWords = result
End Function

Private Function UTL_ConvertGroup(ByVal n As Long, ones() As String, tens() As String) As String
    Dim result As String
    If n >= 1000000 Then
        result = UTL_ConvertGroup(n \ 1000000, ones, tens) & " Million "
        n = n Mod 1000000
    End If
    If n >= 1000 Then
        result = result & UTL_ConvertGroup(n \ 1000, ones, tens) & " Thousand "
        n = n Mod 1000
    End If
    If n >= 100 Then
        result = result & ones(n \ 100) & " Hundred "
        n = n Mod 100
    End If
    If n >= 20 Then
        result = result & tens(n \ 10)
        If n Mod 10 > 0 Then result = result & "-" & ones(n Mod 10)
    ElseIf n > 0 Then
        result = result & ones(n)
    End If
    UTL_ConvertGroup = Trim(result)
End Function
