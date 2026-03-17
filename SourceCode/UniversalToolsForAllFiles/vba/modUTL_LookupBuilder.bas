Attribute VB_Name = "modUTL_LookupBuilder"
'==============================================================================
' modUTL_LookupBuilder — Auto VLOOKUP & INDEX-MATCH Builder
'==============================================================================
' PURPOSE:  Builds lookup formulas automatically. User selects source data,
'           key columns, and value columns — the tool writes the formulas.
'           No formula knowledge required.
'
' PUBLIC SUBS:
'   BuildVLOOKUP       — Build VLOOKUP formulas step by step
'   BuildINDEXMATCH    — Build INDEX-MATCH formulas (more flexible)
'   MatchAndPull       — Match two lists and pull values across
'   FindUnmatched      — Find values in list A that don't exist in list B
'
' DEPENDENCIES: None (standalone). Works in any Excel workbook.
' VERSION:  1.0.0 | DATE: 2026-03-12
'==============================================================================
Option Explicit

'==============================================================================
' PUBLIC: BuildVLOOKUP
' Step-by-step VLOOKUP builder. User selects everything visually.
'==============================================================================
Public Sub BuildVLOOKUP()
    On Error GoTo ErrHandler

    MsgBox "This tool builds VLOOKUP formulas for you." & vbCrLf & vbCrLf & _
           "You will:" & vbCrLf & _
           "  1. Select where your lookup keys are (what to look up)" & vbCrLf & _
           "  2. Select the source table (where to find the answer)" & vbCrLf & _
           "  3. Pick which column has the value you want" & vbCrLf & _
           "  4. The formula will be written for you", _
           vbInformation, "Build VLOOKUP"

    '--- Step 1: Lookup keys ---
    Dim keyRng As Range
    On Error Resume Next
    Set keyRng = Application.InputBox("Step 1: Select the cells containing your LOOKUP KEYS" & vbCrLf & _
                                       "(the values you want to look up):", _
                                       "Build VLOOKUP - Step 1 of 4", Type:=8)
    On Error GoTo ErrHandler
    If keyRng Is Nothing Then Exit Sub

    If keyRng.Columns.Count > 1 Then
        MsgBox "Please select a single column of lookup keys.", vbExclamation, "Build VLOOKUP"
        Exit Sub
    End If

    '--- Step 2: Source table ---
    MsgBox "Now select the SOURCE TABLE (the data to search in)." & vbCrLf & vbCrLf & _
           "IMPORTANT: The lookup key column must be the FIRST column" & vbCrLf & _
           "of the range you select (this is how VLOOKUP works).", _
           vbInformation, "Build VLOOKUP"

    Dim tableRng As Range
    On Error Resume Next
    Set tableRng = Application.InputBox("Step 2: Select the SOURCE TABLE range:" & vbCrLf & _
                                         "(first column must contain the matching keys)", _
                                         "Build VLOOKUP - Step 2 of 4", Type:=8)
    On Error GoTo ErrHandler
    If tableRng Is Nothing Then Exit Sub

    If tableRng.Columns.Count < 2 Then
        MsgBox "Source table must have at least 2 columns.", vbExclamation, "Build VLOOKUP"
        Exit Sub
    End If

    '--- Step 3: Which column to return ---
    Dim colList As String
    colList = "Which column contains the VALUE you want to pull?" & vbCrLf & vbCrLf

    ' Show column headers from first row
    Dim c As Long
    For c = 1 To tableRng.Columns.Count
        Dim hdrVal As String
        hdrVal = CStr(tableRng.Cells(1, c).Value)
        If Len(hdrVal) = 0 Then hdrVal = "(empty)"
        If Len(hdrVal) > 30 Then hdrVal = Left(hdrVal, 30) & "..."
        colList = colList & "  " & c & ". " & hdrVal
        If c = 1 Then colList = colList & " <-- key column"
        colList = colList & vbCrLf
    Next c

    colList = colList & vbCrLf & "Enter column number:"

    Dim colChoice As String
    colChoice = InputBox(colList, "Build VLOOKUP - Step 3 of 4")
    If Len(Trim(colChoice)) = 0 Then Exit Sub
    If Not IsNumeric(colChoice) Then
        MsgBox "Please enter a number.", vbExclamation, "Build VLOOKUP"
        Exit Sub
    End If
    Dim colIndex As Long
    colIndex = CLng(colChoice)
    If colIndex < 2 Or colIndex > tableRng.Columns.Count Then
        MsgBox "Column must be between 2 and " & tableRng.Columns.Count & ".", vbExclamation, "Build VLOOKUP"
        Exit Sub
    End If

    '--- Step 4: Where to put results ---
    Dim resultRng As Range
    On Error Resume Next
    Set resultRng = Application.InputBox("Step 4: Select the cell where the FIRST result should go:" & vbCrLf & _
                                          "(formulas will fill down for all lookup keys)", _
                                          "Build VLOOKUP - Step 4 of 4", Type:=8)
    On Error GoTo ErrHandler
    If resultRng Is Nothing Then Exit Sub

    '--- Build formulas ---
    Application.ScreenUpdating = False

    Dim tableAddr As String
    tableAddr = FormatSheetRef(tableRng.Parent.Name) & "!" & tableRng.Address(True, True)

    Dim r As Long
    For r = 1 To keyRng.Rows.Count
        Dim keyAddr As String
        keyAddr = keyRng.Cells(r, 1).Address(False, False)

        ' Handle cross-sheet reference
        If keyRng.Parent.Name <> resultRng.Parent.Name Then
            keyAddr = FormatSheetRef(keyRng.Parent.Name) & "!" & keyAddr
        End If

        Dim formula As String
        formula = "=IFERROR(VLOOKUP(" & keyAddr & "," & tableAddr & "," & colIndex & ",FALSE),"""")"

        resultRng.Parent.Cells(resultRng.Row + r - 1, resultRng.Column).formula = formula
    Next r

    Application.ScreenUpdating = True

    MsgBox "VLOOKUP formulas created!" & vbCrLf & vbCrLf & _
           "Formulas written: " & keyRng.Rows.Count & vbCrLf & _
           "Starting at: " & resultRng.Address(False, False) & vbCrLf & vbCrLf & _
           "Wrapped in IFERROR — shows blank if no match found.", _
           vbInformation, "Build VLOOKUP"

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Build VLOOKUP"
End Sub

'==============================================================================
' PUBLIC: BuildINDEXMATCH
' More flexible than VLOOKUP — lookup column doesn't need to be first.
'==============================================================================
Public Sub BuildINDEXMATCH()
    On Error GoTo ErrHandler

    MsgBox "INDEX-MATCH is more flexible than VLOOKUP:" & vbCrLf & _
           "  - Lookup column can be anywhere (not just first)" & vbCrLf & _
           "  - Can look left (VLOOKUP can only look right)" & vbCrLf & vbCrLf & _
           "You will:" & vbCrLf & _
           "  1. Select your lookup keys" & vbCrLf & _
           "  2. Select the matching column in the source data" & vbCrLf & _
           "  3. Select the return column (the values you want)" & vbCrLf & _
           "  4. Pick where to put the results", _
           vbInformation, "Build INDEX-MATCH"

    '--- Step 1: Lookup keys ---
    Dim keyRng As Range
    On Error Resume Next
    Set keyRng = Application.InputBox("Step 1: Select your LOOKUP KEYS:", _
                                       "INDEX-MATCH - Step 1 of 4", Type:=8)
    On Error GoTo ErrHandler
    If keyRng Is Nothing Then Exit Sub

    '--- Step 2: Match column ---
    MsgBox "Now select the column in the source data that MATCHES your keys." & vbCrLf & _
           "This is the column where the tool will search for your values.", _
           vbInformation, "INDEX-MATCH"

    Dim matchRng As Range
    On Error Resume Next
    Set matchRng = Application.InputBox("Step 2: Select the MATCH column (where to search):", _
                                         "INDEX-MATCH - Step 2 of 4", Type:=8)
    On Error GoTo ErrHandler
    If matchRng Is Nothing Then Exit Sub

    '--- Step 3: Return column ---
    MsgBox "Now select the column containing the VALUES you want to pull.", _
           vbInformation, "INDEX-MATCH"

    Dim returnRng As Range
    On Error Resume Next
    Set returnRng = Application.InputBox("Step 3: Select the RETURN column (values to pull):", _
                                          "INDEX-MATCH - Step 3 of 4", Type:=8)
    On Error GoTo ErrHandler
    If returnRng Is Nothing Then Exit Sub

    If matchRng.Rows.Count <> returnRng.Rows.Count Then
        MsgBox "Match column and return column must have the same number of rows.", _
               vbExclamation, "INDEX-MATCH"
        Exit Sub
    End If

    '--- Step 4: Output ---
    Dim resultRng As Range
    On Error Resume Next
    Set resultRng = Application.InputBox("Step 4: Select where the FIRST result goes:", _
                                          "INDEX-MATCH - Step 4 of 4", Type:=8)
    On Error GoTo ErrHandler
    If resultRng Is Nothing Then Exit Sub

    '--- Build formulas ---
    Application.ScreenUpdating = False

    Dim matchAddr As String
    matchAddr = FormatRangeRef(matchRng, resultRng.Parent.Name)

    Dim returnAddr As String
    returnAddr = FormatRangeRef(returnRng, resultRng.Parent.Name)

    Dim r As Long
    For r = 1 To keyRng.Rows.Count
        Dim keyAddr As String
        keyAddr = keyRng.Cells(r, 1).Address(False, False)
        If keyRng.Parent.Name <> resultRng.Parent.Name Then
            keyAddr = FormatSheetRef(keyRng.Parent.Name) & "!" & keyAddr
        End If

        Dim formula As String
        formula = "=IFERROR(INDEX(" & returnAddr & ",MATCH(" & keyAddr & "," & matchAddr & ",0)),"""")"

        resultRng.Parent.Cells(resultRng.Row + r - 1, resultRng.Column).formula = formula
    Next r

    Application.ScreenUpdating = True

    MsgBox "INDEX-MATCH formulas created!" & vbCrLf & vbCrLf & _
           "Formulas written: " & keyRng.Rows.Count & vbCrLf & _
           "Starting at: " & resultRng.Address(False, False), _
           vbInformation, "Build INDEX-MATCH"

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Build INDEX-MATCH"
End Sub

'==============================================================================
' PUBLIC: MatchAndPull
' Compare two lists — pull matching data from one to the other.
'==============================================================================
Public Sub MatchAndPull()
    On Error GoTo ErrHandler

    MsgBox "This tool compares two lists and pulls data where keys match." & vbCrLf & vbCrLf & _
           "Example: You have a list of employee IDs on Sheet1," & vbCrLf & _
           "and a table with IDs + departments on Sheet2." & vbCrLf & _
           "This tool pulls the department onto Sheet1 where IDs match.", _
           vbInformation, "Match and Pull"

    '--- Your keys ---
    Dim myKeys As Range
    On Error Resume Next
    Set myKeys = Application.InputBox("Select YOUR key column (the IDs/names you have):", _
                                       "Match and Pull - Step 1 of 4", Type:=8)
    On Error GoTo ErrHandler
    If myKeys Is Nothing Then Exit Sub

    '--- Source keys ---
    Dim srcKeys As Range
    On Error Resume Next
    Set srcKeys = Application.InputBox("Select the SOURCE key column (matching IDs/names in source data):", _
                                        "Match and Pull - Step 2 of 4", Type:=8)
    On Error GoTo ErrHandler
    If srcKeys Is Nothing Then Exit Sub

    '--- Source values ---
    Dim srcVals As Range
    On Error Resume Next
    Set srcVals = Application.InputBox("Select the SOURCE value column (data to pull):", _
                                        "Match and Pull - Step 3 of 4", Type:=8)
    On Error GoTo ErrHandler
    If srcVals Is Nothing Then Exit Sub

    If srcKeys.Rows.Count <> srcVals.Rows.Count Then
        MsgBox "Source key and value columns must have the same number of rows.", _
               vbExclamation, "Match and Pull"
        Exit Sub
    End If

    '--- Output column ---
    Dim output As Range
    On Error Resume Next
    Set output = Application.InputBox("Select where to put the pulled values (single cell, top of output column):", _
                                       "Match and Pull - Step 4 of 4", Type:=8)
    On Error GoTo ErrHandler
    If output Is Nothing Then Exit Sub

    '--- Match and pull ---
    Application.ScreenUpdating = False
    Application.StatusBar = "Matching and pulling values..."

    '--- Build lookup index from source keys for O(1) matching ---
    Dim srcIndex As New Collection
    Dim s As Long
    For s = 1 To srcKeys.Rows.Count
        Dim srcKeyStr As String
        srcKeyStr = CStr(srcKeys.Cells(s, 1).Value)
        If Len(srcKeyStr) > 0 Then
            On Error Resume Next
            srcIndex.Add s, srcKeyStr   ' Key = source key, Value = row index
            Err.Clear
            On Error GoTo ErrHandler
        End If
    Next s

    Dim matched As Long, notFound As Long
    matched = 0
    notFound = 0

    Dim r As Long
    For r = 1 To myKeys.Rows.Count
        Dim myKey As Variant
        myKey = myKeys.Cells(r, 1).Value

        If Not IsEmpty(myKey) Then
            Dim srcRow As Long
            srcRow = 0
            On Error Resume Next
            srcRow = srcIndex(CStr(myKey))
            Err.Clear
            On Error GoTo ErrHandler

            If srcRow > 0 Then
                output.Parent.Cells(output.Row + r - 1, output.Column).Value = srcVals.Cells(srcRow, 1).Value
                matched = matched + 1
            Else
                output.Parent.Cells(output.Row + r - 1, output.Column).Value = ""
                notFound = notFound + 1
            End If
        End If

        If r Mod 500 = 0 Then
            Application.StatusBar = "Processing row " & r & " of " & myKeys.Rows.Count & "..."
        End If
    Next r

    Application.StatusBar = False
    Application.ScreenUpdating = True

    MsgBox "Match and Pull complete!" & vbCrLf & vbCrLf & _
           "Matched: " & matched & vbCrLf & _
           "Not found: " & notFound & vbCrLf & _
           "Total keys: " & myKeys.Rows.Count, _
           vbInformation, "Match and Pull"

    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Match and Pull"
End Sub

'==============================================================================
' PUBLIC: FindUnmatched
' Find values in list A that don't exist in list B.
'==============================================================================
Public Sub FindUnmatched()
    On Error GoTo ErrHandler

    MsgBox "This tool finds values in List A that do NOT exist in List B." & vbCrLf & vbCrLf & _
           "Example: Find customers in your list that aren't in the master list.", _
           vbInformation, "Find Unmatched"

    Dim listA As Range
    On Error Resume Next
    Set listA = Application.InputBox("Select LIST A (the list to check):", _
                                      "Find Unmatched - Step 1 of 2", Type:=8)
    On Error GoTo ErrHandler
    If listA Is Nothing Then Exit Sub

    Dim listB As Range
    On Error Resume Next
    Set listB = Application.InputBox("Select LIST B (the reference list to check against):", _
                                      "Find Unmatched - Step 2 of 2", Type:=8)
    On Error GoTo ErrHandler
    If listB Is Nothing Then Exit Sub

    '--- Ask what to do with unmatched ---
    Dim actionChoice As String
    actionChoice = InputBox("What should I do with unmatched values?" & vbCrLf & vbCrLf & _
                             "  1. Highlight them in yellow (on List A)" & vbCrLf & _
                             "  2. List them on a new sheet" & vbCrLf & _
                             "  3. Both" & vbCrLf & vbCrLf & _
                             "Enter number:", _
                             "Find Unmatched")
    If Len(Trim(actionChoice)) = 0 Then Exit Sub

    Application.ScreenUpdating = False
    Application.StatusBar = "Checking for unmatched values..."

    '--- Build lookup index from list B for O(1) matching ---
    Dim bIndex As New Collection
    Dim cell As Range
    For Each cell In listB.Cells
        If Not IsEmpty(cell.Value) Then
            Dim bKey As String
            bKey = LCase(CStr(cell.Value))
            On Error Resume Next
            bIndex.Add 1, bKey
            Err.Clear
            On Error GoTo ErrHandler
        End If
    Next cell

    '--- Check list A against list B ---
    Dim unmatched() As String
    Dim unmatchedCells() As String
    Dim unmatchedCount As Long
    unmatchedCount = 0
    ReDim unmatched(1 To listA.Rows.Count)
    ReDim unmatchedCells(1 To listA.Rows.Count)

    For Each cell In listA.Cells
        If Not IsEmpty(cell.Value) Then
            Dim aVal As String
            aVal = LCase(CStr(cell.Value))

            Dim foundInB As Long
            foundInB = 0
            On Error Resume Next
            foundInB = bIndex(aVal)
            Err.Clear
            On Error GoTo ErrHandler

            If foundInB = 0 Then
                unmatchedCount = unmatchedCount + 1
                If unmatchedCount > UBound(unmatched) Then
                    ReDim Preserve unmatched(1 To unmatchedCount + 100)
                    ReDim Preserve unmatchedCells(1 To unmatchedCount + 100)
                End If
                unmatched(unmatchedCount) = CStr(cell.Value)
                unmatchedCells(unmatchedCount) = cell.Address(False, False)

                ' Highlight if requested
                If Trim(actionChoice) = "1" Or Trim(actionChoice) = "3" Then
                    cell.Interior.Color = RGB(255, 255, 153)  ' Yellow
                End If
            End If
        End If
    Next cell

    '--- Create report sheet if requested ---
    If Trim(actionChoice) = "2" Or Trim(actionChoice) = "3" Then
        If unmatchedCount > 0 Then
            Dim wsOut As Worksheet
            Dim reportName As String
            reportName = "UTL_Unmatched"

            On Error Resume Next
            Set wsOut = ThisWorkbook.Sheets(reportName)
            On Error GoTo ErrHandler

            If Not wsOut Is Nothing Then
                wsOut.Cells.Clear
            Else
                Set wsOut = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                wsOut.Name = reportName
            End If

            wsOut.Range("A1").Value = "Unmatched Values"
            wsOut.Range("A1").Font.Bold = True
            wsOut.Range("A1").Font.Size = 14
            wsOut.Range("A2").Value = "Found " & unmatchedCount & " values in List A not in List B"

            wsOut.Cells(4, 1).Value = "#"
            wsOut.Cells(4, 2).Value = "Value"
            wsOut.Cells(4, 3).Value = "Cell Address"
            wsOut.Range("A4:C4").Font.Bold = True
            wsOut.Range("A4:C4").Font.Color = RGB(255, 255, 255)
            wsOut.Range("A4:C4").Interior.Color = RGB(11, 71, 121)

            Dim u As Long
            For u = 1 To unmatchedCount
                wsOut.Cells(4 + u, 1).Value = u
                wsOut.Cells(4 + u, 2).Value = unmatched(u)
                wsOut.Cells(4 + u, 3).Value = unmatchedCells(u)
            Next u

            wsOut.Columns("A:C").AutoFit
            wsOut.Activate
        End If
    End If

    Application.StatusBar = False
    Application.ScreenUpdating = True

    MsgBox "Unmatched check complete!" & vbCrLf & vbCrLf & _
           "List A items: " & listA.Cells.Count & vbCrLf & _
           "List B items: " & listB.Cells.Count & vbCrLf & _
           "Unmatched (in A but not B): " & unmatchedCount, _
           vbInformation, "Find Unmatched"

    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Find Unmatched"
End Sub

'==============================================================================
' PRIVATE Helpers
'==============================================================================
Private Function FormatSheetRef(ByVal sheetName As String) As String
    If InStr(sheetName, " ") > 0 Or InStr(sheetName, "'") > 0 Then
        FormatSheetRef = "'" & Replace(sheetName, "'", "''") & "'"
    Else
        FormatSheetRef = sheetName
    End If
End Function

Private Function FormatRangeRef(ByVal rng As Range, ByVal currentSheet As String) As String
    Dim addr As String
    addr = rng.Address(True, True)

    If rng.Parent.Name <> currentSheet Then
        addr = FormatSheetRef(rng.Parent.Name) & "!" & addr
    End If

    FormatRangeRef = addr
End Function
