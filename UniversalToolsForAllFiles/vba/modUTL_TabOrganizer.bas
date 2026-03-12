Attribute VB_Name = "modUTL_TabOrganizer"
'==============================================================================
' modUTL_TabOrganizer — Sheet Tab Organization Tools
'==============================================================================
' PURPOSE:  Color-code tabs, bulk hide/unhide, reorder, and bulk rename.
'           Goes beyond basic sort/rename with color-coding and multi-select.
'
' PUBLIC SUBS:
'   ColorTabsByKeyword    — Color tabs matching a keyword
'   ColorTabsInteractive  — Pick sheets and assign a color
'   BulkHideSheets        — Hide multiple sheets at once
'   BulkUnhideSheets      — Unhide multiple sheets at once
'   ReorderTabs           — Move sheets to front or back
'   BulkRenameTabs        — Find/replace text in tab names
'
' DEPENDENCIES: None (standalone). Works in any Excel workbook.
' VERSION:  1.0.0 | DATE: 2026-03-12
'==============================================================================
Option Explicit

'==============================================================================
' PUBLIC: ColorTabsByKeyword
' Type a keyword — all matching tab names get colored.
'==============================================================================
Public Sub ColorTabsByKeyword()
    On Error GoTo ErrHandler

    Dim keyword As String
    keyword = InputBox("Enter a keyword to match sheet tab names:" & vbCrLf & vbCrLf & _
                        "All tabs containing this text will be colored." & vbCrLf & _
                        "Examples: Q1, 2025, Jan, Summary", _
                        "Color Tabs by Keyword - Step 1 of 2")
    If Len(Trim(keyword)) = 0 Then Exit Sub
    keyword = Trim(keyword)

    '--- Ask for color ---
    Dim colorChoice As String
    colorChoice = InputBox("Choose a tab color:" & vbCrLf & vbCrLf & _
                            "  1. Blue" & vbCrLf & _
                            "  2. Green" & vbCrLf & _
                            "  3. Red" & vbCrLf & _
                            "  4. Orange" & vbCrLf & _
                            "  5. Purple" & vbCrLf & _
                            "  6. Yellow" & vbCrLf & _
                            "  7. No color (remove color)" & vbCrLf & vbCrLf & _
                            "Enter number:", _
                            "Color Tabs by Keyword - Step 2 of 2")
    If Len(Trim(colorChoice)) = 0 Then Exit Sub

    Dim tabColor As Long
    Dim removeColor As Boolean
    removeColor = False

    Select Case Trim(colorChoice)
        Case "1": tabColor = RGB(68, 114, 196)     ' Blue
        Case "2": tabColor = RGB(112, 173, 71)     ' Green
        Case "3": tabColor = RGB(255, 0, 0)        ' Red
        Case "4": tabColor = RGB(237, 125, 49)     ' Orange
        Case "5": tabColor = RGB(112, 48, 160)     ' Purple
        Case "6": tabColor = RGB(255, 217, 102)    ' Yellow
        Case "7": removeColor = True
        Case Else
            MsgBox "Invalid choice.", vbExclamation, "Color Tabs"
            Exit Sub
    End Select

    '--- Apply ---
    Dim count As Long
    count = 0

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, keyword, vbTextCompare) > 0 Then
            If removeColor Then
                ws.Tab.ColorIndex = xlNone
            Else
                ws.Tab.Color = tabColor
            End If
            count = count + 1
        End If
    Next ws

    MsgBox count & " tab(s) matching '" & keyword & "' were colored.", _
           vbInformation, "Color Tabs by Keyword"

    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Color Tabs"
End Sub

'==============================================================================
' PUBLIC: ColorTabsInteractive
' User picks specific sheets from a list and assigns a color.
'==============================================================================
Public Sub ColorTabsInteractive()
    On Error GoTo ErrHandler

    '--- Build list ---
    Dim sheetList As String
    sheetList = "Select sheets to color:" & vbCrLf & String(40, "-") & vbCrLf

    Dim i As Long
    For i = 1 To ThisWorkbook.Sheets.Count
        sheetList = sheetList & "  " & i & ". " & ThisWorkbook.Sheets(i).Name & vbCrLf
    Next i

    sheetList = sheetList & vbCrLf & "Enter sheet numbers (comma-separated):" & vbCrLf & _
                "Example: 1,3,5,6"

    Dim choice As String
    choice = InputBox(sheetList, "Color Tabs - Step 1 of 2")
    If Len(Trim(choice)) = 0 Then Exit Sub

    '--- Ask for color ---
    Dim colorChoice As String
    colorChoice = InputBox("Choose a tab color:" & vbCrLf & vbCrLf & _
                            "  1. Blue     2. Green    3. Red" & vbCrLf & _
                            "  4. Orange   5. Purple   6. Yellow" & vbCrLf & _
                            "  7. No color (remove)" & vbCrLf & vbCrLf & _
                            "Enter number:", _
                            "Color Tabs - Step 2 of 2")
    If Len(Trim(colorChoice)) = 0 Then Exit Sub

    Dim tabColor As Long
    Dim removeColor As Boolean
    removeColor = False

    Select Case Trim(colorChoice)
        Case "1": tabColor = RGB(68, 114, 196)
        Case "2": tabColor = RGB(112, 173, 71)
        Case "3": tabColor = RGB(255, 0, 0)
        Case "4": tabColor = RGB(237, 125, 49)
        Case "5": tabColor = RGB(112, 48, 160)
        Case "6": tabColor = RGB(255, 217, 102)
        Case "7": removeColor = True
        Case Else
            MsgBox "Invalid color.", vbExclamation, "Color Tabs"
            Exit Sub
    End Select

    '--- Apply ---
    Dim parts() As String
    parts = Split(choice, ",")

    Dim count As Long
    count = 0
    Dim p As Long
    For p = LBound(parts) To UBound(parts)
        Dim num As String
        num = Trim(parts(p))
        If IsNumeric(num) Then
            Dim idx As Long
            idx = CLng(num)
            If idx >= 1 And idx <= ThisWorkbook.Sheets.Count Then
                If removeColor Then
                    ThisWorkbook.Sheets(idx).Tab.ColorIndex = xlNone
                Else
                    ThisWorkbook.Sheets(idx).Tab.Color = tabColor
                End If
                count = count + 1
            End If
        End If
    Next p

    MsgBox count & " tab(s) colored.", vbInformation, "Color Tabs"

    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Color Tabs"
End Sub

'==============================================================================
' PUBLIC: BulkHideSheets
' User picks multiple sheets to hide from a numbered list.
'==============================================================================
Public Sub BulkHideSheets()
    On Error GoTo ErrHandler

    '--- Build list of visible sheets ---
    Dim sheetList As String
    sheetList = "Visible sheets:" & vbCrLf & String(40, "-") & vbCrLf

    Dim visibleCount As Long
    visibleCount = 0
    Dim i As Long
    For i = 1 To ThisWorkbook.Sheets.Count
        If ThisWorkbook.Sheets(i).Visible = xlSheetVisible Then
            visibleCount = visibleCount + 1
            sheetList = sheetList & "  " & i & ". " & ThisWorkbook.Sheets(i).Name & vbCrLf
        End If
    Next i

    If visibleCount <= 1 Then
        MsgBox "You must keep at least one sheet visible.", vbExclamation, "Bulk Hide Sheets"
        Exit Sub
    End If

    sheetList = sheetList & vbCrLf & "Enter sheet numbers to HIDE (comma-separated):" & vbCrLf & _
                "Example: 3,5,6" & vbCrLf & vbCrLf & _
                "NOTE: At least 1 sheet must remain visible."

    Dim choice As String
    choice = InputBox(sheetList, "Bulk Hide Sheets")
    If Len(Trim(choice)) = 0 Then Exit Sub

    '--- Parse and count ---
    Dim parts() As String
    parts = Split(choice, ",")

    Dim hideCount As Long
    hideCount = 0

    ' Count how many we'd hide
    Dim p As Long
    For p = LBound(parts) To UBound(parts)
        If IsNumeric(Trim(parts(p))) Then hideCount = hideCount + 1
    Next p

    If hideCount >= visibleCount Then
        MsgBox "Cannot hide all sheets. At least 1 must remain visible.", vbExclamation, "Bulk Hide Sheets"
        Exit Sub
    End If

    '--- Hide ---
    Dim hidden As Long
    hidden = 0
    For p = LBound(parts) To UBound(parts)
        Dim num As String
        num = Trim(parts(p))
        If IsNumeric(num) Then
            Dim idx As Long
            idx = CLng(num)
            If idx >= 1 And idx <= ThisWorkbook.Sheets.Count Then
                If ThisWorkbook.Sheets(idx).Visible = xlSheetVisible Then
                    ThisWorkbook.Sheets(idx).Visible = xlSheetHidden
                    hidden = hidden + 1
                End If
            End If
        End If
    Next p

    MsgBox hidden & " sheet(s) hidden.", vbInformation, "Bulk Hide Sheets"

    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Bulk Hide Sheets"
End Sub

'==============================================================================
' PUBLIC: BulkUnhideSheets
' Shows all hidden sheets and lets user pick which to unhide.
'==============================================================================
Public Sub BulkUnhideSheets()
    On Error GoTo ErrHandler

    '--- Build list of hidden sheets ---
    Dim sheetList As String
    sheetList = "Hidden sheets:" & vbCrLf & String(40, "-") & vbCrLf

    Dim hiddenNames() As String
    Dim hiddenCount As Long
    hiddenCount = 0
    ReDim hiddenNames(1 To ThisWorkbook.Sheets.Count)

    Dim i As Long
    For i = 1 To ThisWorkbook.Sheets.Count
        If ThisWorkbook.Sheets(i).Visible <> xlSheetVisible Then
            hiddenCount = hiddenCount + 1
            hiddenNames(hiddenCount) = ThisWorkbook.Sheets(i).Name

            Dim visType As String
            If ThisWorkbook.Sheets(i).Visible = xlSheetHidden Then
                visType = "Hidden"
            Else
                visType = "Very Hidden"
            End If
            sheetList = sheetList & "  " & hiddenCount & ". " & ThisWorkbook.Sheets(i).Name & _
                       " (" & visType & ")" & vbCrLf
        End If
    Next i

    If hiddenCount = 0 Then
        MsgBox "No hidden sheets found.", vbInformation, "Bulk Unhide Sheets"
        Exit Sub
    End If

    sheetList = sheetList & vbCrLf & "Enter numbers to UNHIDE (comma-separated):" & vbCrLf & _
                "Example: 1,2,3  or  ALL to unhide all"

    Dim choice As String
    choice = InputBox(sheetList, "Bulk Unhide Sheets")
    If Len(Trim(choice)) = 0 Then Exit Sub

    Dim unhidden As Long
    unhidden = 0

    If UCase(Trim(choice)) = "ALL" Then
        For i = 1 To hiddenCount
            ThisWorkbook.Sheets(hiddenNames(i)).Visible = xlSheetVisible
            unhidden = unhidden + 1
        Next i
    Else
        Dim parts() As String
        parts = Split(choice, ",")
        Dim p As Long
        For p = LBound(parts) To UBound(parts)
            Dim num As String
            num = Trim(parts(p))
            If IsNumeric(num) Then
                Dim idx As Long
                idx = CLng(num)
                If idx >= 1 And idx <= hiddenCount Then
                    ThisWorkbook.Sheets(hiddenNames(idx)).Visible = xlSheetVisible
                    unhidden = unhidden + 1
                End If
            End If
        Next p
    End If

    MsgBox unhidden & " sheet(s) unhidden.", vbInformation, "Bulk Unhide Sheets"

    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Bulk Unhide Sheets"
End Sub

'==============================================================================
' PUBLIC: ReorderTabs
' Move selected sheets to the front or back of the workbook.
'==============================================================================
Public Sub ReorderTabs()
    On Error GoTo ErrHandler

    '--- Build list ---
    Dim sheetList As String
    sheetList = "Current sheet order:" & vbCrLf & String(40, "-") & vbCrLf

    Dim i As Long
    For i = 1 To ThisWorkbook.Sheets.Count
        sheetList = sheetList & "  " & i & ". " & ThisWorkbook.Sheets(i).Name & vbCrLf
    Next i

    sheetList = sheetList & vbCrLf & "Enter sheet numbers to move (comma-separated):" & vbCrLf & _
                "Example: 5,3,7"

    Dim choice As String
    choice = InputBox(sheetList, "Reorder Tabs - Step 1 of 2")
    If Len(Trim(choice)) = 0 Then Exit Sub

    '--- Ask where to move ---
    Dim dirChoice As String
    dirChoice = InputBox("Move the selected sheets to:" & vbCrLf & vbCrLf & _
                          "  1. FRONT of workbook (position 1)" & vbCrLf & _
                          "  2. BACK of workbook (last position)" & vbCrLf & _
                          "  3. AFTER a specific sheet (you choose)" & vbCrLf & vbCrLf & _
                          "Enter number:", _
                          "Reorder Tabs - Step 2 of 2")
    If Len(Trim(dirChoice)) = 0 Then Exit Sub

    '--- Parse selections and resolve to sheet NAMES (indices shift during moves) ---
    Dim parts() As String
    parts = Split(choice, ",")

    Dim sheetNames() As String
    ReDim sheetNames(LBound(parts) To UBound(parts))
    Dim validCount As Long
    validCount = 0

    Dim p As Long
    For p = LBound(parts) To UBound(parts)
        Dim numStr As String
        numStr = Trim(parts(p))
        If IsNumeric(numStr) Then
            Dim idx As Long
            idx = CLng(numStr)
            If idx >= 1 And idx <= ThisWorkbook.Sheets.Count Then
                sheetNames(p) = ThisWorkbook.Sheets(idx).Name
                validCount = validCount + 1
            End If
        End If
    Next p

    If validCount = 0 Then
        MsgBox "No valid sheet numbers entered.", vbExclamation, "Reorder Tabs"
        Exit Sub
    End If

    Dim moved As Long
    moved = 0

    Select Case Trim(dirChoice)
        Case "1"  ' Move to front
            Dim p1 As Long
            For p1 = UBound(sheetNames) To LBound(sheetNames) Step -1
                If Len(sheetNames(p1)) > 0 Then
                    ThisWorkbook.Sheets(sheetNames(p1)).Move Before:=ThisWorkbook.Sheets(1)
                    moved = moved + 1
                End If
            Next p1

        Case "2"  ' Move to back
            Dim p2 As Long
            For p2 = LBound(sheetNames) To UBound(sheetNames)
                If Len(sheetNames(p2)) > 0 Then
                    ThisWorkbook.Sheets(sheetNames(p2)).Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                    moved = moved + 1
                End If
            Next p2

        Case "3"  ' After specific sheet
            Dim afterStr As String
            afterStr = InputBox("Move after which sheet number?", "Move After Sheet")
            If Not IsNumeric(afterStr) Then Exit Sub
            Dim afterIdx As Long
            afterIdx = CLng(afterStr)
            If afterIdx < 1 Or afterIdx > ThisWorkbook.Sheets.Count Then Exit Sub
            Dim afterName As String
            afterName = ThisWorkbook.Sheets(afterIdx).Name

            Dim p3 As Long
            For p3 = LBound(sheetNames) To UBound(sheetNames)
                If Len(sheetNames(p3)) > 0 And sheetNames(p3) <> afterName Then
                    ThisWorkbook.Sheets(sheetNames(p3)).Move After:=ThisWorkbook.Sheets(afterName)
                    moved = moved + 1
                End If
            Next p3

        Case Else
            MsgBox "Invalid choice.", vbExclamation, "Reorder Tabs"
            Exit Sub
    End Select

    MsgBox moved & " sheet(s) moved.", vbInformation, "Reorder Tabs"

    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Reorder Tabs"
End Sub

'==============================================================================
' PUBLIC: BulkRenameTabs
' Find and replace text in sheet tab names.
'==============================================================================
Public Sub BulkRenameTabs()
    On Error GoTo ErrHandler

    Dim findText As String
    findText = InputBox("Enter the text to FIND in tab names:" & vbCrLf & vbCrLf & _
                         "Example: 2025, Q1, Draft", _
                         "Bulk Rename Tabs - Step 1 of 2")
    If Len(findText) = 0 Then Exit Sub

    Dim replaceText As String
    replaceText = InputBox("Replace '" & findText & "' with:" & vbCrLf & vbCrLf & _
                            "Example: 2026, Q2, Final" & vbCrLf & _
                            "(Leave empty to remove the text)", _
                            "Bulk Rename Tabs - Step 2 of 2")
    ' Empty replacement is OK (removes text)

    '--- Preview changes ---
    Dim previewMsg As String
    previewMsg = "Preview of tab name changes:" & vbCrLf & String(40, "-") & vbCrLf & vbCrLf

    Dim matchCount As Long
    matchCount = 0

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If InStr(1, ws.Name, findText, vbTextCompare) > 0 Then
            Dim newName As String
            newName = Replace(ws.Name, findText, replaceText, , , vbTextCompare)
            ' Ensure valid name (max 31 chars, no special chars)
            If Len(newName) > 31 Then newName = Left(newName, 31)
            If Len(newName) = 0 Then newName = "Sheet_" & ws.Index

            previewMsg = previewMsg & "  " & ws.Name & "  -->  " & newName & vbCrLf
            matchCount = matchCount + 1
        End If
    Next ws

    If matchCount = 0 Then
        MsgBox "No tab names contain '" & findText & "'.", vbInformation, "Bulk Rename Tabs"
        Exit Sub
    End If

    previewMsg = previewMsg & vbCrLf & "Apply these changes?"

    Dim confirm As VbMsgBoxResult
    confirm = MsgBox(previewMsg, vbYesNo + vbQuestion, "Bulk Rename Tabs")
    If confirm = vbNo Then Exit Sub

    '--- Apply ---
    Dim renamed As Long
    renamed = 0

    For Each ws In ThisWorkbook.Sheets
        If InStr(1, ws.Name, findText, vbTextCompare) > 0 Then
            newName = Replace(ws.Name, findText, replaceText, , , vbTextCompare)
            If Len(newName) > 31 Then newName = Left(newName, 31)
            If Len(newName) = 0 Then newName = "Sheet_" & ws.Index

            On Error Resume Next
            ws.Name = newName
            If Err.Number = 0 Then
                renamed = renamed + 1
            Else
                Err.Clear
            End If
            On Error GoTo ErrHandler
        End If
    Next ws

    MsgBox renamed & " tab(s) renamed.", vbInformation, "Bulk Rename Tabs"

    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Bulk Rename Tabs"
End Sub
