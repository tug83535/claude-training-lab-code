Attribute VB_Name = "modUtilities"
Option Explicit

'===============================================================================
' modUtilities - Sheet & Workbook Utility Macros
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  12 quick-win utility macros for everyday sheet and workbook
'           maintenance. All actions are wired into the Command Center
'           (actions 51-62) and the InputBox fallback menu (page 4).
'
' PUBLIC SUBS:
'   DeleteBlankRows           - Delete completely blank rows in selection
'   UnhideAllSheets           - Make every worksheet visible
'   SortSheetsAlphabetically  - Reorder all tabs A-Z by name
'   ToggleFreezePanes         - Toggle freeze panes on/off (freezes at B2)
'   ConvertToValues           - Replace formulas in selection with values
'   AutoFitAllColumns         - AutoFit every column on the active sheet
'   ProtectAllSheets          - Password-protect every worksheet
'   UnprotectAllSheets        - Remove password protection from every worksheet
'   FindReplaceAllSheets      - Find & replace across every worksheet
'   HighlightHardcodedNumbers - Flag non-formula numbers in blue font
'   TogglePresentationMode    - Hide/show gridlines, headings, formula bar
'   UnmergeAndFillDown        - Unmerge selection and fill blanks from above
'
' VERSION:  2.1.0
' DATE:     2026-02-27
' AUTHOR:   iPipeline Finance & Accounting Demo Project
'===============================================================================


'===============================================================================
' DeleteBlankRows - Delete all completely blank rows in the current selection
'===============================================================================
Public Sub DeleteBlankRows()
    On Error GoTo ErrHandler

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim rng As Range
    Set rng = Selection

    Dim deletedCount As Long: deletedCount = 0
    Dim i As Long

    modPerformance.TurboOn

    ' Loop BACKWARDS — prevents row-shift errors during deletion
    For i = rng.Rows.Count To 1 Step -1
        If Application.WorksheetFunction.CountA(rng.Rows(i)) = 0 Then
            rng.Rows(i).EntireRow.Delete
            deletedCount = deletedCount + 1
        End If
    Next i

    modPerformance.TurboOff

    MsgBox deletedCount & " blank row(s) deleted.", vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    MsgBox "Error deleting blank rows: " & Err.Description, vbExclamation, APP_NAME
End Sub


'===============================================================================
' UnhideAllSheets - Make every worksheet in the workbook visible
'===============================================================================
Public Sub UnhideAllSheets()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim unhiddenCount As Long: unhiddenCount = 0

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
            unhiddenCount = unhiddenCount + 1
        End If
    Next ws

    If unhiddenCount = 0 Then
        MsgBox "All sheets are already visible.", vbInformation, APP_NAME
    Else
        MsgBox unhiddenCount & " sheet(s) unhidden.", vbInformation, APP_NAME
    End If
    Exit Sub

ErrHandler:
    MsgBox "Error unhiding sheets: " & Err.Description, vbExclamation, APP_NAME
End Sub


'===============================================================================
' SortSheetsAlphabetically - Reorder all worksheet tabs A-Z by name
'===============================================================================
Public Sub SortSheetsAlphabetically()
    On Error GoTo ErrHandler

    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("Sort all sheet tabs alphabetically?" & vbCrLf & vbCrLf & _
                     "This will reorder every tab in the workbook.", _
                     vbYesNo + vbQuestion, APP_NAME)
    If confirm = vbNo Then Exit Sub

    modPerformance.TurboOn

    Dim i As Long, j As Long

    ' Bubble sort — case-insensitive
    For i = 1 To ThisWorkbook.Sheets.Count - 1
        For j = i + 1 To ThisWorkbook.Sheets.Count
            If UCase$(ThisWorkbook.Sheets(j).Name) < UCase$(ThisWorkbook.Sheets(i).Name) Then
                ThisWorkbook.Sheets(j).Move Before:=ThisWorkbook.Sheets(i)
            End If
        Next j
    Next i

    modPerformance.TurboOff

    MsgBox "Sheets sorted alphabetically.", vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    MsgBox "Error sorting sheets: " & Err.Description, vbExclamation, APP_NAME
End Sub


'===============================================================================
' ToggleFreezePanes - Toggle freeze panes on or off
' When enabling: freezes Row 1 and Column A (pane anchored at B2)
'===============================================================================
Public Sub ToggleFreezePanes()
    On Error GoTo ErrHandler

    If ActiveWindow.FreezePanes Then
        ActiveWindow.FreezePanes = False
        Application.StatusBar = "Freeze panes OFF"
    Else
        ' Anchor at B2 so Row 1 (headers) and Col A (labels) stay fixed
        ActiveSheet.Range("B2").Select
        ActiveWindow.FreezePanes = True
        Application.StatusBar = "Freeze panes ON  (Row 1 + Col A frozen)"
    End If

    ' Auto-clear the status bar message after 3 seconds
    Application.OnTime Now + TimeValue("00:00:03"), "modUtilities.ClearStatusBar"
    Exit Sub

ErrHandler:
    MsgBox "Error toggling freeze panes: " & Err.Description, vbExclamation, APP_NAME
End Sub

'===============================================================================
' ClearStatusBar - Restore the default Excel status bar (called via OnTime)
'===============================================================================
Public Sub ClearStatusBar()
    Application.StatusBar = False
End Sub


'===============================================================================
' ConvertToValues - Replace all formulas in the selection with their results
' WARNING: This action is irreversible. Always back up before running.
'===============================================================================
Public Sub ConvertToValues()
    On Error GoTo ErrHandler

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim cellCount As Long: cellCount = Selection.Cells.Count

    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("Convert formulas to values in the selected " & cellCount & " cell(s)?" & _
                     vbCrLf & vbCrLf & "WARNING: This cannot be undone.", _
                     vbYesNo + vbExclamation, APP_NAME)
    If confirm = vbNo Then Exit Sub

    modPerformance.TurboOn
    Selection.Value = Selection.Value
    modPerformance.TurboOff

    MsgBox "Formulas converted to values in " & cellCount & " cell(s).", vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    MsgBox "Error converting to values: " & Err.Description, vbExclamation, APP_NAME
End Sub


'===============================================================================
' AutoFitAllColumns - AutoFit every column on the active sheet
'===============================================================================
Public Sub AutoFitAllColumns()
    On Error GoTo ErrHandler

    modPerformance.TurboOn
    ActiveSheet.Cells.EntireColumn.AutoFit
    modPerformance.TurboOff

    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    MsgBox "Error fitting columns: " & Err.Description, vbExclamation, APP_NAME
End Sub


'===============================================================================
' ProtectAllSheets - Password-protect every worksheet in the workbook
'===============================================================================
Public Sub ProtectAllSheets()
    On Error GoTo ErrHandler

    Dim pwd As String
    pwd = InputBox("Enter a password to protect all sheets:" & vbCrLf & vbCrLf & _
                   "(Leave blank to protect without a password — press OK to continue)", _
                   APP_NAME & " - Protect All Sheets")

    ' Cancelled (StrPtr = 0 means user hit Cancel, not just cleared the box)
    If StrPtr(pwd) = 0 Then Exit Sub

    modPerformance.TurboOn

    Dim ws As Worksheet
    Dim protectedCount As Long: protectedCount = 0

    For Each ws In ActiveWorkbook.Worksheets
        ws.Protect Password:=pwd, DrawingObjects:=True, Contents:=True, Scenarios:=True
        protectedCount = protectedCount + 1
    Next ws

    modPerformance.TurboOff

    MsgBox protectedCount & " sheet(s) protected." & vbCrLf & vbCrLf & _
           "Important: Store your password safely — it cannot be recovered.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    MsgBox "Error protecting sheets: " & Err.Description, vbExclamation, APP_NAME
End Sub


'===============================================================================
' UnprotectAllSheets - Remove password protection from every worksheet
'===============================================================================
Public Sub UnprotectAllSheets()
    On Error GoTo ErrHandler

    Dim pwd As String
    pwd = InputBox("Enter the password to unprotect all sheets:" & vbCrLf & vbCrLf & _
                   "(Leave blank if sheets were protected without a password)", _
                   APP_NAME & " - Unprotect All Sheets")

    If StrPtr(pwd) = 0 Then Exit Sub

    modPerformance.TurboOn

    Dim ws As Worksheet
    Dim successCount As Long: successCount = 0
    Dim failCount As Long: failCount = 0

    For Each ws In ActiveWorkbook.Worksheets
        On Error Resume Next
        ws.Unprotect Password:=pwd
        If Err.Number <> 0 Then
            failCount = failCount + 1
            Err.Clear
        Else
            successCount = successCount + 1
        End If
        On Error GoTo ErrHandler
    Next ws

    modPerformance.TurboOff

    If failCount = 0 Then
        MsgBox successCount & " sheet(s) unprotected successfully.", vbInformation, APP_NAME
    Else
        MsgBox successCount & " sheet(s) unprotected." & vbCrLf & _
               failCount & " sheet(s) failed — the password may be incorrect.", _
               vbExclamation, APP_NAME
    End If
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    MsgBox "Error unprotecting sheets: " & Err.Description, vbExclamation, APP_NAME
End Sub


'===============================================================================
' FindReplaceAllSheets - Find and replace text across every worksheet
'===============================================================================
Public Sub FindReplaceAllSheets()
    On Error GoTo ErrHandler

    Dim findText As String
    findText = InputBox("Find what:", APP_NAME & " - Find & Replace All Sheets")
    If StrPtr(findText) = 0 Or findText = "" Then Exit Sub

    Dim replaceText As String
    replaceText = InputBox("Replace with:", APP_NAME & " - Find & Replace All Sheets")
    If StrPtr(replaceText) = 0 Then Exit Sub

    ' Confirm before replacing with blank
    If replaceText = "" Then
        If MsgBox("Replace '" & findText & "' with blank (empty string)?" & vbCrLf & vbCrLf & _
                  "This will erase all matching text.", _
                  vbYesNo + vbExclamation, APP_NAME) = vbNo Then Exit Sub
    End If

    modPerformance.TurboOn

    Dim ws As Worksheet
    Dim replacedSheets As Long: replacedSheets = 0

    For Each ws In ActiveWorkbook.Worksheets
        Dim matchCount As Long
        matchCount = Application.WorksheetFunction.CountIf(ws.Cells, "*" & findText & "*")
        ws.Cells.Replace What:=findText, Replacement:=replaceText, _
                          LookAt:=xlPart, MatchCase:=False, SearchOrder:=xlByRows
        If matchCount > 0 Then replacedSheets = replacedSheets + 1
    Next ws

    modPerformance.TurboOff

    MsgBox "Find & Replace complete." & vbCrLf & _
           "Replaced """ & findText & """ with """ & replaceText & """" & vbCrLf & _
           "across " & replacedSheets & " sheet(s).", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    MsgBox "Error during Find & Replace: " & Err.Description, vbExclamation, APP_NAME
End Sub


'===============================================================================
' HighlightHardcodedNumbers - Change font color to blue for all non-formula numbers
' Standard audit convention: blue = hardcoded input, black = formula-driven
'===============================================================================
Public Sub HighlightHardcodedNumbers()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' SpecialCells(xlCellTypeConstants, xlNumbers) = cells with typed numbers only
    Dim constCells As Range
    On Error Resume Next
    Set constCells = ws.UsedRange.SpecialCells(xlCellTypeConstants, xlNumbers)
    On Error GoTo ErrHandler

    If constCells Is Nothing Then
        MsgBox "No hardcoded numbers found on '" & ws.Name & "'.", vbInformation, APP_NAME
        Exit Sub
    End If

    Dim cellCount As Long: cellCount = constCells.Cells.Count

    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("Found " & cellCount & " hardcoded number(s) on '" & ws.Name & "'." & _
                     vbCrLf & vbCrLf & "Change their font color to blue?", _
                     vbYesNo + vbQuestion, APP_NAME)
    If confirm = vbNo Then Exit Sub

    modPerformance.TurboOn
    constCells.Font.Color = RGB(0, 0, 255)   ' Blue — standard audit color
    modPerformance.TurboOff

    MsgBox cellCount & " hardcoded number(s) highlighted in blue.", vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    MsgBox "Error highlighting hardcoded numbers: " & Err.Description, vbExclamation, APP_NAME
End Sub


'===============================================================================
' TogglePresentationMode - Hide/show gridlines, headings, and formula bar
' Turns the workbook into a clean, professional view for demos and presentations.
' Run once to enter presentation mode. Run again to restore the normal view.
'===============================================================================
Public Sub TogglePresentationMode()
    On Error GoTo ErrHandler

    ' Use gridlines as the toggle indicator:
    '   Gridlines visible  = normal view  -> switch to presentation mode
    '   Gridlines hidden   = pres mode    -> restore normal view
    Dim inPresMode As Boolean
    inPresMode = Not ActiveWindow.DisplayGridlines

    If inPresMode Then
        ' Restore normal view
        ActiveWindow.DisplayGridlines = True
        ActiveWindow.DisplayHeadings = True
        Application.DisplayFormulaBar = True
        Application.StatusBar = "Presentation mode OFF — Normal view restored"
    Else
        ' Enter presentation mode
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
        Application.DisplayFormulaBar = False
        Application.StatusBar = "Presentation mode ON — Run again to restore normal view"
    End If

    Application.OnTime Now + TimeValue("00:00:03"), "modUtilities.ClearStatusBar"
    Exit Sub

ErrHandler:
    MsgBox "Error toggling presentation mode: " & Err.Description, vbExclamation, APP_NAME
End Sub


'===============================================================================
' UnmergeAndFillDown - Unmerge all merged cells in selection, then fill blanks
' Fixes messy ERP/system exports that use merged cells as section headers.
' After unmerging, each blank cell is filled with the value from the row above.
'===============================================================================
Public Sub UnmergeAndFillDown()
    On Error GoTo ErrHandler

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first.", vbExclamation, APP_NAME
        Exit Sub
    End If

    If Selection.Cells.Count < 2 Then
        MsgBox "Please select a range with more than one cell.", vbExclamation, APP_NAME
        Exit Sub
    End If

    modPerformance.TurboOn

    Dim rng As Range
    Set rng = Selection

    ' Step 1: Unmerge all merged cells in the selection
    rng.UnMerge

    ' Step 2: Fill each blank cell with the value from the cell directly above it
    Dim cell As Range
    For Each cell In rng
        If IsEmpty(cell.Value) Then
            If cell.Row > 1 Then
                cell.Value = cell.Offset(-1, 0).Value
            End If
        End If
    Next cell

    modPerformance.TurboOff

    MsgBox "Unmerge and fill down complete.", vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    MsgBox "Error in Unmerge and Fill Down: " & Err.Description, vbExclamation, APP_NAME
End Sub
