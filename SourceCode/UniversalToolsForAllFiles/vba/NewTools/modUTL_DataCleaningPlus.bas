Attribute VB_Name = "modUTL_DataCleaningPlus"
Option Explicit

' ============================================================
' KBT Universal Tools — Data Cleaning Plus Module
' Works on ANY Excel file — no project-specific setup required
' Tools: 3 | All Small effort
' Date: 2026-03-05
' ============================================================
' Tool 01 — Universal Whitespace Cleaner
' Tool 03 — Non-Printable Character Stripper
' Tool 04 — Text Case Standardizer
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
' TOOL 01 — Universal Whitespace Cleaner                [SMALL]
' Removes leading/trailing spaces, double spaces, non-breaking
' spaces (Chr 160), and zero-width chars from all text cells
' on the active sheet. Reports count of cells cleaned.
' ============================================================
Sub UniversalWhitespaceCleaner()
    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim rng As Range
    On Error Resume Next
    Set rng = ws.UsedRange.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo ErrHandler
    If rng Is Nothing Then
        UTL_TurboOff
        MsgBox "No text cells found on '" & ws.Name & "'.", vbInformation
        Exit Sub
    End If

    Dim cleaned As Long: cleaned = 0
    Dim cell As Range
    For Each cell In rng
        Dim original As String: original = cell.Value
        Dim fixed As String: fixed = original

        ' Replace non-breaking space (Chr 160) with regular space
        fixed = Replace(fixed, Chr(160), " ")
        ' Remove zero-width space (Unicode 8203)
        fixed = Replace(fixed, ChrW(8203), "")
        ' Collapse multiple spaces to single
        Do While InStr(fixed, "  ") > 0
            fixed = Replace(fixed, "  ", " ")
        Loop
        ' Trim leading/trailing
        fixed = Trim(fixed)

        If fixed <> original Then
            cell.Value = fixed
            cleaned = cleaned + 1
        End If
    Next cell

    UTL_TurboOff
    MsgBox "Whitespace Cleaner Complete" & vbCrLf & vbCrLf & _
           cleaned & " cell(s) cleaned on '" & ws.Name & "'.", vbInformation
    Exit Sub

ErrHandler:
    UTL_TurboOff
    MsgBox "Whitespace Cleaner error: " & Err.Description, vbCritical
End Sub

' ============================================================
' TOOL 03 — Non-Printable Character Stripper            [SMALL]
' Removes control characters (ASCII 0-31 except Tab/CR/LF),
' soft hyphens, and other invisible characters from all text
' cells on the active sheet.
' ============================================================
Sub NonPrintableCharStripper()
    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim rng As Range
    On Error Resume Next
    Set rng = ws.UsedRange.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo ErrHandler
    If rng Is Nothing Then
        UTL_TurboOff
        MsgBox "No text cells found.", vbInformation
        Exit Sub
    End If

    Dim cleaned As Long: cleaned = 0
    Dim cell As Range
    For Each cell In rng
        Dim original As String: original = cell.Value
        Dim fixed As String: fixed = ""
        Dim i As Long
        For i = 1 To Len(original)
            Dim ch As Long: ch = AscW(Mid(original, i, 1))
            ' Keep printable ASCII, tab (9), LF (10), CR (13), and all Unicode > 31
            ' Exclude soft hyphen (173), zero-width space (8203), BOM (65279)
            If ch >= 32 Or ch = 9 Or ch = 10 Or ch = 13 Then
                If ch <> 173 And ch <> 8203 And ch <> 65279 Then
                    fixed = fixed & Mid(original, i, 1)
                End If
            End If
        Next i

        If fixed <> original Then
            cell.Value = fixed
            cleaned = cleaned + 1
        End If
    Next cell

    UTL_TurboOff
    MsgBox "Non-Printable Stripper Complete" & vbCrLf & vbCrLf & _
           cleaned & " cell(s) cleaned on '" & ws.Name & "'.", vbInformation
    Exit Sub

ErrHandler:
    UTL_TurboOff
    MsgBox "Non-Printable Stripper error: " & Err.Description, vbCritical
End Sub

' ============================================================
' TOOL 04 — Text Case Standardizer                      [SMALL]
' Prompts user for case type (UPPER, lower, Title Case, or
' Sentence case), then applies to all selected cells.
' Works on selection only — does not touch unselected cells.
' ============================================================
Sub TextCaseStandardizer()
    On Error GoTo ErrHandler

    Dim choice As String
    choice = InputBox("Choose case format:" & vbCrLf & vbCrLf & _
                      "1 = UPPER CASE" & vbCrLf & _
                      "2 = lower case" & vbCrLf & _
                      "3 = Title Case" & vbCrLf & _
                      "4 = Sentence case" & vbCrLf & vbCrLf & _
                      "Enter 1, 2, 3, or 4:", _
                      "Text Case Standardizer", "3")
    If choice = "" Then Exit Sub

    UTL_TurboOn

    Dim rng As Range
    Set rng = Selection
    If rng Is Nothing Then
        UTL_TurboOff
        MsgBox "Please select cells first.", vbExclamation
        Exit Sub
    End If

    Dim changed As Long: changed = 0
    Dim cell As Range
    For Each cell In rng
        If VarType(cell.Value) = vbString And Len(cell.Value) > 0 Then
            Dim original As String: original = cell.Value
            Dim result As String

            Select Case choice
                Case "1"
                    result = UCase(original)
                Case "2"
                    result = LCase(original)
                Case "3"
                    result = Application.WorksheetFunction.Proper(original)
                Case "4"
                    ' Sentence case: capitalize first letter, lowercase rest
                    result = UCase(Left(original, 1)) & LCase(Mid(original, 2))
                Case Else
                    UTL_TurboOff
                    MsgBox "Invalid choice. Enter 1, 2, 3, or 4.", vbExclamation
                    Exit Sub
            End Select

            If result <> original Then
                cell.Value = result
                changed = changed + 1
            End If
        End If
    Next cell

    UTL_TurboOff
    MsgBox "Case Standardizer Complete" & vbCrLf & vbCrLf & _
           changed & " cell(s) updated.", vbInformation
    Exit Sub

ErrHandler:
    UTL_TurboOff
    MsgBox "Case Standardizer error: " & Err.Description, vbCritical
End Sub
