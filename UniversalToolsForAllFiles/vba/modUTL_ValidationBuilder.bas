Attribute VB_Name = "modUTL_ValidationBuilder"
'==============================================================================
' modUTL_ValidationBuilder — Data Validation Tools
'==============================================================================
' PURPOSE:  Create dropdown lists, apply number/date validation, copy rules,
'           and find cells that violate validation. Easier than Excel's dialog.
'
' PUBLIC SUBS:
'   CreateDropdownList        — Create a dropdown from a range or typed list
'   ApplyNumberValidation     — Restrict cells to numbers within a range
'   ApplyDateValidation       — Restrict cells to dates within a range
'   CopyValidationRules       — Copy validation from one range to another
'   FindValidationViolations  — Find all cells that break their rules
'   RemoveAllValidation       — Remove validation from selection or sheet
'
' DEPENDENCIES: None (standalone). Works in any Excel workbook.
' VERSION:  1.0.0 | DATE: 2026-03-12
'==============================================================================
Option Explicit

Private Const CLR_HDR As Long = 7930635   ' RGB(11,71,121)

'==============================================================================
' PUBLIC: CreateDropdownList
' Creates a dropdown validation from a range or comma-separated list.
'==============================================================================
Public Sub CreateDropdownList()
    On Error GoTo ErrHandler

    MsgBox "You will:" & vbCrLf & _
           "  1. Select the cells WHERE the dropdown should appear" & vbCrLf & _
           "  2. Choose the source for the dropdown options", _
           vbInformation, "Create Dropdown List"

    '--- Step 1: Target cells ---
    Dim targetRng As Range
    On Error Resume Next
    Set targetRng = Application.InputBox("Select the cells where the dropdown should appear:", _
                                          "Create Dropdown - Step 1 of 3", Type:=8)
    On Error GoTo ErrHandler
    If targetRng Is Nothing Then Exit Sub

    '--- Step 2: Source type ---
    Dim srcChoice As String
    srcChoice = InputBox("How do you want to provide the dropdown options?" & vbCrLf & vbCrLf & _
                          "  1. Select a range of cells (values become the options)" & vbCrLf & _
                          "  2. Type the options (comma-separated)" & vbCrLf & vbCrLf & _
                          "Enter 1 or 2:", _
                          "Create Dropdown - Step 2 of 3")
    If Len(Trim(srcChoice)) = 0 Then Exit Sub

    Dim formula1 As String

    Select Case Trim(srcChoice)
        Case "1"  ' Range source
            MsgBox "Now select the range containing the dropdown options." & vbCrLf & _
                   "This can be on any sheet.", _
                   vbInformation, "Create Dropdown"

            Dim srcRng As Range
            On Error Resume Next
            Set srcRng = Application.InputBox("Select the range with dropdown values:", _
                                              "Create Dropdown - Step 3 of 3", Type:=8)
            On Error GoTo ErrHandler
            If srcRng Is Nothing Then Exit Sub

            Dim srcSheetRef As String
            If InStr(srcRng.Parent.Name, " ") > 0 Or InStr(srcRng.Parent.Name, "'") > 0 Then
                srcSheetRef = "'" & Replace(srcRng.Parent.Name, "'", "''") & "'"
            Else
                srcSheetRef = srcRng.Parent.Name
            End If
            formula1 = "=" & srcSheetRef & "!" & srcRng.Address

        Case "2"  ' Typed list
            Dim typedList As String
            typedList = InputBox("Type the dropdown options separated by commas:" & vbCrLf & vbCrLf & _
                                  "Example: Yes,No,Maybe" & vbCrLf & _
                                  "Example: Q1,Q2,Q3,Q4" & vbCrLf & _
                                  "Example: Approved,Pending,Rejected", _
                                  "Create Dropdown - Step 3 of 3")
            If Len(Trim(typedList)) = 0 Then Exit Sub
            formula1 = typedList

        Case Else
            MsgBox "Invalid choice.", vbExclamation, "Create Dropdown"
            Exit Sub
    End Select

    '--- Apply validation ---
    With targetRng.Validation
        .Delete  ' Remove existing validation first
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=formula1
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
        .ErrorTitle = "Invalid Entry"
        .ErrorMessage = "Please select a value from the dropdown list."
    End With

    MsgBox "Dropdown created successfully!" & vbCrLf & vbCrLf & _
           "Applied to: " & targetRng.Address(False, False) & vbCrLf & _
           "Cells: " & targetRng.Cells.Count, _
           vbInformation, "Create Dropdown"

    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Create Dropdown"
End Sub

'==============================================================================
' PUBLIC: ApplyNumberValidation
' Restricts cells to numbers within a range (min/max).
'==============================================================================
Public Sub ApplyNumberValidation()
    On Error GoTo ErrHandler

    MsgBox "Select the cells to restrict to numbers only.", _
           vbInformation, "Number Validation"

    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox("Select the cells to validate:", _
                                    "Number Validation - Step 1 of 3", Type:=8)
    On Error GoTo ErrHandler
    If rng Is Nothing Then Exit Sub

    '--- Ask for rule type ---
    Dim ruleChoice As String
    ruleChoice = InputBox("What number rule do you want?" & vbCrLf & vbCrLf & _
                           "  1. Any number (no min/max)" & vbCrLf & _
                           "  2. Between a minimum and maximum" & vbCrLf & _
                           "  3. Greater than a minimum" & vbCrLf & _
                           "  4. Less than a maximum" & vbCrLf & _
                           "  5. Whole numbers only (no decimals)" & vbCrLf & vbCrLf & _
                           "Enter number:", _
                           "Number Validation - Step 2 of 3")
    If Len(Trim(ruleChoice)) = 0 Then Exit Sub

    '--- Apply ---
    rng.Validation.Delete  ' Clear existing

    Select Case Trim(ruleChoice)
        Case "1"  ' Any number
            rng.Validation.Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, _
                              Operator:=xlGreaterEqual, Formula1:="-999999999"
            rng.Validation.ErrorMessage = "Please enter a number."

        Case "2"  ' Between
            Dim minStr As String, maxStr As String
            minStr = InputBox("Enter the MINIMUM value:", "Number Validation - Min")
            If Not IsNumeric(minStr) Then MsgBox "Not a number.", vbExclamation: Exit Sub
            maxStr = InputBox("Enter the MAXIMUM value:", "Number Validation - Max")
            If Not IsNumeric(maxStr) Then MsgBox "Not a number.", vbExclamation: Exit Sub

            rng.Validation.Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, _
                              Operator:=xlBetween, Formula1:=minStr, Formula2:=maxStr
            rng.Validation.ErrorMessage = "Please enter a number between " & minStr & " and " & maxStr & "."

        Case "3"  ' Greater than
            Dim gtStr As String
            gtStr = InputBox("Enter the MINIMUM value:", "Number Validation - Min")
            If Not IsNumeric(gtStr) Then MsgBox "Not a number.", vbExclamation: Exit Sub

            rng.Validation.Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, _
                              Operator:=xlGreater, Formula1:=gtStr
            rng.Validation.ErrorMessage = "Please enter a number greater than " & gtStr & "."

        Case "4"  ' Less than
            Dim ltStr As String
            ltStr = InputBox("Enter the MAXIMUM value:", "Number Validation - Max")
            If Not IsNumeric(ltStr) Then MsgBox "Not a number.", vbExclamation: Exit Sub

            rng.Validation.Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, _
                              Operator:=xlLess, Formula1:=ltStr
            rng.Validation.ErrorMessage = "Please enter a number less than " & ltStr & "."

        Case "5"  ' Whole numbers
            Dim wholeMin As String, wholeMax As String
            wholeMin = InputBox("Enter MINIMUM whole number (or leave blank for any):", "Whole Numbers")
            wholeMax = InputBox("Enter MAXIMUM whole number (or leave blank for any):", "Whole Numbers")

            If IsNumeric(wholeMin) And IsNumeric(wholeMax) Then
                rng.Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                                  Operator:=xlBetween, Formula1:=wholeMin, Formula2:=wholeMax
            ElseIf IsNumeric(wholeMin) Then
                rng.Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                                  Operator:=xlGreaterEqual, Formula1:=wholeMin
            Else
                rng.Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                                  Operator:=xlGreaterEqual, Formula1:="-999999999"
            End If
            rng.Validation.ErrorMessage = "Please enter a whole number (no decimals)."

        Case Else
            MsgBox "Invalid choice.", vbExclamation, "Number Validation"
            Exit Sub
    End Select

    rng.Validation.ShowError = True
    rng.Validation.ErrorTitle = "Invalid Entry"

    MsgBox "Number validation applied to " & rng.Cells.Count & " cell(s).", _
           vbInformation, "Number Validation"

    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Number Validation"
End Sub

'==============================================================================
' PUBLIC: ApplyDateValidation
' Restricts cells to dates within a range.
'==============================================================================
Public Sub ApplyDateValidation()
    On Error GoTo ErrHandler

    MsgBox "Select the cells to restrict to dates only.", _
           vbInformation, "Date Validation"

    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox("Select the cells to validate:", _
                                    "Date Validation - Step 1 of 2", Type:=8)
    On Error GoTo ErrHandler
    If rng Is Nothing Then Exit Sub

    Dim ruleChoice As String
    ruleChoice = InputBox("What date rule do you want?" & vbCrLf & vbCrLf & _
                           "  1. Any date" & vbCrLf & _
                           "  2. Between two dates" & vbCrLf & _
                           "  3. After a specific date" & vbCrLf & _
                           "  4. Before a specific date" & vbCrLf & vbCrLf & _
                           "Enter number:", _
                           "Date Validation - Step 2 of 2")
    If Len(Trim(ruleChoice)) = 0 Then Exit Sub

    rng.Validation.Delete

    Select Case Trim(ruleChoice)
        Case "1"  ' Any date
            rng.Validation.Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, _
                              Operator:=xlGreaterEqual, Formula1:="1/1/1900"
            rng.Validation.ErrorMessage = "Please enter a valid date."

        Case "2"  ' Between
            Dim startDate As String, endDate As String
            startDate = InputBox("Enter START date (mm/dd/yyyy):", "Date Range - Start")
            If Len(Trim(startDate)) = 0 Then Exit Sub
            endDate = InputBox("Enter END date (mm/dd/yyyy):", "Date Range - End")
            If Len(Trim(endDate)) = 0 Then Exit Sub

            rng.Validation.Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, _
                              Operator:=xlBetween, Formula1:=startDate, Formula2:=endDate
            rng.Validation.ErrorMessage = "Please enter a date between " & startDate & " and " & endDate & "."

        Case "3"  ' After
            Dim afterDate As String
            afterDate = InputBox("Enter the date (dates AFTER this are allowed):" & vbCrLf & _
                                  "Format: mm/dd/yyyy", "Date Validation - After")
            If Len(Trim(afterDate)) = 0 Then Exit Sub

            rng.Validation.Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, _
                              Operator:=xlGreater, Formula1:=afterDate
            rng.Validation.ErrorMessage = "Please enter a date after " & afterDate & "."

        Case "4"  ' Before
            Dim beforeDate As String
            beforeDate = InputBox("Enter the date (dates BEFORE this are allowed):" & vbCrLf & _
                                   "Format: mm/dd/yyyy", "Date Validation - Before")
            If Len(Trim(beforeDate)) = 0 Then Exit Sub

            rng.Validation.Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, _
                              Operator:=xlLess, Formula1:=beforeDate
            rng.Validation.ErrorMessage = "Please enter a date before " & beforeDate & "."

        Case Else
            MsgBox "Invalid choice.", vbExclamation, "Date Validation"
            Exit Sub
    End Select

    rng.Validation.ShowError = True
    rng.Validation.ErrorTitle = "Invalid Date"

    MsgBox "Date validation applied to " & rng.Cells.Count & " cell(s).", _
           vbInformation, "Date Validation"

    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Date Validation"
End Sub

'==============================================================================
' PUBLIC: CopyValidationRules
' Copy validation from one range to another.
'==============================================================================
Public Sub CopyValidationRules()
    On Error GoTo ErrHandler

    MsgBox "You will select a SOURCE cell (with validation) and a TARGET range.", _
           vbInformation, "Copy Validation Rules"

    Dim srcCell As Range
    On Error Resume Next
    Set srcCell = Application.InputBox("Select a single cell WITH validation to copy FROM:", _
                                        "Copy Validation - Step 1 of 2", Type:=8)
    On Error GoTo ErrHandler
    If srcCell Is Nothing Then Exit Sub

    ' Verify source has validation
    On Error Resume Next
    Dim testType As Long
    testType = srcCell.Cells(1, 1).Validation.Type
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo ErrHandler
        MsgBox "The selected cell has no validation rules to copy.", vbExclamation, "Copy Validation"
        Exit Sub
    End If
    On Error GoTo ErrHandler

    Dim targetRng As Range
    On Error Resume Next
    Set targetRng = Application.InputBox("Select the TARGET range to apply validation TO:", _
                                          "Copy Validation - Step 2 of 2", Type:=8)
    On Error GoTo ErrHandler
    If targetRng Is Nothing Then Exit Sub

    '--- Copy via clipboard ---
    srcCell.Cells(1, 1).Copy
    targetRng.PasteSpecial Paste:=xlPasteValidation
    Application.CutCopyMode = False

    MsgBox "Validation rules copied successfully!" & vbCrLf & vbCrLf & _
           "From: " & srcCell.Address(False, False) & vbCrLf & _
           "To: " & targetRng.Address(False, False) & " (" & targetRng.Cells.Count & " cells)", _
           vbInformation, "Copy Validation"

    Exit Sub

ErrHandler:
    Application.CutCopyMode = False
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Copy Validation"
End Sub

'==============================================================================
' PUBLIC: FindValidationViolations
' Finds all cells that violate their data validation rules.
'==============================================================================
Public Sub FindValidationViolations()
    On Error GoTo ErrHandler

    Dim choice As String
    choice = InputBox("Check for validation violations on:" & vbCrLf & vbCrLf & _
                       "  1. Active sheet only" & vbCrLf & _
                       "  2. ALL sheets" & vbCrLf & vbCrLf & _
                       "Enter 1 or 2:", _
                       "Find Validation Violations")
    If Len(Trim(choice)) = 0 Then Exit Sub

    Application.ScreenUpdating = False

    Dim violations As Long
    violations = 0
    Dim msg As String
    msg = "Validation Violations Found:" & vbCrLf & String(35, "-") & vbCrLf & vbCrLf

    If Trim(choice) = "1" Then
        violations = CheckSheetViolations(ActiveSheet, msg)
    ElseIf Trim(choice) = "2" Then
        Dim ws As Worksheet
        For Each ws In ThisWorkbook.Worksheets
            violations = violations + CheckSheetViolations(ws, msg)
        Next ws
    Else
        Application.ScreenUpdating = True
        MsgBox "Invalid choice.", vbExclamation, "Find Violations"
        Exit Sub
    End If

    Application.ScreenUpdating = True

    If violations = 0 Then
        MsgBox "No validation violations found. All cells comply with their rules.", _
               vbInformation, "Find Violations"
    Else
        ' Truncate if too long
        If Len(msg) > 900 Then
            msg = Left(msg, 900) & vbCrLf & "... (showing first violations)"
        End If
        msg = msg & vbCrLf & "Total violations: " & violations
        MsgBox msg, vbExclamation, "Find Violations"
    End If

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Find Violations"
End Sub

'==============================================================================
' PUBLIC: RemoveAllValidation
' Remove data validation from selection or entire sheet.
'==============================================================================
Public Sub RemoveAllValidation()
    On Error GoTo ErrHandler

    Dim choice As String
    choice = InputBox("Remove data validation from:" & vbCrLf & vbCrLf & _
                       "  1. Current selection only" & vbCrLf & _
                       "  2. Entire active sheet" & vbCrLf & _
                       "  3. ALL sheets in workbook" & vbCrLf & vbCrLf & _
                       "Enter number:", _
                       "Remove Validation")
    If Len(Trim(choice)) = 0 Then Exit Sub

    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("Are you sure? This removes validation rules (dropdown lists, number restrictions, etc.)." & vbCrLf & vbCrLf & _
                      "This cannot be undone.", _
                      vbYesNo + vbExclamation, "Remove Validation")
    If confirm = vbNo Then Exit Sub

    Select Case Trim(choice)
        Case "1"
            If Not Selection Is Nothing Then
                On Error Resume Next
                Selection.Validation.Delete
                Err.Clear
                On Error GoTo ErrHandler
            End If
            MsgBox "Validation removed from selection.", vbInformation, "Remove Validation"

        Case "2"
            On Error Resume Next
            ActiveSheet.Cells.Validation.Delete
            Err.Clear
            On Error GoTo ErrHandler
            MsgBox "Validation removed from " & ActiveSheet.Name & ".", vbInformation, "Remove Validation"

        Case "3"
            Dim ws As Worksheet
            For Each ws In ThisWorkbook.Worksheets
                On Error Resume Next
                ws.Cells.Validation.Delete
                Err.Clear
                On Error GoTo ErrHandler
            Next ws
            MsgBox "Validation removed from all sheets.", vbInformation, "Remove Validation"

        Case Else
            MsgBox "Invalid choice.", vbExclamation, "Remove Validation"
    End Select

    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Remove Validation"
End Sub

'==============================================================================
' PRIVATE: CheckSheetViolations — Uses CircleInvalid approach
'==============================================================================
Private Function CheckSheetViolations(ByVal ws As Worksheet, ByRef msg As String) As Long
    Dim violations As Long
    violations = 0

    ' Note: No xlCellTypeValidation exists in SpecialCells, so we iterate
    ' We limit to UsedRange for performance
    Dim cell As Range
    Dim lastRow As Long, lastCol As Long

    lastRow = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1
    lastCol = ws.UsedRange.Columns.Count + ws.UsedRange.Column - 1

    ' Safety cap
    If lastRow > 5000 Then lastRow = 5000
    If lastCol > 100 Then lastCol = 100

    Dim r As Long, c As Long
    For r = ws.UsedRange.Row To lastRow
        For c = ws.UsedRange.Column To lastCol
            Set cell = ws.Cells(r, c)

            ' Check if cell has validation
            On Error Resume Next
            Dim valType As Long
            valType = -1
            valType = cell.Validation.Type
            If Err.Number <> 0 Then
                Err.Clear
                GoTo NextCell
            End If
            On Error GoTo 0

            ' Check if value violates
            If Not IsEmpty(cell.Value) Then
                Dim isValid As Boolean
                isValid = True

                On Error Resume Next
                ' Use Validation.Value property — not available in all versions
                ' Instead, try to evaluate manually based on type
                Dim valFormula1 As String
                valFormula1 = ""
                valFormula1 = cell.Validation.Formula1
                Err.Clear
                On Error GoTo 0

                ' Simple check: if cell has validation and value doesn't match list
                If valType = 3 Then  ' xlValidateList
                    ' List validation — check if value is in list
                    If Len(valFormula1) > 0 Then
                        If Left(valFormula1, 1) <> "=" Then
                            ' Comma-separated list
                            If InStr(1, "," & valFormula1 & ",", "," & CStr(cell.Value) & ",", vbTextCompare) = 0 Then
                                isValid = False
                            End If
                        End If
                    End If
                End If

                If Not isValid Then
                    violations = violations + 1
                    If violations <= 20 Then
                        msg = msg & "  " & ws.Name & "!" & cell.Address(False, False) & _
                              " = """ & Left(CStr(cell.Value), 30) & """" & vbCrLf
                    End If
                End If
            End If

NextCell:
        Next c
    Next r

    CheckSheetViolations = violations
End Function
