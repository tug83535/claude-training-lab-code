Attribute VB_Name = "modUTL_Highlights"
'==============================================================================
' modUTL_Highlights — Quick Conditional Highlighting Tools
'==============================================================================
' PURPOSE:  One-click highlighting: thresholds, top/bottom N, color scales,
'           outliers, and clear. Simpler than Excel's 6-click CF dialog.
'
' PUBLIC SUBS:
'   HighlightByThreshold   — Highlight cells above/below a value you type
'   HighlightTopBottom     — Highlight top N or bottom N values
'   HighlightDuplicateValues — Highlight duplicate values in a range
'   ApplyColorScale        — Red-Yellow-Green color scale on a range
'   ClearHighlights        — Remove highlighting from selection or sheet
'
' DEPENDENCIES: None (standalone). Works in any Excel workbook.
' VERSION:  1.0.0 | DATE: 2026-03-12
'==============================================================================
Option Explicit

'==============================================================================
' PUBLIC: HighlightByThreshold
' User selects range, types a threshold, picks above/below/both.
'==============================================================================
Public Sub HighlightByThreshold()
    On Error GoTo ErrHandler

    Dim rng As Range

    ' If a multi-cell range is already selected, use it (Director-friendly)
    If Not TypeOf Selection Is Range Then GoTo AskRange
    If Selection.Cells.Count > 1 Then
        Set rng = Selection
        GoTo HaveRange
    End If

AskRange:
    MsgBox "Select the range of numbers to check against a threshold." & vbCrLf & vbCrLf & _
           "Cells meeting your criteria will be highlighted.", _
           vbInformation, "Highlight by Threshold"

    On Error Resume Next
    Set rng = Application.InputBox("Select the range to check:", _
                                    "Highlight by Threshold - Step 1 of 3", Type:=8)
    On Error GoTo ErrHandler
    If rng Is Nothing Then Exit Sub

HaveRange:

    '--- Ask for threshold value ---
    Dim threshStr As String
    threshStr = InputBox("Enter the threshold value:" & vbCrLf & vbCrLf & _
                          "Example: 1000, -50, 15.5", _
                          "Highlight by Threshold - Step 2 of 3")
    If Len(Trim(threshStr)) = 0 Then Exit Sub
    If Not IsNumeric(threshStr) Then
        MsgBox "Please enter a number.", vbExclamation, "Highlight by Threshold"
        Exit Sub
    End If
    Dim threshold As Double
    threshold = CDbl(threshStr)

    '--- Ask for direction ---
    Dim dirChoice As String
    dirChoice = InputBox("How should cells be compared to " & threshold & "?" & vbCrLf & vbCrLf & _
                          "  1. Highlight cells ABOVE " & threshold & " (green)" & vbCrLf & _
                          "  2. Highlight cells BELOW " & threshold & " (red)" & vbCrLf & _
                          "  3. Highlight BOTH (above=green, below=red)" & vbCrLf & _
                          "  4. Highlight cells EQUAL to " & threshold & " (yellow)" & vbCrLf & vbCrLf & _
                          "Enter number:", _
                          "Highlight by Threshold - Step 3 of 3")
    If Len(Trim(dirChoice)) = 0 Then Exit Sub

    Application.ScreenUpdating = False

    Dim cell As Range
    Dim hitCount As Long
    hitCount = 0

    For Each cell In rng.Cells
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            Dim val As Double
            val = CDbl(cell.Value)

            Select Case Trim(dirChoice)
                Case "1"  ' Above
                    If val > threshold Then
                        cell.Interior.Color = RGB(198, 239, 206)  ' Light green
                        hitCount = hitCount + 1
                    End If
                Case "2"  ' Below
                    If val < threshold Then
                        cell.Interior.Color = RGB(255, 199, 206)  ' Light red
                        hitCount = hitCount + 1
                    End If
                Case "3"  ' Both
                    If val > threshold Then
                        cell.Interior.Color = RGB(198, 239, 206)
                        hitCount = hitCount + 1
                    ElseIf val < threshold Then
                        cell.Interior.Color = RGB(255, 199, 206)
                        hitCount = hitCount + 1
                    End If
                Case "4"  ' Equal
                    If Abs(val - threshold) < 0.0001 Then
                        cell.Interior.Color = RGB(255, 255, 153)  ' Yellow
                        hitCount = hitCount + 1
                    End If
            End Select
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Highlighting complete!" & vbCrLf & vbCrLf & _
           "Cells highlighted: " & hitCount & " of " & rng.Cells.Count, _
           vbInformation, "Highlight by Threshold"

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Highlight by Threshold"
End Sub

'==============================================================================
' PUBLIC: HighlightTopBottom
' Highlight the top N or bottom N values in a selected range.
'==============================================================================
Public Sub HighlightTopBottom()
    On Error GoTo ErrHandler

    MsgBox "Select the range of numbers to find top/bottom values.", _
           vbInformation, "Highlight Top/Bottom"

    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox("Select the range:", _
                                    "Highlight Top/Bottom - Step 1 of 3", Type:=8)
    On Error GoTo ErrHandler
    If rng Is Nothing Then Exit Sub

    '--- Top or bottom? ---
    Dim tbChoice As String
    tbChoice = InputBox("What do you want to highlight?" & vbCrLf & vbCrLf & _
                         "  1. TOP N values (highest)" & vbCrLf & _
                         "  2. BOTTOM N values (lowest)" & vbCrLf & _
                         "  3. BOTH top and bottom N" & vbCrLf & vbCrLf & _
                         "Enter number:", _
                         "Highlight Top/Bottom - Step 2 of 3")
    If Len(Trim(tbChoice)) = 0 Then Exit Sub

    '--- How many? ---
    Dim nStr As String
    nStr = InputBox("How many values to highlight?" & vbCrLf & vbCrLf & _
                     "Example: 5, 10, 20", _
                     "Highlight Top/Bottom - Step 3 of 3")
    If Len(Trim(nStr)) = 0 Then Exit Sub
    If Not IsNumeric(nStr) Then
        MsgBox "Please enter a number.", vbExclamation, "Highlight Top/Bottom"
        Exit Sub
    End If
    Dim n As Long
    n = CLng(nStr)
    If n < 1 Then Exit Sub

    '--- Safety cap: prevent overflow on very large ranges ---
    If rng.Cells.CountLarge > 500000 Then
        MsgBox "Selected range has over 500,000 cells. Please select a smaller range.", _
               vbExclamation, "Highlight Top/Bottom"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    '--- Collect all numeric values ---
    Dim vals() As Double
    Dim valCount As Long
    valCount = 0
    ReDim vals(1 To rng.Cells.Count)

    Dim cell As Range
    For Each cell In rng.Cells
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            valCount = valCount + 1
            vals(valCount) = CDbl(cell.Value)
        End If
    Next cell

    If valCount = 0 Then
        Application.ScreenUpdating = True
        MsgBox "No numeric values found.", vbInformation, "Highlight Top/Bottom"
        Exit Sub
    End If

    If n > valCount Then n = valCount

    '--- Sort to find thresholds ---
    ReDim Preserve vals(1 To valCount)
    SortArray vals, valCount

    Dim topThreshold As Double
    Dim bottomThreshold As Double
    topThreshold = vals(valCount - n + 1)      ' Nth highest
    bottomThreshold = vals(n)                   ' Nth lowest

    '--- Apply highlights ---
    Dim hitCount As Long
    hitCount = 0

    For Each cell In rng.Cells
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            Dim v As Double
            v = CDbl(cell.Value)

            Select Case Trim(tbChoice)
                Case "1"  ' Top only
                    If v >= topThreshold Then
                        cell.Interior.Color = RGB(198, 239, 206)
                        hitCount = hitCount + 1
                    End If
                Case "2"  ' Bottom only
                    If v <= bottomThreshold Then
                        cell.Interior.Color = RGB(255, 199, 206)
                        hitCount = hitCount + 1
                    End If
                Case "3"  ' Both
                    If v >= topThreshold Then
                        cell.Interior.Color = RGB(198, 239, 206)
                        hitCount = hitCount + 1
                    ElseIf v <= bottomThreshold Then
                        cell.Interior.Color = RGB(255, 199, 206)
                        hitCount = hitCount + 1
                    End If
            End Select
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Highlighting complete!" & vbCrLf & vbCrLf & _
           "Cells highlighted: " & hitCount, _
           vbInformation, "Highlight Top/Bottom"

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Highlight Top/Bottom"
End Sub

'==============================================================================
' PUBLIC: HighlightDuplicateValues
' Highlights cells with duplicate values in a selected range.
'==============================================================================
Public Sub HighlightDuplicateValues()
    On Error GoTo ErrHandler

    Dim rng As Range

    ' If a multi-cell range is already selected, use it (Director-friendly)
    If Not TypeOf Selection Is Range Then GoTo AskDupRange
    If Selection.Cells.Count > 1 Then
        Set rng = Selection
        GoTo HaveDupRange
    End If

AskDupRange:
    MsgBox "Select the range to check for duplicate values." & vbCrLf & vbCrLf & _
           "Duplicate values will be highlighted in orange.", _
           vbInformation, "Highlight Duplicates"

    On Error Resume Next
    Set rng = Application.InputBox("Select the range:", _
                                    "Highlight Duplicates", Type:=8)
    On Error GoTo ErrHandler
    If rng Is Nothing Then Exit Sub

HaveDupRange:

    '--- Safety cap: prevent slow loop on very large ranges ---
    If rng.Cells.CountLarge > 500000 Then
        MsgBox "Selected range has over 500,000 cells. Please select a smaller range.", _
               vbExclamation, "Highlight Duplicates"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    '--- Count occurrences using Collection for O(1) lookup ---
    Dim cell As Range
    Dim key As String
    Dim dupCount As Long
    dupCount = 0

    ' First pass: count occurrences using Collection
    Dim countCol As New Collection
    For Each cell In rng.Cells
        If Not IsEmpty(cell.Value) Then
            key = LCase(CStr(cell.Value))
            Dim cnt As Long
            cnt = 0
            On Error Resume Next
            cnt = countCol(key)
            On Error GoTo ErrHandler
            If cnt = 0 Then
                countCol.Add 1, key
            Else
                ' Remove and re-add with incremented count
                countCol.Remove key
                countCol.Add cnt + 1, key
            End If
        End If
    Next cell

    ' Second pass: highlight duplicates
    For Each cell In rng.Cells
        If Not IsEmpty(cell.Value) Then
            key = LCase(CStr(cell.Value))
            Dim cellCnt As Long
            cellCnt = 0
            On Error Resume Next
            cellCnt = countCol(key)
            On Error GoTo ErrHandler
            If cellCnt > 1 Then
                cell.Interior.Color = RGB(255, 200, 100)  ' Orange
                dupCount = dupCount + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Duplicate check complete!" & vbCrLf & vbCrLf & _
           "Cells with duplicate values: " & dupCount, _
           vbInformation, "Highlight Duplicates"

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Highlight Duplicates"
End Sub

'==============================================================================
' PUBLIC: ApplyColorScale
' Applies a red-yellow-green gradient based on cell values.
'==============================================================================
Public Sub ApplyColorScale()
    On Error GoTo ErrHandler

    MsgBox "Select a range of numbers to apply a color scale." & vbCrLf & vbCrLf & _
           "Low values = Red, Middle = Yellow, High = Green", _
           vbInformation, "Color Scale"

    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox("Select the range:", _
                                    "Color Scale", Type:=8)
    On Error GoTo ErrHandler
    If rng Is Nothing Then Exit Sub

    '--- Ask direction ---
    Dim dirChoice As String
    dirChoice = InputBox("Color direction:" & vbCrLf & vbCrLf & _
                          "  1. Low=Red, High=Green (default — higher is better)" & vbCrLf & _
                          "  2. Low=Green, High=Red (lower is better, e.g., costs)" & vbCrLf & vbCrLf & _
                          "Enter 1 or 2:", _
                          "Color Scale Direction")
    If Len(Trim(dirChoice)) = 0 Then dirChoice = "1"

    Application.ScreenUpdating = False

    '--- Find min and max ---
    Dim minVal As Double, maxVal As Double
    Dim cell As Range
    Dim firstNum As Boolean
    firstNum = True

    For Each cell In rng.Cells
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            Dim v As Double
            v = CDbl(cell.Value)
            If firstNum Then
                minVal = v
                maxVal = v
                firstNum = False
            Else
                If v < minVal Then minVal = v
                If v > maxVal Then maxVal = v
            End If
        End If
    Next cell

    If firstNum Then
        Application.ScreenUpdating = True
        MsgBox "No numeric values found.", vbInformation, "Color Scale"
        Exit Sub
    End If

    Dim range_val As Double
    range_val = maxVal - minVal
    If range_val = 0 Then range_val = 1  ' Avoid division by zero

    '--- Apply gradient ---
    Dim colorCount As Long
    colorCount = 0

    For Each cell In rng.Cells
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            Dim pct As Double
            pct = (CDbl(cell.Value) - minVal) / range_val  ' 0 to 1

            If Trim(dirChoice) = "2" Then pct = 1 - pct  ' Reverse

            ' Red (0) -> Yellow (0.5) -> Green (1)
            Dim r As Long, g As Long, b As Long
            If pct < 0.5 Then
                ' Red to Yellow
                r = 255
                g = CLng(255 * (pct / 0.5))
                b = 0
            Else
                ' Yellow to Green
                r = CLng(255 * ((1 - pct) / 0.5))
                g = 200
                b = 0
            End If

            ' Lighten the colors for readability
            r = r + CLng((255 - r) * 0.3)
            g = g + CLng((255 - g) * 0.3)
            b = b + CLng((255 - b) * 0.3)

            If r > 255 Then r = 255
            If g > 255 Then g = 255
            If b > 255 Then b = 255

            cell.Interior.Color = RGB(r, g, b)
            colorCount = colorCount + 1
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Color scale applied to " & colorCount & " cells." & vbCrLf & vbCrLf & _
           "Range: " & Format(minVal, "#,##0.00") & " to " & Format(maxVal, "#,##0.00"), _
           vbInformation, "Color Scale"

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Color Scale"
End Sub

'==============================================================================
' PUBLIC: ClearHighlights
' Remove cell highlighting from selection, sheet, or all sheets.
'==============================================================================
Public Sub ClearHighlights()
    On Error GoTo ErrHandler

    Dim choice As String
    choice = InputBox("Clear highlighting from:" & vbCrLf & vbCrLf & _
                       "  1. Current selection only" & vbCrLf & _
                       "  2. Active sheet (all cells)" & vbCrLf & _
                       "  3. ALL sheets in workbook" & vbCrLf & vbCrLf & _
                       "Enter number:", _
                       "Clear Highlights")
    If Len(Trim(choice)) = 0 Then Exit Sub

    Application.ScreenUpdating = False

    Select Case Trim(choice)
        Case "1"
            If Not Selection Is Nothing Then
                Selection.Interior.ColorIndex = xlNone
            End If
            Application.ScreenUpdating = True
            MsgBox "Highlights cleared from selection.", vbInformation, "Clear Highlights"

        Case "2"
            ActiveSheet.Cells.Interior.ColorIndex = xlNone
            Application.ScreenUpdating = True
            MsgBox "Highlights cleared from " & ActiveSheet.Name & ".", vbInformation, "Clear Highlights"

        Case "3"
            Dim confirm As VbMsgBoxResult
            confirm = MsgBox("Clear ALL highlighting from ALL sheets?" & vbCrLf & vbCrLf & _
                              "This removes all cell background colors.", _
                              vbYesNo + vbExclamation, "Clear Highlights")
            If confirm = vbYes Then
                Dim ws As Worksheet
                For Each ws In ThisWorkbook.Worksheets
                    ws.Cells.Interior.ColorIndex = xlNone
                Next ws
                Application.ScreenUpdating = True
                MsgBox "Highlights cleared from all sheets.", vbInformation, "Clear Highlights"
            Else
                Application.ScreenUpdating = True
            End If

        Case Else
            Application.ScreenUpdating = True
            MsgBox "Invalid choice.", vbExclamation, "Clear Highlights"
    End Select

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Clear Highlights"
End Sub

'==============================================================================
' PRIVATE: SortArray — Simple bubble sort for Double array (ascending)
'==============================================================================
Private Sub SortArray(ByRef arr() As Double, ByVal count As Long)
    Dim i As Long, j As Long
    Dim temp As Double

    For i = 1 To count - 1
        For j = 1 To count - i
            If arr(j) > arr(j + 1) Then
                temp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = temp
            End If
        Next j
    Next i
End Sub
