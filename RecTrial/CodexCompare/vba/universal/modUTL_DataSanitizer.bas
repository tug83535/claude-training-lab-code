Attribute VB_Name = "modUTL_DataSanitizer"
Option Explicit

Private Type SanitizerCounters
    CellsScanned As Double
    CellsChanged As Double
    NumbersConverted As Double
    DatesNormalized As Double
    WhitespaceTrimmed As Double
    FloatTailsFixed As Double
End Type

Public Sub RunFullSanitize(Optional ByVal IncludeHidden As Boolean = False)
    Dim targets As Collection
    Dim ws As Worksheet
    Dim stats As SanitizerCounters
    Dim sheetCount As Long

    On Error GoTo CleanFail

    Set targets = UTL_GetTargetSheets(IncludeHidden)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    For Each ws In targets
        If ws.Name <> "UTL_RunLog" Then
            If IncludeHidden Or ws.Visible = xlSheetVisible Then
                SanitizeWorksheet ws, stats
                sheetCount = sheetCount + 1
            End If
        End If
    Next ws

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    UTL_LogAction "modUTL_DataSanitizer", "RunFullSanitize", "PASS", _
                  "Sanitization complete", sheetCount, stats.CellsChanged

    UTL_ShowCompletion "Universal Data Sanitizer", _
        "Completed. Sheets scanned: " & sheetCount & _
        " | Cells changed: " & Format$(stats.CellsChanged, "#,##0") & _
        " | Numbers fixed: " & Format$(stats.NumbersConverted, "#,##0") & _
        " | Dates normalized: " & Format$(stats.DatesNormalized, "#,##0")
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    UTL_LogAction "modUTL_DataSanitizer", "RunFullSanitize", "FAIL", Err.Description
    MsgBox "Sanitizer stopped: " & Err.Description, vbExclamation, "Universal Data Sanitizer"
End Sub

Public Sub PreviewSanitizeChanges(Optional ByVal IncludeHidden As Boolean = False)
    Dim targets As Collection
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim c As Range
    Dim candidateCount As Double
    Dim sheetCount As Long

    On Error GoTo PreviewFail

    Set targets = UTL_GetTargetSheets(IncludeHidden)

    For Each ws In targets
        If ws.Name <> "UTL_RunLog" Then
            Set dataRange = UTL_DetectDataRange(ws)
            For Each c In dataRange.Cells
                If Len(CStr(c.Value2)) > 0 Then
                    If ShouldTrimWhitespace(CStr(c.Value2)) Or IsTextNumber(CStr(c.Value2)) Or IsDateString(CStr(c.Value2)) Or HasFloatingTail(c.Value2) Then
                        candidateCount = candidateCount + 1
                    End If
                End If
            Next c
            sheetCount = sheetCount + 1
        End If
    Next ws

    UTL_LogAction "modUTL_DataSanitizer", "PreviewSanitizeChanges", "PASS", _
                  "Preview complete", sheetCount, candidateCount

    UTL_ShowCompletion "Sanitizer Preview", "Potential fixes found: " & Format$(candidateCount, "#,##0") & " across " & sheetCount & " sheet(s)."
    Exit Sub

PreviewFail:
    UTL_LogAction "modUTL_DataSanitizer", "PreviewSanitizeChanges", "FAIL", Err.Description
    MsgBox "Preview failed: " & Err.Description, vbExclamation, "Sanitizer Preview"
End Sub

Private Sub SanitizeWorksheet(ByVal ws As Worksheet, ByRef stats As SanitizerCounters)
    Dim dataRange As Range
    Dim c As Range
    Dim originalText As String
    Dim cleanedText As String
    Dim changed As Boolean

    Set dataRange = UTL_DetectDataRange(ws)

    For Each c In dataRange.Cells
        stats.CellsScanned = stats.CellsScanned + 1

        If Len(CStr(c.Value2)) = 0 Then GoTo NextCell

        changed = False

        If VarType(c.Value2) = vbString Then
            originalText = CStr(c.Value2)
            cleanedText = Trim$(Replace(Replace(originalText, vbCr, " "), vbLf, " "))

            If cleanedText <> originalText Then
                c.Value = cleanedText
                changed = True
                stats.WhitespaceTrimmed = stats.WhitespaceTrimmed + 1
            End If

            If IsTextNumber(cleanedText) Then
                c.Value = CDbl(Replace(cleanedText, ",", ""))
                changed = True
                stats.NumbersConverted = stats.NumbersConverted + 1
            ElseIf IsDateString(cleanedText) Then
                c.Value = CDate(cleanedText)
                c.NumberFormat = "yyyy-mm-dd"
                changed = True
                stats.DatesNormalized = stats.DatesNormalized + 1
            End If
        ElseIf HasFloatingTail(c.Value2) Then
            c.Value = Round(CDbl(c.Value2), 6)
            changed = True
            stats.FloatTailsFixed = stats.FloatTailsFixed + 1
        End If

        If changed Then stats.CellsChanged = stats.CellsChanged + 1
NextCell:
    Next c
End Sub

Private Function IsTextNumber(ByVal inputText As String) As Boolean
    Dim candidate As String

    candidate = Replace(Trim$(inputText), ",", "")
    If Len(candidate) = 0 Then Exit Function

    If Left$(candidate, 1) = "$" Then candidate = Mid$(candidate, 2)
    If Right$(candidate, 1) = "%" Then Exit Function

    IsTextNumber = IsNumeric(candidate)
End Function

Private Function IsDateString(ByVal inputText As String) As Boolean
    Dim candidate As String

    candidate = Trim$(inputText)
    If Len(candidate) < 6 Then Exit Function

    IsDateString = IsDate(candidate)
End Function

Private Function HasFloatingTail(ByVal valueIn As Variant) As Boolean
    If IsNumeric(valueIn) Then
        HasFloatingTail = (Abs(CDbl(valueIn) - Round(CDbl(valueIn), 6)) > 0.0000001)
    End If
End Function

Private Function ShouldTrimWhitespace(ByVal inputText As String) As Boolean
    ShouldTrimWhitespace = (Trim$(Replace(Replace(inputText, vbCr, " "), vbLf, " ")) <> inputText)
End Function
