Attribute VB_Name = "modUTL_DataSanitizer"
Option Explicit

' ============================================================
' KBT Universal Tools — Enhanced Data Sanitizer Module
' Works on ANY Excel file — no project-specific setup required
' Install in Personal.xlsb to use across all Excel sessions
' Tools: 4 | Tier 1: 4
' ============================================================
' Ported from the project-specific modDataSanitizer, made
' universal (no references to modConfig, modPerformance, etc.)
'
' Three types of number problems fixed:
'   1. TEXT-STORED NUMBERS  - cells like "0.22" stored as text
'   2. FLOATING-POINT TAILS - cells like 9412.300000000001
'   3. INTEGER FORMATS      - cells like 3628 missing 2dp display
'
' Smart detection skips dates, names, IDs, labels, formulas.
' ============================================================

' Minimum decimal places before flagging as floating-point tail
Private Const FP_DECIMAL_THRESHOLD As Long = 5

' Column header keywords that mean "do NOT sanitize this column"
Private Const SKIP_KEYWORDS As String = _
    "id,date,name,code,ref,no.,#,uuid,email,phone,zip,customer,client," & _
    "account,acct,company,vendor,contact,employee,user,member,entity," & _
    "description,desc,category,dept,department,product,type,status," & _
    "label,title,region,country,state,city,address"

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
' TOOL 1 — Full Sanitize (all 3 fixes in order)    [TIER 1]
' Runs: text-numbers -> floating-point -> integer format
' ============================================================
Sub RunFullSanitize()
    On Error GoTo ErrHandler

    If MsgBox("Run full numeric sanitizer on this workbook?" & Chr(10) & Chr(10) & _
              "This will fix:" & Chr(10) & _
              "  1. Text-stored numbers (""0.22"" as text)" & Chr(10) & _
              "  2. Floating-point tails (9412.300000001)" & Chr(10) & _
              "  3. Integer formats (3628 without .00)" & Chr(10) & Chr(10) & _
              "Dates, names, labels, IDs, and formulas are NOT touched." & Chr(10) & _
              "A backup of each sheet will be created first." & Chr(10) & _
              "Tip: Run Preview first to see what will change.", _
              vbYesNo + vbQuestion, "UTL Data Sanitizer") = vbNo Then Exit Sub

    UTL_TurboOn

    ' Create backup of all sheets before sanitization
    Dim bkWs As Worksheet
    For Each bkWs In ActiveWorkbook.Worksheets
        modUTL_Core.UTL_BackupSheet bkWs
    Next bkWs

    Dim t1 As Long
    Dim t2 As Long
    Dim t3 As Long
    t1 = InternalConvertTextNumbers()
    t2 = InternalFixFloatingPoint()
    t3 = InternalNormalizeIntegers()

    UTL_TurboOff

    MsgBox "Sanitizer complete!" & Chr(10) & Chr(10) & _
           "  Text-stored numbers converted:  " & t1 & Chr(10) & _
           "  Floating-point tails rounded:   " & t2 & Chr(10) & _
           "  Integer formats normalized:     " & t3 & Chr(10) & Chr(10) & _
           "All dates, names, labels, and IDs were left unchanged.", _
           vbInformation, "UTL Data Sanitizer"
    Exit Sub

ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Data Sanitizer"
End Sub

' ============================================================
' TOOL 2 — Preview (dry run, no changes)            [TIER 1]
' Scans the workbook and reports what WOULD change
' Creates a "UTL_Sanitizer_Preview" report sheet
' ============================================================
Sub PreviewSanitizeChanges()
    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim rptName As String
    rptName = "UTL_Sanitizer_Preview"

    ' Delete old preview sheet if exists
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Sheets(rptName).Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    Dim wsRpt As Worksheet
    Set wsRpt = ActiveWorkbook.Sheets.Add( _
        After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    wsRpt.Name = rptName

    ' Headers
    With wsRpt
        .Range("A1").Value = "Sanitizer Preview — " & ActiveWorkbook.Name
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A3").Value = "Sheet"
        .Range("B3").Value = "Cell"
        .Range("C3").Value = "Issue Type"
        .Range("D3").Value = "Current Value"
        .Range("E3").Value = "Proposed Value"
        .Range("F3").Value = "Reason"
        .Range("A3:F3").Font.Bold = True
        .Range("A3:F3").Interior.Color = RGB(11, 71, 121)
        .Range("A3:F3").Font.Color = RGB(255, 255, 255)
    End With

    Dim outRow As Long
    outRow = 4

    Dim usedRng As Range
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name = rptName Then GoTo NextWS
        If ws.Visible = xlSheetVeryHidden Then GoTo NextWS

        Set usedRng = Nothing
        On Error Resume Next
        Set usedRng = ws.UsedRange
        On Error GoTo ErrHandler
        If usedRng Is Nothing Then GoTo NextWS

        Dim cell As Range
        For Each cell In usedRng
            If IsEmpty(cell.Value) Then GoTo NextCell
            If cell.HasFormula Then GoTo NextCell

            Dim colHdr As String
            colHdr = GetColHeader(ws, cell.Column)

            ' --- Check 1: Text-stored number ---
            If VarType(cell.Value) = vbString Then
                Dim strVal As String
                strVal = Trim(CStr(cell.Value))
                If IsNumeric(strVal) And Len(strVal) > 0 Then
                    If Not IsSkippedCol(colHdr) And Not IsDateStr(strVal) Then
                        wsRpt.Cells(outRow, 1).Value = ws.Name
                        wsRpt.Cells(outRow, 2).Value = cell.Address
                        wsRpt.Cells(outRow, 3).Value = "Text-Stored Number"
                        wsRpt.Cells(outRow, 4).Value = """" & strVal & """"
                        wsRpt.Cells(outRow, 5).Value = CDbl(strVal)
                        wsRpt.Cells(outRow, 6).Value = "Text string passes IsNumeric()"
                        wsRpt.Cells(outRow, 3).Interior.Color = RGB(255, 200, 200)
                        outRow = outRow + 1
                    End If
                End If
                GoTo NextCell
            End If

            ' --- Only numeric non-string values from here ---
            If Not IsNumeric(cell.Value) Then GoTo NextCell
            If IsDateVal(cell) Then GoTo NextCell
            If IsSkippedCol(colHdr) Then GoTo NextCell

            Dim numVal As Double
            numVal = CDbl(cell.Value)

            ' --- Check 2: Floating-point tail ---
            Dim strNum As String
            strNum = CStr(numVal)
            Dim dotPos As Long
            dotPos = InStr(strNum, ".")
            If dotPos > 0 Then
                Dim decPart As String
                decPart = Mid(strNum, dotPos + 1)
                If Len(decPart) >= FP_DECIMAL_THRESHOLD Then
                    Dim targetDP As Long
                    targetDP = IIf(InStr(cell.NumberFormat, "%") > 0, 4, 2)
                    Dim rounded As Double
                    rounded = Round(numVal, targetDP)
                    If Abs(rounded - numVal) > 0 And Abs(rounded - numVal) < 0.001 Then
                        wsRpt.Cells(outRow, 1).Value = ws.Name
                        wsRpt.Cells(outRow, 2).Value = cell.Address
                        wsRpt.Cells(outRow, 3).Value = "Floating-Point Tail"
                        wsRpt.Cells(outRow, 4).Value = numVal
                        wsRpt.Cells(outRow, 5).Value = rounded
                        wsRpt.Cells(outRow, 6).Value = Len(decPart) & " decimals -> round to " & targetDP & "dp"
                        wsRpt.Cells(outRow, 3).Interior.Color = RGB(255, 235, 180)
                        outRow = outRow + 1
                    End If
                End If
            End If

            ' --- Check 3: Integer that should show 2dp ---
            If numVal = Int(numVal) And Abs(numVal) >= 100 Then
                Dim fmt As String
                fmt = cell.NumberFormat
                If InStr(fmt, ".0") = 0 And InStr(fmt, "#,##0.") = 0 Then
                    If InStr(fmt, "%") = 0 Then
                        wsRpt.Cells(outRow, 1).Value = ws.Name
                        wsRpt.Cells(outRow, 2).Value = cell.Address
                        wsRpt.Cells(outRow, 3).Value = "Integer Format"
                        wsRpt.Cells(outRow, 4).Value = numVal
                        wsRpt.Cells(outRow, 5).Value = Format(numVal, "#,##0.00")
                        wsRpt.Cells(outRow, 6).Value = "Format only - value unchanged"
                        wsRpt.Cells(outRow, 3).Interior.Color = RGB(220, 240, 255)
                        outRow = outRow + 1
                    End If
                End If
            End If

NextCell:
        Next cell
NextWS:
    Next ws

    wsRpt.Columns("A:F").AutoFit
    wsRpt.Activate
    UTL_TurboOff

    Dim issueCount As Long
    issueCount = outRow - 4

    If issueCount = 0 Then
        MsgBox "No numeric issues found. The workbook is already clean.", _
               vbInformation, "UTL Data Sanitizer"
    Else
        MsgBox issueCount & " potential fix(es) found." & Chr(10) & _
               "See '" & rptName & "' for full details." & Chr(10) & Chr(10) & _
               "Red    = Text-stored numbers" & Chr(10) & _
               "Yellow = Floating-point tails" & Chr(10) & _
               "Blue   = Integer format only (value not changed)" & Chr(10) & Chr(10) & _
               "When ready, run RunFullSanitize to apply all fixes.", _
               vbExclamation, "UTL Data Sanitizer"
    End If
    Exit Sub

ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Data Sanitizer"
End Sub

' ============================================================
' TOOL 3 — Fix Floating-Point Tails Only           [TIER 1]
' Rounds cells with excessive decimal digits (FP noise)
' ============================================================
Sub FixFloatingPointTails()
    On Error GoTo ErrHandler
    UTL_TurboOn
    Dim fixCount As Long
    fixCount = InternalFixFloatingPoint()
    UTL_TurboOff
    MsgBox fixCount & " floating-point tail(s) rounded to clean values." & Chr(10) & _
           "Dates, labels, IDs, and formulas were not touched.", _
           vbInformation, "UTL Data Sanitizer"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Data Sanitizer"
End Sub

' ============================================================
' TOOL 4 — Convert Text-Stored Numbers Only        [TIER 1]
' Converts text strings that look like numbers to real numbers
' ============================================================
Sub ConvertTextStoredNumbers()
    On Error GoTo ErrHandler
    UTL_TurboOn
    Dim fixCount As Long
    fixCount = InternalConvertTextNumbers()
    UTL_TurboOff
    MsgBox fixCount & " text-stored number(s) converted." & Chr(10) & _
           "Names, dates, IDs, and labels were not touched.", _
           vbInformation, "UTL Data Sanitizer"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Data Sanitizer"
End Sub


' ============================================================
'  PRIVATE WORKER FUNCTIONS
' ============================================================

Private Function InternalConvertTextNumbers() As Long
    Dim fixCount As Long
    Dim ws As Worksheet

    Dim rng As Range
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVeryHidden Then GoTo NextWS2

        Set rng = Nothing
        On Error Resume Next
        Set rng = ws.UsedRange.SpecialCells(xlCellTypeConstants, xlTextValues)
        On Error GoTo 0
        If rng Is Nothing Then GoTo NextWS2

        Dim cell As Range
        For Each cell In rng
            If cell.HasFormula Then GoTo NextCell2

            Dim colHdr As String
            colHdr = GetColHeader(ws, cell.Column)
            If IsSkippedCol(colHdr) Then GoTo NextCell2

            Dim strVal As String
            strVal = Trim(CStr(cell.Value))
            If Len(strVal) = 0 Then GoTo NextCell2
            If Not IsNumeric(strVal) Then GoTo NextCell2
            If IsDateStr(strVal) Then GoTo NextCell2

            Dim numVal As Double
            numVal = CDbl(strVal)
            cell.Value = numVal

            ' Apply appropriate format
            If InStr(cell.NumberFormat, "%") > 0 Then
                cell.NumberFormat = "0.0%"
            ElseIf Abs(numVal) >= 100 Then
                cell.NumberFormat = "#,##0.00"
            Else
                cell.NumberFormat = "0.0000"
            End If
            fixCount = fixCount + 1
NextCell2:
        Next cell
NextWS2:
    Next ws
    InternalConvertTextNumbers = fixCount
End Function

Private Function InternalFixFloatingPoint() As Long
    Dim fixCount As Long
    Dim ws As Worksheet

    Dim rng As Range
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVeryHidden Then GoTo NextWS3

        Set rng = Nothing
        On Error Resume Next
        Set rng = ws.UsedRange.SpecialCells(xlCellTypeConstants, xlNumbers)
        On Error GoTo 0
        If rng Is Nothing Then GoTo NextWS3

        Dim cell As Range
        For Each cell In rng
            If cell.HasFormula Then GoTo NextCell3
            If IsDateVal(cell) Then GoTo NextCell3

            Dim colHdr As String
            colHdr = GetColHeader(ws, cell.Column)
            If IsSkippedCol(colHdr) Then GoTo NextCell3

            Dim numVal As Double
            numVal = CDbl(cell.Value)
            Dim strNum As String
            strNum = CStr(numVal)
            Dim dotPos As Long
            dotPos = InStr(strNum, ".")
            If dotPos = 0 Then GoTo NextCell3

            Dim decPart As String
            decPart = Mid(strNum, dotPos + 1)
            If Len(decPart) < FP_DECIMAL_THRESHOLD Then GoTo NextCell3

            Dim targetDP As Long
            targetDP = IIf(InStr(cell.NumberFormat, "%") > 0, 4, 2)
            Dim rounded As Double
            rounded = Round(numVal, targetDP)

            If Abs(rounded - numVal) > 0 And Abs(rounded - numVal) < 0.001 Then
                cell.Value = rounded
                fixCount = fixCount + 1
            End If
NextCell3:
        Next cell
NextWS3:
    Next ws
    InternalFixFloatingPoint = fixCount
End Function

Private Function InternalNormalizeIntegers() As Long
    Dim fixCount As Long
    Dim ws As Worksheet

    Dim rng As Range
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVeryHidden Then GoTo NextWS4

        Set rng = Nothing
        On Error Resume Next
        Set rng = ws.UsedRange.SpecialCells(xlCellTypeConstants, xlNumbers)
        On Error GoTo 0
        If rng Is Nothing Then GoTo NextWS4

        Dim cell As Range
        For Each cell In rng
            If cell.HasFormula Then GoTo NextCell4
            If IsDateVal(cell) Then GoTo NextCell4

            Dim colHdr As String
            colHdr = GetColHeader(ws, cell.Column)
            If IsSkippedCol(colHdr) Then GoTo NextCell4

            Dim numVal As Double
            numVal = CDbl(cell.Value)
            If numVal <> Int(numVal) Then GoTo NextCell4
            If Abs(numVal) < 100 Then GoTo NextCell4

            Dim fmt As String
            fmt = cell.NumberFormat
            If InStr(fmt, ".0") > 0 Or InStr(fmt, "#,##0.") > 0 Then GoTo NextCell4
            If InStr(fmt, "%") > 0 Then GoTo NextCell4

            cell.NumberFormat = "#,##0.00"
            fixCount = fixCount + 1
NextCell4:
        Next cell
NextWS4:
    Next ws
    InternalNormalizeIntegers = fixCount
End Function


' ============================================================
'  PRIVATE DETECTION HELPERS
' ============================================================

Private Function IsDateVal(ByVal cell As Range) As Boolean
    IsDateVal = False
    On Error Resume Next

    Dim fmt As String
    fmt = LCase(cell.NumberFormat)
    If InStr(fmt, "yy") > 0 Or InStr(fmt, "yyyy") > 0 Then
        IsDateVal = True
        Exit Function
    End If
    If (InStr(fmt, "dd") > 0 Or InStr(fmt, "d/") > 0) And _
       (InStr(fmt, "mm") > 0 Or InStr(fmt, "m/") > 0) Then
        IsDateVal = True
        Exit Function
    End If

    If IsDate(cell.Value) And IsNumeric(cell.Value) Then
        Dim v As Double
        v = CDbl(cell.Value)
        If v >= 30000 And v <= 60000 Then IsDateVal = True
    End If
    On Error GoTo 0
End Function

Private Function IsDateStr(ByVal s As String) As Boolean
    IsDateStr = False
    On Error Resume Next
    If IsDate(s) Then
        If InStr(s, "/") > 0 Or InStr(s, "-") > 0 Then
            IsDateStr = True
        End If
    End If
    On Error GoTo 0
End Function

Private Function GetColHeader(ByVal ws As Worksheet, ByVal colNum As Long) As String
    GetColHeader = ""
    Dim r As Long
    For r = 1 To 6
        Dim v As String
        v = Trim(CStr(ws.Cells(r, colNum).Value))
        If Len(v) > 0 And Not IsNumeric(v) Then
            GetColHeader = LCase(v)
            Exit Function
        End If
    Next r
End Function

Private Function IsSkippedCol(ByVal headerText As String) As Boolean
    IsSkippedCol = False
    If Len(headerText) = 0 Then Exit Function
    Dim keywords As Variant
    keywords = Split(SKIP_KEYWORDS, ",")
    Dim kw As Variant
    For Each kw In keywords
        If InStr(headerText, Trim(CStr(kw))) > 0 Then
            IsSkippedCol = True
            Exit Function
        End If
    Next kw
End Function

'==============================================================================
' DIRECTOR-ONLY: Silent wrappers for automated recording (no dialogs)
'==============================================================================
Sub DirectorRunFullSanitize()
    On Error Resume Next
    UTL_TurboOn
    Dim t1 As Long, t2 As Long, t3 As Long
    t1 = InternalConvertTextNumbers()
    t2 = InternalFixFloatingPoint()
    t3 = InternalNormalizeIntegers()
    UTL_TurboOff
    Debug.Print "[Director] Sanitize: " & t1 & " text, " & t2 & " FP, " & t3 & " int"
End Sub

Sub DirectorPreviewSanitize()
    ' Replicates PreviewSanitizeChanges core logic without any MsgBox
    On Error Resume Next
    UTL_TurboOn

    Dim rptName As String: rptName = "UTL_Sanitizer_Preview"
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Sheets(rptName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Dim wsRpt As Worksheet
    Set wsRpt = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    wsRpt.Name = rptName

    With wsRpt
        .Range("A1").Value = "Sanitizer Preview - " & ActiveWorkbook.Name
        .Range("A1").Font.Bold = True: .Range("A1").Font.Size = 14
        .Range("A3").Value = "Sheet": .Range("B3").Value = "Cell"
        .Range("C3").Value = "Issue Type": .Range("D3").Value = "Current Value"
        .Range("E3").Value = "Proposed Value": .Range("F3").Value = "Reason"
        .Range("A3:F3").Font.Bold = True
        .Range("A3:F3").Interior.Color = RGB(11, 71, 121)
        .Range("A3:F3").Font.Color = RGB(255, 255, 255)
    End With

    Dim outRow As Long: outRow = 4
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name = rptName Then GoTo NextPrevWS
        If ws.Visible = xlSheetVeryHidden Then GoTo NextPrevWS
        Dim usedRng As Range: Set usedRng = Nothing
        On Error Resume Next
        Set usedRng = ws.UsedRange
        On Error GoTo 0
        If usedRng Is Nothing Then GoTo NextPrevWS
        Dim cell As Range
        For Each cell In usedRng.Cells
            If Not IsEmpty(cell.Value) And Not cell.HasFormula Then
                If cell.NumberFormat = "@" And IsNumeric(cell.Value) Then
                    wsRpt.Cells(outRow, 1).Value = ws.Name
                    wsRpt.Cells(outRow, 2).Value = cell.Address
                    wsRpt.Cells(outRow, 3).Value = "Text-Stored Number"
                    wsRpt.Cells(outRow, 4).Value = cell.Value
                    wsRpt.Cells(outRow, 5).Value = CDbl(cell.Value)
                    wsRpt.Cells(outRow, 6).Value = "Text string passes IsNumeric()"
                    wsRpt.Cells(outRow, 3).Font.Color = RGB(192, 0, 0)
                    outRow = outRow + 1
                End If
            End If
        Next cell
NextPrevWS:
    Next ws

    wsRpt.Columns("A:F").AutoFit
    wsRpt.Activate
    UTL_TurboOff
    Debug.Print "[Director] Preview: " & (outRow - 4) & " issues found"
End Sub
