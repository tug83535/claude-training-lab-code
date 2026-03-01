Attribute VB_Name = "modDataSanitizer"
Option Explicit

'===============================================================================
' modDataSanitizer - Numeric-Only Data Sanitizer
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Cleans up numeric data issues WITHOUT touching dates, names,
'           labels, IDs, or any text content. Safe to run on the full workbook.
'
'           Three types of problems are fixed:
'             1. TEXT-STORED NUMBERS  — cells that look like "0.22" but are
'                stored as text strings. Causes #VALUE! errors in formulas.
'             2. FLOATING-POINT TAILS — cells like 9412.300000000001 caused
'                by Excel's IEEE 754 binary math. Rounded to clean values.
'             3. INTEGER AMOUNTS      — cells like 3628 that should show as
'                3628.00. Number format is applied; the value is unchanged.
'
' HOW IT DECIDES WHAT IS A NUMBER (vs. a date, name, or label):
'   - Cell value must pass IsNumeric() AND not be a string
'   - Cell NumberFormat must NOT contain date indicators (y, d combined with m)
'   - Cell value must NOT fall in Excel date serial range with a date-ish format
'   - Column header must NOT contain keywords: ID, Date, Name, Code, Ref, No.
'
' PUBLIC SUBS:
'   RunFullSanitize          - Master runner: all 3 fixes in one click
'   PreviewSanitizeChanges   - Dry-run report — shows what WOULD change, no edits
'   FixFloatingPointTails    - Fix floating-point noise on the active sheet
'   ConvertTextStoredNumbers - Convert text-numbers to real numbers (all sheets)
'   NormalizeIntegerFormats  - Apply 2dp format to whole-number amounts in GL
'
' VERSION:  2.1.0 (New module — 2026-03-01)
'===============================================================================

'===============================================================================
' PRIVATE CONSTANTS
'===============================================================================
' Minimum decimal places before a value is considered a floating-point tail
Private Const FP_DECIMAL_THRESHOLD As Long = 5

' Keywords in column headers that mean "do NOT sanitize this column"
' Covers IDs, dates, customer/client identifiers, and reference fields
Private Const SKIP_HEADER_KEYWORDS As String = "id,date,name,code,ref,no.,#,uuid,email,phone,zip,customer,client,account,acct,company,vendor,contact,employee,user,member,entity,description,desc,category,dept,department,product,type,status,label,title,region,country,state,city,address"

' Sheets to skip entirely
Private Const SKIP_SHEET_NAMES As String = "VBA_AuditLog,GoldenBaseline,Recon Archive"

'===============================================================================
' RunFullSanitize - Master runner: all 3 fixes in sequence (#67 ref)
' Asks for confirmation, then runs all three sanitizers in the correct order:
'   1. ConvertTextStoredNumbers  (must be first — makes floats available for step 2)
'   2. FixFloatingPointTails     (needs real numbers to work)
'   3. NormalizeIntegerFormats   (cosmetic — always last)
' Writes a summary to the VBA_AuditLog.
'===============================================================================
Public Sub RunFullSanitize()
    On Error GoTo ErrHandler

    Dim msg As String
    msg = "Run full numeric sanitizer on this workbook?" & vbCrLf & vbCrLf & _
          "This will fix three types of number problems:" & vbCrLf & _
          "  1. Text-stored numbers  (e.g., ""0.22"" stored as text)" & vbCrLf & _
          "  2. Floating-point tails (e.g., 9412.300000000001)" & vbCrLf & _
          "  3. Integer formats      (e.g., 3628 displayed without cents)" & vbCrLf & vbCrLf & _
          "Dates, names, labels, IDs, and formulas are NOT touched." & vbCrLf & vbCrLf & _
          "Tip: Run PreviewSanitizeChanges first to see exactly what will change."

    If MsgBox(msg, vbYesNo + vbQuestion, APP_NAME) = vbNo Then Exit Sub

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Sanitizing: converting text numbers...", 0.1

    Dim t1 As Long: Dim t2 As Long: Dim t3 As Long
    t1 = InternalConvertTextNumbers(False)

    modPerformance.UpdateStatus "Sanitizing: fixing floating-point tails...", 0.45
    t2 = InternalFixFloatingPoint(False)

    modPerformance.UpdateStatus "Sanitizing: normalizing integer formats...", 0.75
    t3 = InternalNormalizeIntegers(False)

    modPerformance.TurboOff

    modLogger.LogAction "modDataSanitizer", "RunFullSanitize", _
        "Text-numbers fixed: " & t1 & " | FP tails fixed: " & t2 & " | Integer formats: " & t3

    MsgBox "Numeric sanitizer complete." & vbCrLf & vbCrLf & _
           "  Text-stored numbers converted:  " & t1 & vbCrLf & _
           "  Floating-point tails rounded:   " & t2 & vbCrLf & _
           "  Integer formats normalized:     " & t3 & vbCrLf & vbCrLf & _
           "All dates, names, labels, and IDs were left unchanged.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    MsgBox "RunFullSanitize error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' PreviewSanitizeChanges - Dry-run report showing what WOULD change (#118 ref)
' Scans the workbook exactly like RunFullSanitize but makes NO changes.
' Writes a "Sanitizer Preview" sheet listing every cell that would be fixed,
' with its current value and what the new value would be.
' Run this before committing to any changes.
'===============================================================================
Public Sub PreviewSanitizeChanges()
    On Error GoTo ErrHandler

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Scanning workbook (preview only — no changes)...", 0.1

    Dim rptName As String: rptName = "Sanitizer Preview"
    modConfig.SafeDeleteSheet rptName
    Dim wsRpt As Worksheet
    Set wsRpt = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsRpt.Name = rptName

    modConfig.StyleHeader wsRpt, 1, _
        Array("Sheet", "Cell", "Issue Type", "Current Value", "Proposed Value", "Reason")

    Dim outRow As Long: outRow = 2

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsSkippedSheet(ws.Name) Then GoTo NextWS
        If ws.Visible = xlSheetVeryHidden Then GoTo NextWS

        Dim usedRng As Range
        On Error Resume Next
        Set usedRng = ws.UsedRange
        On Error GoTo ErrHandler
        If usedRng Is Nothing Then GoTo NextWS

        Dim cell As Range
        For Each cell In usedRng

            ' ── Skip non-candidate cells ─────────────────────────────────
            If IsEmpty(cell.Value) Then GoTo NextCell
            If cell.HasFormula Then GoTo NextCell

            Dim colHdr As String: colHdr = GetColumnHeader(ws, cell.Column)

            ' ── Check 1: Text-stored number ──────────────────────────────
            If VarType(cell.Value) = vbString Then
                Dim strVal As String: strVal = Trim(CStr(cell.Value))
                If IsNumeric(strVal) And Len(strVal) > 0 Then
                    If Not IsSkippedHeader(colHdr) Then
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

            ' ── Only numeric (non-string) values from here ────────────────
            If Not IsNumeric(cell.Value) Then GoTo NextCell
            If IsDateCell(cell) Then GoTo NextCell
            If IsSkippedHeader(colHdr) Then GoTo NextCell

            Dim numVal As Double: numVal = CDbl(cell.Value)

            ' ── Check 2: Floating-point tail ─────────────────────────────
            Dim strNum As String: strNum = CStr(numVal)
            Dim dotPos As Long: dotPos = InStr(strNum, ".")
            If dotPos > 0 Then
                Dim decimalPart As String: decimalPart = Mid(strNum, dotPos + 1)
                If Len(decimalPart) >= FP_DECIMAL_THRESHOLD Then
                    Dim targetDP As Long: targetDP = IIf(IsPercentCell(cell), 4, 2)
                    Dim rounded As Double: rounded = Round(numVal, targetDP)
                    If Abs(rounded - numVal) > 0 And Abs(rounded - numVal) < 0.001 Then
                        wsRpt.Cells(outRow, 1).Value = ws.Name
                        wsRpt.Cells(outRow, 2).Value = cell.Address
                        wsRpt.Cells(outRow, 3).Value = "Floating-Point Tail"
                        wsRpt.Cells(outRow, 4).Value = numVal
                        wsRpt.Cells(outRow, 5).Value = rounded
                        wsRpt.Cells(outRow, 6).Value = Len(decimalPart) & " decimal digits → round to " & targetDP & "dp"
                        wsRpt.Cells(outRow, 3).Interior.Color = RGB(255, 235, 180)
                        outRow = outRow + 1
                    End If
                End If
            End If

            ' ── Check 3: Integer that should show 2dp ────────────────────
            If numVal = Int(numVal) And Abs(numVal) >= 100 Then
                Dim fmt As String: fmt = cell.NumberFormat
                ' Only flag if the current format doesn't already show decimals
                If InStr(fmt, ".0") = 0 And InStr(fmt, "#,##0.") = 0 Then
                    If Not IsPercentCell(cell) And Not IsDateCell(cell) Then
                        wsRpt.Cells(outRow, 1).Value = ws.Name
                        wsRpt.Cells(outRow, 2).Value = cell.Address
                        wsRpt.Cells(outRow, 3).Value = "Integer Format"
                        wsRpt.Cells(outRow, 4).Value = numVal
                        wsRpt.Cells(outRow, 5).Value = Format(numVal, "$#,##0.00")
                        wsRpt.Cells(outRow, 6).Value = "Format only — value unchanged"
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
    modPerformance.TurboOff

    Dim issueCount As Long: issueCount = outRow - 2
    modLogger.LogAction "modDataSanitizer", "PreviewSanitizeChanges", _
        issueCount & " issue(s) found (no changes made)"

    If issueCount = 0 Then
        MsgBox "No numeric issues found. The workbook is already clean.", _
               vbInformation, APP_NAME
    Else
        MsgBox issueCount & " potential fix(es) found." & vbCrLf & _
               "See '" & rptName & "' for full details." & vbCrLf & vbCrLf & _
               "Red   = Text-stored numbers" & vbCrLf & _
               "Yellow = Floating-point tails" & vbCrLf & _
               "Blue  = Integer format only (value not changed)" & vbCrLf & vbCrLf & _
               "When ready, run RunFullSanitize to apply all fixes.", _
               vbExclamation, APP_NAME
    End If
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    MsgBox "PreviewSanitizeChanges error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' FixFloatingPointTails - Fix FP noise on every visible sheet
' Public wrapper — calls InternalFixFloatingPoint and reports the result.
'===============================================================================
Public Sub FixFloatingPointTails()
    On Error GoTo ErrHandler
    modPerformance.TurboOn
    Dim fixCount As Long: fixCount = InternalFixFloatingPoint(False)
    modPerformance.TurboOff
    modLogger.LogAction "modDataSanitizer", "FixFloatingPointTails", fixCount & " cell(s) rounded"
    MsgBox fixCount & " floating-point tail(s) rounded to clean values." & vbCrLf & _
           "Dates, labels, IDs, and formulas were not touched.", _
           vbInformation, APP_NAME
    Exit Sub
ErrHandler:
    modPerformance.TurboOff
    MsgBox "FixFloatingPointTails error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' ConvertTextStoredNumbers - Convert text-numbers to real numbers (all sheets)
' Public wrapper — calls InternalConvertTextNumbers and reports the result.
'===============================================================================
Public Sub ConvertTextStoredNumbers()
    On Error GoTo ErrHandler
    modPerformance.TurboOn
    Dim fixCount As Long: fixCount = InternalConvertTextNumbers(False)
    modPerformance.TurboOff
    modLogger.LogAction "modDataSanitizer", "ConvertTextStoredNumbers", fixCount & " cell(s) converted"
    MsgBox fixCount & " text-stored number(s) converted to real numbers." & vbCrLf & _
           "Names, dates, IDs, and labels were not touched.", _
           vbInformation, APP_NAME
    Exit Sub
ErrHandler:
    modPerformance.TurboOff
    MsgBox "ConvertTextStoredNumbers error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' NormalizeIntegerFormats - Apply 2dp display format to whole-number amounts
' Public wrapper — calls InternalNormalizeIntegers and reports the result.
'===============================================================================
Public Sub NormalizeIntegerFormats()
    On Error GoTo ErrHandler
    modPerformance.TurboOn
    Dim fixCount As Long: fixCount = InternalNormalizeIntegers(False)
    modPerformance.TurboOff
    modLogger.LogAction "modDataSanitizer", "NormalizeIntegerFormats", fixCount & " cell(s) formatted"
    MsgBox fixCount & " whole-number amount(s) formatted to show cents (display only)." & vbCrLf & _
           "No values were changed.", _
           vbInformation, APP_NAME
    Exit Sub
ErrHandler:
    modPerformance.TurboOff
    MsgBox "NormalizeIntegerFormats error: " & Err.Description, vbCritical, APP_NAME
End Sub


'===============================================================================
'
' ===  PRIVATE WORKER FUNCTIONS  ===============================================
'
'===============================================================================

'===============================================================================
' InternalConvertTextNumbers
' Scans every non-formula cell. If the cell holds a string that passes
' IsNumeric(), it is converted to a Double and given a number format.
' The column header is checked — columns labelled ID, Name, Code, etc. are skipped.
'===============================================================================
Private Function InternalConvertTextNumbers(ByVal previewOnly As Boolean) As Long
    Dim fixCount As Long: fixCount = 0
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsSkippedSheet(ws.Name) Then GoTo NextWS2
        If ws.Visible = xlSheetVeryHidden Then GoTo NextWS2

        Dim rng As Range
        On Error Resume Next
        Set rng = ws.UsedRange.SpecialCells(xlCellTypeConstants, xlTextValues)
        On Error GoTo 0
        If rng Is Nothing Then GoTo NextWS2

        Dim cell As Range
        For Each cell In rng
            If cell.HasFormula Then GoTo NextCell2
            Dim colHdr As String: colHdr = GetColumnHeader(ws, cell.Column)
            If IsSkippedHeader(colHdr) Then GoTo NextCell2

            Dim strVal As String: strVal = Trim(CStr(cell.Value))
            If Len(strVal) = 0 Then GoTo NextCell2
            If Not IsNumeric(strVal) Then GoTo NextCell2

            ' Additional guard: if it looks like a date string, skip it
            If IsDateString(strVal) Then GoTo NextCell2

            If Not previewOnly Then
                Dim numVal As Double: numVal = CDbl(strVal)
                cell.Value = numVal
                ' Apply appropriate format
                If IsPercentCell(cell) Then
                    cell.NumberFormat = "0.0%"
                ElseIf Abs(numVal) >= 100 Then
                    cell.NumberFormat = "$#,##0.00"
                Else
                    cell.NumberFormat = "0.0000"
                End If
            End If
            fixCount = fixCount + 1
NextCell2:
        Next cell
NextWS2:
    Next ws
    InternalConvertTextNumbers = fixCount
End Function

'===============================================================================
' InternalFixFloatingPoint
' Scans all numeric constant cells. If the raw string representation has
' >= FP_DECIMAL_THRESHOLD digits after the decimal, it is rounded.
' Currency values → 2dp. Percentage cells → 4dp.
' Dates, labels, IDs, and formula cells are always skipped.
'===============================================================================
Private Function InternalFixFloatingPoint(ByVal previewOnly As Boolean) As Long
    Dim fixCount As Long: fixCount = 0
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsSkippedSheet(ws.Name) Then GoTo NextWS3
        If ws.Visible = xlSheetVeryHidden Then GoTo NextWS3

        Dim rng As Range
        On Error Resume Next
        Set rng = ws.UsedRange.SpecialCells(xlCellTypeConstants, xlNumbers)
        On Error GoTo 0
        If rng Is Nothing Then GoTo NextWS3

        Dim cell As Range
        For Each cell In rng
            If cell.HasFormula Then GoTo NextCell3
            If IsDateCell(cell) Then GoTo NextCell3
            Dim colHdr As String: colHdr = GetColumnHeader(ws, cell.Column)
            If IsSkippedHeader(colHdr) Then GoTo NextCell3

            Dim numVal As Double: numVal = CDbl(cell.Value)
            Dim strNum As String: strNum = CStr(numVal)
            Dim dotPos As Long: dotPos = InStr(strNum, ".")
            If dotPos = 0 Then GoTo NextCell3

            Dim decPart As String: decPart = Mid(strNum, dotPos + 1)
            If Len(decPart) < FP_DECIMAL_THRESHOLD Then GoTo NextCell3

            Dim targetDP As Long: targetDP = IIf(IsPercentCell(cell), 4, 2)
            Dim rounded  As Double: rounded = Round(numVal, targetDP)

            ' Only fix if the rounding actually changes the value AND the
            ' change is tiny (genuine FP noise, not a meaningful decimal)
            If Abs(rounded - numVal) > 0 And Abs(rounded - numVal) < 0.001 Then
                If Not previewOnly Then
                    cell.Value = rounded
                End If
                fixCount = fixCount + 1
            End If
NextCell3:
        Next cell
NextWS3:
    Next ws
    InternalFixFloatingPoint = fixCount
End Function

'===============================================================================
' InternalNormalizeIntegers
' Finds numeric constant cells that hold whole numbers >= 100 and whose
' current format does not already show decimal places. Applies $#,##0.00
' (currency) format. The VALUE is never changed — this is display only.
' Percentages and dates are skipped.
'===============================================================================
Private Function InternalNormalizeIntegers(ByVal previewOnly As Boolean) As Long
    Dim fixCount As Long: fixCount = 0

    ' Only run on GL sheet — integer amounts are a GL-specific problem
    If Not modConfig.SheetExists(SH_HIDDEN) Then Exit Function
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_HIDDEN)

    Dim lastRow As Long: lastRow = modConfig.LastRow(ws, COL_GL_ID)
    If lastRow < DATA_ROW_GL Then Exit Function

    ' Only operate on the Amount column (COL_GL_AMOUNT = column G = 7)
    Dim r As Long
    For r = DATA_ROW_GL To lastRow
        Dim cell As Range: Set cell = ws.Cells(r, COL_GL_AMOUNT)
        If IsEmpty(cell.Value) Then GoTo NextCell4
        If Not IsNumeric(cell.Value) Then GoTo NextCell4
        If VarType(cell.Value) = vbString Then GoTo NextCell4
        If IsDateCell(cell) Then GoTo NextCell4

        Dim numVal As Double: numVal = CDbl(cell.Value)
        If numVal <> Int(numVal) Then GoTo NextCell4   ' Already has decimals
        If Abs(numVal) < 100 Then GoTo NextCell4       ' Too small — not a financial amount

        Dim fmt As String: fmt = cell.NumberFormat
        If InStr(fmt, ".0") > 0 Or InStr(fmt, "#,##0.") > 0 Then GoTo NextCell4

        If Not previewOnly Then
            cell.NumberFormat = "$#,##0.00"
        End If
        fixCount = fixCount + 1
NextCell4:
    Next r
    InternalNormalizeIntegers = fixCount
End Function


'===============================================================================
'
' ===  PRIVATE DETECTION HELPERS  ==============================================
'
'===============================================================================

'===============================================================================
' IsDateCell - Returns True if the cell holds a date value
' Uses BOTH the NumberFormat pattern AND a value-range check.
' Excel stores dates as integers (e.g., 45000 = late 2023). If the format
' contains date-indicator characters AND the value is in a plausible date
' serial range, the cell is treated as a date.
'===============================================================================
Private Function IsDateCell(ByVal cell As Range) As Boolean
    IsDateCell = False
    On Error Resume Next

    Dim fmt As String: fmt = LCase(cell.NumberFormat)

    ' Date format indicators: contains year+month, or explicit date patterns
    Dim hasYear  As Boolean: hasYear  = InStr(fmt, "yy") > 0 Or InStr(fmt, "yyyy") > 0
    Dim hasDay   As Boolean: hasDay   = InStr(fmt, "dd") > 0 Or InStr(fmt, "d/") > 0 Or InStr(fmt, "/d") > 0
    Dim hasMonth As Boolean: hasMonth = InStr(fmt, "mm") > 0 Or InStr(fmt, "mmm") > 0 Or _
                                        InStr(fmt, "/m") > 0 Or InStr(fmt, "m/") > 0

    If hasYear Or (hasDay And hasMonth) Then
        IsDateCell = True
        Exit Function
    End If

    ' Secondary check: use VBA's IsDate() on the cell value
    If IsDate(cell.Value) And IsNumeric(cell.Value) Then
        ' Excel date serials: 1 = Jan 1 1900. Plausible range: 30000–60000 (1982–2064)
        Dim v As Double: v = CDbl(cell.Value)
        If v >= 30000 And v <= 60000 Then
            IsDateCell = True
        End If
    End If
    On Error GoTo 0
End Function

'===============================================================================
' IsPercentCell - Returns True if the cell's format is a percentage format
'===============================================================================
Private Function IsPercentCell(ByVal cell As Range) As Boolean
    IsPercentCell = InStr(cell.NumberFormat, "%") > 0
End Function

'===============================================================================
' IsDateString - Returns True if a text string looks like a date
' Guards against converting "7/13/2025" or "2025-12-14" when stored as text.
'===============================================================================
Private Function IsDateString(ByVal s As String) As Boolean
    IsDateString = False
    On Error Resume Next
    ' If VBA can parse it as a date, treat it as one
    If IsDate(s) Then
        ' Extra check: genuine dates have "/" or "-" separators
        If InStr(s, "/") > 0 Or InStr(s, "-") > 0 Then
            IsDateString = True
        End If
    End If
    On Error GoTo 0
End Function

'===============================================================================
' GetColumnHeader - Return the header text for a given column number
' Looks in the first 6 rows for a non-empty string cell in that column.
' Used to check whether a column is labelled "ID", "Date", "Name", etc.
'===============================================================================
Private Function GetColumnHeader(ByVal ws As Worksheet, ByVal colNum As Long) As String
    GetColumnHeader = ""
    Dim r As Long
    For r = 1 To 6
        Dim v As String: v = Trim(CStr(ws.Cells(r, colNum).Value))
        If Len(v) > 0 And Not IsNumeric(v) Then
            GetColumnHeader = LCase(v)
            Exit Function
        End If
    Next r
End Function

'===============================================================================
' IsSkippedHeader - Returns True if the column header matches a skip keyword
' Skip keywords are defined in SKIP_HEADER_KEYWORDS (comma-separated).
'===============================================================================
Private Function IsSkippedHeader(ByVal headerText As String) As Boolean
    IsSkippedHeader = False
    If Len(headerText) = 0 Then Exit Function
    Dim keywords As Variant: keywords = Split(SKIP_HEADER_KEYWORDS, ",")
    Dim kw As Variant
    For Each kw In keywords
        If InStr(headerText, Trim(CStr(kw))) > 0 Then
            IsSkippedHeader = True
            Exit Function
        End If
    Next kw
End Function

'===============================================================================
' IsSkippedSheet - Returns True if the sheet name is in the skip list
'===============================================================================
Private Function IsSkippedSheet(ByVal shName As String) As Boolean
    IsSkippedSheet = InStr("," & SKIP_SHEET_NAMES & ",", "," & shName & ",") > 0
End Function
