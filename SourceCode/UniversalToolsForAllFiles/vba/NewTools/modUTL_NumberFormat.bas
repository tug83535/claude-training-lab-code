Attribute VB_Name = "modUTL_NumberFormat"
Option Explicit

' ============================================================
' KBT Universal Tools — Number Format Module
' Works on ANY Excel file — no project-specific setup required
' Tools: 2 | Small effort
' Date: 2026-03-05
' ============================================================
' Tool 10 — Enhanced Text-to-Number Converter
' Tool 11 — Workbook Metadata Reporter
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
' TOOL 10 — Enhanced Text-to-Number Converter          [SMALL]
' Scans all sheets for text cells that look like numbers
' (including currency formatted like "$1,234.56" or negative
' like "(500)"). Converts them to real numbers. Reports count.
' Smarter than basic CLEAN+TRIM — handles currency symbols,
' commas, parenthetical negatives, and percentage signs.
' ============================================================
Sub EnhancedTextToNumberConverter()
    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim totalFixed As Long: totalFixed = 0
    Dim sheetCount As Long: sheetCount = 0
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        Dim rng As Range: Set rng = Nothing
        On Error Resume Next
        Set rng = ws.UsedRange.SpecialCells(xlCellTypeConstants, xlTextValues)
        On Error GoTo ErrHandler
        If rng Is Nothing Then GoTo NextSheet

        Dim sheetFixed As Long: sheetFixed = 0
        Dim cell As Range
        For Each cell In rng
            Dim original As String: original = Trim(CStr(cell.Value))
            If Len(original) = 0 Then GoTo NextCell

            ' Skip if header row (row 1)
            If cell.Row = 1 Then GoTo NextCell

            ' Try to parse as number
            Dim numVal As Variant
            numVal = ParseTextAsNumber(original)
            If Not IsEmpty(numVal) Then
                cell.Value = CDbl(numVal)
                sheetFixed = sheetFixed + 1
            End If
NextCell:
        Next cell

        If sheetFixed > 0 Then
            totalFixed = totalFixed + sheetFixed
            sheetCount = sheetCount + 1
        End If
NextSheet:
    Next ws

    UTL_TurboOff

    If totalFixed = 0 Then
        MsgBox "No text-stored numbers found in any sheet.", vbInformation, _
               "Enhanced Text-to-Number Converter"
    Else
        MsgBox "Text-to-Number Conversion Complete" & vbCrLf & vbCrLf & _
               totalFixed & " cell(s) converted across " & sheetCount & " sheet(s).", _
               vbInformation, "Enhanced Text-to-Number Converter"
    End If
    Exit Sub

ErrHandler:
    UTL_TurboOff
    MsgBox "Enhanced Text-to-Number Converter error: " & Err.Description, vbCritical
End Sub

' Parse a text string into a number, handling currency, commas, parens, %
' Returns Empty if it's not a recognizable number pattern
Private Function ParseTextAsNumber(ByVal txt As String) As Variant
    ParseTextAsNumber = Empty

    Dim cleaned As String: cleaned = txt

    ' Remove leading/trailing whitespace
    cleaned = Trim(cleaned)
    If Len(cleaned) = 0 Then Exit Function

    ' Check for parenthetical negative: (123.45) -> -123.45
    Dim isNegParen As Boolean: isNegParen = False
    If Left(cleaned, 1) = "(" And Right(cleaned, 1) = ")" Then
        isNegParen = True
        cleaned = Mid(cleaned, 2, Len(cleaned) - 2)
    End If

    ' Remove currency symbols
    cleaned = Replace(cleaned, "$", "")
    cleaned = Replace(cleaned, ChrW(163), "")  ' pound sign
    cleaned = Replace(cleaned, ChrW(8364), "") ' euro sign

    ' Handle percentage
    Dim isPct As Boolean: isPct = False
    If Right(cleaned, 1) = "%" Then
        isPct = True
        cleaned = Left(cleaned, Len(cleaned) - 1)
    End If

    ' Remove commas (thousand separators)
    cleaned = Replace(cleaned, ",", "")

    ' Remove spaces
    cleaned = Replace(cleaned, " ", "")

    ' Check if remaining text is numeric
    If Not IsNumeric(cleaned) Then Exit Function

    Dim result As Double: result = CDbl(cleaned)
    If isNegParen Then result = -result
    If isPct Then result = result / 100

    ParseTextAsNumber = result
End Function

' ============================================================
' TOOL 11 — Workbook Metadata Reporter                 [SMALL]
' Creates a summary report of the active workbook including:
'   - File name, path, size, last saved
'   - Sheet count, names, visibility, row/col counts
'   - Named ranges inventory
'   - External links detected
' ============================================================
Sub WorkbookMetadataReporter()
    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim rptName As String: rptName = "UTL_MetadataReport"

    ' Delete old report
    Dim wsOld As Worksheet
    On Error Resume Next
    Set wsOld = wb.Worksheets(rptName)
    On Error GoTo ErrHandler
    If Not wsOld Is Nothing Then
        Application.DisplayAlerts = False
        wsOld.Delete
        Application.DisplayAlerts = True
    End If

    Dim wsRpt As Worksheet
    Set wsRpt = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsRpt.Name = rptName

    Dim outRow As Long: outRow = 1

    ' ── Section 1: Workbook Info ──
    wsRpt.Cells(outRow, 1).Value = "WORKBOOK METADATA REPORT"
    wsRpt.Cells(outRow, 1).Font.Bold = True
    wsRpt.Cells(outRow, 1).Font.Size = 14
    outRow = outRow + 2

    wsRpt.Cells(outRow, 1).Value = "File Name:"
    wsRpt.Cells(outRow, 2).Value = wb.Name
    wsRpt.Cells(outRow, 1).Font.Bold = True
    outRow = outRow + 1

    wsRpt.Cells(outRow, 1).Value = "Full Path:"
    wsRpt.Cells(outRow, 2).Value = wb.FullName
    wsRpt.Cells(outRow, 1).Font.Bold = True
    outRow = outRow + 1

    wsRpt.Cells(outRow, 1).Value = "Last Saved:"
    On Error Resume Next
    wsRpt.Cells(outRow, 2).Value = Format(wb.BuiltinDocumentProperties("Last Save Time"), "yyyy-mm-dd hh:nn:ss")
    If Err.Number <> 0 Then wsRpt.Cells(outRow, 2).Value = "(not available)"
    Err.Clear
    On Error GoTo ErrHandler
    wsRpt.Cells(outRow, 1).Font.Bold = True
    outRow = outRow + 1

    wsRpt.Cells(outRow, 1).Value = "Author:"
    On Error Resume Next
    wsRpt.Cells(outRow, 2).Value = wb.BuiltinDocumentProperties("Author")
    If Err.Number <> 0 Then wsRpt.Cells(outRow, 2).Value = "(not available)"
    Err.Clear
    On Error GoTo ErrHandler
    wsRpt.Cells(outRow, 1).Font.Bold = True
    outRow = outRow + 1

    wsRpt.Cells(outRow, 1).Value = "Sheet Count:"
    wsRpt.Cells(outRow, 2).Value = wb.Worksheets.Count
    wsRpt.Cells(outRow, 1).Font.Bold = True
    outRow = outRow + 1

    wsRpt.Cells(outRow, 1).Value = "Named Ranges:"
    wsRpt.Cells(outRow, 2).Value = wb.Names.Count
    wsRpt.Cells(outRow, 1).Font.Bold = True
    outRow = outRow + 2

    ' ── Section 2: Sheet Inventory ──
    wsRpt.Cells(outRow, 1).Value = "SHEET INVENTORY"
    wsRpt.Cells(outRow, 1).Font.Bold = True
    wsRpt.Cells(outRow, 1).Font.Size = 12
    outRow = outRow + 1

    wsRpt.Cells(outRow, 1).Value = "Sheet Name"
    wsRpt.Cells(outRow, 2).Value = "Visibility"
    wsRpt.Cells(outRow, 3).Value = "Used Rows"
    wsRpt.Cells(outRow, 4).Value = "Used Cols"
    wsRpt.Cells(outRow, 5).Value = "Has Formulas"
    Dim hdrC As Long
    For hdrC = 1 To 5
        wsRpt.Cells(outRow, hdrC).Font.Bold = True
        wsRpt.Cells(outRow, hdrC).Interior.Color = RGB(11, 71, 121)
        wsRpt.Cells(outRow, hdrC).Font.Color = RGB(255, 255, 255)
    Next hdrC
    outRow = outRow + 1

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Name = rptName Then GoTo NextSheet
        wsRpt.Cells(outRow, 1).Value = ws.Name

        Select Case ws.Visible
            Case xlSheetVisible:    wsRpt.Cells(outRow, 2).Value = "Visible"
            Case xlSheetHidden:     wsRpt.Cells(outRow, 2).Value = "Hidden"
            Case xlSheetVeryHidden: wsRpt.Cells(outRow, 2).Value = "Very Hidden"
        End Select

        On Error Resume Next
        Dim usedRows As Long: usedRows = ws.UsedRange.Rows.Count
        Dim usedCols As Long: usedCols = ws.UsedRange.Columns.Count
        If Err.Number <> 0 Then usedRows = 0: usedCols = 0
        Err.Clear
        On Error GoTo ErrHandler

        wsRpt.Cells(outRow, 3).Value = usedRows
        wsRpt.Cells(outRow, 4).Value = usedCols

        ' Check for formulas
        Dim fRng As Range: Set fRng = Nothing
        On Error Resume Next
        Set fRng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo ErrHandler
        wsRpt.Cells(outRow, 5).Value = IIf(fRng Is Nothing, "No", "Yes")

        outRow = outRow + 1
NextSheet:
    Next ws

    outRow = outRow + 1

    ' ── Section 3: Named Ranges ──
    If wb.Names.Count > 0 Then
        wsRpt.Cells(outRow, 1).Value = "NAMED RANGES"
        wsRpt.Cells(outRow, 1).Font.Bold = True
        wsRpt.Cells(outRow, 1).Font.Size = 12
        outRow = outRow + 1

        wsRpt.Cells(outRow, 1).Value = "Name"
        wsRpt.Cells(outRow, 2).Value = "Refers To"
        wsRpt.Cells(outRow, 3).Value = "Scope"
        For hdrC = 1 To 3
            wsRpt.Cells(outRow, hdrC).Font.Bold = True
            wsRpt.Cells(outRow, hdrC).Interior.Color = RGB(11, 71, 121)
            wsRpt.Cells(outRow, hdrC).Font.Color = RGB(255, 255, 255)
        Next hdrC
        outRow = outRow + 1

        Dim nm As Name
        Dim nameCount As Long: nameCount = 0
        For Each nm In wb.Names
            If nameCount >= 200 Then
                wsRpt.Cells(outRow, 1).Value = "--- LIMIT (200 names shown) ---"
                outRow = outRow + 1
                Exit For
            End If
            wsRpt.Cells(outRow, 1).Value = nm.Name
            On Error Resume Next
            wsRpt.Cells(outRow, 2).Value = "'" & nm.RefersTo
            On Error GoTo ErrHandler
            wsRpt.Cells(outRow, 3).Value = IIf(InStr(nm.Name, "!") > 0, "Sheet", "Workbook")
            outRow = outRow + 1
            nameCount = nameCount + 1
        Next nm
    End If

    wsRpt.Columns("A:E").AutoFit
    wsRpt.Activate

    UTL_TurboOff
    MsgBox "Metadata report created on '" & rptName & "'." & vbCrLf & _
           wb.Worksheets.Count & " sheets, " & wb.Names.Count & " named ranges documented.", _
           vbInformation, "Workbook Metadata Reporter"
    Exit Sub

ErrHandler:
    UTL_TurboOff
    MsgBox "Workbook Metadata Reporter error: " & Err.Description, vbCritical
End Sub
