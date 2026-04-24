Attribute VB_Name = "modUTL_Intelligence"
Option Explicit

Public Sub MaterialityClassifierActiveSheet(Optional ByVal AbsoluteThreshold As Double = 10000, Optional ByVal PercentThreshold As Double = 0.15)
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim headerRow As Long
    Dim currentCol As Long
    Dim priorCol As Long
    Dim outputCol As Long
    Dim pctCol As Long
    Dim r As Long
    Dim deltaValue As Double
    Dim pctChange As Double
    Dim statusLabel As String

    On Error GoTo ClassifierFail

    Set ws = ActiveSheet
    headerRow = UTL_DetectHeaderRow(ws)
    Set dataRange = UTL_DetectDataRange(ws, headerRow)

    currentCol = FindColumnByHeaderText(ws, headerRow, Array("Current", "Actual", "Amount", "Q1 Actual"))
    priorCol = FindColumnByHeaderText(ws, headerRow, Array("Prior", "Budget", "Baseline", "Q1 Budget"))

    If currentCol = 0 Or priorCol = 0 Then
        Err.Raise vbObjectError + 601, "MaterialityClassifierActiveSheet", "Could not detect Current/Prior columns automatically."
    End If

    outputCol = dataRange.Columns.Count + 1
    pctCol = dataRange.Columns.Count + 2

    ws.Cells(headerRow, outputCol).Value = "Materiality Status"
    ws.Cells(headerRow, pctCol).Value = "Variance %"
    ws.Range(ws.Cells(headerRow, outputCol), ws.Cells(headerRow, pctCol)).Font.Bold = True

    For r = headerRow + 1 To dataRange.Row + dataRange.Rows.Count - 1
        If IsNumeric(ws.Cells(r, currentCol).Value2) And IsNumeric(ws.Cells(r, priorCol).Value2) Then
            deltaValue = CDbl(ws.Cells(r, currentCol).Value2) - CDbl(ws.Cells(r, priorCol).Value2)
            pctChange = SafePercent(deltaValue, CDbl(ws.Cells(r, priorCol).Value2))
            statusLabel = LabelMateriality(deltaValue, pctChange, AbsoluteThreshold, PercentThreshold)

            ws.Cells(r, outputCol).Value = statusLabel
            ws.Cells(r, pctCol).Value = pctChange
            ws.Cells(r, pctCol).NumberFormat = "0.0%"
        End If
    Next r

    UTL_LogAction "modUTL_Intelligence", "MaterialityClassifierActiveSheet", "PASS", "Materiality tags applied", 1, dataRange.Rows.Count - 1
    UTL_ShowCompletion "Materiality Classifier", "Classification complete on " & ws.Name
    Exit Sub

ClassifierFail:
    UTL_LogAction "modUTL_Intelligence", "MaterialityClassifierActiveSheet", "FAIL", Err.Description
    MsgBox "Materiality classifier failed: " & Err.Description, vbExclamation, "Materiality Classifier"
End Sub

Public Sub GenerateExceptionNarrativesActiveSheet()
    Dim ws As Worksheet
    Dim headerRow As Long
    Dim dataRange As Range
    Dim statusCol As Long
    Dim nameCol As Long
    Dim amountCol As Long
    Dim outputCol As Long
    Dim r As Long

    On Error GoTo NarrativeFail

    Set ws = ActiveSheet
    headerRow = UTL_DetectHeaderRow(ws)
    Set dataRange = UTL_DetectDataRange(ws, headerRow)

    statusCol = FindColumnByHeaderText(ws, headerRow, Array("Materiality Status", "Status"))
    nameCol = FindColumnByHeaderText(ws, headerRow, Array("Line Item", "Department", "Customer", "Product"))
    amountCol = FindColumnByHeaderText(ws, headerRow, Array("Amount", "Current", "Actual", "Q1 Actual"))

    If statusCol = 0 Then Err.Raise vbObjectError + 602, "GenerateExceptionNarrativesActiveSheet", "Status column not found."
    If nameCol = 0 Then nameCol = 1

    outputCol = dataRange.Columns.Count + 1
    ws.Cells(headerRow, outputCol).Value = "Narrative"
    ws.Cells(headerRow, outputCol).Font.Bold = True

    For r = headerRow + 1 To dataRange.Row + dataRange.Rows.Count - 1
        If Len(Trim$(CStr(ws.Cells(r, statusCol).Value2))) > 0 Then
            ws.Cells(r, outputCol).Value = BuildNarrative(CStr(ws.Cells(r, nameCol).Value2), CStr(ws.Cells(r, statusCol).Value2), ws.Cells(r, amountCol).Value2)
        End If
    Next r

    ws.Columns(outputCol).ColumnWidth = 44

    UTL_LogAction "modUTL_Intelligence", "GenerateExceptionNarrativesActiveSheet", "PASS", "Narratives generated", 1, dataRange.Rows.Count - 1
    UTL_ShowCompletion "Exception Narratives", "Narratives generated on " & ws.Name
    Exit Sub

NarrativeFail:
    UTL_LogAction "modUTL_Intelligence", "GenerateExceptionNarrativesActiveSheet", "FAIL", Err.Description
    MsgBox "Narrative generation failed: " & Err.Description, vbExclamation, "Exception Narratives"
End Sub

Public Sub DataQualityScorecardActiveSheet()
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim headerRow As Long
    Dim c As Range
    Dim totalCells As Double
    Dim blankCells As Double
    Dim errorCells As Double
    Dim numericCells As Double
    Dim score As Double
    Dim outWs As Worksheet

    On Error GoTo ScoreFail

    Set ws = ActiveSheet
    headerRow = UTL_DetectHeaderRow(ws)
    Set dataRange = UTL_DetectDataRange(ws, headerRow)

    For Each c In dataRange.Cells
        totalCells = totalCells + 1
        If Len(CStr(c.Value2)) = 0 Then blankCells = blankCells + 1
        If IsError(c.Value) Then errorCells = errorCells + 1
        If IsNumeric(c.Value2) Then numericCells = numericCells + 1
    Next c

    score = 100 - ((blankCells / totalCells) * 60) - ((errorCells / totalCells) * 40)
    If score < 0 Then score = 0

    Set outWs = GetOrCreateOutputSheet("UTL_QualityScorecard")
    outWs.Cells.Clear
    outWs.Range("A1:B1").Value = Array("Metric", "Value")
    outWs.Rows(1).Font.Bold = True
    outWs.Range("A2:B7").Value = Array( _
        Array("Sheet", ws.Name), _
        Array("Data Range", dataRange.Address(False, False)), _
        Array("Total Cells", totalCells), _
        Array("Blank Cells", blankCells), _
        Array("Error Cells", errorCells), _
        Array("Numeric Cells", numericCells))
    outWs.Range("A8").Value = "Quality Score"
    outWs.Range("B8").Value = score
    outWs.Range("B8").NumberFormat = "0.0"
    outWs.Range("A8:B8").Font.Bold = True
    outWs.Columns("A:B").AutoFit

    UTL_LogAction "modUTL_Intelligence", "DataQualityScorecardActiveSheet", "PASS", "Quality score generated", 1, totalCells
    UTL_ShowCompletion "Data Quality Scorecard", "Scorecard created on UTL_QualityScorecard"
    Exit Sub

ScoreFail:
    UTL_LogAction "modUTL_Intelligence", "DataQualityScorecardActiveSheet", "FAIL", Err.Description
    MsgBox "Quality scorecard failed: " & Err.Description, vbExclamation, "Data Quality Scorecard"
End Sub

Private Function FindColumnByHeaderText(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal candidates As Variant) As Long
    Dim col As Long
    Dim lastCol As Long
    Dim candidate As Variant
    Dim headerText As String

    lastCol = UTL_LastUsedColumn(ws)

    For col = 1 To lastCol
        headerText = LCase$(Trim$(CStr(ws.Cells(headerRow, col).Value2)))
        For Each candidate In candidates
            If InStr(1, headerText, LCase$(CStr(candidate)), vbTextCompare) > 0 Then
                FindColumnByHeaderText = col
                Exit Function
            End If
        Next candidate
    Next col
End Function

Private Function SafePercent(ByVal delta As Double, ByVal baseline As Double) As Double
    If baseline = 0 Then
        If delta = 0 Then
            SafePercent = 0
        Else
            SafePercent = 1
        End If
    Else
        SafePercent = delta / baseline
    End If
End Function

Private Function LabelMateriality(ByVal deltaValue As Double, ByVal pctChange As Double, ByVal absoluteThreshold As Double, ByVal percentThreshold As Double) As String
    If Abs(deltaValue) >= absoluteThreshold And Abs(pctChange) >= percentThreshold Then
        If deltaValue > 0 Then
            LabelMateriality = "Material increase"
        Else
            LabelMateriality = "Material decrease"
        End If
    ElseIf Abs(deltaValue) >= absoluteThreshold Or Abs(pctChange) >= percentThreshold Then
        LabelMateriality = "Watch"
    Else
        LabelMateriality = "Normal"
    End If
End Function

Private Function BuildNarrative(ByVal lineName As String, ByVal statusText As String, ByVal amountValue As Variant) As String
    Dim amountText As String

    If IsNumeric(amountValue) Then
        amountText = Format$(CDbl(amountValue), "$#,##0")
    Else
        amountText = CStr(amountValue)
    End If

    Select Case statusText
        Case "Material increase"
            BuildNarrative = lineName & " increased materially and requires owner confirmation. Current value: " & amountText & "."
        Case "Material decrease"
            BuildNarrative = lineName & " decreased materially and should be validated before close. Current value: " & amountText & "."
        Case "Watch"
            BuildNarrative = lineName & " is near materiality thresholds. Validate assumptions. Current value: " & amountText & "."
        Case Else
            BuildNarrative = lineName & " is within normal range. Current value: " & amountText & "."
    End Select
End Function

Private Function GetOrCreateOutputSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If

    Set GetOrCreateOutputSheet = ws
End Function
