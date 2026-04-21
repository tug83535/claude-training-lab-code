Attribute VB_Name = "modUTL_Intelligence"
Option Explicit

' ============================================================
' KBT Universal Tools — Intelligence Module
' Three universal tools that work on ANY active sheet:
'   - MaterialityClassifierActiveSheet: tags rows Material / Watch / Normal
'   - GenerateExceptionNarrativesActiveSheet: plain-English sentence per row
'   - DataQualityScorecardActiveSheet: 0-100 score on blanks + error cells
' Cherry-picked from Codex comparison (Batch 2, 2026-04-20).
' Install in Personal.xlsb to use across all Excel sessions.
' Tools: 3 | Tier 1: 3
' ============================================================
' iPipeline Brand Colors used (per modUTL_Branding.bas documentation):
'   iPipeline Blue:   RGB(11,  71,  121)
'   Navy Blue:        RGB(17,  46,  81)
'   Arctic White:     RGB(249, 249, 249)
'   Charcoal:         RGB(22,  22,  22)
' ============================================================
' DEPENDENCIES:
'   modUTL_Core: UTL_LastRow, UTL_LastCol, UTL_DetectHeaderRow
' ============================================================

' ============================================================
' TOOL 1 — Materiality Classifier                  [TIER 1]
' Labels each row of the active sheet as Material increase /
' Material decrease / Watch / Normal based on an absolute $
' threshold and a % threshold. Finds Current and Prior columns
' automatically by header text.
' Writes two new columns: "Materiality Status" and "Variance %".
' ============================================================
Public Sub MaterialityClassifierActiveSheet(Optional ByVal AbsoluteThreshold As Double = 10000, _
                                            Optional ByVal PercentThreshold As Double = 0.15)
    Dim ws As Worksheet
    Dim headerRow As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim currentCol As Long
    Dim priorCol As Long
    Dim outputCol As Long
    Dim pctCol As Long
    Dim r As Long
    Dim deltaValue As Double
    Dim pctChange As Double
    Dim statusLabel As String
    Dim taggedCount As Long

    On Error GoTo ClassifierFail

    Set ws = ActiveSheet
    headerRow = UTL_DetectHeaderRow(ws)
    lastCol = UTL_LastCol(ws, headerRow)
    lastRow = UTL_LastRow(ws, 1)
    If lastRow < headerRow + 1 Or lastCol < 2 Then
        Err.Raise vbObjectError + 600, "MaterialityClassifierActiveSheet", _
                  "Active sheet has no data rows below the detected header row."
    End If

    currentCol = FindColumnByHeaderText(ws, headerRow, lastCol, _
                    Array("Current", "Actual", "Amount", "Q1 Actual"))
    priorCol = FindColumnByHeaderText(ws, headerRow, lastCol, _
                    Array("Prior", "Budget", "Baseline", "Q1 Budget"))

    If currentCol = 0 Or priorCol = 0 Then
        Err.Raise vbObjectError + 601, "MaterialityClassifierActiveSheet", _
                  "Could not detect Current/Prior columns by header text. " & _
                  "Add a column named Current/Actual/Amount and another named Prior/Budget/Baseline."
    End If

    outputCol = lastCol + 1
    pctCol = lastCol + 2

    ws.Cells(headerRow, outputCol).Value = "Materiality Status"
    ws.Cells(headerRow, pctCol).Value = "Variance %"
    With ws.Range(ws.Cells(headerRow, outputCol), ws.Cells(headerRow, pctCol))
        .Font.Bold = True
        .Font.Name = "Arial"
        .Font.Color = RGB(249, 249, 249)      ' Arctic White
        .Interior.Color = RGB(11, 71, 121)    ' iPipeline Blue
    End With

    For r = headerRow + 1 To lastRow
        If IsNumeric(ws.Cells(r, currentCol).Value2) And IsNumeric(ws.Cells(r, priorCol).Value2) Then
            deltaValue = CDbl(ws.Cells(r, currentCol).Value2) - CDbl(ws.Cells(r, priorCol).Value2)
            pctChange = SafePercent(deltaValue, CDbl(ws.Cells(r, priorCol).Value2))
            statusLabel = LabelMateriality(deltaValue, pctChange, AbsoluteThreshold, PercentThreshold)

            ws.Cells(r, outputCol).Value = statusLabel
            ws.Cells(r, pctCol).Value = pctChange
            ws.Cells(r, pctCol).NumberFormat = "0.0%"
            taggedCount = taggedCount + 1
        End If
    Next r

    ws.Columns(outputCol).AutoFit
    ws.Columns(pctCol).AutoFit

    Debug.Print "[UTL Intelligence] MaterialityClassifier: " & taggedCount & " row(s) tagged on " & ws.Name
    MsgBox "Materiality classifier complete." & Chr(10) & Chr(10) & _
           taggedCount & " row(s) tagged on '" & ws.Name & "'." & Chr(10) & _
           "Thresholds: $" & Format(AbsoluteThreshold, "#,##0") & " / " & Format(PercentThreshold, "0.0%"), _
           vbInformation, "UTL Intelligence - Materiality"
    Exit Sub

ClassifierFail:
    MsgBox "Materiality classifier failed: " & Err.Description, vbExclamation, "UTL Intelligence - Materiality"
End Sub

' ============================================================
' TOOL 2 — Generate Exception Narratives            [TIER 1]
' Reads the Materiality Status column (from Tool 1) and writes
' a plain-English narrative sentence in a new "Narrative" column.
' If Tool 1 hasn't run yet, this sub warns the user.
' ============================================================
Public Sub GenerateExceptionNarrativesActiveSheet()
    Dim ws As Worksheet
    Dim headerRow As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim statusCol As Long
    Dim nameCol As Long
    Dim amountCol As Long
    Dim outputCol As Long
    Dim r As Long
    Dim narrativeCount As Long

    On Error GoTo NarrativeFail

    Set ws = ActiveSheet
    headerRow = UTL_DetectHeaderRow(ws)
    lastCol = UTL_LastCol(ws, headerRow)
    lastRow = UTL_LastRow(ws, 1)
    If lastRow < headerRow + 1 Or lastCol < 1 Then
        Err.Raise vbObjectError + 610, "GenerateExceptionNarrativesActiveSheet", _
                  "Active sheet has no data rows below the detected header row."
    End If

    statusCol = FindColumnByHeaderText(ws, headerRow, lastCol, _
                   Array("Materiality Status", "Status"))
    nameCol = FindColumnByHeaderText(ws, headerRow, lastCol, _
                 Array("Line Item", "Department", "Customer", "Product"))
    amountCol = FindColumnByHeaderText(ws, headerRow, lastCol, _
                   Array("Amount", "Current", "Actual", "Q1 Actual"))

    If statusCol = 0 Then
        Err.Raise vbObjectError + 611, "GenerateExceptionNarrativesActiveSheet", _
                  "Status column not found. Run the Materiality Classifier first."
    End If
    If nameCol = 0 Then nameCol = 1

    outputCol = lastCol + 1
    ws.Cells(headerRow, outputCol).Value = "Narrative"
    With ws.Cells(headerRow, outputCol)
        .Font.Bold = True
        .Font.Name = "Arial"
        .Font.Color = RGB(249, 249, 249)      ' Arctic White
        .Interior.Color = RGB(11, 71, 121)    ' iPipeline Blue
    End With

    For r = headerRow + 1 To lastRow
        If Len(Trim$(CStr(ws.Cells(r, statusCol).Value2))) > 0 Then
            ws.Cells(r, outputCol).Value = BuildNarrative( _
                CStr(ws.Cells(r, nameCol).Value2), _
                CStr(ws.Cells(r, statusCol).Value2), _
                ws.Cells(r, amountCol).Value2)
            narrativeCount = narrativeCount + 1
        End If
    Next r

    ws.Columns(outputCol).ColumnWidth = 44

    Debug.Print "[UTL Intelligence] ExceptionNarratives: " & narrativeCount & " narrative(s) on " & ws.Name
    MsgBox "Narratives generated." & Chr(10) & Chr(10) & _
           narrativeCount & " row(s) on '" & ws.Name & "'.", _
           vbInformation, "UTL Intelligence - Narratives"
    Exit Sub

NarrativeFail:
    MsgBox "Narrative generation failed: " & Err.Description, vbExclamation, "UTL Intelligence - Narratives"
End Sub

' ============================================================
' TOOL 3 — Data Quality Scorecard                   [TIER 1]
' Scores the active sheet 0-100. Score = 100 - (blank% * 60) -
' (error% * 40). Writes a summary sheet UTL_QualityScorecard.
' Works on any sheet — no column detection required.
' ============================================================
Public Sub DataQualityScorecardActiveSheet()
    Dim ws As Worksheet
    Dim headerRow As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long, c As Long
    Dim totalCells As Double
    Dim blankCells As Double
    Dim errorCells As Double
    Dim numericCells As Double
    Dim cellVal As Variant
    Dim score As Double
    Dim outWs As Worksheet

    On Error GoTo ScoreFail

    Set ws = ActiveSheet
    headerRow = UTL_DetectHeaderRow(ws)
    lastCol = UTL_LastCol(ws, headerRow)
    lastRow = UTL_LastRow(ws, 1)
    If lastRow < headerRow Or lastCol < 1 Then
        Err.Raise vbObjectError + 620, "DataQualityScorecardActiveSheet", _
                  "Active sheet is empty or has no detectable data."
    End If

    For r = headerRow To lastRow
        For c = 1 To lastCol
            totalCells = totalCells + 1
            If IsError(ws.Cells(r, c).Value) Then
                errorCells = errorCells + 1
            Else
                cellVal = ws.Cells(r, c).Value2
                If Len(CStr(cellVal)) = 0 Then blankCells = blankCells + 1
                If IsNumeric(cellVal) Then numericCells = numericCells + 1
            End If
        Next c
    Next r

    If totalCells <= 0 Then totalCells = 1    ' Guard against div/0
    score = 100 - ((blankCells / totalCells) * 60) - ((errorCells / totalCells) * 40)
    If score < 0 Then score = 0

    Set outWs = GetOrCreateOutputSheet("UTL_QualityScorecard")
    outWs.Cells.Clear

    outWs.Range("A1").Value = "Data Quality Scorecard"
    outWs.Range("A1").Font.Bold = True
    outWs.Range("A1").Font.Size = 14
    outWs.Range("A1").Font.Name = "Arial"
    outWs.Range("A1").Font.Color = RGB(17, 46, 81)        ' Navy

    outWs.Range("A3").Value = "Metric"
    outWs.Range("B3").Value = "Value"
    With outWs.Range("A3:B3")
        .Font.Bold = True
        .Font.Name = "Arial"
        .Font.Color = RGB(249, 249, 249)      ' Arctic White
        .Interior.Color = RGB(11, 71, 121)    ' iPipeline Blue
    End With

    outWs.Range("A4").Value = "Sheet"
    outWs.Range("B4").Value = ws.Name
    outWs.Range("A5").Value = "Data Range"
    outWs.Range("B5").Value = ws.Cells(headerRow, 1).Address(False, False) & ":" & _
                              ws.Cells(lastRow, lastCol).Address(False, False)
    outWs.Range("A6").Value = "Total Cells"
    outWs.Range("B6").Value = totalCells
    outWs.Range("A7").Value = "Blank Cells"
    outWs.Range("B7").Value = blankCells
    outWs.Range("A8").Value = "Error Cells"
    outWs.Range("B8").Value = errorCells
    outWs.Range("A9").Value = "Numeric Cells"
    outWs.Range("B9").Value = numericCells

    outWs.Range("A10").Value = "Quality Score (0-100)"
    outWs.Range("B10").Value = score
    outWs.Range("B10").NumberFormat = "0.0"
    With outWs.Range("A10:B10")
        .Font.Bold = True
        .Font.Name = "Arial"
        .Interior.Color = RGB(240, 240, 238)
    End With

    ' Color the score cell based on severity
    Select Case score
        Case Is >= 90: outWs.Range("B10").Font.Color = RGB(0, 128, 0)          ' Strong green
        Case Is >= 75: outWs.Range("B10").Font.Color = RGB(17, 46, 81)         ' Navy (ok)
        Case Is >= 60: outWs.Range("B10").Font.Color = RGB(200, 100, 0)        ' Warning orange
        Case Else:     outWs.Range("B10").Font.Color = RGB(200, 0, 0)          ' Red
    End Select

    outWs.Range("A4:A10").Font.Name = "Arial"
    outWs.Range("A4:A10").Font.Color = RGB(22, 22, 22)    ' Charcoal
    outWs.Range("B4:B9").Font.Name = "Arial"
    outWs.Range("B4:B9").Font.Color = RGB(22, 22, 22)

    outWs.Columns("A").ColumnWidth = 26
    outWs.Columns("B").ColumnWidth = 40
    outWs.Activate
    outWs.Range("A1").Select

    Debug.Print "[UTL Intelligence] QualityScorecard: " & Format(score, "0.0") & " on " & ws.Name
    MsgBox "Data Quality Scorecard created." & Chr(10) & Chr(10) & _
           "Sheet: " & ws.Name & Chr(10) & _
           "Score: " & Format(score, "0.0") & " / 100" & Chr(10) & Chr(10) & _
           "See UTL_QualityScorecard for the breakdown.", _
           vbInformation, "UTL Intelligence - Scorecard"
    Exit Sub

ScoreFail:
    MsgBox "Quality scorecard failed: " & Err.Description, vbExclamation, "UTL Intelligence - Scorecard"
End Sub

' ============================================================
' Private helpers
' ============================================================

' Find the first column whose header text contains any of the candidate strings.
' Case-insensitive substring match. Returns 0 if not found.
Private Function FindColumnByHeaderText(ByVal ws As Worksheet, _
                                         ByVal headerRow As Long, _
                                         ByVal lastCol As Long, _
                                         ByVal candidates As Variant) As Long
    Dim col As Long
    Dim candidate As Variant
    Dim headerText As String

    For col = 1 To lastCol
        headerText = LCase$(Trim$(CStr(ws.Cells(headerRow, col).Value2)))
        If Len(headerText) > 0 Then
            For Each candidate In candidates
                If InStr(1, headerText, LCase$(CStr(candidate)), vbTextCompare) > 0 Then
                    FindColumnByHeaderText = col
                    Exit Function
                End If
            Next candidate
        End If
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

Private Function LabelMateriality(ByVal deltaValue As Double, _
                                   ByVal pctChange As Double, _
                                   ByVal absoluteThreshold As Double, _
                                   ByVal percentThreshold As Double) As String
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

Private Function BuildNarrative(ByVal lineName As String, _
                                 ByVal statusText As String, _
                                 ByVal amountValue As Variant) As String
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

' Returns an existing sheet by name, or creates it if missing. Uses
' ActiveWorkbook so the tool works whether installed in the workbook
' itself or in a Personal.xlsb add-in.
Private Function GetOrCreateOutputSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add( _
                    After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If

    Set GetOrCreateOutputSheet = ws
End Function
