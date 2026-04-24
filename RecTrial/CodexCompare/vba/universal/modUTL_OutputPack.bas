Attribute VB_Name = "modUTL_OutputPack"
Option Explicit

Public Sub BuildExecutiveOnePagerFromActiveSheet()
    Dim src As Worksheet
    Dim outWs As Worksheet
    Dim headerRow As Long
    Dim dataRange As Range
    Dim numericCol As Long
    Dim lastRow As Long
    Dim sumValue As Double
    Dim avgValue As Double
    Dim maxValue As Double
    Dim minValue As Double

    On Error GoTo BuildFail

    Set src = ActiveSheet
    headerRow = UTL_DetectHeaderRow(src)
    Set dataRange = UTL_DetectDataRange(src, headerRow)

    numericCol = FindFirstNumericColumn(src, headerRow)
    If numericCol = 0 Then Err.Raise vbObjectError + 701, "BuildExecutiveOnePagerFromActiveSheet", "Could not find a numeric column."

    lastRow = dataRange.Row + dataRange.Rows.Count - 1

    sumValue = Application.WorksheetFunction.Sum(src.Range(src.Cells(headerRow + 1, numericCol), src.Cells(lastRow, numericCol)))
    avgValue = Application.WorksheetFunction.Average(src.Range(src.Cells(headerRow + 1, numericCol), src.Cells(lastRow, numericCol)))
    maxValue = Application.WorksheetFunction.Max(src.Range(src.Cells(headerRow + 1, numericCol), src.Cells(lastRow, numericCol)))
    minValue = Application.WorksheetFunction.Min(src.Range(src.Cells(headerRow + 1, numericCol), src.Cells(lastRow, numericCol)))

    Set outWs = GetOrCreateOutputSheet("UTL_ExecutiveOnePager")
    outWs.Cells.Clear

    ApplyOnePagerBrandHeader outWs, src.Name

    outWs.Range("B7:C7").Value = Array("Metric", "Value")
    outWs.Range("B7:C7").Font.Bold = True

    outWs.Range("B8:C11").Value = Array( _
        Array("Total", sumValue), _
        Array("Average", avgValue), _
        Array("Maximum", maxValue), _
        Array("Minimum", minValue))

    outWs.Range("C8:C11").NumberFormat = "$#,##0;($#,##0);""-"""
    outWs.Columns("B:C").AutoFit

    UTL_LogAction "modUTL_OutputPack", "BuildExecutiveOnePagerFromActiveSheet", "PASS", "Executive one-pager built", 1, dataRange.Rows.Count - 1
    UTL_ShowCompletion "Executive One-Pager", "Created sheet UTL_ExecutiveOnePager from " & src.Name
    Exit Sub

BuildFail:
    UTL_LogAction "modUTL_OutputPack", "BuildExecutiveOnePagerFromActiveSheet", "FAIL", Err.Description
    MsgBox "One-pager build failed: " & Err.Description, vbExclamation, "Executive One-Pager"
End Sub

Public Sub ExportExecutivePackPDF(Optional ByVal OutputFolder As String = "")
    Dim wsMain As Worksheet
    Dim wsBrief As Worksheet
    Dim exportPath As String
    Dim fileName As String

    On Error GoTo ExportFail

    If Len(OutputFolder) = 0 Then
        exportPath = ThisWorkbook.Path
    Else
        exportPath = OutputFolder
    End If

    If Len(exportPath) = 0 Then Err.Raise vbObjectError + 702, "ExportExecutivePackPDF", "Workbook path is empty. Save workbook first."

    fileName = exportPath & Application.PathSeparator & "Executive_Pack_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf"

    Set wsMain = ActiveSheet

    On Error Resume Next
    Set wsBrief = ThisWorkbook.Worksheets("UTL_ExecutiveOnePager")
    On Error GoTo ExportFail

    If wsBrief Is Nothing Then
        BuildExecutiveOnePagerFromActiveSheet
        Set wsBrief = ThisWorkbook.Worksheets("UTL_ExecutiveOnePager")
    End If

    ThisWorkbook.Worksheets(Array(wsMain.Name, wsBrief.Name)).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fileName, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False
    wsMain.Select

    UTL_LogAction "modUTL_OutputPack", "ExportExecutivePackPDF", "PASS", "PDF exported: " & fileName, 2, 0
    UTL_ShowCompletion "Export Executive Pack", "PDF created: " & fileName
    Exit Sub

ExportFail:
    UTL_LogAction "modUTL_OutputPack", "ExportExecutivePackPDF", "FAIL", Err.Description
    MsgBox "PDF export failed: " & Err.Description, vbExclamation, "Export Executive Pack"
End Sub

Public Sub CreateRunReceiptSheet(ByVal FeatureName As String, ByVal Notes As String)
    Dim ws As Worksheet
    Dim nextRow As Long

    On Error GoTo ReceiptFail

    Set ws = GetOrCreateOutputSheet("UTL_RunReceipt")

    If Len(Trim$(CStr(ws.Range("A1").Value2))) = 0 Then
        ws.Range("A1:F1").Value = Array("Timestamp", "User", "Workbook", "Feature", "Notes", "Status")
        ws.Rows(1).Font.Bold = True
    End If

    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(nextRow, 2).Value = Environ$("Username")
    ws.Cells(nextRow, 3).Value = ThisWorkbook.Name
    ws.Cells(nextRow, 4).Value = FeatureName
    ws.Cells(nextRow, 5).Value = Notes
    ws.Cells(nextRow, 6).Value = "Recorded"

    ws.Columns("A:F").AutoFit

    UTL_LogAction "modUTL_OutputPack", "CreateRunReceiptSheet", "PASS", "Run receipt updated", 1, 1
    UTL_ShowCompletion "Run Receipt", "Receipt row added to UTL_RunReceipt"
    Exit Sub

ReceiptFail:
    UTL_LogAction "modUTL_OutputPack", "CreateRunReceiptSheet", "FAIL", Err.Description
    MsgBox "Run receipt update failed: " & Err.Description, vbExclamation, "Run Receipt"
End Sub

Private Function FindFirstNumericColumn(ByVal ws As Worksheet, ByVal headerRow As Long) As Long
    Dim lastCol As Long
    Dim lastRow As Long
    Dim c As Long
    Dim r As Long

    lastCol = UTL_LastUsedColumn(ws)
    lastRow = UTL_LastUsedRow(ws)

    For c = 1 To lastCol
        For r = headerRow + 1 To Application.Min(lastRow, headerRow + 25)
            If IsNumeric(ws.Cells(r, c).Value2) And Len(CStr(ws.Cells(r, c).Value2)) > 0 Then
                FindFirstNumericColumn = c
                Exit Function
            End If
        Next r
    Next c
End Function

Private Sub ApplyOnePagerBrandHeader(ByVal ws As Worksheet, ByVal sourceName As String)
    ws.Range("B2:E2").Merge
    ws.Range("B2").Value = "iPipeline"
    ws.Range("B2").Font.Name = "Arial"
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Size = 20
    ws.Range("B2").Font.Color = RGB(11, 71, 121)

    ws.Range("B3:E3").Merge
    ws.Range("B3").Value = "Finance & Accounting"
    ws.Range("B3").Font.Name = "Arial"
    ws.Range("B3").Font.Size = 10
    ws.Range("B3").Font.Color = RGB(17, 46, 81)

    ws.Range("B4:E4").Merge
    ws.Range("B4").Value = "Executive One-Pager"
    ws.Range("B4").Font.Name = "Arial"
    ws.Range("B4").Font.Size = 14
    ws.Range("B4").Font.Bold = True
    ws.Range("B4").Interior.Color = RGB(11, 71, 121)
    ws.Range("B4").Font.Color = RGB(249, 249, 249)

    ws.Range("B5").Value = "Source Sheet: " & sourceName
    ws.Range("C5").Value = "Generated: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

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
