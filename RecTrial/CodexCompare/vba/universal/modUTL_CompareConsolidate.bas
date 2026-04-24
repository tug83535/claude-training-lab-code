Attribute VB_Name = "modUTL_CompareConsolidate"
Option Explicit

Public Sub CompareActiveSheetToSheet(ByVal TargetSheetName As String)
    Dim wsLeft As Worksheet
    Dim wsRight As Worksheet
    Dim reportWs As Worksheet
    Dim headerRowLeft As Long
    Dim headerRowRight As Long
    Dim leftMap As Object
    Dim rightMap As Object
    Dim reportRow As Long
    Dim key As Variant

    On Error GoTo CompareFail

    Set wsLeft = ActiveSheet
    Set wsRight = ThisWorkbook.Worksheets(TargetSheetName)

    headerRowLeft = UTL_DetectHeaderRow(wsLeft)
    headerRowRight = UTL_DetectHeaderRow(wsRight)

    Set leftMap = BuildRowHashMap(wsLeft, headerRowLeft)
    Set rightMap = BuildRowHashMap(wsRight, headerRowRight)

    Set reportWs = GetOrCreateOutputSheet("UTL_CompareReport")
    reportWs.Cells.Clear
    reportWs.Range("A1:E1").Value = Array("Status", "Row Key", "Source Sheet", "Target Sheet", "Notes")
    reportWs.Rows(1).Font.Bold = True

    reportRow = 2

    For Each key In leftMap.Keys
        If Not rightMap.Exists(CStr(key)) Then
            reportWs.Cells(reportRow, 1).Value = "Missing in target"
            reportWs.Cells(reportRow, 2).Value = key
            reportWs.Cells(reportRow, 3).Value = wsLeft.Name
            reportWs.Cells(reportRow, 4).Value = wsRight.Name
            reportWs.Cells(reportRow, 5).Value = "Row signature exists on source but not target."
            reportRow = reportRow + 1
        End If
    Next key

    For Each key In rightMap.Keys
        If Not leftMap.Exists(CStr(key)) Then
            reportWs.Cells(reportRow, 1).Value = "Missing in source"
            reportWs.Cells(reportRow, 2).Value = key
            reportWs.Cells(reportRow, 3).Value = wsLeft.Name
            reportWs.Cells(reportRow, 4).Value = wsRight.Name
            reportWs.Cells(reportRow, 5).Value = "Row signature exists on target but not source."
            reportRow = reportRow + 1
        End If
    Next key

    reportWs.Columns("A:E").AutoFit

    UTL_LogAction "modUTL_CompareConsolidate", "CompareActiveSheetToSheet", "PASS", _
                  "Comparison complete", 2, reportRow - 2
    UTL_ShowCompletion "Compare Sheets", "Comparison report created on UTL_CompareReport. Differences: " & (reportRow - 2)
    Exit Sub

CompareFail:
    UTL_LogAction "modUTL_CompareConsolidate", "CompareActiveSheetToSheet", "FAIL", Err.Description
    MsgBox "Compare failed: " & Err.Description, vbExclamation, "Compare Sheets"
End Sub

Public Sub ConsolidateVisibleSheetsByHeader()
    Dim targets As Collection
    Dim ws As Worksheet
    Dim outWs As Worksheet
    Dim headerRow As Long
    Dim dataRange As Range
    Dim outRow As Long
    Dim writeHeader As Boolean

    On Error GoTo ConsolidateFail

    Set targets = UTL_GetTargetSheets(False)
    Set outWs = GetOrCreateOutputSheet("UTL_Consolidated")
    outWs.Cells.Clear

    outRow = 1
    writeHeader = True

    For Each ws In targets
        If ws.Name <> outWs.Name And ws.Name <> "UTL_RunLog" And ws.Name <> "UTL_CommandCenter" Then
            headerRow = UTL_DetectHeaderRow(ws)
            Set dataRange = UTL_DetectDataRange(ws, headerRow)

            If writeHeader Then
                dataRange.Rows(1).Copy outWs.Cells(outRow, 1)
                outWs.Cells(outRow, dataRange.Columns.Count + 1).Value = "SourceSheet"
                outRow = outRow + 1
                writeHeader = False
            End If

            If dataRange.Rows.Count > 1 Then
                dataRange.Offset(1, 0).Resize(dataRange.Rows.Count - 1, dataRange.Columns.Count).Copy outWs.Cells(outRow, 1)
                FillSourceSheetTag outWs, outRow, dataRange.Rows.Count - 1, dataRange.Columns.Count + 1, ws.Name
                outRow = outRow + dataRange.Rows.Count - 1
            End If
        End If
    Next ws

    outWs.Rows(1).Font.Bold = True
    outWs.Columns.AutoFit

    UTL_LogAction "modUTL_CompareConsolidate", "ConsolidateVisibleSheetsByHeader", "PASS", _
                  "Consolidation complete", targets.Count, outRow - 2
    UTL_ShowCompletion "Consolidate Sheets", "Consolidated rows written: " & Application.Max(0, outRow - 2)
    Exit Sub

ConsolidateFail:
    UTL_LogAction "modUTL_CompareConsolidate", "ConsolidateVisibleSheetsByHeader", "FAIL", Err.Description
    MsgBox "Consolidation failed: " & Err.Description, vbExclamation, "Consolidate Sheets"
End Sub

Private Function BuildRowHashMap(ByVal ws As Worksheet, ByVal headerRow As Long) As Object
    Dim dataRange As Range
    Dim r As Long
    Dim c As Long
    Dim rowText As String
    Dim dict As Object

    Set dict = CreateObject("Scripting.Dictionary")
    Set dataRange = UTL_DetectDataRange(ws, headerRow)

    If dataRange.Rows.Count <= 1 Then
        Set BuildRowHashMap = dict
        Exit Function
    End If

    For r = 2 To dataRange.Rows.Count
        rowText = ""
        For c = 1 To dataRange.Columns.Count
            rowText = rowText & "|" & Trim$(CStr(dataRange.Cells(r, c).Value2))
        Next c
        If Not dict.Exists(rowText) Then dict.Add rowText, True
    Next r

    Set BuildRowHashMap = dict
End Function

Private Sub FillSourceSheetTag(ByVal ws As Worksheet, ByVal startRow As Long, ByVal rowCount As Long, ByVal colNum As Long, ByVal sheetName As String)
    Dim r As Long
    For r = startRow To startRow + rowCount - 1
        ws.Cells(r, colNum).Value = sheetName
    Next r
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
