Attribute VB_Name = "modMultiWorkbookDiff"
'===============================================================================
' modMultiWorkbookDiff
' PURPOSE: Compare any 2 or more Excel workbooks cell-by-cell and produce a
'          full change report (Added / Changed / Removed) across every sheet.
'
' WHY THIS IS NOT NATIVE: "Compare and Merge Workbooks" and "Spreadsheet
'          Compare" are legacy tools - only 2 files at a time, awkward UI,
'          and only work on workbooks previously Shared. OneDrive version
'          history shows who edited but not WHAT at a cell level across
'          an arbitrary set of files.
'
' USE CASE (software business):
'   - A template goes out to 40 regional offices. Each returns a completed copy.
'     Finance wants to know: "What did each region actually change vs the master?"
'   - Audit wants a diff of the SOX control matrix between Q1 and Q4.
'
' INPUT:
'   Either select files via dialog, or list them in a sheet called "FilesToDiff"
'   column A. The first file listed is the BASELINE; all others are compared to it.
'===============================================================================
Option Explicit

Private Type DiffRow
    filePath As String
    sheetName As String
    cellAddr As String
    changeType As String   ' "Changed" | "Added" | "Removed"
    baselineValue As Variant
    compareValue As Variant
End Type

Public Sub RunMultiWorkbookDiff()
    Dim files() As String, i As Long
    files = CollectFiles()
    If UBound(files) < 1 Then
        MsgBox "Select at least 2 files (baseline + 1 or more).", vbExclamation
        Exit Sub
    End If

    Dim report As Worksheet
    Set report = EnsureReportSheet()

    Dim reportRow As Long: reportRow = 2
    Dim baselineData As Object
    Set baselineData = LoadWorkbookData(files(0))

    For i = 1 To UBound(files)
        reportRow = DiffAgainstBaseline(files(i), baselineData, report, reportRow)
        Application.StatusBar = "Diffed " & i & " of " & UBound(files) & " files..."
    Next i

    report.Columns("A:F").AutoFit
    report.Range("A1").Select
    Application.StatusBar = False
    MsgBox "Diff complete. " & (reportRow - 2) & " rows written.", vbInformation
End Sub

Private Function CollectFiles() As String()
    Dim arr() As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("FilesToDiff")
    On Error GoTo 0
    If Not ws Is Nothing Then
        Dim lastRow As Long, r As Long, n As Long
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        ReDim arr(0 To lastRow - 2)
        For r = 2 To lastRow
            arr(r - 2) = CStr(ws.Cells(r, "A").Value)
        Next r
        CollectFiles = arr
        Exit Function
    End If

    ' Fallback: file dialog
    Dim fd As FileDialog, j As Long
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Pick files - FIRST file is the baseline"
    fd.AllowMultiSelect = True
    fd.Filters.Clear
    fd.Filters.Add "Excel", "*.xlsx; *.xlsm; *.xls"
    If fd.Show <> -1 Then Exit Function
    ReDim arr(0 To fd.SelectedItems.Count - 1)
    For j = 1 To fd.SelectedItems.Count
        arr(j - 1) = fd.SelectedItems(j)
    Next j
    CollectFiles = arr
End Function

Private Function LoadWorkbookData(ByVal path As String) As Object
    ' Key: "SheetName|A1" -> value
    Dim dict As Object, wb As Workbook, ws As Worksheet, rng As Range, cell As Range
    Set dict = CreateObject("Scripting.Dictionary")
    Set wb = Workbooks.Open(Filename:=path, ReadOnly:=True, UpdateLinks:=0)

    For Each ws In wb.Worksheets
        If ws.UsedRange.Rows.Count * ws.UsedRange.Columns.Count > 500000 Then
            ' Skip extreme sheets - log and continue
            dict.Add ws.Name & "|__SKIPPED__", "Too large"
            GoTo NextSheet
        End If
        For Each cell In ws.UsedRange
            If Len(CStr(cell.Value)) > 0 Or Len(CStr(cell.Formula)) > 0 Then
                dict(ws.Name & "|" & cell.Address(False, False)) = cell.Value
            End If
        Next cell
NextSheet:
    Next ws

    wb.Close SaveChanges:=False
    Set LoadWorkbookData = dict
End Function

Private Function DiffAgainstBaseline(ByVal path As String, baseline As Object, _
                                      report As Worksheet, ByVal startRow As Long) As Long
    Dim wb As Workbook, ws As Worksheet, cell As Range
    Dim compKey As String, row As Long, visited As Object
    Set visited = CreateObject("Scripting.Dictionary")
    row = startRow

    Set wb = Workbooks.Open(Filename:=path, ReadOnly:=True, UpdateLinks:=0)

    For Each ws In wb.Worksheets
        For Each cell In ws.UsedRange
            If Len(CStr(cell.Value)) > 0 Then
                compKey = ws.Name & "|" & cell.Address(False, False)
                visited(compKey) = True
                If baseline.Exists(compKey) Then
                    If CStr(baseline(compKey)) <> CStr(cell.Value) Then
                        WriteRow report, row, path, ws.Name, cell.Address(False, False), _
                                 "Changed", baseline(compKey), cell.Value
                        row = row + 1
                    End If
                Else
                    WriteRow report, row, path, ws.Name, cell.Address(False, False), _
                             "Added", "", cell.Value
                    row = row + 1
                End If
            End If
        Next cell
    Next ws

    ' Removed rows: exist in baseline but not in comp
    Dim k As Variant
    For Each k In baseline.Keys
        If Not visited.Exists(k) And InStr(CStr(k), "__SKIPPED__") = 0 Then
            Dim parts() As String
            parts = Split(CStr(k), "|")
            WriteRow report, row, path, parts(0), parts(1), _
                     "Removed", baseline(k), ""
            row = row + 1
        End If
    Next k

    wb.Close SaveChanges:=False
    DiffAgainstBaseline = row
End Function

Private Function EnsureReportSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("DiffReport")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "DiffReport"
    End If
    ws.Cells.Clear
    ws.Range("A1:F1").Value = Array("File", "Sheet", "Cell", "Change", "Baseline", "Compare")
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.Color = RGB(11, 71, 121)
    ws.Rows(1).Font.Color = vbWhite
    Set EnsureReportSheet = ws
End Function

Private Sub WriteRow(ws As Worksheet, r As Long, path As String, sheet As String, _
                     addr As String, change As String, b As Variant, c As Variant)
    ws.Cells(r, 1).Value = Mid(path, InStrRev(path, "\") + 1)
    ws.Cells(r, 2).Value = sheet
    ws.Cells(r, 3).Value = addr
    ws.Cells(r, 4).Value = change
    ws.Cells(r, 5).Value = b
    ws.Cells(r, 6).Value = c
    Select Case change
        Case "Changed": ws.Cells(r, 4).Interior.Color = RGB(255, 245, 200)
        Case "Added":   ws.Cells(r, 4).Interior.Color = RGB(220, 245, 220)
        Case "Removed": ws.Cells(r, 4).Interior.Color = RGB(255, 220, 220)
    End Select
End Sub
