Attribute VB_Name = "modUTL_Core"
Option Explicit

' ============================================================
' KBT Universal Tools — Core Shared Utilities Module
' Shared helper functions used across all UTL modules
' Eliminates duplicate code (TurboOn/Off, SafeDelete, etc.)
' Date: 2026-03-05
' ============================================================
' PUBLIC SUBS/FUNCTIONS:
'   UTL_TurboOn          — Disable screen updating, calculation, events
'   UTL_TurboOff         — Re-enable screen updating, calculation, events
'   UTL_SafeDeleteSheet  — Delete a sheet by name (no error if missing)
'   UTL_LastRow          — Find last used row in a column
'   UTL_LastCol          — Find last used column in a row
'   UTL_SafeNum          — Safe CDbl conversion (returns 0 on error)
'   UTL_SafeStr          — Safe CStr conversion (returns "" on error)
'   UTL_StyleHeader      — Write and style a header row (iPipeline blue)
'   UTL_BackupSheet      — Create a backup copy of a sheet before changes
' ============================================================

' ============================================================
' UTL_TurboOn — Disable screen updates for speed
' ============================================================
Public Sub UTL_TurboOn()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

' ============================================================
' UTL_TurboOff — Re-enable screen updates
' ============================================================
Public Sub UTL_TurboOff()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

' ============================================================
' UTL_SafeDeleteSheet — Delete a sheet by name, no error if missing
' ============================================================
Public Sub UTL_SafeDeleteSheet(ByVal sheetName As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
End Sub

' ============================================================
' UTL_LastRow — Find last used row in a specific column
' ============================================================
Public Function UTL_LastRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    UTL_LastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

' ============================================================
' UTL_LastCol — Find last used column in a specific row
' ============================================================
Public Function UTL_LastCol(ByVal ws As Worksheet, ByVal rowNum As Long) As Long
    UTL_LastCol = ws.Cells(rowNum, ws.Columns.Count).End(xlToLeft).Column
End Function

' ============================================================
' UTL_SafeNum — Safe numeric conversion (returns 0 on error)
' ============================================================
Public Function UTL_SafeNum(ByVal v As Variant) As Double
    On Error Resume Next
    UTL_SafeNum = CDbl(v)
    If Err.Number <> 0 Then UTL_SafeNum = 0
    On Error GoTo 0
End Function

' ============================================================
' UTL_SafeStr — Safe string conversion (returns "" on error)
' ============================================================
Public Function UTL_SafeStr(ByVal v As Variant) As String
    On Error Resume Next
    UTL_SafeStr = Trim(CStr(v))
    If Err.Number <> 0 Then UTL_SafeStr = ""
    On Error GoTo 0
End Function

' ============================================================
' UTL_StyleHeader — Write header labels and apply iPipeline styling
'   ws       — target worksheet
'   rowNum   — header row number
'   headers  — Array of header label strings
' ============================================================
Public Sub UTL_StyleHeader(ByVal ws As Worksheet, ByVal rowNum As Long, _
                           ByRef headers As Variant)
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        Dim col As Long: col = i - LBound(headers) + 1
        ws.Cells(rowNum, col).Value = headers(i)
        ws.Cells(rowNum, col).Font.Bold = True
        ws.Cells(rowNum, col).Interior.Color = RGB(11, 71, 121)   ' iPipeline Blue
        ws.Cells(rowNum, col).Font.Color = RGB(255, 255, 255)
    Next i
End Sub

' ============================================================
' UTL_BackupSheet — Create a backup copy of a sheet
'   Copies the sheet to the end of the workbook with a
'   "_BACKUP_yyyymmdd_hhnnss" suffix. Returns the backup sheet.
'   Use before any destructive operation.
' ============================================================
Public Function UTL_BackupSheet(ByVal ws As Worksheet) As Worksheet
    Dim backupName As String
    backupName = Left(ws.Name, 21) & "_BK_" & Format(Now, "yymmdd_hhnnss")

    ' Ensure name is 31 chars or less
    If Len(backupName) > 31 Then
        backupName = Left(backupName, 31)
    End If

    ws.Copy After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
    Dim newWs As Worksheet
    Set newWs = ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)

    On Error Resume Next
    newWs.Name = backupName
    If Err.Number <> 0 Then
        ' Name conflict — add a random suffix
        Err.Clear
        newWs.Name = Left(backupName, 27) & "_" & Right(Format(Timer * 100, "0000"), 3)
    End If
    On Error GoTo 0

    newWs.Visible = xlSheetHidden
    Set UTL_BackupSheet = newWs
End Function


' ============================================================
' UTL_DetectHeaderRow — Find the most-likely header row on a sheet
' Scans the first maxScanRows rows and returns the row with the
' highest count of non-empty cells. Ties broken by earliest row.
' Returns 1 if the sheet is empty.
' Cherry-picked from Codex comparison (Batch 1, 2026-04-20).
' ============================================================
Public Function UTL_DetectHeaderRow(ByVal ws As Worksheet, _
                                    Optional ByVal maxScanRows As Long = 25) As Long
    UTL_DetectHeaderRow = 1
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then Exit Function

    Dim scanUpTo As Long
    scanUpTo = maxScanRows
    If scanUpTo > lastRow Then scanUpTo = lastRow

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If ws.UsedRange.Columns.Count + ws.UsedRange.Column - 1 > lastCol Then
        lastCol = ws.UsedRange.Columns.Count + ws.UsedRange.Column - 1
    End If
    If lastCol < 1 Then lastCol = 1

    Dim bestRow As Long: bestRow = 1
    Dim bestCount As Long: bestCount = -1
    Dim r As Long, c As Long
    For r = 1 To scanUpTo
        Dim filled As Long: filled = 0
        For c = 1 To lastCol
            If Len(Trim(CStr(ws.Cells(r, c).Value))) > 0 Then filled = filled + 1
        Next c
        If filled > bestCount Then
            bestCount = filled
            bestRow = r
        End If
    Next r

    UTL_DetectHeaderRow = bestRow
End Function
