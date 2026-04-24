Attribute VB_Name = "modUTL_Core"
Option Explicit

Private Const LOG_SHEET_NAME As String = "UTL_RunLog"

Public Type UTL_RunStats
    SheetsTouched As Long
    CellsScanned As Double
    CellsChanged As Double
    Warnings As Long
    StartedAt As Date
    FinishedAt As Date
End Type

Public Function UTL_GetTargetSheets(Optional ByVal IncludeHidden As Boolean = False) As Collection
    Dim ws As Worksheet
    Dim targets As New Collection

    For Each ws In ThisWorkbook.Worksheets
        If IncludeHidden Then
            targets.Add ws
        ElseIf ws.Visible = xlSheetVisible Then
            targets.Add ws
        End If
    Next ws

    Set UTL_GetTargetSheets = targets
End Function

Public Function UTL_LastUsedRow(ByVal ws As Worksheet) As Long
    Dim foundCell As Range

    Set foundCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If foundCell Is Nothing Then
        UTL_LastUsedRow = 1
    Else
        UTL_LastUsedRow = foundCell.Row
    End If
End Function

Public Function UTL_LastUsedColumn(ByVal ws As Worksheet) As Long
    Dim foundCell As Range

    Set foundCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If foundCell Is Nothing Then
        UTL_LastUsedColumn = 1
    Else
        UTL_LastUsedColumn = foundCell.Column
    End If
End Function

Public Function UTL_DetectHeaderRow(ByVal ws As Worksheet, Optional ByVal MaxScanRows As Long = 25) As Long
    Dim r As Long
    Dim c As Long
    Dim usedCols As Long
    Dim currentScore As Long
    Dim bestScore As Long
    Dim bestRow As Long

    usedCols = UTL_LastUsedColumn(ws)
    If usedCols < 1 Then usedCols = 1

    bestScore = -1
    bestRow = 1

    For r = 1 To MaxScanRows
        currentScore = 0
        For c = 1 To usedCols
            If Len(Trim$(CStr(ws.Cells(r, c).Value2))) > 0 Then
                currentScore = currentScore + 1
            End If
        Next c

        If currentScore > bestScore Then
            bestScore = currentScore
            bestRow = r
        End If
    Next r

    UTL_DetectHeaderRow = bestRow
End Function

Public Function UTL_DetectDataRange(ByVal ws As Worksheet, Optional ByVal HeaderRow As Long = 0) As Range
    Dim lastRow As Long
    Dim lastCol As Long

    If HeaderRow = 0 Then HeaderRow = UTL_DetectHeaderRow(ws)

    lastRow = UTL_LastUsedRow(ws)
    lastCol = UTL_LastUsedColumn(ws)

    If lastRow < HeaderRow Then lastRow = HeaderRow
    If lastCol < 1 Then lastCol = 1

    Set UTL_DetectDataRange = ws.Range(ws.Cells(HeaderRow, 1), ws.Cells(lastRow, lastCol))
End Function

Public Sub UTL_EnsureRunLogSheet()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = LOG_SHEET_NAME
        ws.Range("A1:H1").Value = Array("Timestamp", "User", "Module", "Procedure", "Status", "Message", "Sheets", "Cells Changed")
        ws.Rows(1).Font.Bold = True
    End If
End Sub

Public Sub UTL_LogAction(ByVal ModuleName As String, ByVal ProcedureName As String, ByVal Status As String, ByVal Message As String, Optional ByVal SheetsTouched As Long = 0, Optional ByVal CellsChanged As Double = 0)
    Dim ws As Worksheet
    Dim nextRow As Long

    UTL_EnsureRunLogSheet
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET_NAME)

    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(nextRow, 2).Value = Environ$("Username")
    ws.Cells(nextRow, 3).Value = ModuleName
    ws.Cells(nextRow, 4).Value = ProcedureName
    ws.Cells(nextRow, 5).Value = Status
    ws.Cells(nextRow, 6).Value = Message
    ws.Cells(nextRow, 7).Value = SheetsTouched
    ws.Cells(nextRow, 8).Value = CellsChanged
End Sub

Public Sub UTL_ShowCompletion(ByVal FeatureName As String, ByVal StatusMessage As String)
    Application.StatusBar = FeatureName & " — " & StatusMessage
    MsgBox StatusMessage, vbInformation, FeatureName
End Sub

Public Function UTL_IsWorksheetUsable(ByVal ws As Worksheet) As Boolean
    If ws.Name = LOG_SHEET_NAME Then
        UTL_IsWorksheetUsable = False
    ElseIf ws.Visible <> xlSheetVisible Then
        UTL_IsWorksheetUsable = False
    ElseIf UTL_LastUsedRow(ws) = 1 And UTL_LastUsedColumn(ws) = 1 And Len(Trim$(CStr(ws.Cells(1, 1).Value2))) = 0 Then
        UTL_IsWorksheetUsable = False
    Else
        UTL_IsWorksheetUsable = True
    End If
End Function
