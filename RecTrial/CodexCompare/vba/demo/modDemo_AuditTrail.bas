Attribute VB_Name = "modDemo_AuditTrail"
Option Explicit

Public Sub DemoLog(ByVal ProcedureName As String, ByVal StatusText As String, ByVal MessageText As String)
    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = DemoGetOrCreateAuditSheet()
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(nextRow, 2).Value = Environ$("Username")
    ws.Cells(nextRow, 3).Value = "Demo"
    ws.Cells(nextRow, 4).Value = ProcedureName
    ws.Cells(nextRow, 5).Value = MessageText
    ws.Cells(nextRow, 6).Value = StatusText

    UTL_LogAction "modDemo_AuditTrail", ProcedureName, StatusText, MessageText, 1, 0
End Sub

Public Function DemoGetOrCreateAuditSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DEMO_SHEET_AUDIT)
    On Error GoTo 0

    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets("DEMO_AuditLog")
        On Error GoTo 0

        If ws Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = "DEMO_AuditLog"
            ws.Visible = xlSheetVeryHidden
        End If
    End If

    If Len(Trim$(CStr(ws.Cells(1, 1).Value2))) = 0 Then
        ws.Range("A1:F1").Value = Array("Timestamp", "User", "Module", "Procedure", "Message", "Status")
        ws.Rows(1).Font.Bold = True
    End If

    Set DemoGetOrCreateAuditSheet = ws
End Function
