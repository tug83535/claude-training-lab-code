Attribute VB_Name = "modMailMerge_WithAttachments"
'===============================================================================
' modMailMerge_WithAttachments
' PURPOSE: Send personalized emails via Outlook with PER-ROW attachments.
' WHY THIS IS NOT NATIVE: Word mail merge cannot attach different files per row.
'                         Outlook cannot read an Excel list and attach specific
'                         PDFs/invoices to each recipient. This macro bridges both.
'
' USE CASE (software business):
'   Send every customer their own renewal quote PDF, their own usage report,
'   or their own signed SOW — 500+ recipients in one click.
'
' INPUT SHEET LAYOUT (sheet name: "MailMerge"):
'   A: Email        B: First Name   C: Subject      D: Body Template
'   E: Attachment Path (full path or blank)         F: CC           G: BCC
'   H: Status (written back by macro: "Sent" / "Failed - reason")
'
' BODY TEMPLATE TOKENS: {FirstName}, {Today}, {Any column header in { })
'===============================================================================
Option Explicit

Private Const SHEET_NAME As String = "MailMerge"
Private Const HDR_ROW As Long = 1

'-------------------------------------------------------------------------------
' SendAllMerged - Loops every row and sends email with attachment.
'-------------------------------------------------------------------------------
Public Sub SendAllMerged()
    Dim ws As Worksheet, ol As Object, mail As Object
    Dim lastRow As Long, r As Long
    Dim sent As Long, failed As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow <= HDR_ROW Then
        MsgBox "No rows to send.", vbInformation
        Exit Sub
    End If

    If MsgBox("About to send " & (lastRow - HDR_ROW) & " emails via Outlook." & _
             vbCrLf & vbCrLf & "Continue?", vbYesNo + vbQuestion, "Mail Merge") <> vbYes Then
        Exit Sub
    End If

    On Error Resume Next
    Set ol = GetObject(, "Outlook.Application")
    If ol Is Nothing Then Set ol = CreateObject("Outlook.Application")
    On Error GoTo 0
    If ol Is Nothing Then
        MsgBox "Outlook is not available.", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    For r = HDR_ROW + 1 To lastRow
        On Error Resume Next
        Set mail = ol.CreateItem(0)  ' 0 = olMailItem
        mail.To = ws.Cells(r, "A").Value
        mail.CC = ws.Cells(r, "F").Value
        mail.BCC = ws.Cells(r, "G").Value
        mail.Subject = RenderTokens(CStr(ws.Cells(r, "C").Value), ws, r)
        mail.HTMLBody = TextToHTML(RenderTokens(CStr(ws.Cells(r, "D").Value), ws, r))

        Dim att As String
        att = CStr(ws.Cells(r, "E").Value)
        If Len(att) > 0 Then
            If Dir(att) <> "" Then
                mail.Attachments.Add att
            Else
                ws.Cells(r, "H").Value = "Failed - attachment not found: " & att
                failed = failed + 1
                GoTo NextRow
            End If
        End If

        mail.Send
        If Err.Number = 0 Then
            ws.Cells(r, "H").Value = "Sent " & Format(Now, "yyyy-mm-dd hh:nn")
            sent = sent + 1
        Else
            ws.Cells(r, "H").Value = "Failed - " & Err.Description
            failed = failed + 1
        End If
NextRow:
        Err.Clear
        Set mail = Nothing
        On Error GoTo 0
        Application.StatusBar = "Sent " & sent & " of " & (lastRow - HDR_ROW) & "..."
    Next r

    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Done." & vbCrLf & "Sent: " & sent & vbCrLf & "Failed: " & failed, vbInformation
End Sub

'-------------------------------------------------------------------------------
' RenderTokens - Replaces {ColumnHeader} tokens with row values.
'-------------------------------------------------------------------------------
Private Function RenderTokens(ByVal template As String, ws As Worksheet, rowIdx As Long) As String
    Dim lastCol As Long, c As Long, header As String, token As String
    Dim result As String
    result = template
    lastCol = ws.Cells(HDR_ROW, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        header = CStr(ws.Cells(HDR_ROW, c).Value)
        If Len(header) > 0 Then
            token = "{" & header & "}"
            result = Replace(result, token, CStr(ws.Cells(rowIdx, c).Value))
        End If
    Next c
    result = Replace(result, "{Today}", Format(Date, "mmmm d, yyyy"))
    RenderTokens = result
End Function

Private Function TextToHTML(ByVal s As String) As String
    TextToHTML = "<div style='font-family:Arial,sans-serif;font-size:11pt;'>" & _
                 Replace(s, vbCrLf, "<br>") & "</div>"
End Function
