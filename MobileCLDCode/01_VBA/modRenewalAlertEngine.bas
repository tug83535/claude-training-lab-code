Attribute VB_Name = "modRenewalAlertEngine"
'===============================================================================
' modRenewalAlertEngine
' PURPOSE: Scan a contract list for upcoming renewals and send personalized
'          Outlook reminders to the contract owner. Escalates as the date nears.
'
' WHY THIS IS NOT NATIVE: Excel conditional formatting highlights dates but
'          can't send per-owner emails. OneDrive notifications only fire on
'          file edit, not on calendar-based thresholds.
'
' USE CASE (software business):
'   Legal + Vendor Management tracks 400+ SaaS, hosting, and service contracts.
'   Each contract has an owner, a renewal date, and a notice-window.
'   This macro finds which contracts fall inside 90/60/30/7-day windows and
'   emails the owner + escalates to their manager after 30 days.
'
' SHEET LAYOUT ("Contracts"):
'   A: Contract ID    B: Vendor    C: Category    D: Annual Value
'   E: Renewal Date   F: Notice Days    G: Owner Email    H: Manager Email
'   I: Last Alert Sent (written)  J: Alert Level (written: 90/60/30/7)
'===============================================================================
Option Explicit

'-------------------------------------------------------------------------------
' ScanAndAlert - Main entry. Call from a button, a Workbook_Open event,
'                or via Windows Task Scheduler + a VBS launcher.
'-------------------------------------------------------------------------------
Public Sub ScanAndAlert()
    Dim ws As Worksheet, r As Long, lastRow As Long
    Dim daysOut As Long, level As Long
    Dim sent As Long, escalated As Long

    Set ws = ThisWorkbook.Worksheets("Contracts")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For r = 2 To lastRow
        If IsDate(ws.Cells(r, "E").Value) Then
            daysOut = DateDiff("d", Date, CDate(ws.Cells(r, "E").Value))
            level = ClassifyLevel(daysOut, CLng(ws.Cells(r, "F").Value))

            If level > 0 And level <> CLng(Val(ws.Cells(r, "J").Value)) Then
                SendRenewalEmail ws, r, daysOut, level
                ws.Cells(r, "I").Value = Now
                ws.Cells(r, "J").Value = level
                sent = sent + 1

                If level = 30 And daysOut <= 30 Then
                    EscalateToManager ws, r, daysOut
                    escalated = escalated + 1
                End If
            End If

            ColorByUrgency ws.Cells(r, "E"), daysOut
        End If
    Next r

    MsgBox "Renewal scan complete." & vbCrLf & _
           "Alerts sent: " & sent & vbCrLf & _
           "Escalations: " & escalated, vbInformation
End Sub

Private Function ClassifyLevel(daysOut As Long, notice As Long) As Long
    ' Return the bucket crossed. Larger number = earlier warning.
    ' Owner only gets the highest bucket they've just crossed into.
    Select Case True
        Case daysOut <= 7:  ClassifyLevel = 7
        Case daysOut <= 30: ClassifyLevel = 30
        Case daysOut <= 60: ClassifyLevel = 60
        Case daysOut <= 90: ClassifyLevel = 90
        Case Else:          ClassifyLevel = 0
    End Select
End Function

Private Sub SendRenewalEmail(ws As Worksheet, r As Long, daysOut As Long, level As Long)
    Dim ol As Object, mail As Object, subj As String, body As String
    Dim vendor As String, value As String, owner As String

    vendor = CStr(ws.Cells(r, "B").Value)
    value = Format(ws.Cells(r, "D").Value, "$#,##0")
    owner = CStr(ws.Cells(r, "G").Value)

    subj = "[" & level & "-day] Contract Renewal: " & vendor & " - " & Format(daysOut, "0") & " days out"

    body = "<div style='font-family:Arial;font-size:11pt;'>" & _
           "<p>Hi,</p>" & _
           "<p>Your <b>" & vendor & "</b> contract (" & ws.Cells(r, "C").Value & _
           ") renews on <b>" & Format(ws.Cells(r, "E").Value, "mmm d, yyyy") & _
           "</b> - that's <b>" & daysOut & " days</b> away.</p>" & _
           "<p><b>Annual Value:</b> " & value & "<br>" & _
           "<b>Notice Required:</b> " & ws.Cells(r, "F").Value & " days</p>" & _
           "<p>Please confirm renewal intent and update any price changes.</p>" & _
           "<p><i>Auto-sent by the Contract Renewal Engine.</i></p></div>"

    On Error Resume Next
    Set ol = GetObject(, "Outlook.Application")
    If ol Is Nothing Then Set ol = CreateObject("Outlook.Application")
    Set mail = ol.CreateItem(0)
    mail.To = owner
    mail.Subject = subj
    mail.HTMLBody = body
    If level <= 30 Then mail.Importance = 2   ' High
    mail.Send
    On Error GoTo 0
End Sub

Private Sub EscalateToManager(ws As Worksheet, r As Long, daysOut As Long)
    Dim ol As Object, mail As Object, mgr As String
    mgr = CStr(ws.Cells(r, "H").Value)
    If Len(mgr) = 0 Then Exit Sub

    On Error Resume Next
    Set ol = GetObject(, "Outlook.Application")
    If ol Is Nothing Then Set ol = CreateObject("Outlook.Application")
    Set mail = ol.CreateItem(0)
    mail.To = mgr
    mail.CC = CStr(ws.Cells(r, "G").Value)
    mail.Subject = "[ESCALATION] " & ws.Cells(r, "B").Value & " renews in " & daysOut & " days"
    mail.HTMLBody = "<div style='font-family:Arial;font-size:11pt;'>" & _
                    "FYI - the renewal decision for <b>" & ws.Cells(r, "B").Value & _
                    "</b> (" & Format(ws.Cells(r, "D").Value, "$#,##0") & ") is now " & _
                    "inside the 30-day window. Please confirm the owner is on track.</div>"
    mail.Importance = 2
    mail.Send
    On Error GoTo 0
End Sub

Private Sub ColorByUrgency(cell As Range, daysOut As Long)
    With cell.Interior
        Select Case True
            Case daysOut <= 7:  .Color = RGB(255, 180, 180)
            Case daysOut <= 30: .Color = RGB(255, 220, 180)
            Case daysOut <= 60: .Color = RGB(255, 245, 200)
            Case daysOut <= 90: .Color = RGB(235, 245, 255)
            Case Else:          .ColorIndex = xlNone
        End Select
    End With
End Sub
