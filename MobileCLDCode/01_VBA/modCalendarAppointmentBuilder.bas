Attribute VB_Name = "modCalendarAppointmentBuilder"
'===============================================================================
' modCalendarAppointmentBuilder
' PURPOSE: Create Outlook calendar appointments, meetings with invitees, and
'          recurring blocks directly from rows in an Excel sheet.
'
' WHY THIS IS NOT NATIVE: Outlook won't read Excel rows as appointment sources.
'          Copy-pasting into "Quick Event" one row at a time is how everyone
'          currently does it. This macro does 200 in one click.
'
' USE CASE (software business):
'   - Implementation manager plans a 6-month customer rollout: 48 kickoff
'     meetings, 48 UAT reviews, 48 go-lives. One sheet, one click.
'   - PMO builds the company-wide training calendar from a roster of 120 sessions.
'
' SHEET LAYOUT ("Appointments"):
'   A: Subject           B: Start (date+time)   C: End (date+time)
'   D: Location          E: Body/Agenda         F: Required Attendees (csv)
'   G: Optional Attendees (csv)                 H: Reminder (minutes)
'   I: Recurrence (none|daily|weekly|monthly)  J: Recur Until (date)
'   K: Category (color)  L: Result (written back)
'===============================================================================
Option Explicit

' Outlook enums we need (using literals so we don't need early binding)
Private Const olAppointmentItem As Long = 1
Private Const olMeeting As Long = 1
Private Const olRecursDaily As Long = 0
Private Const olRecursWeekly As Long = 1
Private Const olRecursMonthly As Long = 2

Public Sub CreateAllAppointments()
    Dim ws As Worksheet, ol As Object, r As Long, lastRow As Long
    Dim appt As Object, rec As Object
    Dim created As Long, failed As Long

    Set ws = ThisWorkbook.Worksheets("Appointments")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    On Error Resume Next
    Set ol = GetObject(, "Outlook.Application")
    If ol Is Nothing Then Set ol = CreateObject("Outlook.Application")
    On Error GoTo 0
    If ol Is Nothing Then
        MsgBox "Outlook is not available.", vbCritical
        Exit Sub
    End If

    For r = 2 To lastRow
        On Error Resume Next
        Set appt = ol.CreateItem(olAppointmentItem)
        appt.Subject = CStr(ws.Cells(r, "A").Value)
        appt.Start = CDate(ws.Cells(r, "B").Value)
        appt.End = CDate(ws.Cells(r, "C").Value)
        appt.Location = CStr(ws.Cells(r, "D").Value)
        appt.Body = CStr(ws.Cells(r, "E").Value)
        appt.ReminderMinutesBeforeStart = IIf(IsNumeric(ws.Cells(r, "H").Value), _
                                              CLng(ws.Cells(r, "H").Value), 15)
        appt.ReminderSet = True

        ' Meeting with attendees
        If Len(ws.Cells(r, "F").Value) > 0 Or Len(ws.Cells(r, "G").Value) > 0 Then
            appt.MeetingStatus = olMeeting
            AddAttendees appt, CStr(ws.Cells(r, "F").Value), 1   ' required
            AddAttendees appt, CStr(ws.Cells(r, "G").Value), 2   ' optional
        End If

        ' Recurrence
        Dim rtype As String: rtype = LCase(CStr(ws.Cells(r, "I").Value))
        If rtype <> "" And rtype <> "none" Then
            Set rec = appt.GetRecurrencePattern
            Select Case rtype
                Case "daily":   rec.RecurrenceType = olRecursDaily
                Case "weekly":  rec.RecurrenceType = olRecursWeekly
                Case "monthly": rec.RecurrenceType = olRecursMonthly
            End Select
            rec.PatternStartDate = CDate(ws.Cells(r, "B").Value)
            If IsDate(ws.Cells(r, "J").Value) Then
                rec.PatternEndDate = CDate(ws.Cells(r, "J").Value)
            End If
        End If

        ' Category
        If Len(ws.Cells(r, "K").Value) > 0 Then
            appt.Categories = CStr(ws.Cells(r, "K").Value)
        End If

        appt.Send   ' for meetings; for solo appt use .Save instead
        If Err.Number <> 0 Then appt.Save
        If Err.Number = 0 Then
            ws.Cells(r, "L").Value = "Created " & Format(Now, "yyyy-mm-dd hh:nn")
            created = created + 1
        Else
            ws.Cells(r, "L").Value = "Failed: " & Err.Description
            failed = failed + 1
        End If
        Err.Clear
        On Error GoTo 0
    Next r

    MsgBox "Done." & vbCrLf & "Created: " & created & vbCrLf & "Failed: " & failed, vbInformation
End Sub

Private Sub AddAttendees(appt As Object, ByVal csv As String, ByVal recType As Long)
    If Len(csv) = 0 Then Exit Sub
    Dim parts() As String, i As Long, rec As Object
    parts = Split(csv, ",")
    For i = LBound(parts) To UBound(parts)
        Set rec = appt.Recipients.Add(Trim(parts(i)))
        rec.Type = recType
    Next i
End Sub
