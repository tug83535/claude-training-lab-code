Attribute VB_Name = "modTeamsNotifier"
'===============================================================================
' modTeamsNotifier
' PURPOSE: Post rich Adaptive Card messages to a Microsoft Teams channel via
'          an Incoming Webhook. Supports threshold-based alerts from any sheet.
'
' WHY THIS IS NOT NATIVE: Excel cannot post to Teams. OneDrive cannot evaluate
'          a threshold on refresh and alert a channel with a formatted card.
'
' USE CASE (software business):
'   - Alert the finance Teams channel the moment monthly burn exceeds budget
'   - Post a card to the on-call channel when any API SLO falls below target
'   - Notify the sales channel when a pipeline stage age exceeds 30 days
'
' SETUP:
'   1. In Teams, add an "Incoming Webhook" connector to the target channel.
'   2. Paste the URL into the WEBHOOK_URL constant below (or into a named range
'      called "TeamsWebhookUrl" on a hidden Settings sheet).
'===============================================================================
Option Explicit

Private Const WEBHOOK_URL As String = "https://outlook.office.com/webhook/REPLACE-ME"

'-------------------------------------------------------------------------------
' PostTeamsCard - Sends a titled card with an optional fact list and color.
'-------------------------------------------------------------------------------
Public Sub PostTeamsCard(ByVal title As String, _
                         ByVal summary As String, _
                         Optional ByVal themeColor As String = "0B4779", _
                         Optional ByVal factNames As Variant, _
                         Optional ByVal factValues As Variant)
    Dim json As String, http As Object
    json = BuildCardJSON(title, summary, themeColor, factNames, factValues)

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", GetWebhookUrl(), False
    http.setRequestHeader "Content-Type", "application/json"
    http.send json

    If http.Status < 200 Or http.Status >= 300 Then
        MsgBox "Teams post failed: HTTP " & http.Status & vbCrLf & http.responseText, vbExclamation
    End If
End Sub

'-------------------------------------------------------------------------------
' RunThresholdWatcher - Reads a "Watchers" sheet and posts any breaches to Teams.
' Watchers sheet columns:
'   A: Name   B: Cell Ref (e.g. Summary!B12)   C: Operator (>,<,>=,<=,=,<>)
'   D: Threshold   E: Message Template   F: Theme Color   G: Last Alert (written)
'-------------------------------------------------------------------------------
Public Sub RunThresholdWatcher()
    Dim ws As Worksheet, lastRow As Long, r As Long
    Dim val As Double, threshold As Double, op As String
    Dim breached As Boolean, msg As String
    Dim rngRef As Range

    Set ws = ThisWorkbook.Worksheets("Watchers")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For r = 2 To lastRow
        On Error Resume Next
        Set rngRef = Application.Range(CStr(ws.Cells(r, "B").Value))
        If rngRef Is Nothing Then
            ws.Cells(r, "G").Value = "ERROR - bad cell ref"
            GoTo NextWatcher
        End If
        val = CDbl(rngRef.Value)
        threshold = CDbl(ws.Cells(r, "D").Value)
        op = CStr(ws.Cells(r, "C").Value)

        Select Case op
            Case ">":  breached = val > threshold
            Case "<":  breached = val < threshold
            Case ">=": breached = val >= threshold
            Case "<=": breached = val <= threshold
            Case "=":  breached = val = threshold
            Case "<>": breached = val <> threshold
            Case Else: breached = False
        End Select

        If breached Then
            msg = Replace(CStr(ws.Cells(r, "E").Value), "{Value}", Format(val, "#,##0.00"))
            msg = Replace(msg, "{Threshold}", Format(threshold, "#,##0.00"))
            PostTeamsCard CStr(ws.Cells(r, "A").Value), msg, _
                          CStr(ws.Cells(r, "F").Value), _
                          Array("Value", "Threshold", "Operator"), _
                          Array(Format(val, "#,##0.00"), Format(threshold, "#,##0.00"), op)
            ws.Cells(r, "G").Value = "Alerted " & Format(Now, "yyyy-mm-dd hh:nn")
        Else
            ws.Cells(r, "G").Value = "OK " & Format(Now, "yyyy-mm-dd hh:nn")
        End If
NextWatcher:
        Set rngRef = Nothing
        On Error GoTo 0
    Next r
End Sub

Private Function GetWebhookUrl() As String
    On Error Resume Next
    Dim nm As String
    nm = ThisWorkbook.Names("TeamsWebhookUrl").RefersToRange.Value
    If Len(nm) > 10 Then
        GetWebhookUrl = nm
    Else
        GetWebhookUrl = WEBHOOK_URL
    End If
End Function

Private Function BuildCardJSON(title As String, summary As String, color As String, _
                               factNames As Variant, factValues As Variant) As String
    Dim s As String, i As Long
    s = "{""@type"":""MessageCard"",""@context"":""http://schema.org/extensions""," & _
        """themeColor"":""" & color & """,""summary"":""" & JsonEscape(title) & """," & _
        """title"":""" & JsonEscape(title) & """," & _
        """sections"":[{""activityTitle"":""" & JsonEscape(summary) & """"
    If Not IsMissing(factNames) Then
        s = s & ",""facts"":["
        For i = LBound(factNames) To UBound(factNames)
            If i > LBound(factNames) Then s = s & ","
            s = s & "{""name"":""" & JsonEscape(CStr(factNames(i))) & """," & _
                    """value"":""" & JsonEscape(CStr(factValues(i))) & """}"
        Next i
        s = s & "]"
    End If
    s = s & "}]}"
    BuildCardJSON = s
End Function

Private Function JsonEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JsonEscape = s
End Function
