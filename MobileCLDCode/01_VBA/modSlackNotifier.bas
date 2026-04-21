Attribute VB_Name = "modSlackNotifier"
'===============================================================================
' modSlackNotifier
' PURPOSE: Post Block Kit messages to Slack channels via Incoming Webhook or
'          via a Bot token (chat.postMessage). Supports rich blocks, mentions,
'          and table-style fields.
'
' WHY THIS IS NOT NATIVE: Excel cannot post to Slack. Power Automate has a
'          Slack connector but is a separate licensing product, not available
'          to every team, and doesn't handle ad-hoc row-by-row posting easily.
'
' USE CASE (software business):
'   - Sales ops: every time a deal flips to Closed Won in Excel, post a
'     celebration card to #sales-wins with the rep, amount, and logo.
'   - Product ops: weekly incident digest - build a Block Kit table and drop
'     it in #eng-weekly from the incidents tracker sheet.
'
' SETUP:
'   Named range "SlackWebhookUrl" = https://hooks.slack.com/services/...
'   OR
'   Named range "SlackBotToken"   = xoxb-...
'   Named range "SlackDefaultChannel" = #sales-wins
'===============================================================================
Option Explicit

'-------------------------------------------------------------------------------
' PostSlackSimple - One-liner. Pass text, optional channel override.
'-------------------------------------------------------------------------------
Public Sub PostSlackSimple(ByVal text As String, Optional ByVal channel As String = "")
    Dim payload As String
    payload = "{""text"":""" & JsonEsc(text) & """"
    If Len(channel) > 0 Then payload = payload & ",""channel"":""" & channel & """"
    payload = payload & "}"
    SendToSlack payload
End Sub

'-------------------------------------------------------------------------------
' PostSlackDealWon - Business-specific helper. Call from a Closed Won macro.
'-------------------------------------------------------------------------------
Public Sub PostSlackDealWon(ByVal repName As String, ByVal customer As String, _
                            ByVal amount As Double, ByVal product As String)
    Dim payload As String
    payload = "{""blocks"":[" & _
        "{""type"":""header"",""text"":{""type"":""plain_text"",""text"":" & _
              """:tada: Closed Won - " & JsonEsc(customer) & """}}," & _
        "{""type"":""section"",""fields"":[" & _
            "{""type"":""mrkdwn"",""text"":""*Rep:*\n" & JsonEsc(repName) & """}," & _
            "{""type"":""mrkdwn"",""text"":""*Amount:*\n$" & Format(amount, "#,##0") & """}," & _
            "{""type"":""mrkdwn"",""text"":""*Product:*\n" & JsonEsc(product) & """}," & _
            "{""type"":""mrkdwn"",""text"":""*Closed:*\n" & Format(Now, "yyyy-mm-dd hh:nn") & """}" & _
            "]}," & _
        "{""type"":""context"",""elements"":[{""type"":""mrkdwn"",""text"":" & _
              """Posted by the Deal Desk automation :robot_face:""}]}" & _
        "]}"
    SendToSlack payload
End Sub

'-------------------------------------------------------------------------------
' PostSheetAsTable - Sends the active sheet (or a named range) as a Block Kit
'                    plain-text table. Great for weekly digests.
'-------------------------------------------------------------------------------
Public Sub PostSheetAsTable(ByVal title As String, Optional ByVal rangeName As String = "")
    Dim rng As Range
    If Len(rangeName) > 0 Then
        Set rng = Application.Range(rangeName)
    Else
        Set rng = ActiveSheet.UsedRange
    End If
    If rng Is Nothing Then Exit Sub

    Dim widths() As Long, r As Long, c As Long
    ReDim widths(1 To rng.Columns.Count)
    For r = 1 To rng.Rows.Count
        For c = 1 To rng.Columns.Count
            If Len(CStr(rng.Cells(r, c).Value)) > widths(c) Then
                widths(c) = Len(CStr(rng.Cells(r, c).Value))
            End If
        Next c
    Next r

    Dim tbl As String
    tbl = "```"
    For r = 1 To rng.Rows.Count
        For c = 1 To rng.Columns.Count
            tbl = tbl & PadRight(CStr(rng.Cells(r, c).Value), widths(c) + 2)
        Next c
        tbl = tbl & vbLf
        If r = 1 Then
            For c = 1 To rng.Columns.Count
                tbl = tbl & String(widths(c) + 2, "-")
            Next c
            tbl = tbl & vbLf
        End If
    Next r
    tbl = tbl & "```"

    Dim payload As String
    payload = "{""blocks"":[" & _
        "{""type"":""header"",""text"":{""type"":""plain_text"",""text"":""" & JsonEsc(title) & """}}," & _
        "{""type"":""section"",""text"":{""type"":""mrkdwn"",""text"":""" & JsonEsc(tbl) & """}}" & _
        "]}"
    SendToSlack payload
End Sub

'-------------------------------------------------------------------------------
' Transport
'-------------------------------------------------------------------------------
Private Sub SendToSlack(ByVal payload As String)
    Dim http As Object, webhook As String, token As String
    Set http = CreateObject("MSXML2.XMLHTTP")

    webhook = GetNamed("SlackWebhookUrl")
    token = GetNamed("SlackBotToken")

    If Len(webhook) > 10 Then
        http.Open "POST", webhook, False
        http.setRequestHeader "Content-Type", "application/json"
        http.send payload
    ElseIf Len(token) > 10 Then
        http.Open "POST", "https://slack.com/api/chat.postMessage", False
        http.setRequestHeader "Authorization", "Bearer " & token
        http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
        If InStr(payload, """channel""") = 0 Then
            payload = Left(payload, 1) & """channel"":""" & _
                      GetNamed("SlackDefaultChannel") & """," & Mid(payload, 2)
        End If
        http.send payload
    Else
        MsgBox "Neither SlackWebhookUrl nor SlackBotToken is set.", vbCritical
        Exit Sub
    End If

    If http.Status >= 400 Then
        Debug.Print "Slack error: " & http.Status & " - " & http.responseText
    End If
End Sub

Private Function GetNamed(ByVal nm As String) As String
    On Error Resume Next
    GetNamed = CStr(ThisWorkbook.Names(nm).RefersToRange.Value)
End Function

Private Function JsonEsc(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JsonEsc = s
End Function

Private Function PadRight(ByVal s As String, ByVal width As Long) As String
    If Len(s) >= width Then
        PadRight = Left(s, width)
    Else
        PadRight = s & String(width - Len(s), " ")
    End If
End Function
