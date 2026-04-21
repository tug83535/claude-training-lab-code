Attribute VB_Name = "modJiraBridge"
'===============================================================================
' modJiraBridge
' PURPOSE: Create, update, and query JIRA tickets directly from Excel via the
'          JIRA Cloud REST API v3.
'
' WHY THIS IS NOT NATIVE: Excel has no JIRA integration. Pulling ticket data
'          from JIRA, or creating tickets in bulk from a spreadsheet, requires
'          HTTP + JSON + Basic auth with an API token.
'
' USE CASE (software business):
'   - Release manager pastes 50 QA bugs into a sheet, clicks one button, and
'     every row becomes a JIRA ticket in the right project with the right labels.
'   - Program manager pulls every open ticket for a sprint into Excel for an
'     ad-hoc pivot that JIRA's native reports can't build.
'
' SETUP:
'   Create a named range "JiraBaseUrl"  = https://yourcompany.atlassian.net
'   Create a named range "JiraEmail"    = your.email@ipipeline.com
'   Create a named range "JiraApiToken" = (from https://id.atlassian.com/manage-profile/security/api-tokens)
'===============================================================================
Option Explicit

'-------------------------------------------------------------------------------
' CreateJiraTicketsFromSheet
' Sheet "JiraCreate" columns:
'   A: Project Key    B: Issue Type   C: Summary      D: Description
'   E: Priority       F: Assignee     G: Labels (csv) H: Result (ticket key)
'-------------------------------------------------------------------------------
Public Sub CreateJiraTicketsFromSheet()
    Dim ws As Worksheet, lastRow As Long, r As Long
    Dim payload As String, response As String, key As String

    Set ws = ThisWorkbook.Worksheets("JiraCreate")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    For r = 2 To lastRow
        If Len(ws.Cells(r, "H").Value) > 0 Then GoTo NextTicket   ' skip already-sent

        payload = BuildCreatePayload( _
            CStr(ws.Cells(r, "A").Value), _
            CStr(ws.Cells(r, "B").Value), _
            CStr(ws.Cells(r, "C").Value), _
            CStr(ws.Cells(r, "D").Value), _
            CStr(ws.Cells(r, "E").Value), _
            CStr(ws.Cells(r, "F").Value), _
            CStr(ws.Cells(r, "G").Value))

        response = JiraPost("/rest/api/3/issue", payload)
        key = ExtractJsonValue(response, "key")
        If Len(key) > 0 Then
            ws.Cells(r, "H").Value = key
            ws.Cells(r, "H").Hyperlinks.Add _
                Anchor:=ws.Cells(r, "H"), _
                Address:=GetNamed("JiraBaseUrl") & "/browse/" & key, _
                TextToDisplay:=key
        Else
            ws.Cells(r, "H").Value = "ERROR: " & Left(response, 200)
        End If
        Application.StatusBar = "Created " & (r - 1) & " of " & (lastRow - 1) & " tickets..."
NextTicket:
    Next r

    Application.StatusBar = False
    MsgBox "Done creating JIRA tickets.", vbInformation
End Sub

'-------------------------------------------------------------------------------
' FetchOpenTicketsToSheet - Uses JQL to pull ticket data into a sheet.
' Example JQL: project = FIN AND status != Done AND assignee = currentUser()
'-------------------------------------------------------------------------------
Public Sub FetchOpenTicketsToSheet(Optional ByVal jql As String = "")
    Dim ws As Worksheet
    Dim response As String, encoded As String
    Dim issues As Variant, i As Long

    If Len(jql) = 0 Then
        jql = InputBox("Enter JQL query:", "JIRA Fetch", _
                       "project = FIN AND status != Done ORDER BY created DESC")
        If Len(jql) = 0 Then Exit Sub
    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("JiraResults")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "JiraResults"
    End If
    On Error GoTo 0

    ws.Cells.Clear
    ws.Range("A1:G1").Value = Array("Key", "Type", "Status", "Priority", _
                                     "Summary", "Assignee", "Updated")
    ws.Range("A1:G1").Font.Bold = True

    encoded = URLEncode(jql)
    response = JiraGet("/rest/api/3/search?jql=" & encoded & "&maxResults=500")

    ' Simple tolerant JSON walk - real code should use a full parser.
    Dim lines() As String, row As Long
    row = 2
    lines = SplitIssues(response)
    For i = 0 To UBound(lines)
        If Len(lines(i)) > 10 Then
            ws.Cells(row, 1).Value = ExtractJsonValue(lines(i), "key")
            ws.Cells(row, 2).Value = ExtractJsonValue(lines(i), "issuetype.name")
            ws.Cells(row, 3).Value = ExtractJsonValue(lines(i), "status.name")
            ws.Cells(row, 4).Value = ExtractJsonValue(lines(i), "priority.name")
            ws.Cells(row, 5).Value = ExtractJsonValue(lines(i), "summary")
            ws.Cells(row, 6).Value = ExtractJsonValue(lines(i), "assignee.displayName")
            ws.Cells(row, 7).Value = ExtractJsonValue(lines(i), "updated")
            row = row + 1
        End If
    Next i

    ws.Columns("A:G").AutoFit
    MsgBox "Fetched " & (row - 2) & " tickets.", vbInformation
End Sub

'-------------------------------------------------------------------------------
' Helpers
'-------------------------------------------------------------------------------
Private Function JiraPost(ByVal path As String, ByVal body As String) As String
    Dim http As Object, url As String
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = GetNamed("JiraBaseUrl") & path
    http.Open "POST", url, False
    http.setRequestHeader "Authorization", "Basic " & BasicAuthHeader()
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    http.send body
    JiraPost = http.responseText
End Function

Private Function JiraGet(ByVal path As String) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", GetNamed("JiraBaseUrl") & path, False
    http.setRequestHeader "Authorization", "Basic " & BasicAuthHeader()
    http.setRequestHeader "Accept", "application/json"
    http.send
    JiraGet = http.responseText
End Function

Private Function BasicAuthHeader() As String
    Dim s As String
    s = GetNamed("JiraEmail") & ":" & GetNamed("JiraApiToken")
    BasicAuthHeader = Base64Encode(s)
End Function

Private Function Base64Encode(ByVal s As String) As String
    Dim xml As Object, node As Object, bytes() As Byte
    bytes = StrConv(s, vbFromUnicode)
    Set xml = CreateObject("MSXML2.DOMDocument")
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = bytes
    Base64Encode = Replace(node.Text, vbLf, "")
End Function

Private Function URLEncode(ByVal s As String) As String
    Dim i As Long, c As String, result As String
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        Select Case Asc(c)
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126
                result = result & c
            Case 32
                result = result & "+"
            Case Else
                result = result & "%" & Right("0" & Hex(Asc(c)), 2)
        End Select
    Next i
    URLEncode = result
End Function

Private Function GetNamed(ByVal nm As String) As String
    On Error Resume Next
    GetNamed = CStr(ThisWorkbook.Names(nm).RefersToRange.Value)
End Function

Private Function BuildCreatePayload(pk As String, itype As String, summ As String, _
                                    desc As String, prio As String, assignee As String, _
                                    labels As String) As String
    Dim labelArr As String, parts() As String, i As Long
    If Len(labels) > 0 Then
        parts = Split(labels, ",")
        For i = LBound(parts) To UBound(parts)
            If i > 0 Then labelArr = labelArr & ","
            labelArr = labelArr & """" & Trim(parts(i)) & """"
        Next i
    End If

    BuildCreatePayload = "{""fields"":{" & _
        """project"":{""key"":""" & pk & """}," & _
        """issuetype"":{""name"":""" & itype & """}," & _
        """summary"":""" & JsonEsc(summ) & """," & _
        """description"":{""type"":""doc"",""version"":1,""content"":[{""type"":""paragraph""," & _
        """content"":[{""type"":""text"",""text"":""" & JsonEsc(desc) & """}]}]}," & _
        IIf(Len(prio) > 0, """priority"":{""name"":""" & prio & """},", "") & _
        IIf(Len(assignee) > 0, """assignee"":{""accountId"":""" & assignee & """},", "") & _
        """labels"":[" & labelArr & "]" & _
        "}}"
End Function

Private Function JsonEsc(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbLf, "\n")
    JsonEsc = s
End Function

Private Function ExtractJsonValue(ByVal json As String, ByVal dotPath As String) As String
    ' Minimal dot-path JSON extractor ("key", "status.name", etc.)
    Dim keys() As String, i As Long, pos As Long, tmp As String
    keys = Split(dotPath, ".")
    tmp = json
    For i = LBound(keys) To UBound(keys)
        pos = InStr(tmp, """" & keys(i) & """")
        If pos = 0 Then Exit Function
        tmp = Mid(tmp, pos + Len(keys(i)) + 2)
        pos = InStr(tmp, ":")
        If pos = 0 Then Exit Function
        tmp = Trim(Mid(tmp, pos + 1))
    Next i
    If Left(tmp, 1) = """" Then
        ExtractJsonValue = Mid(tmp, 2, InStr(2, tmp, """") - 2)
    Else
        pos = InStr(tmp, ",")
        If pos = 0 Then pos = InStr(tmp, "}")
        If pos > 0 Then ExtractJsonValue = Trim(Left(tmp, pos - 1))
    End If
End Function

Private Function SplitIssues(ByVal json As String) As String()
    Dim arr() As String, chunks() As String
    chunks = Split(json, """self""")
    ReDim arr(UBound(chunks) - 1)
    Dim i As Long
    For i = 1 To UBound(chunks)
        arr(i - 1) = """self""" & chunks(i)
    Next i
    SplitIssues = arr
End Function
