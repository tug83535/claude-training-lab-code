Attribute VB_Name = "modSharePointSync"
'===============================================================================
' modSharePointSync
' PURPOSE: Pull and push SharePoint list data via the SharePoint REST API.
'          Works with modern SharePoint Online lists (not legacy linked tables).
'
' WHY THIS IS NOT NATIVE: Excel's "Export to SharePoint list" is deprecated and
'          one-way. The modern "From SharePoint List" Power Query connector is
'          read-only and needs a full refresh. This macro supports two-way sync
'          via selective POST/PATCH, using the MERGE/IF-MATCH pattern.
'
' USE CASE (software business):
'   - Finance team keeps a master vendor list on SharePoint for Procurement.
'     An Excel workbook pulls the list, edits it offline, and pushes only the
'     changed rows back - without Power Automate or SharePoint Designer.
'   - Legal pulls every open contract from a SharePoint library, adds renewal
'     forecasting columns locally, and syncs the renewal dates back.
'
' SETUP:
'   Named ranges: "SP_Site"     = https://company.sharepoint.com/sites/Finance
'                 "SP_ListName" = Vendor Master
'                 "SP_AccessToken" = Bearer token (issued via Azure AD app reg)
'===============================================================================
Option Explicit

Private Const ROW_TAG_COL As Long = 1  ' __etag written to column A (hidden)
Private Const ROW_ID_COL As Long = 2   ' SharePoint Id
Private Const DATA_START_COL As Long = 3

'-------------------------------------------------------------------------------
' PullListToSheet - Fetches all items into the active sheet.
'-------------------------------------------------------------------------------
Public Sub PullListToSheet()
    Dim ws As Worksheet, url As String, response As String
    Dim items() As String, i As Long, fieldLine() As String, f As Long
    Dim headers As Collection

    Set ws = ThisWorkbook.ActiveSheet
    url = GetNamed("SP_Site") & "/_api/web/lists/getbytitle('" & _
          URLEncode(GetNamed("SP_ListName")) & "')/items?$top=5000"

    response = SharePointGet(url)

    Set headers = GatherFieldNames(response)
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "_etag"
    ws.Cells(1, 2).Value = "Id"
    Dim h As Variant, col As Long
    col = DATA_START_COL
    For Each h In headers
        ws.Cells(1, col).Value = h
        col = col + 1
    Next h
    ws.Columns(1).Hidden = True
    ws.Rows(1).Font.Bold = True

    items = SplitJsonArray(response)
    Dim row As Long: row = 2
    For i = 0 To UBound(items)
        If Len(items(i)) > 10 Then
            ws.Cells(row, 1).Value = ExtractValue(items(i), "__metadata.etag")
            ws.Cells(row, 2).Value = ExtractValue(items(i), "Id")
            For col = DATA_START_COL To col - 1
                ws.Cells(row, col).Value = ExtractValue(items(i), CStr(ws.Cells(1, col).Value))
            Next col
            row = row + 1
        End If
    Next i
    ws.Columns.AutoFit
    MsgBox "Pulled " & (row - 2) & " items.", vbInformation
End Sub

'-------------------------------------------------------------------------------
' PushChangesToList - Pushes rows flagged in a "__dirty" column back to SharePoint.
' Column A holds etag, B holds Id. Edit any row and put "x" in a __dirty column.
'-------------------------------------------------------------------------------
Public Sub PushChangesToList()
    Dim ws As Worksheet, lastRow As Long, lastCol As Long, r As Long, c As Long
    Dim dirtyCol As Long, spID As String, body As String, headerName As String
    Dim updated As Long, created As Long, failed As Long

    Set ws = ThisWorkbook.ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    dirtyCol = FindHeader(ws, "__dirty")
    If dirtyCol = 0 Then
        MsgBox "Add a column with header '__dirty' and mark x on rows to sync.", vbExclamation
        Exit Sub
    End If

    For r = 2 To lastRow
        If LCase(CStr(ws.Cells(r, dirtyCol).Value)) = "x" Then
            body = "{""__metadata"":{""type"":""SP.Data." & SafeListType() & """}"
            For c = DATA_START_COL To lastCol
                headerName = CStr(ws.Cells(1, c).Value)
                If headerName <> "__dirty" And Len(headerName) > 0 And Left(headerName, 1) <> "_" Then
                    body = body & ",""" & headerName & """:" & JsonVal(ws.Cells(r, c).Value)
                End If
            Next c
            body = body & "}"

            spID = CStr(ws.Cells(r, 2).Value)
            If Len(spID) = 0 Then
                ' Insert new item
                If SharePointPost("/_api/web/lists/getbytitle('" & _
                    URLEncode(GetNamed("SP_ListName")) & "')/items", body) Then
                    created = created + 1
                Else
                    failed = failed + 1
                End If
            Else
                ' Update existing (MERGE)
                If SharePointMerge("/_api/web/lists/getbytitle('" & _
                    URLEncode(GetNamed("SP_ListName")) & "')/items(" & spID & ")", _
                    body, CStr(ws.Cells(r, 1).Value)) Then
                    updated = updated + 1
                Else
                    failed = failed + 1
                End If
            End If
            ws.Cells(r, dirtyCol).Value = ""
        End If
    Next r

    MsgBox "Push complete." & vbCrLf & "Updated: " & updated & vbCrLf & _
           "Created: " & created & vbCrLf & "Failed: " & failed, vbInformation
End Sub

'-------------------------------------------------------------------------------
' HTTP wrappers
'-------------------------------------------------------------------------------
Private Function SharePointGet(ByVal fullUrl As String) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", fullUrl, False
    http.setRequestHeader "Authorization", "Bearer " & GetNamed("SP_AccessToken")
    http.setRequestHeader "Accept", "application/json;odata=verbose"
    http.send
    SharePointGet = http.responseText
End Function

Private Function SharePointPost(ByVal path As String, ByVal body As String) As Boolean
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", GetNamed("SP_Site") & path, False
    http.setRequestHeader "Authorization", "Bearer " & GetNamed("SP_AccessToken")
    http.setRequestHeader "Accept", "application/json;odata=verbose"
    http.setRequestHeader "Content-Type", "application/json;odata=verbose"
    http.setRequestHeader "X-RequestDigest", GetFormDigest()
    http.send body
    SharePointPost = (http.Status >= 200 And http.Status < 300)
End Function

Private Function SharePointMerge(ByVal path As String, ByVal body As String, _
                                 ByVal etag As String) As Boolean
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", GetNamed("SP_Site") & path, False
    http.setRequestHeader "Authorization", "Bearer " & GetNamed("SP_AccessToken")
    http.setRequestHeader "Accept", "application/json;odata=verbose"
    http.setRequestHeader "Content-Type", "application/json;odata=verbose"
    http.setRequestHeader "X-RequestDigest", GetFormDigest()
    http.setRequestHeader "X-HTTP-Method", "MERGE"
    http.setRequestHeader "If-Match", IIf(Len(etag) > 0, etag, "*")
    http.send body
    SharePointMerge = (http.Status >= 200 And http.Status < 300)
End Function

Private Function GetFormDigest() As String
    Dim http As Object, response As String
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", GetNamed("SP_Site") & "/_api/contextinfo", False
    http.setRequestHeader "Authorization", "Bearer " & GetNamed("SP_AccessToken")
    http.setRequestHeader "Accept", "application/json;odata=verbose"
    http.send
    GetFormDigest = ExtractValue(http.responseText, "FormDigestValue")
End Function

'-------------------------------------------------------------------------------
' Utilities
'-------------------------------------------------------------------------------
Private Function FindHeader(ws As Worksheet, ByVal name As String) As Long
    Dim c As Long, lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If CStr(ws.Cells(1, c).Value) = name Then
            FindHeader = c
            Exit Function
        End If
    Next c
End Function

Private Function SafeListType() As String
    ' SharePoint requires SP.Data.<ListNameWithCapsAndNoSpaces>ListItem
    Dim s As String
    s = GetNamed("SP_ListName")
    s = Replace(s, " ", "_x0020_")
    SafeListType = s & "ListItem"
End Function

Private Function JsonVal(v As Variant) As String
    If IsNumeric(v) And Not IsDate(v) Then
        JsonVal = CStr(v)
    ElseIf IsDate(v) Then
        JsonVal = """" & Format(v, "yyyy-mm-ddThh:nn:ss") & "Z"""
    Else
        JsonVal = """" & Replace(Replace(CStr(v), "\", "\\"), """", "\""") & """"
    End If
End Function

Private Function GetNamed(ByVal nm As String) As String
    On Error Resume Next
    GetNamed = CStr(ThisWorkbook.Names(nm).RefersToRange.Value)
End Function

Private Function URLEncode(ByVal s As String) As String
    Dim i As Long, c As String, result As String
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        Select Case Asc(c)
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126
                result = result & c
            Case 32
                result = result & "%20"
            Case Else
                result = result & "%" & Right("0" & Hex(Asc(c)), 2)
        End Select
    Next i
    URLEncode = result
End Function

Private Function ExtractValue(ByVal json As String, ByVal dotPath As String) As String
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
        ExtractValue = Mid(tmp, 2, InStr(2, tmp, """") - 2)
    Else
        pos = InStr(tmp, ",")
        If pos = 0 Then pos = InStr(tmp, "}")
        If pos > 0 Then ExtractValue = Trim(Left(tmp, pos - 1))
    End If
End Function

Private Function GatherFieldNames(ByVal json As String) As Collection
    Dim result As New Collection
    Dim candidates As Variant, c As Variant, name As String
    candidates = Array("Title", "Id", "Modified", "Created", "Author", "Status", _
                       "Category", "AssignedTo", "Description", "DueDate", _
                       "Amount", "VendorName", "ContractID", "RenewalDate")
    For Each c In candidates
        name = CStr(c)
        If InStr(json, """" & name & """:") > 0 Then
            On Error Resume Next
            result.Add name, name
            On Error GoTo 0
        End If
    Next c
    Set GatherFieldNames = result
End Function

Private Function SplitJsonArray(ByVal json As String) As String()
    Dim arr() As String, chunks() As String, i As Long
    chunks = Split(json, """__metadata""")
    ReDim arr(UBound(chunks) - 1)
    For i = 1 To UBound(chunks)
        arr(i - 1) = """__metadata""" & chunks(i)
    Next i
    SplitJsonArray = arr
End Function
