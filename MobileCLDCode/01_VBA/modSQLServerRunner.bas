Attribute VB_Name = "modSQLServerRunner"
'===============================================================================
' modSQLServerRunner
' PURPOSE: Run parameterized SQL queries against SQL Server and land the result
'          set directly into a worksheet. Supports .sql files, stored procs,
'          and parameters taken from named ranges.
'
' WHY THIS IS NOT NATIVE: Power Query can connect to SQL, but it (a) caches
'          results in a model rather than writing raw data to a visible sheet,
'          (b) doesn't easily handle ad-hoc parameterized queries, and (c)
'          cannot run a batch of queries each to its own sheet in one click.
'
' USE CASE (software business):
'   - Finance analyst maintains a library of 30 .sql files (AR aging, churn,
'     license util, etc.) and refreshes them all to Excel sheets every morning.
'   - Sales ops runs the same pipeline query with different region parameters
'     and gets 5 regional sheets in one click.
'
' SETUP:
'   Named range "SQLConnString":
'       Driver={ODBC Driver 17 for SQL Server};Server=tcp:yourserver,1433;
'       Database=YourDB;Trusted_Connection=yes;
'   Sheet "QueryBook":
'       A: Name    B: SQL File Path or inline    C: Destination Sheet
'       D: Parameters (k=v;k=v)    E: Last Run   F: Row Count   G: Elapsed
'===============================================================================
Option Explicit

'-------------------------------------------------------------------------------
' RefreshAllQueries - Iterates the QueryBook and runs every query.
'-------------------------------------------------------------------------------
Public Sub RefreshAllQueries()
    Dim ws As Worksheet, r As Long, lastRow As Long
    Dim t0 As Double, rowsReturned As Long
    Dim sql As String, params As String

    Set ws = ThisWorkbook.Worksheets("QueryBook")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For r = 2 To lastRow
        t0 = Timer
        sql = LoadQueryText(CStr(ws.Cells(r, "B").Value))
        params = CStr(ws.Cells(r, "D").Value)
        sql = ApplyParameters(sql, params)

        rowsReturned = RunSQLToSheet(sql, CStr(ws.Cells(r, "C").Value))

        ws.Cells(r, "E").Value = Format(Now, "yyyy-mm-dd hh:nn")
        ws.Cells(r, "F").Value = rowsReturned
        ws.Cells(r, "G").Value = Format(Timer - t0, "0.00") & "s"
        Application.StatusBar = "Ran " & CStr(ws.Cells(r, "A").Value) & " - " & rowsReturned & " rows"
    Next r

    Application.StatusBar = False
    MsgBox "All queries refreshed.", vbInformation
End Sub

'-------------------------------------------------------------------------------
' RunSQLToSheet - Executes SQL and writes the result set to destSheet.
'-------------------------------------------------------------------------------
Public Function RunSQLToSheet(ByVal sql As String, ByVal destSheet As String) As Long
    Dim conn As Object, rs As Object, ws As Worksheet
    Dim i As Long, row As Long

    Set conn = CreateObject("ADODB.Connection")
    conn.Open GetNamed("SQLConnString")
    conn.CommandTimeout = 180

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 0, 1   ' forward-only, read-only

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(destSheet)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = Left(destSheet, 31)
    End If
    On Error GoTo 0
    ws.Cells.Clear

    ' Write headers
    For i = 0 To rs.Fields.Count - 1
        ws.Cells(1, i + 1).Value = rs.Fields(i).Name
    Next i
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.Color = RGB(11, 71, 121)
    ws.Rows(1).Font.Color = vbWhite

    ' Dump recordset
    If Not rs.EOF Then ws.Cells(2, 1).CopyFromRecordset rs

    row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If row > 1 Then
        ws.Range(ws.Cells(1, 1), ws.Cells(1, rs.Fields.Count)).AutoFilter
        ws.Columns.AutoFit
    End If
    RunSQLToSheet = row - 1

    rs.Close: conn.Close
    Set rs = Nothing: Set conn = Nothing
End Function

'-------------------------------------------------------------------------------
' LoadQueryText - Reads a .sql file, or returns the inline SQL string.
'-------------------------------------------------------------------------------
Private Function LoadQueryText(ByVal spec As String) As String
    If Right(LCase(spec), 4) = ".sql" And Dir(spec) <> "" Then
        Dim fh As Integer, content As String, line As String
        fh = FreeFile
        Open spec For Input As #fh
        Do While Not EOF(fh)
            Line Input #fh, line
            content = content & line & vbCrLf
        Loop
        Close #fh
        LoadQueryText = content
    Else
        LoadQueryText = spec
    End If
End Function

'-------------------------------------------------------------------------------
' ApplyParameters - Replaces @name tokens in the SQL with values from:
'   - named ranges (preferred, keeps them centralized)
'   - or the D-column value: "region='US';year=2026"
'-------------------------------------------------------------------------------
Private Function ApplyParameters(ByVal sql As String, ByVal params As String) As String
    Dim parts() As String, kv() As String, i As Long, val As String
    Dim nm As Variant
    ' Params from named ranges
    For Each nm In ThisWorkbook.Names
        If Left(nm.Name, 6) = "param_" Then
            sql = Replace(sql, "@" & Mid(nm.Name, 7), CStr(nm.RefersToRange.Value))
        End If
    Next nm
    ' Params from the D cell
    If Len(params) > 0 Then
        parts = Split(params, ";")
        For i = LBound(parts) To UBound(parts)
            If InStr(parts(i), "=") > 0 Then
                kv = Split(parts(i), "=", 2)
                sql = Replace(sql, "@" & Trim(kv(0)), Trim(kv(1)))
            End If
        Next i
    End If
    ApplyParameters = sql
End Function

Private Function GetNamed(ByVal nm As String) As String
    On Error Resume Next
    GetNamed = CStr(ThisWorkbook.Names(nm).RefersToRange.Value)
End Function
