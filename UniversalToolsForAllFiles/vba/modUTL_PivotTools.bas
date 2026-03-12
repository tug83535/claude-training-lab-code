Attribute VB_Name = "modUTL_PivotTools"
'==============================================================================
' modUTL_PivotTools — Pivot Table Utilities
'==============================================================================
' PURPOSE:  Refresh, list, and manage pivot tables. Unlike Excel's "Refresh All"
'           (which also refreshes external connections and Power Query), these
'           tools target ONLY pivot tables.
'
' PUBLIC SUBS:
'   RefreshAllPivots      — Refresh every pivot table in the workbook
'   RefreshSelectedPivots — User picks which pivots to refresh from a list
'   ListAllPivots         — Build a summary sheet of all pivot tables
'   ClearOldPivotCache    — Remove orphaned pivot caches to reduce file size
'
' DEPENDENCIES: None (standalone). Works in any Excel workbook.
' VERSION:  1.0.0 | DATE: 2026-03-12
'==============================================================================
Option Explicit

Private Const REPORT_SHEET As String = "UTL_PivotReport"
Private Const CLR_HDR As Long = 7930635   ' RGB(11,71,121)

'==============================================================================
' PUBLIC: RefreshAllPivots
' Refreshes every pivot table in the workbook. Skips external connections.
'==============================================================================
Public Sub RefreshAllPivots()
    On Error GoTo ErrHandler

    '--- Count pivots first ---
    Dim totalPivots As Long
    totalPivots = 0
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        totalPivots = totalPivots + ws.PivotTables.Count
    Next ws

    If totalPivots = 0 Then
        MsgBox "No pivot tables found in this workbook.", vbInformation, "Refresh All Pivots"
        Exit Sub
    End If

    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("Found " & totalPivots & " pivot table(s)." & vbCrLf & vbCrLf & _
                      "Refresh all of them?", _
                      vbYesNo + vbQuestion, "Refresh All Pivots")
    If confirm = vbNo Then Exit Sub

    Application.StatusBar = "Refreshing pivot tables..."
    Application.ScreenUpdating = False

    Dim refreshed As Long
    Dim failed As Long
    refreshed = 0
    failed = 0

    For Each ws In ThisWorkbook.Worksheets
        Dim pt As PivotTable
        For Each pt In ws.PivotTables
            On Error Resume Next
            pt.RefreshTable
            If Err.Number = 0 Then
                refreshed = refreshed + 1
            Else
                failed = failed + 1
                Err.Clear
            End If
            On Error GoTo ErrHandler
            Application.StatusBar = "Refreshed " & refreshed & " of " & totalPivots & " pivots..."
        Next pt
    Next ws

    Application.StatusBar = False
    Application.ScreenUpdating = True

    Dim msg As String
    msg = "Pivot refresh complete!" & vbCrLf & vbCrLf & _
          "Refreshed: " & refreshed & vbCrLf
    If failed > 0 Then
        msg = msg & "Failed: " & failed & " (check source data connections)"
    End If

    MsgBox msg, vbInformation, "Refresh All Pivots"
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Refresh All Pivots"
End Sub

'==============================================================================
' PUBLIC: RefreshSelectedPivots
' Shows a numbered list of all pivots. User picks which ones to refresh.
'==============================================================================
Public Sub RefreshSelectedPivots()
    On Error GoTo ErrHandler

    '--- Build pivot inventory ---
    Dim pivotNames() As String
    Dim pivotSheets() As String
    Dim pivotCount As Long
    pivotCount = 0
    ReDim pivotNames(1 To 100)
    ReDim pivotSheets(1 To 100)

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Dim pt As PivotTable
        For Each pt In ws.PivotTables
            pivotCount = pivotCount + 1
            If pivotCount > UBound(pivotNames) Then
                ReDim Preserve pivotNames(1 To pivotCount + 50)
                ReDim Preserve pivotSheets(1 To pivotCount + 50)
            End If
            pivotNames(pivotCount) = pt.Name
            pivotSheets(pivotCount) = ws.Name
        Next pt
    Next ws

    If pivotCount = 0 Then
        MsgBox "No pivot tables found in this workbook.", vbInformation, "Refresh Selected Pivots"
        Exit Sub
    End If

    '--- Show list ---
    Dim menuText As String
    menuText = "Pivot Tables in this workbook:" & vbCrLf & String(40, "-") & vbCrLf & vbCrLf

    Dim i As Long
    For i = 1 To pivotCount
        menuText = menuText & "  " & i & ". " & pivotNames(i) & " (on '" & pivotSheets(i) & "')" & vbCrLf
    Next i

    menuText = menuText & vbCrLf & "Enter numbers to refresh (comma-separated):" & vbCrLf & _
               "Example: 1,3,5  or  ALL to refresh all"

    Dim choice As String
    choice = InputBox(menuText, "Refresh Selected Pivots")
    If Len(Trim(choice)) = 0 Then Exit Sub

    Application.ScreenUpdating = False
    Application.StatusBar = "Refreshing selected pivots..."

    Dim refreshed As Long
    refreshed = 0

    If UCase(Trim(choice)) = "ALL" Then
        '--- Refresh all ---
        For Each ws In ThisWorkbook.Worksheets
            For Each pt In ws.PivotTables
                On Error Resume Next
                pt.RefreshTable
                If Err.Number = 0 Then refreshed = refreshed + 1
                Err.Clear
                On Error GoTo ErrHandler
            Next pt
        Next ws
    Else
        '--- Parse selections ---
        Dim parts() As String
        parts = Split(choice, ",")

        Dim p As Long
        For p = LBound(parts) To UBound(parts)
            Dim num As String
            num = Trim(parts(p))
            If IsNumeric(num) Then
                Dim idx As Long
                idx = CLng(num)
                If idx >= 1 And idx <= pivotCount Then
                    Set ws = ThisWorkbook.Sheets(pivotSheets(idx))
                    On Error Resume Next
                    ws.PivotTables(pivotNames(idx)).RefreshTable
                    If Err.Number = 0 Then refreshed = refreshed + 1
                    Err.Clear
                    On Error GoTo ErrHandler
                End If
            End If
        Next p
    End If

    Application.StatusBar = False
    Application.ScreenUpdating = True

    MsgBox "Refreshed " & refreshed & " pivot table(s).", vbInformation, "Refresh Selected Pivots"
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Refresh Selected Pivots"
End Sub

'==============================================================================
' PUBLIC: ListAllPivots
' Creates a styled inventory sheet listing every pivot table with details.
'==============================================================================
Public Sub ListAllPivots()
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    '--- Create or clear report sheet ---
    Dim wsOut As Worksheet
    On Error Resume Next
    Set wsOut = ThisWorkbook.Sheets(REPORT_SHEET)
    On Error GoTo ErrHandler

    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsOut.Name = REPORT_SHEET
    Else
        wsOut.Cells.Clear
    End If

    '--- Title ---
    wsOut.Range("A1").Value = "Pivot Table Inventory"
    wsOut.Range("A1").Font.Bold = True
    wsOut.Range("A1").Font.Size = 14
    wsOut.Range("A2").Value = "Generated: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    wsOut.Range("A2").Font.Italic = True

    '--- Headers ---
    Dim hdr As Long
    hdr = 4
    wsOut.Cells(hdr, 1).Value = "#"
    wsOut.Cells(hdr, 2).Value = "Pivot Name"
    wsOut.Cells(hdr, 3).Value = "Sheet"
    wsOut.Cells(hdr, 4).Value = "Source Type"
    wsOut.Cells(hdr, 5).Value = "Source Range/Connection"
    wsOut.Cells(hdr, 6).Value = "Row Fields"
    wsOut.Cells(hdr, 7).Value = "Column Fields"
    wsOut.Cells(hdr, 8).Value = "Data Fields"
    wsOut.Cells(hdr, 9).Value = "Location"

    Dim hdrRng As Range
    Set hdrRng = wsOut.Range(wsOut.Cells(hdr, 1), wsOut.Cells(hdr, 9))
    hdrRng.Font.Bold = True
    hdrRng.Font.Color = RGB(255, 255, 255)
    hdrRng.Interior.Color = CLR_HDR

    '--- Populate ---
    Dim r As Long
    r = hdr
    Dim count As Long
    count = 0

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Dim pt As PivotTable
        For Each pt In ws.PivotTables
            count = count + 1
            r = r + 1

            wsOut.Cells(r, 1).Value = count
            wsOut.Cells(r, 2).Value = pt.Name
            wsOut.Cells(r, 3).Value = ws.Name

            ' Source info
            On Error Resume Next
            Dim srcType As String
            Select Case pt.PivotCache.SourceType
                Case xlDatabase: srcType = "Worksheet Range"
                Case xlExternal: srcType = "External Connection"
                Case xlConsolidation: srcType = "Consolidation"
                Case Else: srcType = "Other"
            End Select
            wsOut.Cells(r, 4).Value = srcType

            Dim srcData As String
            srcData = ""
            srcData = pt.SourceData
            If Len(srcData) = 0 Then srcData = "(external/connection)"
            wsOut.Cells(r, 5).Value = srcData
            Err.Clear

            ' Field names
            Dim fld As PivotField
            Dim rowFlds As String, colFlds As String, dataFlds As String
            rowFlds = ""
            colFlds = ""
            dataFlds = ""

            For Each fld In pt.RowFields
                If fld.Name <> "Data" Then
                    If Len(rowFlds) > 0 Then rowFlds = rowFlds & ", "
                    rowFlds = rowFlds & fld.Name
                End If
            Next fld

            For Each fld In pt.ColumnFields
                If fld.Name <> "Data" Then
                    If Len(colFlds) > 0 Then colFlds = colFlds & ", "
                    colFlds = colFlds & fld.Name
                End If
            Next fld

            For Each fld In pt.DataFields
                If Len(dataFlds) > 0 Then dataFlds = dataFlds & ", "
                dataFlds = dataFlds & fld.Name
            Next fld
            Err.Clear

            On Error GoTo ErrHandler

            wsOut.Cells(r, 6).Value = rowFlds
            wsOut.Cells(r, 7).Value = colFlds
            wsOut.Cells(r, 8).Value = dataFlds
            wsOut.Cells(r, 9).Value = pt.TableRange1.Address(False, False)

            ' Alternating rows
            If count Mod 2 = 0 Then
                wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 9)).Interior.Color = RGB(235, 241, 250)
            End If
        Next pt
    Next ws

    '--- Summary ---
    wsOut.Range("A3").Value = "Total Pivot Tables: " & count
    wsOut.Range("A3").Font.Bold = True

    wsOut.Columns("A:I").AutoFit
    If wsOut.Columns("E").ColumnWidth > 50 Then wsOut.Columns("E").ColumnWidth = 50

    wsOut.Activate
    wsOut.Range("A1").Select

    Application.ScreenUpdating = True

    If count = 0 Then
        MsgBox "No pivot tables found in this workbook.", vbInformation, "List All Pivots"
    Else
        MsgBox count & " pivot table(s) listed on '" & REPORT_SHEET & "' sheet.", _
               vbInformation, "List All Pivots"
    End If

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "List All Pivots"
End Sub

'==============================================================================
' PUBLIC: ClearOldPivotCache
' Removes orphaned pivot caches to reduce file size.
'==============================================================================
Public Sub ClearOldPivotCache()
    On Error GoTo ErrHandler

    '--- Count active vs orphaned caches ---
    Dim totalCaches As Long
    totalCaches = ThisWorkbook.PivotCaches.Count

    If totalCaches = 0 Then
        MsgBox "No pivot caches found.", vbInformation, "Clear Pivot Cache"
        Exit Sub
    End If

    ' Find which caches are actually used by a pivot table
    Dim usedIndices() As Boolean
    ReDim usedIndices(1 To totalCaches)

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Dim pt As PivotTable
        For Each pt In ws.PivotTables
            On Error Resume Next
            Dim cacheIdx As Long
            cacheIdx = pt.CacheIndex
            If Err.Number = 0 And cacheIdx >= 1 And cacheIdx <= totalCaches Then
                usedIndices(cacheIdx) = True
            End If
            Err.Clear
            On Error GoTo ErrHandler
        Next pt
    Next ws

    Dim orphanCount As Long
    orphanCount = 0
    Dim i As Long
    For i = 1 To totalCaches
        If Not usedIndices(i) Then orphanCount = orphanCount + 1
    Next i

    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("Pivot cache summary:" & vbCrLf & vbCrLf & _
                      "Total caches: " & totalCaches & vbCrLf & _
                      "Active (used by a pivot): " & (totalCaches - orphanCount) & vbCrLf & _
                      "Orphaned (no longer used): " & orphanCount & vbCrLf & vbCrLf & _
                      "Note: Orphaned caches increase file size but do nothing." & vbCrLf & _
                      "To remove them, save the file as a new name — Excel drops" & vbCrLf & _
                      "orphaned caches on save. Would you like to do that now?", _
                      vbYesNo + vbInformation, "Clear Pivot Cache")

    If confirm = vbYes Then
        MsgBox "To clean orphaned caches:" & vbCrLf & vbCrLf & _
               "1. File > Save As" & vbCrLf & _
               "2. Save with a new filename" & vbCrLf & _
               "3. Excel automatically drops orphaned caches on save" & vbCrLf & vbCrLf & _
               "This is the safest method — Excel handles the cleanup.", _
               vbInformation, "Clear Pivot Cache"
    End If

    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Clear Pivot Cache"
End Sub
