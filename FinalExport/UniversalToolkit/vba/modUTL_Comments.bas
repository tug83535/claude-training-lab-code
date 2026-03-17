Attribute VB_Name = "modUTL_Comments"
'==============================================================================
' modUTL_Comments — Comment & Note Manager
'==============================================================================
' PURPOSE:  Extract, summarize, and bulk-manage cell comments/notes.
'           Handles both legacy Comments and threaded Comments (Excel 365).
'
' PUBLIC SUBS:
'   ExtractAllComments   — Export all comments to a summary sheet
'   DeleteSheetComments  — Delete comments from user-selected sheets
'   DeleteAllComments    — Delete all comments in the workbook (with confirmation)
'   CountComments        — Quick count of comments per sheet
'
' DEPENDENCIES: None (standalone). Works in any Excel workbook.
' VERSION:  1.0.0 | DATE: 2026-03-12
'==============================================================================
Option Explicit

Private Const REPORT_SHEET As String = "UTL_CommentReport"
Private Const CLR_HDR As Long = 7948043   ' RGB(11,71,121)

'==============================================================================
' PUBLIC: ExtractAllComments
' Exports every comment in the workbook to a styled summary sheet.
'==============================================================================
Public Sub ExtractAllComments()
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.StatusBar = "Extracting comments..."

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
    wsOut.Range("A1").Value = "Comment Inventory"
    wsOut.Range("A1").Font.Bold = True
    wsOut.Range("A1").Font.Size = 14
    wsOut.Range("A2").Value = "Generated: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    wsOut.Range("A2").Font.Italic = True

    '--- Headers ---
    Dim hdr As Long
    hdr = 4
    wsOut.Cells(hdr, 1).Value = "#"
    wsOut.Cells(hdr, 2).Value = "Sheet"
    wsOut.Cells(hdr, 3).Value = "Cell"
    wsOut.Cells(hdr, 4).Value = "Cell Value"
    wsOut.Cells(hdr, 5).Value = "Comment Author"
    wsOut.Cells(hdr, 6).Value = "Comment Text"

    Dim hdrRng As Range
    Set hdrRng = wsOut.Range(wsOut.Cells(hdr, 1), wsOut.Cells(hdr, 6))
    hdrRng.Font.Bold = True
    hdrRng.Font.Color = RGB(255, 255, 255)
    hdrRng.Interior.Color = CLR_HDR

    '--- Extract comments from all sheets ---
    Dim r As Long
    r = hdr
    Dim count As Long
    count = 0

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = REPORT_SHEET Then GoTo NextSheet

        Dim cmt As Comment
        For Each cmt In ws.Comments
            count = count + 1
            r = r + 1

            wsOut.Cells(r, 1).Value = count
            wsOut.Cells(r, 2).Value = ws.Name
            wsOut.Cells(r, 3).Value = cmt.Parent.Address(False, False)

            On Error Resume Next
            wsOut.Cells(r, 4).Value = Left(CStr(cmt.Parent.Value), 100)
            wsOut.Cells(r, 5).Value = cmt.Author
            wsOut.Cells(r, 6).Value = cmt.Text
            Err.Clear
            On Error GoTo ErrHandler

            ' Alternating rows
            If count Mod 2 = 0 Then
                wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 6)).Interior.Color = RGB(235, 241, 250)
            End If
        Next cmt

NextSheet:
    Next ws

    '--- Summary ---
    wsOut.Range("A3").Value = "Total Comments: " & count
    wsOut.Range("A3").Font.Bold = True

    wsOut.Columns("A:F").AutoFit
    If wsOut.Columns("F").ColumnWidth > 60 Then wsOut.Columns("F").ColumnWidth = 60

    wsOut.Activate
    wsOut.Range("A1").Select

    Application.StatusBar = False
    Application.ScreenUpdating = True

    If count = 0 Then
        MsgBox "No comments found in this workbook.", vbInformation, "Extract Comments"
    Else
        MsgBox count & " comment(s) extracted to '" & REPORT_SHEET & "' sheet.", _
               vbInformation, "Extract Comments"
    End If

    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Extract Comments"
End Sub

'==============================================================================
' PUBLIC: DeleteSheetComments
' User picks which sheets to delete comments from.
'==============================================================================
Public Sub DeleteSheetComments()
    On Error GoTo ErrHandler

    '--- Build sheet list with comment counts ---
    Dim sheetList As String
    sheetList = "Select sheets to delete comments from:" & vbCrLf & String(40, "-") & vbCrLf

    Dim totalComments As Long
    totalComments = 0

    Dim i As Long
    For i = 1 To ThisWorkbook.Worksheets.Count
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(i)
        Dim cmtCount As Long
        cmtCount = ws.Comments.Count
        totalComments = totalComments + cmtCount
        sheetList = sheetList & "  " & i & ". " & ws.Name & " (" & cmtCount & " comments)" & vbCrLf
    Next i

    If totalComments = 0 Then
        MsgBox "No comments found in any sheet.", vbInformation, "Delete Sheet Comments"
        Exit Sub
    End If

    sheetList = sheetList & vbCrLf & "Enter sheet numbers (comma-separated):" & vbCrLf & _
                "Example: 1,3,5  or  ALL for all sheets"

    Dim choice As String
    choice = InputBox(sheetList, "Delete Sheet Comments")
    If Len(Trim(choice)) = 0 Then Exit Sub

    '--- Confirm ---
    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("Are you sure you want to delete comments from the selected sheets?" & vbCrLf & vbCrLf & _
                      "This cannot be undone. Consider running ExtractAllComments first" & vbCrLf & _
                      "to save a copy.", _
                      vbYesNo + vbExclamation, "Delete Sheet Comments")
    If confirm = vbNo Then Exit Sub

    Dim deleted As Long
    deleted = 0

    If UCase(Trim(choice)) = "ALL" Then
        For Each ws In ThisWorkbook.Worksheets
            deleted = deleted + ws.Comments.Count
            ws.Cells.ClearComments
        Next ws
    Else
        Dim parts() As String
        parts = Split(choice, ",")

        Dim p As Long
        For p = LBound(parts) To UBound(parts)
            Dim num As String
            num = Trim(parts(p))
            If IsNumeric(num) Then
                Dim idx As Long
                idx = CLng(num)
                If idx >= 1 And idx <= ThisWorkbook.Worksheets.Count Then
                    Set ws = ThisWorkbook.Worksheets(idx)
                    deleted = deleted + ws.Comments.Count
                    ws.Cells.ClearComments
                End If
            End If
        Next p
    End If

    MsgBox deleted & " comment(s) deleted.", vbInformation, "Delete Sheet Comments"
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Delete Sheet Comments"
End Sub

'==============================================================================
' PUBLIC: DeleteAllComments
' Deletes every comment in the entire workbook after double confirmation.
'==============================================================================
Public Sub DeleteAllComments()
    On Error GoTo ErrHandler

    '--- Count first ---
    Dim totalComments As Long
    totalComments = 0

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        totalComments = totalComments + ws.Comments.Count
    Next ws

    If totalComments = 0 Then
        MsgBox "No comments found in this workbook.", vbInformation, "Delete All Comments"
        Exit Sub
    End If

    '--- First confirmation ---
    Dim confirm1 As VbMsgBoxResult
    confirm1 = MsgBox("This will delete ALL " & totalComments & " comment(s) from EVERY sheet." & vbCrLf & vbCrLf & _
                       "This cannot be undone." & vbCrLf & vbCrLf & _
                       "Tip: Run ExtractAllComments first to save a backup." & vbCrLf & vbCrLf & _
                       "Continue?", _
                       vbYesNo + vbExclamation, "Delete All Comments")
    If confirm1 = vbNo Then Exit Sub

    '--- Second confirmation ---
    Dim confirm2 As String
    confirm2 = InputBox("Type DELETE to confirm removing all " & totalComments & " comments:", _
                         "Delete All Comments - Final Confirmation")
    If UCase(Trim(confirm2)) <> "DELETE" Then
        MsgBox "Cancelled. No comments were deleted.", vbInformation, "Delete All Comments"
        Exit Sub
    End If

    '--- Delete ---
    Dim deleted As Long
    deleted = 0

    For Each ws In ThisWorkbook.Worksheets
        deleted = deleted + ws.Comments.Count
        ws.Cells.ClearComments
    Next ws

    MsgBox deleted & " comment(s) deleted from all sheets.", vbInformation, "Delete All Comments"
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Delete All Comments"
End Sub

'==============================================================================
' PUBLIC: CountComments
' Quick summary of comment counts per sheet.
'==============================================================================
Public Sub CountComments()
    On Error GoTo ErrHandler

    Dim msg As String
    msg = "Comment Count by Sheet:" & vbCrLf & String(35, "-") & vbCrLf & vbCrLf

    Dim total As Long
    total = 0

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Dim cnt As Long
        cnt = ws.Comments.Count
        total = total + cnt
        If cnt > 0 Then
            msg = msg & "  " & ws.Name & ": " & cnt & vbCrLf
        End If
    Next ws

    If total = 0 Then
        msg = msg & "  (no comments found)" & vbCrLf
    End If

    msg = msg & vbCrLf & "Total: " & total & " comment(s)"

    MsgBox msg, vbInformation, "Comment Count"
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Comment Count"
End Sub
