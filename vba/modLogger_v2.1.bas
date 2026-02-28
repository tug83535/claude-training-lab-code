Attribute VB_Name = "modLogger"
Option Explicit

'===============================================================================
' modLogger - Runtime Action Logger
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Provides a central logging function used by every other VBA module.
'           Writes timestamped entries to a hidden audit log sheet (VBA_AuditLog).
'           If the sheet does not exist, it is created automatically on first call.
'
' PUBLIC SUBS / FUNCTIONS:
'   LogAction      - Write a log entry (module, procedure, message)
'   ClearLog       - Erase all entries from the audit log sheet
'   ExportLog      - Copy the audit log to a new workbook for review
'   GetLogSheet    - Returns (or creates) the audit log worksheet
'
' CALLED BY:
'   modFormBuilder, modAdmin, modAllocation, modIntegrationTest, and others
'
' LOG SHEET:
'   Sheet name:  SH_LOG = "VBA_AuditLog"  (defined in modConfig)
'   Visibility:  xlSheetVeryHidden — does not appear in the tab bar
'   Columns:     A=Timestamp  B=User  C=Module  D=Procedure  E=Message  F=Status
'
' VERSION:  2.1.0
' DATE:     2026-02-27
' AUTHOR:   iPipeline Finance & Accounting Demo Project
'===============================================================================

' --- Log Sheet Column Layout ---
Private Const COL_TIMESTAMP  As Long = 1   ' A
Private Const COL_USER       As Long = 2   ' B
Private Const COL_MODULE     As Long = 3   ' C
Private Const COL_PROCEDURE  As Long = 4   ' D
Private Const COL_MESSAGE    As Long = 5   ' E
Private Const COL_STATUS     As Long = 6   ' F

Private Const LOG_HEADER_ROW As Long = 1
Private Const LOG_DATA_ROW   As Long = 2   ' First data row (below header)

Private Const MAX_LOG_ROWS   As Long = 5000 ' Auto-trim when exceeded


'===============================================================================
' LogAction - Write one entry to the audit log
'
' Parameters:
'   moduleName  - Name of the VBA module calling this (e.g. "modFormBuilder")
'   procName    - Name of the procedure or action  (e.g. "BuildCommandCenter")
'   message     - Description of what happened     (e.g. "Form created successfully")
'   Optional status - "OK" (default), "ERROR", "WARN", "INFO"
'===============================================================================
Public Sub LogAction(ByVal moduleName As String, _
                     ByVal procName   As String, _
                     ByVal message    As String, _
                     Optional ByVal status As String = "OK")

    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = GetLogSheet()

    ' Find next empty row
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, COL_TIMESTAMP).End(xlUp).Row + 1
    If nextRow < LOG_DATA_ROW Then nextRow = LOG_DATA_ROW

    ' Write the log entry
    ws.Cells(nextRow, COL_TIMESTAMP).Value = Now()
    ws.Cells(nextRow, COL_USER).Value      = Application.UserName
    ws.Cells(nextRow, COL_MODULE).Value    = moduleName
    ws.Cells(nextRow, COL_PROCEDURE).Value = procName
    ws.Cells(nextRow, COL_MESSAGE).Value   = message
    ws.Cells(nextRow, COL_STATUS).Value    = UCase$(status)

    ' Color-code the Status cell for quick visual scanning
    Select Case UCase$(status)
        Case "ERROR"
            ws.Cells(nextRow, COL_STATUS).Interior.Color = RGB(255, 199, 206)  ' Light red
            ws.Cells(nextRow, COL_STATUS).Font.Color     = RGB(156, 0, 6)
        Case "WARN"
            ws.Cells(nextRow, COL_STATUS).Interior.Color = RGB(255, 235, 156)  ' Light yellow
            ws.Cells(nextRow, COL_STATUS).Font.Color     = RGB(156, 87, 0)
        Case "INFO"
            ws.Cells(nextRow, COL_STATUS).Interior.Color = RGB(189, 215, 238)  ' Light blue
            ws.Cells(nextRow, COL_STATUS).Font.Color     = RGB(31, 73, 125)
        Case Else  ' "OK"
            ws.Cells(nextRow, COL_STATUS).Interior.ColorIndex = xlNone
            ws.Cells(nextRow, COL_STATUS).Font.Color = RGB(0, 97, 0)           ' Dark green
    End Select

    ' Auto-trim if the log is getting too large
    If nextRow > MAX_LOG_ROWS + LOG_DATA_ROW Then TrimOldEntries ws

    Exit Sub

ErrHandler:
    ' Logging should never crash the calling macro — silently swallow errors
    On Error Resume Next
    Debug.Print "modLogger.LogAction ERROR: " & Err.Description
    On Error GoTo 0
End Sub


'===============================================================================
' ClearLog - Erase all log data (keeps the header row)
' Called by modAdmin "Clear Audit Log" (Action #43)
'===============================================================================
Public Sub ClearLog()

    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("Clear the entire audit log?" & vbCrLf & vbCrLf & _
                     "All logged entries will be permanently deleted.", _
                     vbYesNo + vbExclamation, APP_NAME)
    If confirm = vbNo Then Exit Sub

    Dim ws As Worksheet
    Set ws = GetLogSheet()

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_TIMESTAMP).End(xlUp).Row

    If lastRow >= LOG_DATA_ROW Then
        ws.Rows(LOG_DATA_ROW & ":" & lastRow).Delete
    End If

    MsgBox "Audit log cleared.", vbInformation, APP_NAME
    LogAction "modLogger", "ClearLog", "Audit log cleared by user"

End Sub


'===============================================================================
' ExportLog - Copy the audit log to a new workbook for archiving / sharing
' Called by modAdmin "Export Audit Log" (Action #42)
'===============================================================================
Public Sub ExportLog()

    Dim ws As Worksheet
    Set ws = GetLogSheet()

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_TIMESTAMP).End(xlUp).Row

    If lastRow < LOG_DATA_ROW Then
        MsgBox "The audit log is empty — nothing to export.", vbInformation, APP_NAME
        Exit Sub
    End If

    ' Copy the entire log sheet to a new workbook
    Dim newWB As Workbook
    ws.Copy   ' Creates new workbook with just this sheet
    Set newWB = ActiveWorkbook

    ' Make the log sheet visible in the export workbook
    newWB.Sheets(1).Visible = xlSheetVisible
    newWB.Sheets(1).Name = "Audit Log Export"

    ' Widen columns for readability
    newWB.Sheets(1).Cells.EntireColumn.AutoFit

    ' Prompt user to save
    Dim savePath As String
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="AuditLog_Export_" & Format(Now, "YYYY-MM-DD"), _
        FileFilter:="Excel Workbook (*.xlsx),*.xlsx", _
        Title:="Save Audit Log Export")

    If savePath <> "False" Then
        Application.DisplayAlerts = False
        newWB.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
        Application.DisplayAlerts = True
        MsgBox "Audit log exported to:" & vbCrLf & savePath, vbInformation, APP_NAME
    Else
        ' User cancelled — close the temp workbook without saving
        Application.DisplayAlerts = False
        newWB.Close SaveChanges:=False
        Application.DisplayAlerts = True
    End If

    LogAction "modLogger", "ExportLog", "Audit log exported"

End Sub


'===============================================================================
' GetLogSheet - Returns the VBA_AuditLog worksheet, creating it if needed
' This is the only place the log sheet is created — always use this function.
'===============================================================================
Public Function GetLogSheet() As Worksheet

    Dim ws As Worksheet
    Dim found As Boolean: found = False

    ' Check if the sheet already exists
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    found = (Not ws Is Nothing)
    On Error GoTo 0

    If found Then
        Set GetLogSheet = ws
        Exit Function
    End If

    ' Sheet doesn't exist — create it now
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = SH_LOG

    ' Write the header row
    With ws
        .Cells(LOG_HEADER_ROW, COL_TIMESTAMP).Value  = "Timestamp"
        .Cells(LOG_HEADER_ROW, COL_USER).Value        = "User"
        .Cells(LOG_HEADER_ROW, COL_MODULE).Value      = "Module"
        .Cells(LOG_HEADER_ROW, COL_PROCEDURE).Value   = "Procedure"
        .Cells(LOG_HEADER_ROW, COL_MESSAGE).Value     = "Message"
        .Cells(LOG_HEADER_ROW, COL_STATUS).Value      = "Status"

        ' Style the header row
        With .Rows(LOG_HEADER_ROW)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 73, 125)  ' Dark blue header
        End With

        ' Format timestamp column
        .Columns(COL_TIMESTAMP).NumberFormat = "yyyy-mm-dd hh:mm:ss"
        .Columns(COL_TIMESTAMP).ColumnWidth = 20

        ' Set column widths
        .Columns(COL_USER).ColumnWidth      = 20
        .Columns(COL_MODULE).ColumnWidth    = 22
        .Columns(COL_PROCEDURE).ColumnWidth = 28
        .Columns(COL_MESSAGE).ColumnWidth   = 55
        .Columns(COL_STATUS).ColumnWidth    = 10

        ' Freeze the header row
        .Rows(LOG_DATA_ROW).Select
        ActiveWindow.FreezePanes = True

        ' Auto-filter for easy searching
        .Rows(LOG_HEADER_ROW).AutoFilter

        ' Hide the sheet — xlSheetVeryHidden hides it from the tab bar
        ' AND from Format > Sheet > Unhide — only VBA can reveal it
        .Visible = xlSheetVeryHidden
    End With

    Set GetLogSheet = ws

End Function


'===============================================================================
' TrimOldEntries - Remove the oldest entries to keep log size manageable
' Deletes the oldest 500 rows when MAX_LOG_ROWS is exceeded
'===============================================================================
Private Sub TrimOldEntries(ws As Worksheet)

    On Error Resume Next
    Const TRIM_ROWS As Long = 500
    ws.Rows(LOG_DATA_ROW & ":" & (LOG_DATA_ROW + TRIM_ROWS - 1)).Delete
    Debug.Print "modLogger: Trimmed " & TRIM_ROWS & " old log entries."
    On Error GoTo 0

End Sub
