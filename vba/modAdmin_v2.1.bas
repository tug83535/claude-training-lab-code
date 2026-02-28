Attribute VB_Name = "modAdmin"
Option Explicit

'===============================================================================
' modAdmin - Auto-Documentation & Change Management
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Generates technical documentation of the workbook (sheet inventory,
'           named ranges, VBA modules) and provides a change management system
'           for tracking modification requests.
'
' PUBLIC SUBS:
'   GenerateDocumentation    - Auto-document workbook structure (Action #36)
'   ShowChangeMenu           - Display change management status (Action #37)
'   AddChangeRequest         - Log a new change request (Action #38)
'   UpdateChangeStatus       - Update status of a change request (Action #39)
'   ChangeManagementSummary  - Summary report of all CRs (Action #40)
'
' DEPENDENCIES: modConfig, modPerformance, modLogger
' VERSION:  2.1.0
'===============================================================================

'===============================================================================
' GenerateDocumentation - Auto-document workbook sheets & structure
'===============================================================================
Public Sub GenerateDocumentation()
    On Error GoTo ErrHandler

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Generating documentation...", 0.05

    ' Create output sheet
    modConfig.SafeDeleteSheet SH_TECH_DOC
    Dim wsDoc As Worksheet
    Set wsDoc = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsDoc.Name = SH_TECH_DOC

    ' Title
    wsDoc.Range("A1").Value = "TECHNICAL DOCUMENTATION"
    wsDoc.Range("A1").Font.Bold = True
    wsDoc.Range("A1").Font.Size = 14
    wsDoc.Range("A1").Font.Color = CLR_NAVY
    wsDoc.Range("A2").Value = "Auto-generated: " & Format(Now, "yyyy-mm-dd hh:mm") & _
        "  |  " & APP_NAME & " v" & APP_VERSION

    ' Section 1: Sheet Inventory
    Dim outRow As Long: outRow = 4
    wsDoc.Cells(outRow, 1).Value = "SHEET INVENTORY"
    wsDoc.Cells(outRow, 1).Font.Bold = True
    wsDoc.Cells(outRow, 1).Font.Size = 12
    outRow = outRow + 1

    modConfig.StyleHeader wsDoc, outRow, _
        Array("Sheet Name", "Visible", "Rows Used", "Cols Used", "Type", "Tab Color")
    outRow = outRow + 1

    Dim ws As Worksheet
    Dim sheetCount As Long: sheetCount = 0
    For Each ws In ThisWorkbook.Worksheets
        modPerformance.UpdateStatus "Documenting: " & ws.Name, sheetCount / ThisWorkbook.Worksheets.Count

        wsDoc.Cells(outRow, 1).Value = ws.Name

        Select Case ws.Visible
            Case xlSheetVisible: wsDoc.Cells(outRow, 2).Value = "Visible"
            Case xlSheetHidden: wsDoc.Cells(outRow, 2).Value = "Hidden"
            Case xlSheetVeryHidden: wsDoc.Cells(outRow, 2).Value = "VeryHidden"
        End Select

        Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim lc As Long: lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        wsDoc.Cells(outRow, 3).Value = lr
        wsDoc.Cells(outRow, 4).Value = lc

        ' Determine sheet type
        Dim shType As String: shType = "Data"
        If InStr(LCase(ws.Name), "p&l") > 0 Or InStr(LCase(ws.Name), "trend") > 0 Then shType = "Report"
        If InStr(LCase(ws.Name), "summary") > 0 Then shType = "Summary"
        If InStr(LCase(ws.Name), "check") > 0 Then shType = "Validation"
        If InStr(LCase(ws.Name), "chart") > 0 Then shType = "Charts"
        If ws.Name = SH_REPORT Then shType = "Dashboard"
        If ws.Name = SH_ASSUMPTIONS Then shType = "Config"
        If ws.Name = SH_DATADICT Then shType = "Reference"
        If ws.Name = SH_LOG Then shType = "Audit Log"
        wsDoc.Cells(outRow, 5).Value = shType

        ' Tab color
        On Error Resume Next
        If ws.Tab.Color <> 0 Then
            wsDoc.Cells(outRow, 6).Interior.Color = ws.Tab.Color
            wsDoc.Cells(outRow, 6).Value = "Custom"
        Else
            wsDoc.Cells(outRow, 6).Value = "Default"
        End If
        On Error GoTo ErrHandler

        If outRow Mod 2 = 0 Then
            wsDoc.Range("A" & outRow & ":F" & outRow).Interior.Color = CLR_ALT_ROW
        End If

        sheetCount = sheetCount + 1
        outRow = outRow + 1
    Next ws

    ' Section 2: Named Ranges
    outRow = outRow + 2
    wsDoc.Cells(outRow, 1).Value = "NAMED RANGES"
    wsDoc.Cells(outRow, 1).Font.Bold = True
    wsDoc.Cells(outRow, 1).Font.Size = 12
    outRow = outRow + 1

    If ThisWorkbook.Names.Count > 0 Then
        modConfig.StyleHeader wsDoc, outRow, Array("Name", "Refers To", "Scope")
        outRow = outRow + 1
        Dim nm As Name
        For Each nm In ThisWorkbook.Names
            wsDoc.Cells(outRow, 1).Value = nm.Name
            On Error Resume Next
            wsDoc.Cells(outRow, 2).Value = "'" & nm.RefersTo
            On Error GoTo ErrHandler
            wsDoc.Cells(outRow, 3).Value = IIf(InStr(nm.Name, "!") > 0, "Sheet", "Workbook")
            outRow = outRow + 1
        Next nm
    Else
        wsDoc.Cells(outRow, 1).Value = "No named ranges defined."
        wsDoc.Cells(outRow, 1).Font.Italic = True
        outRow = outRow + 1
    End If

    ' Section 3: Summary Stats
    outRow = outRow + 2
    wsDoc.Cells(outRow, 1).Value = "WORKBOOK SUMMARY"
    wsDoc.Cells(outRow, 1).Font.Bold = True
    wsDoc.Cells(outRow, 1).Font.Size = 12
    outRow = outRow + 1
    wsDoc.Cells(outRow, 1).Value = "Total Sheets:"
    wsDoc.Cells(outRow, 2).Value = ThisWorkbook.Worksheets.Count
    outRow = outRow + 1
    wsDoc.Cells(outRow, 1).Value = "Named Ranges:"
    wsDoc.Cells(outRow, 2).Value = ThisWorkbook.Names.Count
    outRow = outRow + 1
    wsDoc.Cells(outRow, 1).Value = "File Size:"
    On Error Resume Next
    wsDoc.Cells(outRow, 2).Value = Format(FileLen(ThisWorkbook.FullName) / 1024, "#,##0") & " KB"
    On Error GoTo ErrHandler
    outRow = outRow + 1
    wsDoc.Cells(outRow, 1).Value = "Toolkit Version:"
    wsDoc.Cells(outRow, 2).Value = APP_VERSION

    wsDoc.Columns("A").ColumnWidth = 35
    wsDoc.Columns("B").ColumnWidth = 14
    wsDoc.Columns("C:F").ColumnWidth = 14
    wsDoc.Tab.Color = RGB(112, 48, 160)
    wsDoc.Activate

    Dim elapsed As Double: elapsed = modPerformance.ElapsedSeconds()
    modPerformance.TurboOff

    modLogger.LogAction "modAdmin", "GenerateDocumentation", _
        sheetCount & " sheets documented in " & Format(elapsed, "0.0") & "s"

    MsgBox "Documentation generated on '" & SH_TECH_DOC & "' sheet." & vbCrLf & _
           sheetCount & " sheets documented.", vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modAdmin", "ERROR", Err.Description
    MsgBox "Documentation error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' ShowChangeMenu - Display change management status
'===============================================================================
Public Sub ShowChangeMenu()
    Dim crCount As Long: crCount = GetCRCount()

    MsgBox "CHANGE MANAGEMENT" & vbCrLf & String(25, "=") & vbCrLf & vbCrLf & _
           "Change Requests:  " & crCount & vbCrLf & vbCrLf & _
           "Available Actions:" & vbCrLf & _
           "  #38 - Add Change Request" & vbCrLf & _
           "  #39 - Update CR Status" & vbCrLf & _
           "  #40 - CR Summary Report", _
           vbInformation, APP_NAME

    modLogger.LogAction "modAdmin", "ShowChangeMenu", crCount & " change requests tracked"
End Sub

'===============================================================================
' AddChangeRequest - Log a new change request
'===============================================================================
Public Sub AddChangeRequest()
    On Error GoTo ErrHandler

    Dim crTitle As String
    crTitle = InputBox("Change Request title:" & vbCrLf & _
                       "(e.g., Add DocFast Q2 allocation, Fix Jan variance)", _
                       APP_NAME & " - New Change Request")
    If crTitle = "" Then Exit Sub

    Dim crDesc As String
    crDesc = InputBox("Description / reason for change:", _
                      APP_NAME & " - CR Description")
    If crDesc = "" Then crDesc = "(No description provided)"

    Dim crPriority As String
    crPriority = InputBox("Priority (1=Critical, 2=High, 3=Medium, 4=Low):", _
                          APP_NAME & " - CR Priority", "3")

    ' Ensure change log sheet exists
    Dim wsCR As Worksheet: Set wsCR = EnsureChangeLogSheet()
    Dim nextRow As Long: nextRow = modConfig.LastRow(wsCR, 1) + 1
    If nextRow < 2 Then nextRow = 2

    Dim crNum As Long: crNum = nextRow - 1

    wsCR.Cells(nextRow, 1).Value = "CR-" & Format(crNum, "000")
    wsCR.Cells(nextRow, 2).Value = crTitle
    wsCR.Cells(nextRow, 3).Value = crDesc
    wsCR.Cells(nextRow, 4).Value = Format(Now, "yyyy-mm-dd")
    wsCR.Cells(nextRow, 5).Value = Application.UserName
    wsCR.Cells(nextRow, 6).Value = "Open"
    wsCR.Cells(nextRow, 6).Font.Color = RGB(255, 165, 0)
    wsCR.Cells(nextRow, 7).Value = "P" & crPriority

    modLogger.LogAction "modAdmin", "AddChangeRequest", _
        "CR-" & Format(crNum, "000") & ": " & crTitle

    MsgBox "Change Request logged:" & vbCrLf & vbCrLf & _
           "ID:       CR-" & Format(crNum, "000") & vbCrLf & _
           "Title:    " & crTitle & vbCrLf & _
           "Priority: P" & crPriority & vbCrLf & _
           "Status:   Open", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modLogger.LogAction "modAdmin", "ERROR", Err.Description
    MsgBox "Add CR error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' UpdateChangeStatus - Update the status of an existing CR
'===============================================================================
Public Sub UpdateChangeStatus()
    On Error GoTo ErrHandler

    Dim crCount As Long: crCount = GetCRCount()
    If crCount = 0 Then
        MsgBox "No change requests exist. Use Action #38 first.", vbInformation, APP_NAME
        Exit Sub
    End If

    ' Build CR list
    Dim wsCR As Worksheet: Set wsCR = ThisWorkbook.Worksheets(SH_CHANGE_LOG)
    Dim crList As String: crList = ""
    Dim lr As Long: lr = modConfig.LastRow(wsCR, 1)
    Dim r As Long
    For r = 2 To lr
        If Trim(CStr(wsCR.Cells(r, 1).Value)) <> "" Then
            crList = crList & wsCR.Cells(r, 1).Value & " - " & _
                     wsCR.Cells(r, 2).Value & " [" & wsCR.Cells(r, 6).Value & "]" & vbCrLf
        End If
    Next r

    Dim crID As String
    crID = InputBox("Enter CR ID to update:" & vbCrLf & vbCrLf & crList, _
                    APP_NAME & " - Update CR")
    If crID = "" Then Exit Sub

    ' Find the CR
    Dim found As Boolean: found = False
    For r = 2 To lr
        If LCase(Trim(CStr(wsCR.Cells(r, 1).Value))) = LCase(Trim(crID)) Then
            found = True

            Dim newStatus As String
            newStatus = InputBox("Current status: " & wsCR.Cells(r, 6).Value & vbCrLf & vbCrLf & _
                                 "New status (Open / In Progress / Testing / Closed / Rejected):", _
                                 APP_NAME, CStr(wsCR.Cells(r, 6).Value))
            If newStatus = "" Then Exit Sub

            wsCR.Cells(r, 6).Value = newStatus

            ' Color by status
            Select Case LCase(newStatus)
                Case "open": wsCR.Cells(r, 6).Font.Color = RGB(255, 165, 0)
                Case "in progress": wsCR.Cells(r, 6).Font.Color = RGB(0, 0, 192)
                Case "testing": wsCR.Cells(r, 6).Font.Color = RGB(128, 0, 128)
                Case "closed": wsCR.Cells(r, 6).Font.Color = RGB(0, 128, 0)
                Case "rejected": wsCR.Cells(r, 6).Font.Color = RGB(192, 0, 0)
            End Select

            modLogger.LogAction "modAdmin", "UpdateChangeStatus", crID & " -> " & newStatus
            MsgBox crID & " updated to: " & newStatus, vbInformation, APP_NAME
            Exit For
        End If
    Next r

    If Not found Then MsgBox "CR '" & crID & "' not found.", vbExclamation, APP_NAME
    Exit Sub

ErrHandler:
    modLogger.LogAction "modAdmin", "ERROR", Err.Description
    MsgBox "Update CR error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' ChangeManagementSummary - Show summary of all CRs
'===============================================================================
Public Sub ChangeManagementSummary()
    On Error GoTo ErrHandler

    Dim crCount As Long: crCount = GetCRCount()
    If crCount = 0 Then
        MsgBox "No change requests logged.", vbInformation, APP_NAME
        Exit Sub
    End If

    Dim wsCR As Worksheet: Set wsCR = ThisWorkbook.Worksheets(SH_CHANGE_LOG)
    wsCR.Visible = xlSheetVisible
    wsCR.Activate
    wsCR.Columns("A:G").AutoFit

    ' Count by status
    Dim lr As Long: lr = modConfig.LastRow(wsCR, 1)
    Dim openCount As Long, closedCount As Long, otherCount As Long
    Dim r As Long
    For r = 2 To lr
        Select Case LCase(Trim(CStr(wsCR.Cells(r, 6).Value)))
            Case "open": openCount = openCount + 1
            Case "closed": closedCount = closedCount + 1
            Case Else: otherCount = otherCount + 1
        End Select
    Next r

    modLogger.LogAction "modAdmin", "ChangeManagementSummary", _
        crCount & " CRs: " & openCount & " open, " & closedCount & " closed"

    MsgBox "CR SUMMARY" & vbCrLf & String(20, "=") & vbCrLf & vbCrLf & _
           "Total CRs:   " & crCount & vbCrLf & _
           "Open:        " & openCount & vbCrLf & _
           "Closed:      " & closedCount & vbCrLf & _
           "Other:       " & otherCount & vbCrLf & vbCrLf & _
           "Details on '" & SH_CHANGE_LOG & "' sheet.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    MsgBox "Summary error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' PRIVATE HELPERS
'===============================================================================
Private Function EnsureChangeLogSheet() As Worksheet
    If modConfig.SheetExists(SH_CHANGE_LOG) Then
        Set EnsureChangeLogSheet = ThisWorkbook.Worksheets(SH_CHANGE_LOG)
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = SH_CHANGE_LOG

    modConfig.StyleHeader ws, 1, _
        Array("CR ID", "Title", "Description", "Date", "Requested By", "Status", "Priority")
    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 40
    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 16
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 10
    ws.Visible = xlSheetVeryHidden

    Set EnsureChangeLogSheet = ws
End Function

Private Function GetCRCount() As Long
    If Not modConfig.SheetExists(SH_CHANGE_LOG) Then
        GetCRCount = 0
        Exit Function
    End If
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_CHANGE_LOG)
    Dim lr As Long: lr = modConfig.LastRow(ws, 1)
    If lr < 2 Then
        GetCRCount = 0
    Else
        GetCRCount = lr - 1
    End If
End Function
