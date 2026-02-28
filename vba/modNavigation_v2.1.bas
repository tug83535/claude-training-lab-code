Attribute VB_Name = "modNavigation"
Option Explicit

'===============================================================================
' modNavigation - Sheet Navigation & Table of Contents
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Quick-jump to any sheet, refresh the Report--> hyperlinks,
'           and provide keyboard-shortcut-driven navigation.
'
' VERSION:  2.1.0
' CHANGES:  v2.0 -> v2.1:
'           + ISSUE-004 (BUG-011): Replaced Application.MacroOptions with
'             Application.OnKey. MacroOptions set Ctrl+H/J/R/M which
'             overwrote Excel built-ins (Find & Replace, etc.). OnKey with
'             Ctrl+Shift combos avoids conflicts.
'           + Ctrl+Shift+M now routes to LaunchCommandCenter (v2.1 entry)
'           + Added ClearShortcuts to unbind on workbook close
'           + Removed unnecessary .Select calls (kept .Activate for navigation
'             commands where sheet activation IS the purpose)
'===============================================================================

'===============================================================================
' RefreshTableOfContents - Rebuild hyperlinks on Report--> sheet
'===============================================================================
Public Sub RefreshTableOfContents()
    On Error GoTo ErrHandler
    
    If Not modConfig.SheetExists(SH_REPORT) Then
        MsgBox "Report--> sheet not found.", vbCritical, APP_NAME
        Exit Sub
    End If
    
    Dim wsReport As Worksheet: Set wsReport = ThisWorkbook.Worksheets(SH_REPORT)
    
    ' Find the TOC area (look for "Sheet" or first hyperlink)
    Dim tocStartRow As Long: tocStartRow = 0
    Dim r As Long
    For r = 1 To 30
        If InStr(1, CStr(wsReport.Cells(r, 1).Value), "Sheet", vbTextCompare) > 0 Or _
           wsReport.Cells(r, 1).Hyperlinks.Count > 0 Then
            tocStartRow = r
            Exit For
        End If
    Next r
    
    If tocStartRow = 0 Then tocStartRow = 8  ' Default position
    
    ' Clear existing TOC entries (but keep header)
    Dim clearEnd As Long: clearEnd = tocStartRow + ThisWorkbook.Worksheets.Count + 5
    wsReport.Range("A" & (tocStartRow + 1) & ":C" & clearEnd).ClearContents
    wsReport.Range("A" & (tocStartRow + 1) & ":C" & clearEnd).Hyperlinks.Delete
    
    ' Write fresh hyperlinks
    Dim row As Long: row = tocStartRow + 1
    Dim ws As Worksheet
    Dim sheetNum As Long: sheetNum = 1
    
    For Each ws In ThisWorkbook.Worksheets
        ' Skip hidden sheets and the log sheet
        If ws.Visible = xlSheetVisible And ws.Name <> SH_LOG Then
            wsReport.Cells(row, 1).Value = sheetNum
            
            wsReport.Hyperlinks.Add _
                Anchor:=wsReport.Cells(row, 2), _
                Address:="", _
                SubAddress:="'" & ws.Name & "'!A1", _
                TextToDisplay:=ws.Name
            
            wsReport.Cells(row, 2).Font.Color = RGB(31, 78, 121)
            wsReport.Cells(row, 2).Font.Underline = xlUnderlineStyleSingle
            
            ' Sheet description
            Select Case ws.Name
                Case SH_ASSUMPTIONS: wsReport.Cells(row, 3).Value = "Driver table & allocation methodology"
                Case SH_DATADICT: wsReport.Cells(row, 3).Value = "Products, departments, vendors reference"
                Case SH_AWS: wsReport.Cells(row, 3).Value = "AWS cost allocation model"
                Case SH_PL_TREND: wsReport.Cells(row, 3).Value = "Consolidated monthly P&L"
                Case SH_PROD_SUMMARY: wsReport.Cells(row, 3).Value = "Product-level P&L & expenses"
                Case SH_FUNC_TREND: wsReport.Cells(row, 3).Value = "Core calculation engine"
                Case SH_NATURAL: wsReport.Cells(row, 3).Value = "Natural expense detail by department"
                Case SH_CHECKS: wsReport.Cells(row, 3).Value = "Cross-sheet reconciliation"
                Case Else
                    If InStr(ws.Name, "Functional P&L Summary") > 0 Then
                        wsReport.Cells(row, 3).Value = "Monthly snapshot"
                    End If
            End Select
            
            sheetNum = sheetNum + 1
            row = row + 1
        End If
    Next ws
    
    wsReport.Columns("A:C").AutoFit
    wsReport.Activate
    
    modLogger.LogAction "modNavigation", "RefreshTableOfContents", (sheetNum - 1) & " sheets linked"
    MsgBox "Table of contents refreshed with " & (sheetNum - 1) & " sheet links.", _
           vbInformation, APP_NAME
    Exit Sub
    
ErrHandler:
    MsgBox "TOC refresh error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' GoHome - Navigate to Report--> sheet
'===============================================================================
Public Sub GoHome()
    If modConfig.SheetExists(SH_REPORT) Then
        Application.GoTo ThisWorkbook.Worksheets(SH_REPORT).Range("A1"), Scroll:=True
    End If
End Sub

'===============================================================================
' QuickJump - Show sheet list and jump to selection
'===============================================================================
Public Sub QuickJump()
    Dim sheetList As String: sheetList = ""
    Dim ws As Worksheet
    Dim i As Long: i = 1
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            sheetList = sheetList & i & ". " & ws.Name & vbCrLf
            i = i + 1
        End If
    Next ws
    
    Dim choice As String
    choice = InputBox("Enter sheet number to navigate:" & vbCrLf & vbCrLf & sheetList, _
                      APP_NAME & " - Quick Jump")
    
    If choice = "" Then Exit Sub
    If Not IsNumeric(choice) Then Exit Sub
    
    Dim idx As Long: idx = CLng(choice)
    Dim j As Long: j = 1
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            If j = idx Then
                Application.GoTo ws.Range("A1"), Scroll:=True
                Exit Sub
            End If
            j = j + 1
        End If
    Next ws
    
    MsgBox "Invalid selection.", vbExclamation, APP_NAME
End Sub

'===============================================================================
' AssignShortcuts - Bind keyboard shortcuts (call from Workbook_Open)
' FIX (v2.1): Uses Application.OnKey with Ctrl+Shift combos instead of
' Application.MacroOptions which was overriding Excel built-in shortcuts
' (Ctrl+H = Find & Replace, Ctrl+R = Fill Right, etc.)
'===============================================================================
Public Sub AssignShortcuts()
    On Error Resume Next
    Application.OnKey "^+h", "GoHome"              ' Ctrl+Shift+H
    Application.OnKey "^+j", "QuickJump"           ' Ctrl+Shift+J
    Application.OnKey "^+r", "modReconciliation.RunAllChecks"  ' Ctrl+Shift+R
    Application.OnKey "^+m", "LaunchCommandCenter" ' Ctrl+Shift+M
    On Error GoTo 0
End Sub

'===============================================================================
' ClearShortcuts - Unbind all custom shortcuts (call from Workbook_BeforeClose)
' Restores default Excel key behavior when the workbook is closed.
'===============================================================================
'===============================================================================
' ToggleExecutiveMode - Hide utility/working sheets, show only report sheets
' Called by ExecuteAction (Action #48)
'===============================================================================
Public Sub ToggleExecutiveMode()
    On Error GoTo ErrHandler

    ' Define which sheets are "executive view" (visible in exec mode)
    Dim execSheets As Variant
    execSheets = Array(SH_REPORT, SH_PL_TREND, SH_PROD_SUMMARY, SH_CHECKS)

    ' Check current state: if working sheets are hidden, we're IN exec mode
    Dim inExecMode As Boolean: inExecMode = False
    If SheetExists(SH_HIDDEN) Then
        inExecMode = (ThisWorkbook.Worksheets(SH_HIDDEN).Visible <> xlSheetVisible)
    End If

    If inExecMode Then
        ' EXIT exec mode — unhide everything
        Dim ws As Worksheet
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name <> SH_LOG Then ws.Visible = xlSheetVisible
        Next ws
        modLogger.LogAction "modNavigation", "ToggleExecutiveMode", "Executive mode OFF — all sheets visible"
        MsgBox "Executive Mode OFF" & vbCrLf & "All sheets are now visible.", vbInformation, APP_NAME
    Else
        ' ENTER exec mode — hide non-executive sheets
        Dim s As Long
        For Each ws In ThisWorkbook.Worksheets
            Dim isExec As Boolean: isExec = False
            For s = LBound(execSheets) To UBound(execSheets)
                If ws.Name = CStr(execSheets(s)) Then isExec = True: Exit For
            Next s
            ' Also keep monthly summary sheets visible
            If InStr(ws.Name, "Functional P&L Summary") > 0 Then isExec = True
            If Not isExec Then ws.Visible = xlSheetVeryHidden
        Next ws

        ' Make sure Report sheet is active
        If SheetExists(SH_REPORT) Then
            ThisWorkbook.Worksheets(SH_REPORT).Activate
        End If
        modLogger.LogAction "modNavigation", "ToggleExecutiveMode", "Executive mode ON — showing report sheets only"
        MsgBox "Executive Mode ON" & vbCrLf & "Only report sheets are visible." & vbCrLf & _
               "Run again to restore all sheets.", vbInformation, APP_NAME
    End If
    Exit Sub

ErrHandler:
    MsgBox "Executive Mode error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' ClearShortcuts - Unbind all custom shortcuts (call from Workbook_BeforeClose)
    On Error Resume Next
    Application.OnKey "^+h"    ' Reset to default
    Application.OnKey "^+j"    ' Reset to default
    Application.OnKey "^+r"    ' Reset to default
    Application.OnKey "^+m"    ' Reset to default
    On Error GoTo 0
End Sub
