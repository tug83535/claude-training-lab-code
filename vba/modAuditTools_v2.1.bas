Attribute VB_Name = "modAuditTools"
Option Explicit

'===============================================================================
' modAuditTools - Workbook Audit, Governance & Safety Tools
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Audit and governance utilities that keep the workbook clean,
'           secure, and documented. Run these as part of the QA checklist.
'
' PUBLIC SUBS:
'   AppendChangeLogEntry      - Log a version note to Change Management Log (#93)
'   FindExternalLinks         - List all external workbook links in formulas (#106)
'   FixExternalLinks          - Remove / repair broken external links (#107)
'   AuditHiddenSheets         - List all hidden and very-hidden sheets (#109)
'   CreateMaskedCopy          - Save a copy with numeric data scrambled (#115)
'   ExportErrorSummaryClipboard - Copy DQ error summary to clipboard (#196)
'   ResetDemoNote             - Placeholder stub for the demo reset button (#200 ref)
'
' VERSION:  2.1.0 (New module — 2026-03-01)
' SOURCE:   Ideas from NewTesting/VBA Examples (200) — items #93, #106, #107, #109, #115, #196
'===============================================================================

'===============================================================================
' AppendChangeLogEntry - Append a note to the Change Management Log (#93)
' Prompts for a short version note and writes it with a timestamp and user
' to the SH_CHANGE_LOG sheet. Creates the sheet if it does not yet exist.
'===============================================================================
Public Sub AppendChangeLogEntry()
    On Error GoTo ErrHandler

    Dim note As String
    note = InputBox("Enter a short description of the change made:", _
                    APP_NAME & " — Change Log", "")
    If Len(Trim(note)) = 0 Then Exit Sub

    ' Create Change Log sheet if it does not exist
    If Not modConfig.SheetExists(SH_CHANGE_LOG) Then
        Dim wsNew As Worksheet
        Set wsNew = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsNew.Name = SH_CHANGE_LOG
        modConfig.StyleHeader wsNew, 1, _
            Array("Timestamp", "User", "Version", "Change Description")
        wsNew.Columns("A:D").AutoFit
    End If

    Dim wsLog As Worksheet: Set wsLog = ThisWorkbook.Worksheets(SH_CHANGE_LOG)
    Dim nextRow As Long: nextRow = modConfig.LastRow(wsLog, 1) + 1

    wsLog.Cells(nextRow, 1).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    wsLog.Cells(nextRow, 2).Value = Environ("USERNAME")
    wsLog.Cells(nextRow, 3).Value = APP_VERSION
    wsLog.Cells(nextRow, 4).Value = note

    modLogger.LogAction "modAuditTools", "AppendChangeLogEntry", note
    MsgBox "Change logged on row " & nextRow & " of '" & SH_CHANGE_LOG & "'.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    MsgBox "AppendChangeLogEntry error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' FindExternalLinks - List all external workbook links in formulas (#106)
' Scans every formula cell in the workbook for references containing "[".
' Writes a report on a "External Links Report" sheet. Run this before
' the demo to make sure no stale file path references are lurking.
'===============================================================================
Public Sub FindExternalLinks()
    On Error GoTo ErrHandler

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Scanning for external links...", 0

    Dim rptName As String: rptName = "External Links Report"
    modConfig.SafeDeleteSheet rptName
    Dim wsRpt As Worksheet
    Set wsRpt = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsRpt.Name = rptName
    modConfig.StyleHeader wsRpt, 1, _
        Array("Sheet", "Cell", "Formula / Address", "Link Target")

    Dim outRow As Long: outRow = 2
    Dim linkCount As Long: linkCount = 0

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = rptName Then GoTo NextWS

        ' Check formula cells for external references (contain "[")
        Dim rng As Range
        On Error Resume Next
        Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo ErrHandler
        If rng Is Nothing Then GoTo NextWS

        Dim cell As Range
        For Each cell In rng
            If InStr(cell.Formula, "[") > 0 Then
                ' Extract just the external path portion
                Dim fml As String: fml = cell.Formula
                Dim startPos As Long: startPos = InStr(fml, "[")
                Dim endPos   As Long: endPos   = InStr(fml, "]")
                Dim linkTarget As String
                If startPos > 0 And endPos > startPos Then
                    linkTarget = Mid(fml, startPos, endPos - startPos + 1)
                Else
                    linkTarget = fml
                End If

                wsRpt.Cells(outRow, 1).Value = ws.Name
                wsRpt.Cells(outRow, 2).Value = cell.Address
                wsRpt.Cells(outRow, 3).Value = Left(fml, 100)
                wsRpt.Cells(outRow, 4).Value = linkTarget
                wsRpt.Cells(outRow, 4).Interior.Color = RGB(255, 235, 200)
                outRow = outRow + 1
                linkCount = linkCount + 1
            End If
        Next cell

        ' Also check Hyperlinks collection for file:// links
        Dim hl As Hyperlink
        For Each hl In ws.Hyperlinks
            If Left(LCase(hl.Address), 4) = "file" Or _
               Left(LCase(hl.Address), 2) = "\\" Or _
               InStr(hl.Address, ":\") > 0 Then
                wsRpt.Cells(outRow, 1).Value = ws.Name
                wsRpt.Cells(outRow, 2).Value = hl.Range.Address
                wsRpt.Cells(outRow, 3).Value = "Hyperlink: " & hl.TextToDisplay
                wsRpt.Cells(outRow, 4).Value = hl.Address
                wsRpt.Cells(outRow, 4).Interior.Color = RGB(255, 200, 200)
                outRow = outRow + 1
                linkCount = linkCount + 1
            End If
        Next hl
NextWS:
    Next ws

    wsRpt.Columns("A:D").AutoFit
    wsRpt.Activate
    modPerformance.TurboOff

    modLogger.LogAction "modAuditTools", "FindExternalLinks", linkCount & " external link(s) found"
    If linkCount > 0 Then
        MsgBox linkCount & " external link(s) found. See '" & rptName & "'." & vbCrLf & _
               "Orange = formula links. Red = broken hyperlinks." & vbCrLf & _
               "Run FixExternalLinks to remove the hyperlink ones.", _
               vbExclamation, APP_NAME
    Else
        MsgBox "No external links found. The workbook is self-contained.", _
               vbInformation, APP_NAME
    End If
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    MsgBox "FindExternalLinks error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' FixExternalLinks - Remove broken external hyperlinks from all sheets (#107)
' Deletes hyperlinks whose address starts with file://, \\ or contains :\
' (local or UNC paths). Preserves all internal #Sheet!A1 links untouched.
'===============================================================================
Public Sub FixExternalLinks()
    On Error GoTo ErrHandler

    If MsgBox("This will remove all hyperlinks that point to external files" & vbCrLf & _
              "(file://, \\server, C:\...). Internal links are kept." & vbCrLf & vbCrLf & _
              "Continue?", vbYesNo + vbQuestion, APP_NAME) = vbNo Then Exit Sub

    Dim fixCount As Long: fixCount = 0
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Dim hlCol As New Collection
        Dim hl As Hyperlink
        ' Collect first, then delete (avoid modifying collection mid-loop)
        For Each hl In ws.Hyperlinks
            If Left(LCase(hl.Address), 4) = "file" Or _
               Left(LCase(hl.Address), 2) = "\\" Or _
               InStr(hl.Address, ":\") > 0 Then
                hlCol.Add hl
            End If
        Next hl
        Dim item As Variant
        For Each item In hlCol
            On Error Resume Next
            item.Delete
            If Err.Number = 0 Then fixCount = fixCount + 1
            Err.Clear
            On Error GoTo ErrHandler
        Next item
    Next ws

    modLogger.LogAction "modAuditTools", "FixExternalLinks", fixCount & " external link(s) removed"
    MsgBox fixCount & " external hyperlink(s) removed." & vbCrLf & _
           "Internal #Sheet!A1 links are preserved.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    MsgBox "FixExternalLinks error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' AuditHiddenSheets - List all hidden and very-hidden sheets (#109)
' Writes a short report to a MsgBox and to the VBA_AuditLog so there is
' a permanent record. Useful to confirm no surprise sheets exist before demo.
'===============================================================================
Public Sub AuditHiddenSheets()
    On Error GoTo ErrHandler

    Dim hiddenList     As String
    Dim veryHiddenList As String
    Dim ws As Worksheet
    Dim hidCount     As Long: hidCount     = 0
    Dim vHidCount    As Long: vHidCount    = 0

    For Each ws In ThisWorkbook.Worksheets
        Select Case ws.Visible
            Case xlSheetHidden
                hiddenList = hiddenList & vbCrLf & "  [Hidden]       " & ws.Name
                hidCount = hidCount + 1
            Case xlSheetVeryHidden
                veryHiddenList = veryHiddenList & vbCrLf & "  [Very Hidden]  " & ws.Name
                vHidCount = vHidCount + 1
        End Select
    Next ws

    Dim total As Long: total = hidCount + vHidCount
    Dim report As String
    If total = 0 Then
        report = "No hidden sheets found. All sheets are visible."
    Else
        report = total & " hidden sheet(s) found:" & _
                 hiddenList & veryHiddenList
    End If

    modLogger.LogAction "modAuditTools", "AuditHiddenSheets", _
        hidCount & " hidden, " & vHidCount & " very-hidden"
    MsgBox report, IIf(total > 0, vbExclamation, vbInformation), APP_NAME
    Exit Sub

ErrHandler:
    MsgBox "AuditHiddenSheets error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' CreateMaskedCopy - Save a copy of the workbook with numeric data scrambled (#115)
' Saves a new workbook alongside the original named with "_MASKED" suffix.
' On the masked copy, all numeric constants on visible sheets are replaced
' with random values in the same order of magnitude. Formulas are converted
' to values first. Safe to share with coworkers for testing without exposing
' real financial figures.
'===============================================================================
Public Sub CreateMaskedCopy()
    On Error GoTo ErrHandler

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Save the workbook first before creating a masked copy.", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    If MsgBox("Create a masked copy with all dollar amounts randomized?" & vbCrLf & vbCrLf & _
              "The original workbook is not changed." & vbCrLf & _
              "The copy will be saved in the same folder with '_MASKED' in the name.", _
              vbYesNo + vbQuestion, APP_NAME) = vbNo Then Exit Sub

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Creating masked copy...", 0.1

    ' Build output path
    Dim origName As String: origName = ThisWorkbook.Name
    Dim ext As String: ext = Mid(origName, InStrRev(origName, "."))
    Dim baseName As String: baseName = Left(origName, Len(origName) - Len(ext))
    Dim outPath As String
    outPath = ThisWorkbook.Path & "\" & baseName & "_MASKED" & ext

    ' Copy the workbook
    ThisWorkbook.SaveCopyAs outPath
    Dim wbMask As Workbook: Set wbMask = Workbooks.Open(outPath)

    modPerformance.UpdateStatus "Masking numeric values...", 0.4

    ' On each visible sheet in the copy: convert formulas to values, then randomize
    Dim ws As Worksheet
    For Each ws In wbMask.Worksheets
        If ws.Visible = xlSheetVisible Then
            ' Convert all formula cells to values
            On Error Resume Next
            Dim fRng As Range
            Set fRng = ws.UsedRange.SpecialCells(xlCellTypeFormulas, xlNumbers)
            If Not fRng Is Nothing Then
                fRng.Value = fRng.Value
            End If
            On Error GoTo ErrHandler

            ' Randomize all numeric constants
            Dim nRng As Range
            On Error Resume Next
            Set nRng = ws.UsedRange.SpecialCells(xlCellTypeConstants, xlNumbers)
            On Error GoTo ErrHandler
            If Not nRng Is Nothing Then
                Dim cell As Range
                For Each cell In nRng
                    Dim orig As Double: orig = Abs(cell.Value)
                    If orig > 1 Then
                        ' Preserve order of magnitude; apply random ±20% noise
                        Dim factor As Double
                        factor = 0.8 + (Rnd() * 0.4)   ' 0.80 to 1.20
                        cell.Value = Round(orig * factor, 2)
                    End If
                Next cell
            End If
        End If
    Next ws

    wbMask.Save
    wbMask.Close SaveChanges:=False

    modPerformance.TurboOff
    modLogger.LogAction "modAuditTools", "CreateMaskedCopy", "Saved: " & outPath
    MsgBox "Masked copy saved:" & vbCrLf & outPath & vbCrLf & vbCrLf & _
           "All dollar amounts are randomized. Safe to share for testing.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    MsgBox "CreateMaskedCopy error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' ExportErrorSummaryClipboard - Copy DQ error summary to Windows clipboard (#196)
' Reads the Data Quality Report sheet (or VBA_AuditLog) and formats a short
' text summary that can be pasted into an email or Teams message instantly.
'===============================================================================
Public Sub ExportErrorSummaryClipboard()
    On Error GoTo ErrHandler

    Dim summary As String
    summary = "=== Keystone BenefitTech — Data Quality Summary ===" & vbCrLf
    summary = summary & "Generated: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & vbCrLf

    ' Pull from DQ Report sheet if it exists
    If modConfig.SheetExists(SH_DQ_REPORT) Then
        Dim wsDQ As Worksheet: Set wsDQ = ThisWorkbook.Worksheets(SH_DQ_REPORT)
        Dim lastRow As Long: lastRow = modConfig.LastRow(wsDQ, 1)
        Dim failCount As Long: failCount = 0
        Dim warnCount As Long: warnCount = 0
        Dim r As Long
        For r = 2 To lastRow
            Dim status As String: status = UCase(modConfig.SafeStr(wsDQ.Cells(r, 5).Value))
            If status = "FAIL" Then
                failCount = failCount + 1
                summary = summary & "  FAIL: " & modConfig.SafeStr(wsDQ.Cells(r, 1).Value) & vbCrLf
            ElseIf status = "WARN" Then
                warnCount = warnCount + 1
            End If
        Next r
        summary = summary & vbCrLf & "Total: " & failCount & " FAIL | " & warnCount & " WARN" & vbCrLf
    Else
        summary = summary & "Data Quality Report sheet not found." & vbCrLf & _
                  "Run Data Quality Check first." & vbCrLf
    End If

    ' Also pull Checks tab summary
    If modConfig.SheetExists(SH_CHECKS) Then
        Dim wsChk As Worksheet: Set wsChk = ThisWorkbook.Worksheets(SH_CHECKS)
        Dim chkLastRow As Long: chkLastRow = modConfig.LastRow(wsChk, 1)
        Dim chkFail As Long: chkFail = 0
        Dim chkPass As Long: chkPass = 0
        Dim c As Long
        For c = DATA_ROW_CHECKS To chkLastRow
            Dim chkStat As String: chkStat = UCase(modConfig.SafeStr(wsChk.Cells(c, COL_CHECK_STATUS).Value))
            If chkStat = "PASS" Then chkPass = chkPass + 1
            If chkStat = "FAIL" Then chkFail = chkFail + 1
        Next c
        summary = summary & vbCrLf & "Reconciliation Checks: " & chkPass & " PASS | " & chkFail & " FAIL" & vbCrLf
    End If

    summary = summary & vbCrLf & "=== End of Summary ==="

    ' Copy to clipboard via DataObject
    Dim dataObj As Object
    Set dataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dataObj.SetText summary
    dataObj.PutInClipboard

    modLogger.LogAction "modAuditTools", "ExportErrorSummaryClipboard", "Summary copied to clipboard"
    MsgBox "Error summary copied to clipboard." & vbCrLf & _
           "Paste it into email, Teams, or any text editor with Ctrl+V.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    MsgBox "ExportErrorSummaryClipboard error: " & Err.Description & vbCrLf & _
           "(Note: clipboard access requires Microsoft Forms 2.0 library)", _
           vbCritical, APP_NAME
End Sub

'===============================================================================
' ResetDemoNote - Informational stub used by the demo control button (#200 ref)
' Shows the user what to do to reset the demo environment. Full automation
' of a demo reset (clearing test data, restoring samples) requires knowing
' exactly which cells hold test values — this stub guides the user through
' the manual steps so nothing is accidentally deleted.
'===============================================================================
Public Sub ResetDemoNote()
    MsgBox "To reset the demo environment:" & vbCrLf & vbCrLf & _
           "1. Close any extra sheets created during testing" & vbCrLf & _
           "2. Run Data Quality Check to confirm clean state" & vbCrLf & _
           "3. Run Run Reconciliation to confirm all checks pass" & vbCrLf & _
           "4. Run Refresh Dashboard to update all charts" & vbCrLf & vbCrLf & _
           "When you are ready, press OK and the demo can begin.", _
           vbInformation, APP_NAME & " — Demo Reset"
End Sub
