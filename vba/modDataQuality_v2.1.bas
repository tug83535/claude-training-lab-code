Attribute VB_Name = "modDataQuality"
Option Explicit

'===============================================================================
' modDataQuality - Data Cleaning Scanner & Fixer
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Scan the CrossfireHiddenWorksheet for all known data quality issues,
'           report findings, and optionally fix them.
'
' VERSION:  2.1.0
' CHANGES:  v2.0 -> v2.1:
'           + ISSUE-003 (BUG-018): FixTextNumbers was iterating the ENTIRE
'             workbook and converting every text-stored number — including GL
'             IDs ("001"), fiscal year references ("2025"), and column headers.
'             Now uses a pre-flagged m_TextNumberCells collection populated
'             during ScanAll/ScanTextNumbers/ScanAssumptionsTextValues.
'             FixTextNumbers refuses to run unless ScanAll has been run first.
'           + Uses centralized SH_DQ_REPORT from modConfig instead of
'             Private Const REPORT_SHEET
'           + Uses modConfig.SafeDeleteSheet and StyleHeader helpers
'===============================================================================

Private Type DQIssue
    Sheet       As String
    CellRef     As String
    IssueType   As String
    CurrentVal  As String
    Severity    As String
    FixAvail    As Boolean
End Type

Private m_Issues() As DQIssue
Private m_IssueCount As Long

' v2.1: Track cells flagged as text-stored numbers during scan.
' FixTextNumbers will ONLY convert cells in this collection.
Private m_TextNumberCells As Collection

'===============================================================================
' ScanAll - Full workbook data quality scan
'===============================================================================
Public Sub ScanAll()
    On Error GoTo ErrHandler
    
    m_IssueCount = 0
    Erase m_Issues
    
    ' v2.1: Initialize the text-number tracking collection fresh each scan
    Set m_TextNumberCells = New Collection
    
    modPerformance.TurboOn
    modPerformance.UpdateStatus "Scanning for data quality issues...", 0
    
    ' 1. Scan CrossfireHiddenWorksheet
    If modConfig.SheetExists(SH_HIDDEN) Then
        Dim wsHidden As Worksheet: Set wsHidden = ThisWorkbook.Worksheets(SH_HIDDEN)
        wsHidden.Visible = xlSheetVisible
        
        modPerformance.UpdateStatus "Scanning staging data...", 0.1
        ScanDuplicateRows wsHidden
        
        modPerformance.UpdateStatus "Checking date formats...", 0.25
        ScanMixedDates wsHidden
        
        modPerformance.UpdateStatus "Finding text-stored numbers...", 0.4
        ScanTextNumbers wsHidden
        
        wsHidden.Visible = xlSheetHidden
    End If
    
    ' 2. Scan Assumptions for text-stored numbers
    If modConfig.SheetExists(SH_ASSUMPTIONS) Then
        modPerformance.UpdateStatus "Scanning assumptions...", 0.55
        ScanAssumptionsTextValues
    End If
    
    ' 3. Scan Product Line Summary for misspelling
    If modConfig.SheetExists(SH_PROD_SUMMARY) Then
        modPerformance.UpdateStatus "Checking product names...", 0.7
        ScanMisspellings
    End If
    
    ' 4. Scan Natural P&L for blank AWS cells
    If modConfig.SheetExists(SH_NATURAL) Then
        modPerformance.UpdateStatus "Checking for blank cells...", 0.85
        ScanBlankCells
    End If
    
    modPerformance.UpdateStatus "Generating report...", 0.95
    WriteDataQualityReport
    
    modPerformance.TurboOff
    
    Dim critCount As Long, warnCount As Long, infoCount As Long
    Dim i As Long
    For i = 0 To m_IssueCount - 1
        Select Case m_Issues(i).Severity
            Case "Critical": critCount = critCount + 1
            Case "Warning": warnCount = warnCount + 1
            Case "Info": infoCount = infoCount + 1
        End Select
    Next i
    
    modLogger.LogAction "modDataQuality", "ScanAll", _
                        m_IssueCount & " issues (" & critCount & " critical, " & _
                        warnCount & " warning, " & infoCount & " info)", _
                        modPerformance.ElapsedSeconds()
    
    MsgBox "Data Quality Scan Complete" & vbCrLf & vbCrLf & _
           "Issues Found: " & m_IssueCount & vbCrLf & _
           "  Critical: " & critCount & vbCrLf & _
           "  Warning:  " & warnCount & vbCrLf & _
           "  Info:     " & infoCount & vbCrLf & vbCrLf & _
           IIf(m_TextNumberCells.Count > 0, _
               m_TextNumberCells.Count & " text-stored numbers flagged (use Fix Text Numbers to convert)." & vbCrLf & vbCrLf, _
               "") & _
           "See '" & SH_DQ_REPORT & "' sheet for details.", _
           IIf(critCount > 0, vbExclamation, vbInformation), APP_NAME
    Exit Sub
    
ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modDataQuality", "ERROR", Err.Description
    MsgBox "Scan error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' AddIssue - Record a data quality finding
'===============================================================================
Private Sub AddIssue(ByVal sheet As String, ByVal cellRef As String, _
                      ByVal issueType As String, ByVal currentVal As String, _
                      ByVal severity As String, ByVal fixAvail As Boolean)
    ReDim Preserve m_Issues(m_IssueCount)
    With m_Issues(m_IssueCount)
        .Sheet = sheet
        .CellRef = cellRef
        .IssueType = issueType
        .CurrentVal = Left(currentVal, 200)
        .Severity = severity
        .FixAvail = fixAvail
    End With
    m_IssueCount = m_IssueCount + 1
End Sub

'===============================================================================
' ScanDuplicateRows - Find duplicate rows in GL data
'===============================================================================
Private Sub ScanDuplicateRows(ByVal ws As Worksheet)
    Dim lastRow As Long: lastRow = modConfig.LastRow(ws, 1)
    Dim lastCol As Long: lastCol = modConfig.LastCol(ws, HDR_ROW_GL)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    
    Dim r As Long, c As Long, key As String
    For r = DATA_ROW_GL To lastRow
        key = ""
        For c = 1 To lastCol
            key = key & "|" & CStr(ws.Cells(r, c).Value)
        Next c
        
        If dict.Exists(key) Then
            AddIssue ws.Name, "Row " & r, "Duplicate Row", _
                     "Duplicate of row " & dict(key), "Warning", True
            ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = RGB(255, 255, 200)
        Else
            dict.Add key, r
        End If
    Next r
End Sub

'===============================================================================
' ScanMixedDates - Find text-stored or inconsistent dates in GL
'===============================================================================
Private Sub ScanMixedDates(ByVal ws As Worksheet)
    Dim lastRow As Long: lastRow = modConfig.LastRow(ws, 1)
    Dim dateCol As Long: dateCol = COL_GL_DATE
    Dim r As Long
    
    For r = DATA_ROW_GL To lastRow
        Dim cellVal As String: cellVal = CStr(ws.Cells(r, dateCol).Value)
        
        If Not IsEmpty(ws.Cells(r, dateCol).Value) Then
            If VarType(ws.Cells(r, dateCol).Value) = vbString Then
                If cellVal Like "####-##-##" Then
                    AddIssue ws.Name, ws.Cells(r, dateCol).Address, _
                             "ISO Date Format (text)", cellVal, "Warning", True
                ElseIf cellVal Like "##/##/####" Then
                    AddIssue ws.Name, ws.Cells(r, dateCol).Address, _
                             "Zero-padded Date (text)", cellVal, "Info", True
                ElseIf cellVal Like "#/##/####" Or cellVal Like "##/#/####" Then
                    AddIssue ws.Name, ws.Cells(r, dateCol).Address, _
                             "Non-padded Date (text)", cellVal, "Info", True
                End If
            End If
        End If
    Next r
End Sub

'===============================================================================
' ScanTextNumbers - Find amounts stored as text in GL Amount column
' v2.1: Also records each flagged cell address in m_TextNumberCells
'===============================================================================
Private Sub ScanTextNumbers(ByVal ws As Worksheet)
    Dim lastRow As Long: lastRow = modConfig.LastRow(ws, 1)
    Dim amtCol As Long: amtCol = COL_GL_AMOUNT
    Dim r As Long
    
    For r = DATA_ROW_GL To lastRow
        If VarType(ws.Cells(r, amtCol).Value) = vbString Then
            If IsNumeric(ws.Cells(r, amtCol).Value) Then
                AddIssue ws.Name, ws.Cells(r, amtCol).Address, _
                         "Amount Stored as Text", CStr(ws.Cells(r, amtCol).Value), "Critical", True
                ws.Cells(r, amtCol).Interior.Color = RGB(255, 200, 200)
                
                ' v2.1: Track this cell for safe FixTextNumbers
                m_TextNumberCells.Add ws.Cells(r, amtCol).Address(True, True, xlA1, True)
            End If
        End If
    Next r
End Sub

'===============================================================================
' ScanAssumptionsTextValues - Find driver values stored as text
' v2.1: Also records each flagged cell address in m_TextNumberCells
'===============================================================================
Private Sub ScanAssumptionsTextValues()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_ASSUMPTIONS)
    Dim lastRow As Long: lastRow = modConfig.LastRow(ws, 1)
    Dim r As Long
    
    For r = DATA_ROW_ASSUME To lastRow
        If VarType(ws.Cells(r, 2).Value) = vbString Then
            If IsNumeric(ws.Cells(r, 2).Value) Then
                AddIssue ws.Name, ws.Cells(r, 2).Address, _
                         "Driver Value Stored as Text", CStr(ws.Cells(r, 2).Value), "Critical", True
                
                ' v2.1: Track this cell for safe FixTextNumbers
                m_TextNumberCells.Add ws.Cells(r, 2).Address(True, True, xlA1, True)
            End If
        End If
    Next r
End Sub

'===============================================================================
' ScanMisspellings - Check product name casing on Product Line Summary
'===============================================================================
Private Sub ScanMisspellings()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_PROD_SUMMARY)
    Dim validProducts As Variant: validProducts = modConfig.GetProducts()
    Dim lastRow As Long: lastRow = modConfig.LastRow(ws, 1)
    Dim lastCol As Long: lastCol = modConfig.LastCol(ws, HDR_ROW_REPORT)
    
    Dim r As Long, c As Long
    For r = 1 To lastRow
        For c = 1 To lastCol
            Dim cellVal As String: cellVal = Trim(CStr(ws.Cells(r, c).Value))
            If Len(cellVal) > 2 Then
                Dim p As Long
                For p = 0 To UBound(validProducts)
                    If StrComp(cellVal, validProducts(p), vbTextCompare) = 0 And _
                       StrComp(cellVal, validProducts(p), vbBinaryCompare) <> 0 Then
                        AddIssue ws.Name, ws.Cells(r, c).Address, _
                                 "Product Name Misspelling", _
                                 cellVal & " (should be " & validProducts(p) & ")", "Warning", True
                        ws.Cells(r, c).Font.Color = RGB(255, 0, 0)
                    End If
                Next p
            End If
        Next c
    Next r
End Sub

'===============================================================================
' ScanBlankCells - Check for empty AWS expense cells in Natural P&L
'===============================================================================
Private Sub ScanBlankCells()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_NATURAL)
    Dim lastRow As Long: lastRow = modConfig.LastRow(ws, 1)
    Dim r As Long
    
    For r = DATA_ROW_REPORT To lastRow
        If InStr(1, CStr(ws.Cells(r, 1).Value), "AWS", vbTextCompare) > 0 Then
            Dim c As Long
            For c = 2 To 5
                If IsEmpty(ws.Cells(r, c).Value) Or Trim(CStr(ws.Cells(r, c).Value)) = "" Then
                    AddIssue ws.Name, ws.Cells(r, c).Address, _
                             "Blank AWS Expense Cell", "(empty)", "Critical", False
                    ws.Cells(r, c).Interior.Color = RGB(255, 230, 230)
                End If
            Next c
        End If
    Next r
End Sub

'===============================================================================
' WriteDataQualityReport - Output findings to a dedicated sheet
' v2.1: Uses SH_DQ_REPORT from modConfig; uses SafeDeleteSheet & StyleHeader
'===============================================================================
Private Sub WriteDataQualityReport()
    modConfig.SafeDeleteSheet SH_DQ_REPORT
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = SH_DQ_REPORT
    ws.Tab.Color = RGB(255, 165, 0)
    
    ws.Range("A1").Value = "Data Quality Report - Keystone BenefitTech"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Scan Date: " & Format(Now, "mmmm d, yyyy h:mm AM/PM") & _
                           " | Issues Found: " & m_IssueCount
    ws.Range("A2").Font.Italic = True
    
    Dim headers As Variant
    headers = Array("Sheet", "Cell", "Issue Type", "Current Value", "Severity", "Fix Available")
    modConfig.StyleHeader ws, 4, headers
    
    Dim r As Long: r = 5
    Dim i As Long
    For i = 0 To m_IssueCount - 1
        ws.Cells(r, 1).Value = m_Issues(i).Sheet
        ws.Cells(r, 2).Value = m_Issues(i).CellRef
        ws.Cells(r, 3).Value = m_Issues(i).IssueType
        ws.Cells(r, 4).Value = m_Issues(i).CurrentVal
        ws.Cells(r, 5).Value = m_Issues(i).Severity
        ws.Cells(r, 6).Value = IIf(m_Issues(i).FixAvail, "Yes", "No")
        
        Select Case m_Issues(i).Severity
            Case "Critical"
                ws.Cells(r, 5).Interior.Color = RGB(255, 199, 206)
                ws.Cells(r, 5).Font.Color = RGB(156, 0, 6)
            Case "Warning"
                ws.Cells(r, 5).Interior.Color = RGB(255, 235, 156)
                ws.Cells(r, 5).Font.Color = RGB(156, 101, 0)
            Case "Info"
                ws.Cells(r, 5).Interior.Color = RGB(198, 239, 206)
                ws.Cells(r, 5).Font.Color = RGB(0, 97, 0)
        End Select
        
        r = r + 1
    Next i
    
    ws.Columns("A:F").AutoFit
End Sub

'===============================================================================
' FixTextNumbers - Convert ONLY pre-flagged text-stored numbers
'
' FIX (v2.1 — ISSUE-003 / BUG-018):
' The v2.0 version iterated the ENTIRE workbook and converted every text-stored
' number, including:
'   - GL IDs like "001" (losing leading zeros)
'   - Fiscal year strings like "2025" in headers
'   - Column labels that happened to be numeric
'
' The v2.1 version ONLY converts cells that were specifically flagged during
' ScanAll. If ScanAll hasn't been run, it refuses to proceed.
'===============================================================================
Public Sub FixTextNumbers()
    On Error GoTo ErrHandler
    
    If m_TextNumberCells Is Nothing Then
        MsgBox "Run Scan Data Quality first (Menu 7) to identify text-stored numbers." & _
               vbCrLf & vbCrLf & "FixTextNumbers only converts cells flagged by the scanner.", _
               vbInformation, APP_NAME
        Exit Sub
    End If
    
    If m_TextNumberCells.Count = 0 Then
        MsgBox "No text-stored numbers were flagged during the last scan." & vbCrLf & _
               "Run Scan Data Quality first (Menu 7) to identify issues.", _
               vbInformation, APP_NAME
        Exit Sub
    End If
    
    If MsgBox("Convert " & m_TextNumberCells.Count & " flagged text-stored numbers to actual numbers?" & _
              vbCrLf & vbCrLf & "Only cells identified by the last Scan Data Quality run will be changed.", _
              vbYesNo + vbQuestion, APP_NAME) = vbNo Then Exit Sub
    
    modPerformance.TurboOn
    Dim fixCount As Long: fixCount = 0
    Dim item As Variant
    
    For Each item In m_TextNumberCells
        On Error Resume Next
        Dim cell As Range
        Set cell = Nothing
        Set cell = Range(CStr(item))
        If Not cell Is Nothing Then
            If VarType(cell.Value) = vbString And IsNumeric(cell.Value) Then
                cell.Value = CDbl(cell.Value)
                fixCount = fixCount + 1
            End If
        End If
        On Error GoTo 0
    Next item
    
    modPerformance.TurboOff
    modLogger.LogAction "modDataQuality", "FixTextNumbers", fixCount & " cells converted (from scan list)"
    MsgBox fixCount & " text-stored numbers converted." & vbCrLf & _
           "Run Scan again to verify.", vbInformation, APP_NAME
    Exit Sub
    
ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modDataQuality", "ERROR", "FixTextNumbers: " & Err.Description
    MsgBox "Fix error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' FixDuplicates - Remove duplicate rows from GL staging data
'===============================================================================
Public Sub FixDuplicates()
    On Error GoTo ErrHandler
    
    If Not modConfig.SheetExists(SH_HIDDEN) Then Exit Sub
    
    If MsgBox("Remove duplicate rows from staging data?", _
              vbYesNo + vbQuestion, APP_NAME) = vbNo Then Exit Sub
    
    modPerformance.TurboOn
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_HIDDEN)
    ws.Visible = xlSheetVisible
    
    Dim lastRow As Long: lastRow = modConfig.LastRow(ws, 1)
    Dim lastCol As Long: lastCol = modConfig.LastCol(ws, HDR_ROW_GL)
    Dim beforeCount As Long: beforeCount = lastRow - 1
    
    ws.Range("A1").CurrentRegion.RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, 7), Header:=xlYes
    
    lastRow = modConfig.LastRow(ws, 1)
    Dim afterCount As Long: afterCount = lastRow - 1
    Dim removed As Long: removed = beforeCount - afterCount
    
    ws.Visible = xlSheetHidden
    modPerformance.TurboOff
    
    modLogger.LogAction "modDataQuality", "FixDuplicates", removed & " duplicates removed"
    MsgBox removed & " duplicate rows removed." & vbCrLf & _
           "Before: " & beforeCount & " rows" & vbCrLf & _
           "After:  " & afterCount & " rows", vbInformation, APP_NAME
    Exit Sub
    
ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modDataQuality", "ERROR", "FixDuplicates: " & Err.Description
    MsgBox "Fix error: " & Err.Description, vbCritical, APP_NAME
End Sub
