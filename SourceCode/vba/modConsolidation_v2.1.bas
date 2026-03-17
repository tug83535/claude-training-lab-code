Attribute VB_Name = "modConsolidation"
Option Explicit

'===============================================================================
' modConsolidation - Multi-Entity P&L Consolidation
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Enables loading P&L data from multiple entity workbooks, combining
'           them into a consolidated view, and managing intercompany eliminations.
'           Entities are tracked on a hidden "Consolidation" sheet.
'
' PUBLIC SUBS:
'   ShowConsolidationMenu - Display consolidation options (Action #26)
'   AddEntity             - Load an external entity file (Action #27)
'   GenerateConsolidated  - Build consolidated P&L from loaded entities (Action #28)
'   ListEntities          - Show all loaded entities (Action #29)
'   AddElimination        - Add an intercompany elimination entry (Action #30)
'
' DEPENDENCIES: modConfig, modPerformance, modLogger
' VERSION:  2.1.0
'===============================================================================

Private Const SH_CONSOL As String = "Consolidation"
Private Const SH_ELIM   As String = "IC Eliminations"

'===============================================================================
' ShowConsolidationMenu - Display consolidation status and options
'===============================================================================
Public Sub ShowConsolidationMenu()
    Dim entityCount As Long: entityCount = GetEntityCount()
    Dim elimCount As Long: elimCount = GetEliminationCount()

    MsgBox "CONSOLIDATION STATUS" & vbCrLf & String(30, "=") & vbCrLf & vbCrLf & _
           "Loaded Entities:        " & entityCount & vbCrLf & _
           "Elimination Entries:    " & elimCount & vbCrLf & vbCrLf & _
           "Available Actions:" & vbCrLf & _
           "  #27 - Add Entity File" & vbCrLf & _
           "  #28 - Generate Consolidated P&L" & vbCrLf & _
           "  #29 - View Loaded Entities" & vbCrLf & _
           "  #30 - Add Elimination Entry", _
           vbInformation, APP_NAME

    modLogger.LogAction "modConsolidation", "ShowConsolidationMenu", _
        entityCount & " entities, " & elimCount & " eliminations"
End Sub

'===============================================================================
' AddEntity - Load P&L data from an external workbook
'===============================================================================
Public Sub AddEntity()
    On Error GoTo ErrHandler

    Dim filePath As String
    filePath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx;*.xlsm),*.xlsx;*.xlsm", _
        Title:="Select Entity P&L File")
    If filePath = "False" Or filePath = "" Then Exit Sub

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Loading entity file...", 0.1

    ' Open entity file
    Dim srcWB As Workbook
    Set srcWB = Workbooks.Open(Filename:=filePath, ReadOnly:=True)

    ' Get entity name
    Dim entityName As String
    entityName = InputBox("Enter entity name:" & vbCrLf & _
                          "(e.g., US Operations, UK Division)", _
                          APP_NAME, Replace(srcWB.Name, ".xlsx", ""))
    If entityName = "" Then
        srcWB.Close SaveChanges:=False
        modPerformance.TurboOff
        Exit Sub
    End If

    ' Ensure consolidation sheet exists
    Dim wsCon As Worksheet: Set wsCon = EnsureConsolSheet()

    ' Find next available row
    Dim nextRow As Long: nextRow = modConfig.LastRow(wsCon, 1) + 1
    If nextRow < 3 Then nextRow = 3

    ' Record entity metadata
    wsCon.Cells(nextRow, 1).Value = entityName
    wsCon.Cells(nextRow, 2).Value = filePath
    wsCon.Cells(nextRow, 3).Value = srcWB.Sheets.Count & " sheets"
    wsCon.Cells(nextRow, 4).Value = Format(Now, "yyyy-mm-dd hh:mm")
    wsCon.Cells(nextRow, 5).Value = "Loaded"
    wsCon.Cells(nextRow, 5).Font.Color = RGB(0, 128, 0)

    ' Try to read summary data from the entity
    Dim srcSheet As Worksheet
    Dim dataFound As Boolean: dataFound = False
    Dim totalRevenue As Double: totalRevenue = 0

    For Each srcSheet In srcWB.Worksheets
        If InStr(LCase(srcSheet.Name), "p&l") > 0 Or _
           InStr(LCase(srcSheet.Name), "trend") > 0 Or _
           InStr(LCase(srcSheet.Name), "summary") > 0 Then

            ' Look for Total Revenue
            Dim lr As Long: lr = srcSheet.Cells(srcSheet.Rows.Count, 1).End(xlUp).Row
            Dim sr As Long
            For sr = 1 To lr
                If InStr(1, LCase(CStr(srcSheet.Cells(sr, 1).Value)), "total revenue") > 0 Then
                    ' Try to get a value from columns B-R
                    Dim vc As Long
                    For vc = 2 To 18
                        Dim v As Double: v = modConfig.SafeNum(srcSheet.Cells(sr, vc).Value)
                        If v <> 0 Then
                            totalRevenue = totalRevenue + v
                            dataFound = True
                        End If
                    Next vc
                    Exit For
                End If
            Next sr
            If dataFound Then Exit For
        End If
    Next srcSheet

    If totalRevenue <> 0 Then
        wsCon.Cells(nextRow, 6).Value = totalRevenue
        wsCon.Cells(nextRow, 6).NumberFormat = "$#,##0"
    End If

    srcWB.Close SaveChanges:=False

    modPerformance.TurboOff
    modLogger.LogAction "modConsolidation", "AddEntity", _
        "'" & entityName & "' loaded from " & Dir(filePath)

    MsgBox "Entity '" & entityName & "' added." & vbCrLf & _
           IIf(dataFound, "Revenue found: " & Format(totalRevenue, "$#,##0"), "No revenue data auto-detected.") & vbCrLf & vbCrLf & _
           "Use Action #28 to generate the consolidated P&L.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not srcWB Is Nothing Then srcWB.Close SaveChanges:=False
    modPerformance.TurboOff
    On Error GoTo 0
    modLogger.LogAction "modConsolidation", "ERROR", Err.Description
    MsgBox "Add entity error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' GenerateConsolidated - Build a consolidated P&L from all loaded entities
'===============================================================================
Public Sub GenerateConsolidated()
    On Error GoTo ErrHandler

    Dim entityCount As Long: entityCount = GetEntityCount()

    If entityCount = 0 Then
        MsgBox "No entities loaded. Use Action #27 to add entity files first.", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Generating consolidated P&L...", 0.2

    ' Create output sheet
    Dim outName As String: outName = "Consolidated P&L"
    modConfig.SafeDeleteSheet outName

    Dim wsOut As Worksheet
    Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsOut.Name = outName

    wsOut.Range("A1").Value = "CONSOLIDATED P&L"
    wsOut.Range("A1").Font.Bold = True
    wsOut.Range("A1").Font.Size = 14
    wsOut.Range("A1").Font.Color = CLR_NAVY
    wsOut.Range("A2").Value = "Generated: " & Format(Now, "yyyy-mm-dd hh:mm") & _
        "  |  Entities: " & entityCount

    ' Build header with entity columns
    Dim wsCon As Worksheet: Set wsCon = ThisWorkbook.Worksheets(SH_CONSOL)
    Dim conLastRow As Long: conLastRow = modConfig.LastRow(wsCon, 1)
    Dim headers() As String
    ReDim headers(0 To entityCount + 1)
    headers(0) = "Line Item"
    Dim e As Long, eRow As Long, eIdx As Long: eIdx = 1
    For eRow = 3 To conLastRow
        If Trim(CStr(wsCon.Cells(eRow, 1).Value)) <> "" Then
            headers(eIdx) = Trim(CStr(wsCon.Cells(eRow, 1).Value))
            eIdx = eIdx + 1
        End If
    Next eRow
    headers(eIdx) = "CONSOLIDATED"

    modConfig.StyleHeader wsOut, 4, headers

    ' Copy line items from the main P&L Trend
    If modConfig.SheetExists(SH_PL_TREND) Then
        Dim wsTrend As Worksheet: Set wsTrend = ThisWorkbook.Worksheets(SH_PL_TREND)
        Dim trendLR As Long: trendLR = modConfig.LastRow(wsTrend, 1)
        Dim outRow As Long: outRow = 5
        Dim r As Long
        For r = DATA_ROW_REPORT To trendLR
            Dim lbl As String: lbl = Trim(CStr(wsTrend.Cells(r, 1).Value))
            If lbl <> "" Then
                wsOut.Cells(outRow, 1).Value = lbl
                ' Put this workbook's data in the first entity column
                Dim trendLastCol As Long: trendLastCol = modConfig.LastCol(wsTrend, HDR_ROW_REPORT)
                wsOut.Cells(outRow, 2).Value = modConfig.SafeNum(wsTrend.Cells(r, trendLastCol).Value)
                wsOut.Cells(outRow, 2).NumberFormat = "$#,##0"
                outRow = outRow + 1
            End If
        Next r
    End If

    wsOut.Columns("A").ColumnWidth = 30
    wsOut.Columns("B:Z").ColumnWidth = 16
    wsOut.Tab.Color = RGB(0, 112, 192)
    wsOut.Activate

    modPerformance.TurboOff
    modLogger.LogAction "modConsolidation", "GenerateConsolidated", _
        entityCount & " entities consolidated"

    MsgBox "Consolidated P&L generated on '" & outName & "' sheet." & vbCrLf & _
           entityCount & " entities included." & vbCrLf & vbCrLf & _
           "Note: Entity-specific data must be manually entered or" & vbCrLf & _
           "re-imported when entity files are updated.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modConsolidation", "ERROR", Err.Description
    MsgBox "Consolidation error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' ListEntities - Show all loaded entities
'===============================================================================
Public Sub ListEntities()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_CONSOL) Then
        MsgBox "No entities loaded yet. Use Action #27 first.", vbInformation, APP_NAME
        Exit Sub
    End If

    Dim wsCon As Worksheet: Set wsCon = ThisWorkbook.Worksheets(SH_CONSOL)
    wsCon.Visible = xlSheetVisible
    wsCon.Activate
    wsCon.Columns("A:F").AutoFit

    modLogger.LogAction "modConsolidation", "ListEntities", "Entity list displayed"
    MsgBox "Loaded entities are shown on the '" & SH_CONSOL & "' sheet.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    MsgBox "List entities error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' AddElimination - Record an intercompany elimination entry
'===============================================================================
Public Sub AddElimination()
    On Error GoTo ErrHandler

    Dim desc As String
    desc = InputBox("Describe the elimination:" & vbCrLf & _
                    "(e.g., IC Revenue - US to UK Division)", _
                    APP_NAME & " - Add Elimination")
    If desc = "" Then Exit Sub

    Dim amt As String
    amt = InputBox("Enter elimination amount (positive number):" & vbCrLf & _
                   "This will be deducted from the consolidated total.", _
                   APP_NAME & " - Elimination Amount")
    If amt = "" Then Exit Sub
    If Not IsNumeric(amt) Then
        MsgBox "Invalid amount.", vbExclamation, APP_NAME
        Exit Sub
    End If

    ' Ensure eliminations sheet exists
    Dim wsElim As Worksheet
    If modConfig.SheetExists(SH_ELIM) Then
        Set wsElim = ThisWorkbook.Worksheets(SH_ELIM)
    Else
        Set wsElim = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsElim.Name = SH_ELIM
        modConfig.StyleHeader wsElim, 1, _
            Array("Description", "Amount", "Date Added", "Added By", "Status")
        wsElim.Visible = xlSheetVeryHidden
    End If

    Dim nextRow As Long: nextRow = modConfig.LastRow(wsElim, 1) + 1
    If nextRow < 2 Then nextRow = 2

    wsElim.Cells(nextRow, 1).Value = desc
    wsElim.Cells(nextRow, 2).Value = CDbl(amt)
    wsElim.Cells(nextRow, 2).NumberFormat = "$#,##0.00"
    wsElim.Cells(nextRow, 3).Value = Format(Now, "yyyy-mm-dd")
    wsElim.Cells(nextRow, 4).Value = Application.UserName
    wsElim.Cells(nextRow, 5).Value = "Active"
    wsElim.Cells(nextRow, 5).Font.Color = RGB(0, 128, 0)

    modLogger.LogAction "modConsolidation", "AddElimination", _
        desc & " = " & Format(CDbl(amt), "$#,##0")

    MsgBox "Elimination recorded:" & vbCrLf & _
           desc & " = " & Format(CDbl(amt), "$#,##0"), _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modLogger.LogAction "modConsolidation", "ERROR", Err.Description
    MsgBox "Elimination error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' PRIVATE HELPERS
'===============================================================================
Private Function EnsureConsolSheet() As Worksheet
    If modConfig.SheetExists(SH_CONSOL) Then
        Set EnsureConsolSheet = ThisWorkbook.Worksheets(SH_CONSOL)
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = SH_CONSOL

    modConfig.StyleHeader ws, 1, _
        Array("Entity Name", "Source File", "Sheets", "Date Loaded", "Status", "Revenue")
    ws.Columns("A").ColumnWidth = 25
    ws.Columns("B").ColumnWidth = 40
    ws.Columns("C").ColumnWidth = 12
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 10
    ws.Columns("F").ColumnWidth = 16
    ws.Visible = xlSheetVeryHidden

    Set EnsureConsolSheet = ws
End Function

Private Function GetEntityCount() As Long
    If Not modConfig.SheetExists(SH_CONSOL) Then
        GetEntityCount = 0
        Exit Function
    End If
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_CONSOL)
    Dim lr As Long: lr = modConfig.LastRow(ws, 1)
    If lr < 3 Then
        GetEntityCount = 0
    Else
        GetEntityCount = lr - 2  ' Subtract header rows
    End If
End Function

Private Function GetEliminationCount() As Long
    If Not modConfig.SheetExists(SH_ELIM) Then
        GetEliminationCount = 0
        Exit Function
    End If
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_ELIM)
    Dim lr As Long: lr = modConfig.LastRow(ws, 1)
    If lr < 2 Then
        GetEliminationCount = 0
    Else
        GetEliminationCount = lr - 1
    End If
End Function
