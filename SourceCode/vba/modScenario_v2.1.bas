Attribute VB_Name = "modScenario"
Option Explicit

'===============================================================================
' modScenario - Scenario Management (Save / Load / Compare / Delete)
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Saves snapshots of the Assumptions sheet as named scenarios on a
'           hidden "Scenarios" sheet. Users can save, load, compare, and
'           delete scenarios for what-if planning.
'
' PUBLIC SUBS:
'   SaveScenario     - Save current assumptions as named scenario (Action #20)
'   LoadScenario     - Restore assumptions from a saved scenario (Action #21)
'   CompareScenarios - Side-by-side comparison of two scenarios (Action #22)
'   DeleteScenario   - Remove a saved scenario (Action #23)
'
' DEPENDENCIES: modConfig, modPerformance, modLogger
' VERSION:  2.1.0
'===============================================================================

Private Const SH_SCENARIOS As String = "Scenarios"
Private Const SCN_HDR_ROW  As Long = 1
Private Const SCN_DATA_ROW As Long = 2

'===============================================================================
' SaveScenario - Snapshot current Assumptions to the Scenarios sheet
'===============================================================================
Public Sub SaveScenario()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_ASSUMPTIONS) Then
        MsgBox "Assumptions sheet not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    Dim scenarioName As String
    scenarioName = InputBox("Enter a name for this scenario:" & vbCrLf & vbCrLf & _
                            "Examples: Base Case, Optimistic, Q2 Revised", _
                            APP_NAME & " - Save Scenario")
    If scenarioName = "" Then Exit Sub

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Saving scenario: " & scenarioName, 0.1

    ' Ensure Scenarios sheet exists
    Dim wsSc As Worksheet: Set wsSc = EnsureScenariosSheet()

    ' Find next available column for this scenario
    Dim nextCol As Long: nextCol = wsSc.Cells(SCN_HDR_ROW, wsSc.Columns.Count).End(xlToLeft).Column + 1
    If nextCol < 3 Then nextCol = 3  ' Col A=Driver, B=Base, C+ = scenarios

    ' Check if name already exists
    Dim c As Long
    For c = 3 To nextCol - 1
        If LCase(Trim(CStr(wsSc.Cells(SCN_HDR_ROW, c).Value))) = LCase(scenarioName) Then
            If MsgBox("Scenario '" & scenarioName & "' already exists." & vbCrLf & _
                       "Overwrite?", vbYesNo + vbQuestion, APP_NAME) = vbNo Then
                modPerformance.TurboOff
                Exit Sub
            End If
            nextCol = c
            Exit For
        End If
    Next c

    ' Read Assumptions and write to scenario column
    Dim wsA As Worksheet: Set wsA = ThisWorkbook.Worksheets(SH_ASSUMPTIONS)
    Dim lastRow As Long: lastRow = modConfig.LastRow(wsA, 1)

    ' Write scenario header
    wsSc.Cells(SCN_HDR_ROW, nextCol).Value = scenarioName
    wsSc.Cells(SCN_HDR_ROW, nextCol).Font.Bold = True
    wsSc.Cells(SCN_HDR_ROW, nextCol).Interior.Color = CLR_NAVY
    wsSc.Cells(SCN_HDR_ROW, nextCol).Font.Color = CLR_WHITE

    Dim driverCount As Long: driverCount = 0
    Dim r As Long, outRow As Long: outRow = SCN_DATA_ROW

    For r = DATA_ROW_ASSUME To lastRow
        Dim dName As String: dName = Trim(CStr(wsA.Cells(r, 1).Value))
        If dName <> "" Then
            ' Write driver name (col A) and base value (col B) if first scenario
            wsSc.Cells(outRow, 1).Value = dName
            wsSc.Cells(outRow, 2).Value = wsA.Cells(r, 2).Value

            ' Write scenario value
            wsSc.Cells(outRow, nextCol).Value = wsA.Cells(r, 2).Value
            wsSc.Cells(outRow, nextCol).NumberFormat = "#,##0.00"

            driverCount = driverCount + 1
            outRow = outRow + 1
        End If
    Next r

    ' Add metadata row
    wsSc.Cells(outRow + 1, nextCol).Value = "Saved: " & Format(Now, "yyyy-mm-dd hh:mm")
    wsSc.Cells(outRow + 1, nextCol).Font.Italic = True
    wsSc.Cells(outRow + 1, nextCol).Font.Size = 8

    modPerformance.TurboOff
    modLogger.LogAction "modScenario", "SaveScenario", "'" & scenarioName & "' saved with " & driverCount & " drivers"
    MsgBox "Scenario '" & scenarioName & "' saved." & vbCrLf & _
           driverCount & " driver values captured.", vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modScenario", "ERROR", Err.Description
    MsgBox "Save scenario error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' LoadScenario - Restore Assumptions from a saved scenario
'===============================================================================
Public Sub LoadScenario()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_SCENARIOS) Then
        MsgBox "No scenarios saved yet. Use 'Save Scenario' first.", vbInformation, APP_NAME
        Exit Sub
    End If
    If Not modConfig.SheetExists(SH_ASSUMPTIONS) Then
        MsgBox "Assumptions sheet not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    Dim wsSc As Worksheet: Set wsSc = ThisWorkbook.Worksheets(SH_SCENARIOS)
    Dim lastCol As Long: lastCol = wsSc.Cells(SCN_HDR_ROW, wsSc.Columns.Count).End(xlToLeft).Column

    If lastCol < 3 Then
        MsgBox "No scenarios saved yet.", vbInformation, APP_NAME
        Exit Sub
    End If

    ' Build list of scenarios
    Dim scList As String: scList = ""
    Dim scCount As Long: scCount = 0
    Dim c As Long
    For c = 3 To lastCol
        Dim scName As String: scName = Trim(CStr(wsSc.Cells(SCN_HDR_ROW, c).Value))
        If scName <> "" Then
            scCount = scCount + 1
            scList = scList & scCount & ". " & scName & vbCrLf
        End If
    Next c

    Dim choice As String
    choice = InputBox("Select scenario to load:" & vbCrLf & vbCrLf & scList, _
                      APP_NAME & " - Load Scenario")
    If choice = "" Then Exit Sub
    If Not IsNumeric(choice) Then Exit Sub

    Dim selIdx As Long: selIdx = CLng(choice)
    If selIdx < 1 Or selIdx > scCount Then
        MsgBox "Invalid selection.", vbExclamation, APP_NAME
        Exit Sub
    End If

    ' Find the selected column
    Dim selCol As Long: selCol = 0
    Dim cnt As Long: cnt = 0
    For c = 3 To lastCol
        If Trim(CStr(wsSc.Cells(SCN_HDR_ROW, c).Value)) <> "" Then
            cnt = cnt + 1
            If cnt = selIdx Then selCol = c: Exit For
        End If
    Next c

    If selCol = 0 Then Exit Sub

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Loading scenario...", 0.3

    ' Write scenario values back to Assumptions
    Dim wsA As Worksheet: Set wsA = ThisWorkbook.Worksheets(SH_ASSUMPTIONS)
    Dim lastRow As Long: lastRow = modConfig.LastRow(wsA, 1)
    Dim loadCount As Long: loadCount = 0

    Dim r As Long
    For r = DATA_ROW_ASSUME To lastRow
        Dim dName As String: dName = Trim(CStr(wsA.Cells(r, 1).Value))
        If dName = "" Then GoTo NextLoadRow

        ' Find matching driver in scenarios sheet
        Dim sr As Long
        For sr = SCN_DATA_ROW To modConfig.LastRow(wsSc, 1)
            If LCase(Trim(CStr(wsSc.Cells(sr, 1).Value))) = LCase(dName) Then
                wsA.Cells(r, 2).Value = wsSc.Cells(sr, selCol).Value
                loadCount = loadCount + 1
                Exit For
            End If
        Next sr
NextLoadRow:
    Next r

    wsA.Activate
    modPerformance.TurboOff

    Dim loadedName As String: loadedName = Trim(CStr(wsSc.Cells(SCN_HDR_ROW, selCol).Value))
    modLogger.LogAction "modScenario", "LoadScenario", "'" & loadedName & "' loaded: " & loadCount & " drivers"
    MsgBox "Scenario '" & loadedName & "' loaded." & vbCrLf & _
           loadCount & " driver values restored.", vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modScenario", "ERROR", Err.Description
    MsgBox "Load scenario error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' CompareScenarios - Side-by-side comparison of saved scenarios
'===============================================================================
Public Sub CompareScenarios()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_SCENARIOS) Then
        MsgBox "No scenarios saved yet.", vbInformation, APP_NAME
        Exit Sub
    End If

    Dim wsSc As Worksheet: Set wsSc = ThisWorkbook.Worksheets(SH_SCENARIOS)

    ' Make the scenarios sheet visible and navigate to it
    wsSc.Visible = xlSheetVisible
    wsSc.Activate
    wsSc.Columns("A:Z").AutoFit

    modLogger.LogAction "modScenario", "CompareScenarios", "Scenarios sheet displayed"
    MsgBox "Scenarios sheet is now visible for comparison." & vbCrLf & _
           "Each column is a saved scenario." & vbCrLf & vbCrLf & _
           "The sheet will be re-hidden when you run another action.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    MsgBox "Compare error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' DeleteScenario - Remove a saved scenario
'===============================================================================
Public Sub DeleteScenario()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_SCENARIOS) Then
        MsgBox "No scenarios saved yet.", vbInformation, APP_NAME
        Exit Sub
    End If

    Dim wsSc As Worksheet: Set wsSc = ThisWorkbook.Worksheets(SH_SCENARIOS)
    Dim lastCol As Long: lastCol = wsSc.Cells(SCN_HDR_ROW, wsSc.Columns.Count).End(xlToLeft).Column

    If lastCol < 3 Then
        MsgBox "No scenarios to delete.", vbInformation, APP_NAME
        Exit Sub
    End If

    ' Build scenario list
    Dim scList As String: scList = ""
    Dim scCount As Long: scCount = 0
    Dim c As Long
    For c = 3 To lastCol
        Dim scName As String: scName = Trim(CStr(wsSc.Cells(SCN_HDR_ROW, c).Value))
        If scName <> "" Then
            scCount = scCount + 1
            scList = scList & scCount & ". " & scName & vbCrLf
        End If
    Next c

    Dim choice As String
    choice = InputBox("Select scenario to DELETE:" & vbCrLf & vbCrLf & scList, _
                      APP_NAME & " - Delete Scenario")
    If choice = "" Then Exit Sub
    If Not IsNumeric(choice) Then Exit Sub

    Dim selIdx As Long: selIdx = CLng(choice)
    If selIdx < 1 Or selIdx > scCount Then
        MsgBox "Invalid selection.", vbExclamation, APP_NAME
        Exit Sub
    End If

    ' Find the column
    Dim selCol As Long: selCol = 0
    Dim cnt As Long: cnt = 0
    For c = 3 To lastCol
        If Trim(CStr(wsSc.Cells(SCN_HDR_ROW, c).Value)) <> "" Then
            cnt = cnt + 1
            If cnt = selIdx Then selCol = c: Exit For
        End If
    Next c

    If selCol = 0 Then Exit Sub

    Dim delName As String: delName = Trim(CStr(wsSc.Cells(SCN_HDR_ROW, selCol).Value))
    If MsgBox("Delete scenario '" & delName & "'?" & vbCrLf & "This cannot be undone.", _
              vbYesNo + vbExclamation, APP_NAME) = vbNo Then Exit Sub

    wsSc.Columns(selCol).Delete
    modLogger.LogAction "modScenario", "DeleteScenario", "'" & delName & "' deleted"
    MsgBox "Scenario '" & delName & "' deleted.", vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modLogger.LogAction "modScenario", "ERROR", Err.Description
    MsgBox "Delete error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' EnsureScenariosSheet - Create hidden Scenarios sheet if it doesn't exist
'===============================================================================
Private Function EnsureScenariosSheet() As Worksheet
    If modConfig.SheetExists(SH_SCENARIOS) Then
        Set EnsureScenariosSheet = ThisWorkbook.Worksheets(SH_SCENARIOS)
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = SH_SCENARIOS

    ' Setup headers
    ws.Cells(SCN_HDR_ROW, 1).Value = "Driver Name"
    ws.Cells(SCN_HDR_ROW, 2).Value = "Base Value"
    ws.Rows(SCN_HDR_ROW).Font.Bold = True
    ws.Rows(SCN_HDR_ROW).Interior.Color = CLR_NAVY
    ws.Rows(SCN_HDR_ROW).Font.Color = CLR_WHITE
    ws.Columns("A").ColumnWidth = 30
    ws.Columns("B").ColumnWidth = 14

    ws.Visible = xlSheetVeryHidden
    Set EnsureScenariosSheet = ws
End Function
