Attribute VB_Name = "modVersionControl"
Option Explicit

'===============================================================================
' modVersionControl - Workbook Version Management
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Saves timestamped snapshots of the workbook as versioned copies.
'           Tracks version metadata on a hidden "Version History" sheet.
'           Supports save, compare, restore, and list operations.
'
' PUBLIC SUBS:
'   ShowVersionMenu  - Display version control status (Action #31)
'   SaveVersion      - Save current state as a new version (Action #32)
'   CompareVersions  - Compare two saved versions (Action #33)
'   RestoreVersion   - Restore from a previous version (Action #34)
'   ListVersions     - Show all saved versions (Action #35)
'
' DEPENDENCIES: modConfig, modPerformance, modLogger
' VERSION:  2.1.0
'===============================================================================

Private Const SH_VERSIONS As String = "Version History"

'===============================================================================
' ShowVersionMenu - Display current version control status
'===============================================================================
Public Sub ShowVersionMenu()
    Dim vCount As Long: vCount = GetVersionCount()

    MsgBox "VERSION CONTROL" & vbCrLf & String(25, "=") & vbCrLf & vbCrLf & _
           "Current Version: " & APP_VERSION & vbCrLf & _
           "Saved Snapshots: " & vCount & vbCrLf & vbCrLf & _
           "Available Actions:" & vbCrLf & _
           "  #32 - Save Version" & vbCrLf & _
           "  #33 - Compare Versions" & vbCrLf & _
           "  #34 - Restore Version" & vbCrLf & _
           "  #35 - List Versions", _
           vbInformation, APP_NAME

    modLogger.LogAction "modVersionControl", "ShowVersionMenu", vCount & " versions tracked"
End Sub

'===============================================================================
' SaveVersion - Save current workbook as a timestamped copy
'===============================================================================
Public Sub SaveVersion()
    On Error GoTo ErrHandler

    Dim versionNote As String
    versionNote = InputBox("Enter a note for this version:" & vbCrLf & _
                           "(e.g., Q1 final, Pre-board review, Post AWS update)", _
                           APP_NAME & " - Save Version")
    If versionNote = "" Then versionNote = "Manual save"

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Saving version...", 0.1

    ' Build version filename
    Dim ts As String: ts = Format(Now, "yyyymmdd_hhmmss")
    Dim vNum As Long: vNum = GetVersionCount() + 1
    Dim basePath As String: basePath = ThisWorkbook.Path
    If basePath = "" Then basePath = Environ("USERPROFILE") & "\Desktop"

    Dim versionFile As String
    versionFile = basePath & "\versions\"
    ' Create versions folder if needed
    If Dir(basePath & "\versions", vbDirectory) = "" Then MkDir basePath & "\versions"

    versionFile = versionFile & "v" & vNum & "_" & ts & "_" & _
                  Replace(Left(versionNote, 20), " ", "_") & ".xlsx"

    modPerformance.UpdateStatus "Copying workbook...", 0.4

    ' Save a copy
    ThisWorkbook.SaveCopyAs versionFile

    ' Record in version history
    Dim wsVer As Worksheet: Set wsVer = EnsureVersionSheet()
    Dim nextRow As Long: nextRow = modConfig.LastRow(wsVer, 1) + 1
    If nextRow < 2 Then nextRow = 2

    wsVer.Cells(nextRow, 1).Value = vNum
    wsVer.Cells(nextRow, 2).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    wsVer.Cells(nextRow, 3).Value = versionNote
    wsVer.Cells(nextRow, 4).Value = Application.UserName
    wsVer.Cells(nextRow, 5).Value = versionFile
    wsVer.Cells(nextRow, 6).Value = ThisWorkbook.Worksheets.Count & " sheets"

    modPerformance.TurboOff

    modLogger.LogAction "modVersionControl", "SaveVersion", _
        "v" & vNum & " saved: " & versionNote

    MsgBox "VERSION SAVED" & vbCrLf & String(20, "=") & vbCrLf & vbCrLf & _
           "Version:  v" & vNum & vbCrLf & _
           "Note:     " & versionNote & vbCrLf & _
           "File:     " & Dir(versionFile) & vbCrLf & _
           "Location: " & basePath & "\versions\", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modVersionControl", "ERROR", Err.Description
    MsgBox "Save version error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' CompareVersions - Show version history for comparison
'===============================================================================
Public Sub CompareVersions()
    On Error GoTo ErrHandler

    Dim vCount As Long: vCount = GetVersionCount()
    If vCount < 2 Then
        MsgBox "Need at least 2 saved versions to compare." & vbCrLf & _
               "Current versions: " & vCount, vbExclamation, APP_NAME
        Exit Sub
    End If

    ' Show the version history sheet
    Dim wsVer As Worksheet: Set wsVer = ThisWorkbook.Worksheets(SH_VERSIONS)
    wsVer.Visible = xlSheetVisible
    wsVer.Activate
    wsVer.Columns("A:F").AutoFit

    MsgBox "Version history is now visible." & vbCrLf & vbCrLf & _
           "To compare two versions:" & vbCrLf & _
           "1. Note the file paths in column E" & vbCrLf & _
           "2. Open both files in separate Excel windows" & vbCrLf & _
           "3. Use View > View Side by Side" & vbCrLf & vbCrLf & _
           "The sheet will be re-hidden when you run another action.", _
           vbInformation, APP_NAME

    modLogger.LogAction "modVersionControl", "CompareVersions", "Version history displayed"
    Exit Sub

ErrHandler:
    MsgBox "Compare error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' RestoreVersion - Open a previous version for manual restore
'===============================================================================
Public Sub RestoreVersion()
    On Error GoTo ErrHandler

    Dim vCount As Long: vCount = GetVersionCount()
    If vCount = 0 Then
        MsgBox "No saved versions. Use Action #32 first.", vbInformation, APP_NAME
        Exit Sub
    End If

    ' Build version list
    Dim wsVer As Worksheet: Set wsVer = ThisWorkbook.Worksheets(SH_VERSIONS)
    Dim vList As String: vList = ""
    Dim lr As Long: lr = modConfig.LastRow(wsVer, 1)
    Dim r As Long
    For r = 2 To lr
        If Trim(CStr(wsVer.Cells(r, 1).Value)) <> "" Then
            vList = vList & "v" & wsVer.Cells(r, 1).Value & " - " & _
                    wsVer.Cells(r, 2).Value & " - " & wsVer.Cells(r, 3).Value & vbCrLf
        End If
    Next r

    Dim choice As String
    choice = InputBox("Select version number to restore:" & vbCrLf & vbCrLf & vList, _
                      APP_NAME & " - Restore Version")
    If choice = "" Then Exit Sub
    If Not IsNumeric(choice) Then Exit Sub

    Dim selVer As Long: selVer = CLng(choice)

    ' Find the file path
    Dim filePath As String: filePath = ""
    For r = 2 To lr
        If modConfig.SafeNum(wsVer.Cells(r, 1).Value) = selVer Then
            filePath = Trim(CStr(wsVer.Cells(r, 5).Value))
            Exit For
        End If
    Next r

    If filePath = "" Then
        MsgBox "Version v" & selVer & " not found.", vbExclamation, APP_NAME
        Exit Sub
    End If

    ' Check if file exists
    If Dir(filePath) = "" Then
        MsgBox "Version file not found:" & vbCrLf & filePath, vbCritical, APP_NAME
        Exit Sub
    End If

    If MsgBox("Open version v" & selVer & " in a new window?" & vbCrLf & vbCrLf & _
              "You can then manually copy data from the old version." & vbCrLf & _
              "The current workbook will NOT be modified.", _
              vbYesNo + vbQuestion, APP_NAME) = vbNo Then Exit Sub

    Workbooks.Open filePath, ReadOnly:=True
    modLogger.LogAction "modVersionControl", "RestoreVersion", "v" & selVer & " opened for review"

    MsgBox "Version v" & selVer & " is now open in a separate window." & vbCrLf & _
           "Copy any data you need, then close the old file.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modLogger.LogAction "modVersionControl", "ERROR", Err.Description
    MsgBox "Restore error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' ListVersions - Show all saved versions
'===============================================================================
Public Sub ListVersions()
    On Error GoTo ErrHandler

    Dim vCount As Long: vCount = GetVersionCount()
    If vCount = 0 Then
        MsgBox "No versions saved yet. Use Action #32 to save your first version.", _
               vbInformation, APP_NAME
        Exit Sub
    End If

    Dim wsVer As Worksheet: Set wsVer = ThisWorkbook.Worksheets(SH_VERSIONS)
    wsVer.Visible = xlSheetVisible
    wsVer.Activate
    wsVer.Columns("A:F").AutoFit

    modLogger.LogAction "modVersionControl", "ListVersions", vCount & " versions listed"
    MsgBox vCount & " versions saved." & vbCrLf & _
           "Version history is now visible on the '" & SH_VERSIONS & "' sheet.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    MsgBox "List error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' PRIVATE HELPERS
'===============================================================================
Private Function EnsureVersionSheet() As Worksheet
    If modConfig.SheetExists(SH_VERSIONS) Then
        Set EnsureVersionSheet = ThisWorkbook.Worksheets(SH_VERSIONS)
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = SH_VERSIONS

    modConfig.StyleHeader ws, 1, _
        Array("Version #", "Date Saved", "Note", "Saved By", "File Path", "Sheets")
    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 16
    ws.Columns("E").ColumnWidth = 50
    ws.Columns("F").ColumnWidth = 12
    ws.Visible = xlSheetVeryHidden

    Set EnsureVersionSheet = ws
End Function

Private Function GetVersionCount() As Long
    If Not modConfig.SheetExists(SH_VERSIONS) Then
        GetVersionCount = 0
        Exit Function
    End If
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_VERSIONS)
    Dim lr As Long: lr = modConfig.LastRow(ws, 1)
    If lr < 2 Then
        GetVersionCount = 0
    Else
        GetVersionCount = lr - 1
    End If
End Function
