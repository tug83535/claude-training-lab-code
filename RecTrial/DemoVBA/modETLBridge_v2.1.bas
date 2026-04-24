Attribute VB_Name = "modETLBridge"
Option Explicit

'===============================================================================
' modETLBridge - Python ETL Integration Bridge
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Lets users run the Python ETL pipeline and import its output
'           without leaving Excel. One button triggers the script; a second
'           button pulls the cleaned data back into the workbook.
'
' PUBLIC SUBS:
'   TriggerETLLocally     - Run kbt_etl_pipeline.py via Windows Shell (#119)
'   ImportETLOutput       - Load cleaned Excel output into workbook (#120)
'
' REQUIREMENTS:
'   - Python 3.x must be installed and on the system PATH
'   - kbt_etl_pipeline.py must be in the same folder as this workbook
'     (or the user is prompted to locate it)
'   - pip install openpyxl pandas (once, done outside Excel)
'
' VERSION:  2.1.0 (New module — 2026-03-01)
' SOURCE:   Ideas from NewTesting/VBA Examples (200) — items #119, #120
'===============================================================================

' Default output file name produced by the ETL script
Private Const ETL_SCRIPT_NAME  As String = "kbt_etl_pipeline.py"
Private Const ETL_OUTPUT_NAME  As String = "KBT_Cleaned.xlsx"
Private Const ETL_SOURCE_SHEET As String = "CleanedTransactions"

'===============================================================================
' TriggerETLLocally - Run the Python ETL pipeline from inside Excel (#119)
' Builds the Shell command, shows the user the expected output path, and
' launches the script in a visible command window so progress is visible.
' The user must wait for the window to close, then run ImportETLOutput.
'===============================================================================
Public Sub TriggerETLLocally()
    On Error GoTo ErrHandler

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Save the workbook first so Excel knows which folder to look in.", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim wbFolder   As String: wbFolder   = ThisWorkbook.Path
    Dim scriptPath As String: scriptPath = wbFolder & "\" & ETL_SCRIPT_NAME
    Dim inputPath  As String: inputPath  = wbFolder & "\" & ThisWorkbook.Name
    Dim outputPath As String: outputPath = wbFolder & "\" & ETL_OUTPUT_NAME

    ' Check that the script exists in the same folder
    If Dir(scriptPath) = "" Then
        Dim altResult As Variant
        altResult = Application.GetOpenFilename( _
            "Python Scripts (*.py),*.py", , _
            "Locate " & ETL_SCRIPT_NAME)
        If VarType(altResult) = vbBoolean Then Exit Sub
        If Len(CStr(altResult)) = 0 Then Exit Sub
        scriptPath = CStr(altResult)
    End If

    ' Confirm before running
    If MsgBox("Run the ETL pipeline now?" & vbCrLf & vbCrLf & _
              "Script:  " & scriptPath & vbCrLf & _
              "Input:   " & inputPath & vbCrLf & _
              "Output:  " & outputPath & vbCrLf & vbCrLf & _
              "A command window will open. Wait for it to close, then run ImportETLOutput.", _
              vbYesNo + vbQuestion, APP_NAME) = vbNo Then Exit Sub

    ' Build the Shell command: cmd /k keeps window open until user closes it
    Dim shellCmd As String
    shellCmd = "cmd /k python """ & scriptPath & """ """ & inputPath & _
               """ --output """ & outputPath & """"

    Shell "cmd.exe /c start """ & APP_NAME & " ETL"" " & shellCmd, vbNormalFocus

    modLogger.LogAction "modETLBridge", "TriggerETLLocally", _
        "Shell launched: " & ETL_SCRIPT_NAME & " | Output: " & ETL_OUTPUT_NAME
    MsgBox "ETL script launched in a command window." & vbCrLf & vbCrLf & _
           "When it finishes (window shows 'DQ log -> ...' and stops scrolling)," & vbCrLf & _
           "close the command window and run ImportETLOutput.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    MsgBox "TriggerETLLocally error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' ImportETLOutput - Load cleaned transaction data from the ETL output file (#120)
' Opens KBT_Cleaned.xlsx (or asks the user to locate it), copies the
' CleanedTransactions sheet data, and pastes it as values into the
' CrossfireHiddenWorksheet, replacing the old raw data.
' A confirmation prompt is shown before any data is overwritten.
'===============================================================================
Public Sub ImportETLOutput()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_HIDDEN) Then
        MsgBox "GL destination sheet '" & SH_HIDDEN & "' not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    ' Locate the ETL output file
    Dim outputPath As String
    If Len(ThisWorkbook.Path) > 0 Then
        outputPath = ThisWorkbook.Path & "\" & ETL_OUTPUT_NAME
    End If

    If Len(outputPath) = 0 Or Dir(outputPath) = "" Then
        Dim fileResult As Variant
        fileResult = Application.GetOpenFilename( _
            "Excel Files (*.xlsx),*.xlsx", , _
            "Locate " & ETL_OUTPUT_NAME)
        If VarType(fileResult) = vbBoolean Then Exit Sub
        If Len(CStr(fileResult)) = 0 Then Exit Sub
        outputPath = CStr(fileResult)
    End If

    ' Confirm before overwriting GL data
    If MsgBox("Import cleaned data from:" & vbCrLf & outputPath & vbCrLf & vbCrLf & _
              "This will REPLACE all data in '" & SH_HIDDEN & "'." & vbCrLf & _
              "Make sure the ETL script ran successfully first.", _
              vbYesNo + vbExclamation, APP_NAME) = vbNo Then Exit Sub

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Opening ETL output file...", 0.1

    ' Open the ETL output (read only)
    Dim wbSrc As Workbook
    Set wbSrc = Workbooks.Open(outputPath, ReadOnly:=True)

    If Not WorkbookHasSheet(wbSrc, ETL_SOURCE_SHEET) Then
        wbSrc.Close SaveChanges:=False
        modPerformance.TurboOff
        MsgBox "Sheet '" & ETL_SOURCE_SHEET & "' not found in " & ETL_OUTPUT_NAME & "." & vbCrLf & _
               "Make sure the ETL script ran successfully before importing.", _
               vbCritical, APP_NAME
        Exit Sub
    End If

    Dim wsSrc As Worksheet: Set wsSrc = wbSrc.Worksheets(ETL_SOURCE_SHEET)
    Dim srcLastRow As Long: srcLastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    Dim srcLastCol As Long: srcLastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

    modPerformance.UpdateStatus "Copying cleaned data to GL sheet...", 0.5

    Dim wsDest As Worksheet: Set wsDest = ThisWorkbook.Worksheets(SH_HIDDEN)
    wsDest.Visible = xlSheetVisible

    ' Clear the old data (keep header row 1 intact)
    If modConfig.LastRow(wsDest, COL_GL_ID) > HDR_ROW_GL Then
        wsDest.Range(wsDest.Cells(DATA_ROW_GL, 1), _
                     wsDest.Cells(modConfig.LastRow(wsDest, 1), srcLastCol)).ClearContents
    End If

    ' Copy headers from source (row 1)
    wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(1, srcLastCol)).Copy
    wsDest.Cells(HDR_ROW_GL, 1).PasteSpecial xlPasteValues

    ' Copy data rows
    If srcLastRow > 1 Then
        wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(srcLastRow, srcLastCol)).Copy
        wsDest.Cells(DATA_ROW_GL, 1).PasteSpecial xlPasteValues
    End If

    Application.CutCopyMode = False
    wbSrc.Close SaveChanges:=False

    ' Stamp the import timestamp
    Dim importedRows As Long: importedRows = srcLastRow - 1
    wsDest.Cells(1, srcLastCol + 2).Value = "Imported: " & Format(Now, "yyyy-mm-dd hh:nn:ss")

    modPerformance.TurboOff
    modLogger.LogAction "modETLBridge", "ImportETLOutput", _
        importedRows & " rows imported from " & ETL_OUTPUT_NAME
    MsgBox importedRows & " clean rows imported into '" & SH_HIDDEN & "'." & vbCrLf & _
           "Run Data Quality Check and Reconciliation to verify.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    On Error GoTo 0
    MsgBox "ImportETLOutput error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' PRIVATE HELPER - WorkbookHasSheet
'===============================================================================
Private Function WorkbookHasSheet(ByVal wb As Workbook, ByVal shName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(shName)
    WorkbookHasSheet = Not ws Is Nothing
    On Error GoTo 0
End Function
