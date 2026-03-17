Attribute VB_Name = "modImport"
Option Explicit

'===============================================================================
' modImport - GL Data Import Pipeline
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Imports raw GL transaction data from an external CSV or Excel file
'           into the CrossfireHiddenWorksheet (GL detail sheet). Validates
'           column structure, checks for duplicates, and appends or replaces.
'
' PUBLIC SUBS:
'   ImportDataPipeline - Main entry (Action #17)
'
' DEPENDENCIES: modConfig, modPerformance, modLogger
' VERSION:  2.1.0
'===============================================================================

' Expected GL columns (must match CrossfireHiddenWorksheet layout)
Private Const EXPECTED_COLS As Long = 7  ' ID, Date, Dept, Product, Category, Vendor, Amount

'===============================================================================
' ImportDataPipeline - Import GL data from external file
'===============================================================================
Public Sub ImportDataPipeline()
    On Error GoTo ErrHandler

    ' Prompt user for file
    Dim filePath As String
    filePath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx;*.xls;*.csv),*.xlsx;*.xls;*.csv", _
        Title:="Select GL Data File to Import")

    If filePath = "False" Or filePath = "" Then Exit Sub

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Opening import file...", 0.05

    ' Determine file type and open
    Dim srcWB As Workbook
    Dim srcWS As Worksheet
    Dim isCSV As Boolean: isCSV = (LCase(Right(filePath, 4)) = ".csv")

    Set srcWB = Workbooks.Open(Filename:=filePath, ReadOnly:=True)
    Set srcWS = srcWB.Sheets(1)

    ' Validate column count
    Dim srcLastCol As Long: srcLastCol = srcWS.Cells(1, srcWS.Columns.Count).End(xlToLeft).Column
    Dim srcLastRow As Long: srcLastRow = srcWS.Cells(srcWS.Rows.Count, 1).End(xlUp).Row

    If srcLastCol < EXPECTED_COLS Then
        srcWB.Close SaveChanges:=False
        modPerformance.TurboOff
        MsgBox "Import file has " & srcLastCol & " columns." & vbCrLf & _
               "Expected at least " & EXPECTED_COLS & " columns:" & vbCrLf & _
               "ID, Date, Dept, Product, Category, Vendor, Amount", _
               vbCritical, APP_NAME
        Exit Sub
    End If

    modPerformance.UpdateStatus "Validating import data...", 0.2

    ' Validate header names (flexible matching)
    Dim hdrIssues As String: hdrIssues = ""
    Dim expectedHeaders As Variant
    expectedHeaders = Array("id", "date", "dept", "product", "category", "vendor", "amount")

    Dim c As Long
    For c = 1 To EXPECTED_COLS
        Dim srcHdr As String: srcHdr = LCase(Trim(CStr(srcWS.Cells(1, c).Value)))
        If InStr(srcHdr, CStr(expectedHeaders(c - 1))) = 0 Then
            hdrIssues = hdrIssues & "  Column " & c & ": Expected '" & CStr(expectedHeaders(c - 1)) & _
                        "', found '" & srcHdr & "'" & vbCrLf
        End If
    Next c

    ' Ask user how to import
    Dim importMode As VbMsgBoxResult
    importMode = MsgBox("Import " & (srcLastRow - 1) & " rows from:" & vbCrLf & _
                        filePath & vbCrLf & vbCrLf & _
                        IIf(hdrIssues <> "", "COLUMN WARNINGS:" & vbCrLf & hdrIssues & vbCrLf, "") & _
                        "YES = Replace existing GL data" & vbCrLf & _
                        "NO = Append to existing data", _
                        vbYesNoCancel + vbQuestion, APP_NAME & " - Import")

    If importMode = vbCancel Then
        srcWB.Close SaveChanges:=False
        modPerformance.TurboOff
        Exit Sub
    End If

    ' Get target sheet
    If Not modConfig.SheetExists(SH_GL) Then
        srcWB.Close SaveChanges:=False
        modPerformance.TurboOff
        MsgBox "Target sheet '" & SH_GL & "' not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    Dim wsGL As Worksheet: Set wsGL = ThisWorkbook.Worksheets(SH_GL)

    modPerformance.UpdateStatus "Importing data...", 0.4

    Dim startRow As Long
    If importMode = vbYes Then
        ' Replace: clear existing data
        Dim glLastRow As Long: glLastRow = wsGL.Cells(wsGL.Rows.Count, 1).End(xlUp).Row
        If glLastRow >= DATA_ROW_GL Then
            wsGL.Rows(DATA_ROW_GL & ":" & glLastRow).ClearContents
        End If
        startRow = DATA_ROW_GL
    Else
        ' Append
        startRow = wsGL.Cells(wsGL.Rows.Count, 1).End(xlUp).Row + 1
        If startRow < DATA_ROW_GL Then startRow = DATA_ROW_GL
    End If

    ' Copy data row by row (safer than bulk copy for mixed formats)
    Dim importCount As Long: importCount = 0
    Dim dupCount As Long: dupCount = 0
    Dim r As Long

    For r = 2 To srcLastRow
        ' Skip blank rows
        If Trim(CStr(srcWS.Cells(r, 1).Value)) <> "" Then
            ' Copy each column
            wsGL.Cells(startRow, COL_GL_ID).Value = srcWS.Cells(r, 1).Value
            wsGL.Cells(startRow, COL_GL_DATE).Value = srcWS.Cells(r, 2).Value
            wsGL.Cells(startRow, COL_GL_DEPT).Value = srcWS.Cells(r, 3).Value
            wsGL.Cells(startRow, COL_GL_PRODUCT).Value = srcWS.Cells(r, 4).Value
            wsGL.Cells(startRow, COL_GL_CATEGORY).Value = srcWS.Cells(r, 5).Value
            wsGL.Cells(startRow, COL_GL_VENDOR).Value = srcWS.Cells(r, 6).Value
            wsGL.Cells(startRow, COL_GL_AMOUNT).Value = srcWS.Cells(r, 7).Value

            ' Format the amount column
            wsGL.Cells(startRow, COL_GL_AMOUNT).NumberFormat = "#,##0.00"

            startRow = startRow + 1
            importCount = importCount + 1
        End If

        If r Mod 50 = 0 Then
            modPerformance.UpdateStatus "Importing row " & r & " of " & srcLastRow & "...", _
                0.4 + 0.5 * (r / srcLastRow)
        End If
    Next r

    ' Close source file
    srcWB.Close SaveChanges:=False

    ' Format date column
    wsGL.Columns(COL_GL_DATE).NumberFormat = "yyyy-mm-dd"

    Dim elapsed As Double: elapsed = modPerformance.ElapsedSeconds()
    modPerformance.TurboOff

    modLogger.LogAction "modImport", "ImportDataPipeline", _
        importCount & " rows imported (" & IIf(importMode = vbYes, "replace", "append") & _
        ") from " & Dir(filePath)

    MsgBox "GL DATA IMPORT COMPLETE" & vbCrLf & String(30, "=") & vbCrLf & vbCrLf & _
           "Rows Imported:  " & importCount & vbCrLf & _
           "Mode:           " & IIf(importMode = vbYes, "Replace", "Append") & vbCrLf & _
           "Source:         " & Dir(filePath) & vbCrLf & _
           "Time:           " & Format(elapsed, "0.0") & " seconds", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not srcWB Is Nothing Then srcWB.Close SaveChanges:=False
    modPerformance.TurboOff
    On Error GoTo 0
    modLogger.LogAction "modImport", "ERROR", Err.Description
    MsgBox "Import error: " & Err.Description, vbCritical, APP_NAME
End Sub
