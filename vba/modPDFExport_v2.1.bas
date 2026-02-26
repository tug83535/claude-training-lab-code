Attribute VB_Name = "modPDFExport"
Option Explicit

'===============================================================================
' modPDFExport - Professional Batch PDF Export
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Export selected report sheets to individual or combined PDFs.
'           Applies print settings (landscape, fit-to-page, headers/footers)
'           before exporting. Auto-names files with date stamps.
'
' VERSION:  2.1.0
' CHANGES:  v2.0 -> v2.1:
'           + ISSUE-007: Replaced hardcoded REPORT_SHEETS constant that
'             contained literal "Functional P&L Summary - Jan 25" etc.
'             Now builds the list dynamically from modConfig constants
'             (SH_PL_TREND, SH_PROD_SUMMARY, SH_FUNC_TREND, SH_FUNC_JAN,
'             SH_FUNC_FEB, SH_FUNC_MAR, SH_CHECKS). This means fiscal year
'             rollover only requires changing FISCAL_YEAR in modConfig.
'           + Added proper error handler and TurboOff to ExportSingleSheet
'===============================================================================

'===============================================================================
' GetReportSheetList - Build the report package sheet list dynamically
' FIX (v2.1 — ISSUE-007):
' v2.0 used a hardcoded Private Const with literal "Jan 25" etc.
' v2.1 builds from modConfig constants so fiscal year rollover works.
'===============================================================================
Private Function GetReportSheetList() As Variant
    Dim sheets(0 To 6) As String
    sheets(0) = SH_PL_TREND
    sheets(1) = SH_PROD_SUMMARY
    sheets(2) = SH_FUNC_TREND
    sheets(3) = SH_FUNC_JAN
    sheets(4) = SH_FUNC_FEB
    sheets(5) = SH_FUNC_MAR
    sheets(6) = SH_CHECKS
    GetReportSheetList = sheets
End Function

'===============================================================================
' ExportReportPackage - Export all report sheets to a single PDF
'===============================================================================
Public Sub ExportReportPackage()
    On Error GoTo ErrHandler
    
    Dim exportPath As String
    exportPath = GetExportPath("KBT_Report_Package")
    If exportPath = "" Then Exit Sub
    
    modPerformance.TurboOn
    modPerformance.UpdateStatus "Preparing report package...", 0
    
    ' Build array of sheets to export (v2.1: dynamic, not hardcoded)
    Dim sheetNames As Variant: sheetNames = GetReportSheetList()
    Dim validSheets() As String
    Dim count As Long: count = 0
    
    Dim i As Long
    For i = 0 To UBound(sheetNames)
        If modConfig.SheetExists(sheetNames(i)) Then
            ReDim Preserve validSheets(count)
            validSheets(count) = sheetNames(i)
            count = count + 1
        End If
    Next i
    
    If count = 0 Then
        modPerformance.TurboOff
        MsgBox "No report sheets found to export.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' Apply print settings to each sheet
    For i = 0 To UBound(validSheets)
        modPerformance.UpdateStatus "Formatting " & validSheets(i) & "...", i / count
        ApplyPrintSettings ThisWorkbook.Worksheets(validSheets(i))
    Next i
    
    ' Select all sheets for combined export
    ThisWorkbook.Worksheets(validSheets(0)).Select
    For i = 1 To UBound(validSheets)
        ThisWorkbook.Worksheets(validSheets(i)).Select Replace:=False
    Next i
    
    ' Export
    modPerformance.UpdateStatus "Exporting PDF...", 0.9
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=exportPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    
    ' Deselect
    ThisWorkbook.Worksheets(validSheets(0)).Select
    
    modPerformance.TurboOff
    modLogger.LogAction "modPDFExport", "ExportReportPackage", _
                        count & " sheets exported to: " & exportPath, _
                        modPerformance.ElapsedSeconds()
    
    MsgBox "Report package exported (" & count & " sheets):" & vbCrLf & exportPath, _
           vbInformation, APP_NAME
    Exit Sub
    
ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modPDFExport", "ERROR", Err.Description
    MsgBox "PDF export error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' ExportSingleSheet - Export the active sheet to PDF
'===============================================================================
Public Sub ExportSingleSheet()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim exportPath As String
    exportPath = GetExportPath("KBT_" & CleanSheetName(ws.Name))
    If exportPath = "" Then Exit Sub
    
    modPerformance.TurboOn
    ApplyPrintSettings ws
    
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=exportPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    
    modPerformance.TurboOff
    modLogger.LogAction "modPDFExport", "ExportSingleSheet", ws.Name & " -> " & exportPath
    Exit Sub
    
ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modPDFExport", "ERROR", "ExportSingleSheet: " & Err.Description
    MsgBox "Export error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' ApplyPrintSettings - Professional print configuration
'===============================================================================
Private Sub ApplyPrintSettings(ByVal ws As Worksheet)
    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperLetter
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False  ' Let it flow to multiple pages vertically
        
        ' Professional headers/footers
        .LeftHeader = "&""Calibri,Bold""&10Keystone BenefitTech, Inc."
        .CenterHeader = "&""Calibri,Bold""&10" & ws.Name
        .RightHeader = "&""Calibri""&9CONFIDENTIAL"
        .LeftFooter = "&""Calibri""&8Printed: &D &T"
        .CenterFooter = "&""Calibri""&8Page &P of &N"
        .RightFooter = "&""Calibri""&8" & APP_NAME & " v" & APP_VERSION
        
        ' Margins (inches)
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        
        ' Repeat row 1 on every page
        .PrintTitleRows = "$1:$1"
        .PrintGridlines = False
        .BlackAndWhite = False
    End With
End Sub

'===============================================================================
' GetExportPath - Build export file path with SaveAs dialog fallback
'===============================================================================
Private Function GetExportPath(ByVal baseName As String) As String
    Dim defaultPath As String
    defaultPath = Environ("USERPROFILE") & "\Desktop\" & _
                  baseName & "_" & Format(Now, "yyyymmdd") & ".pdf"
    
    GetExportPath = Application.GetSaveAsFilename( _
        InitialFileName:=defaultPath, _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Export PDF Report")
    
    If GetExportPath = "False" Then GetExportPath = ""
End Function

'===============================================================================
' CleanSheetName - Remove invalid filename characters
'===============================================================================
Private Function CleanSheetName(ByVal s As String) As String
    Dim invalid As Variant: invalid = Array("/", "\", ":", "*", "?", """", "<", ">", "|", " ")
    Dim i As Long
    CleanSheetName = s
    For i = 0 To UBound(invalid)
        CleanSheetName = Replace(CleanSheetName, CStr(invalid(i)), "_")
    Next i
End Function
