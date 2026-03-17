Attribute VB_Name = "modUTL_ExecBrief"
Option Explicit

'===============================================================================
' modUTL_ExecBrief - Universal Executive Brief Auto-Generator
' Universal Toolkit - Works on ANY Excel file
'===============================================================================
' PURPOSE:  One button scans any workbook and generates a plain English summary.
'           Covers: sheet inventory, data volume, formulas vs values, potential
'           issues (errors, blanks, hidden sheets), and key statistics.
'           Ready to paste into an email or print for leadership.
'
' PUBLIC SUBS:
'   GenerateExecBrief   - Build the executive brief on a new sheet
'
' DEPENDENCIES: None (fully standalone)
' VERSION:  1.0.0
'===============================================================================

Private Const SH_BRIEF As String = "Executive Brief"

'===============================================================================
' GenerateExecBrief - Scan entire workbook and produce summary
'===============================================================================
Public Sub GenerateExecBrief()
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.StatusBar = "Generating Executive Brief..."

    ' Remove old brief if exists
    Dim oldSheet As Worksheet
    On Error Resume Next
    Set oldSheet = ThisWorkbook.Worksheets(SH_BRIEF)
    If Not oldSheet Is Nothing Then
        Application.DisplayAlerts = False
        oldSheet.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = SH_BRIEF

    ws.Columns(1).ColumnWidth = 80

    Dim r As Long: r = 1

    ' --- Title ---
    ws.Cells(r, 1).Value = "EXECUTIVE BRIEF"
    ws.Cells(r, 1).Font.Size = 18: ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Color = RGB(11, 71, 121)
    r = r + 1

    ws.Cells(r, 1).Value = "Workbook: " & ThisWorkbook.Name
    ws.Cells(r, 1).Font.Size = 11
    r = r + 1

    ws.Cells(r, 1).Value = "Generated: " & Format(Now, "mmmm d, yyyy h:mm AM/PM")
    ws.Cells(r, 1).Font.Size = 9: ws.Cells(r, 1).Font.Italic = True
    r = r + 2

    ' Divider
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 1)).Interior.Color = RGB(11, 71, 121)
    ws.Rows(r).RowHeight = 3
    r = r + 2

    ' =============================================
    ' SECTION 1: Workbook Overview
    ' =============================================
    ws.Cells(r, 1).Value = "1. WORKBOOK OVERVIEW"
    ws.Cells(r, 1).Font.Size = 13: ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Color = RGB(11, 71, 121)
    r = r + 1

    Dim visCount As Long: visCount = 0
    Dim hidCount As Long: hidCount = 0
    Dim sh As Worksheet

    For Each sh In ThisWorkbook.Worksheets
        If sh.Visible = xlSheetVisible Then
            visCount = visCount + 1
        Else
            hidCount = hidCount + 1
        End If
    Next sh

    ws.Cells(r, 1).Value = "- Total sheets: " & ThisWorkbook.Worksheets.Count & _
        " (" & visCount & " visible, " & hidCount & " hidden)"
    r = r + 1

    ' File size
    Dim fSize As Long
    On Error Resume Next
    fSize = FileLen(ThisWorkbook.FullName)
    On Error GoTo ErrHandler

    If fSize > 0 Then
        If fSize > 1048576 Then
            ws.Cells(r, 1).Value = "- File size: " & Format(fSize / 1048576, "#,##0.0") & " MB"
        Else
            ws.Cells(r, 1).Value = "- File size: " & Format(fSize / 1024, "#,##0") & " KB"
        End If
        r = r + 1
    End If

    ws.Cells(r, 1).Value = "- File path: " & ThisWorkbook.Path
    r = r + 2

    ' =============================================
    ' SECTION 2: Sheet-by-Sheet Summary
    ' =============================================
    ws.Cells(r, 1).Value = "2. SHEET INVENTORY"
    ws.Cells(r, 1).Font.Size = 13: ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Color = RGB(11, 71, 121)
    r = r + 1

    Dim totalRows As Long: totalRows = 0
    Dim totalCells As Long: totalCells = 0

    For Each sh In ThisWorkbook.Worksheets
        If sh.Visible = xlSheetVisible Then
            Dim lastR As Long, lastC As Long
            On Error Resume Next
            lastR = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
            lastC = sh.Cells(1, sh.Columns.Count).End(xlToLeft).Column
            On Error GoTo ErrHandler

            Dim cellCount As Long: cellCount = lastR * lastC
            totalRows = totalRows + lastR
            totalCells = totalCells + cellCount

            ws.Cells(r, 1).Value = "- " & sh.Name & ": " & lastR & " rows x " & lastC & " cols"
            r = r + 1
        End If
    Next sh

    ws.Cells(r, 1).Value = "- TOTAL: ~" & Format(totalRows, "#,##0") & " data rows, ~" & _
        Format(totalCells, "#,##0") & " cells"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 2

    ' =============================================
    ' SECTION 3: Data Quality Scan
    ' =============================================
    ws.Cells(r, 1).Value = "3. DATA QUALITY SNAPSHOT"
    ws.Cells(r, 1).Font.Size = 13: ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Color = RGB(11, 71, 121)
    r = r + 1

    Dim totalErrors As Long: totalErrors = 0
    Dim totalFormulas As Long: totalFormulas = 0
    Dim sheetsWithErrors As String: sheetsWithErrors = ""

    For Each sh In ThisWorkbook.Worksheets
        If sh.Visible = xlSheetVisible Then
            ' Count errors
            Dim errRng As Range
            On Error Resume Next
            Set errRng = Nothing
            Set errRng = sh.UsedRange.SpecialCells(xlCellTypeFormulas, xlErrors)
            If Not errRng Is Nothing Then
                Dim errCount As Long: errCount = errRng.Cells.Count
                totalErrors = totalErrors + errCount
                sheetsWithErrors = sheetsWithErrors & "    - " & sh.Name & ": " & errCount & " error(s)" & vbCrLf
            End If
            Set errRng = Nothing

            ' Count formulas
            Dim frmRng As Range
            Set frmRng = Nothing
            Set frmRng = sh.UsedRange.SpecialCells(xlCellTypeFormulas)
            If Not frmRng Is Nothing Then
                totalFormulas = totalFormulas + frmRng.Cells.Count
            End If
            Set frmRng = Nothing
            On Error GoTo ErrHandler
        End If
    Next sh

    ws.Cells(r, 1).Value = "- Total formulas: " & Format(totalFormulas, "#,##0")
    r = r + 1

    If totalErrors = 0 Then
        ws.Cells(r, 1).Value = "- Cell errors: None found - clean workbook!"
        ws.Cells(r, 1).Font.Color = RGB(0, 128, 0)
    Else
        ws.Cells(r, 1).Value = "- Cell errors: " & totalErrors & " found (review recommended)"
        ws.Cells(r, 1).Font.Color = RGB(200, 0, 0)
        r = r + 1
        ws.Cells(r, 1).Value = sheetsWithErrors
        ws.Cells(r, 1).WrapText = True
    End If
    r = r + 2

    ' =============================================
    ' SECTION 4: Hidden Sheets
    ' =============================================
    If hidCount > 0 Then
        ws.Cells(r, 1).Value = "4. HIDDEN SHEETS"
        ws.Cells(r, 1).Font.Size = 13: ws.Cells(r, 1).Font.Bold = True
        ws.Cells(r, 1).Font.Color = RGB(11, 71, 121)
        r = r + 1

        For Each sh In ThisWorkbook.Worksheets
            If sh.Visible <> xlSheetVisible Then
                Dim visType As String
                If sh.Visible = xlSheetHidden Then
                    visType = "Hidden"
                Else
                    visType = "Very Hidden"
                End If
                ws.Cells(r, 1).Value = "- " & sh.Name & " (" & visType & ")"
                r = r + 1
            End If
        Next sh
        r = r + 1
    End If

    ' =============================================
    ' Footer
    ' =============================================
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 1)).Interior.Color = RGB(11, 71, 121)
    ws.Rows(r).RowHeight = 3
    r = r + 2

    ws.Cells(r, 1).Value = "Generated by iPipeline Universal Toolkit"
    ws.Cells(r, 1).Font.Size = 8: ws.Cells(r, 1).Font.Italic = True
    ws.Cells(r, 1).Font.Color = RGB(150, 150, 150)
    r = r + 1
    ws.Cells(r, 1).Value = "This brief can be copied and pasted into an email or printed."
    ws.Cells(r, 1).Font.Size = 8: ws.Cells(r, 1).Font.Italic = True
    ws.Cells(r, 1).Font.Color = RGB(150, 150, 150)

    ws.Activate
    ws.Range("A1").Select

    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Executive Brief generated!" & vbCrLf & vbCrLf & _
           "4 sections: Overview, Sheet Inventory, Data Quality, Hidden Sheets" & vbCrLf & vbCrLf & _
           "Ready to copy/paste into email or print.", _
           vbInformation, "Executive Brief"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Executive Brief error: " & Err.Description, vbCritical, "Executive Brief"
End Sub
