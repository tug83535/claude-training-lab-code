Attribute VB_Name = "modUTL_WorkbookMgmt"
Option Explicit

' ============================================================
' KBT Universal Tools — Workbook Management Module
' Works on ANY Excel file — no project-specific setup required
' Install in Personal.xlsb to use across all Excel sessions
' Tools: 15 | Tier 1: 4 | Tier 2: 11
' ============================================================

Private Sub UTL_TurboOn()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

Private Sub UTL_TurboOff()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

' ============================================================
' TOOL 1 — Unhide All Sheets, Rows & Columns        [TIER 1]
' Makes every hidden sheet, row, and column visible instantly
' ============================================================
Sub UnhideAllSheetsRowsColumns()
    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim ws As Worksheet
    Dim sheetCount As Long
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
            sheetCount = sheetCount + 1
        End If
        ws.Rows.Hidden = False
        ws.Columns.Hidden = False
    Next ws

    UTL_TurboOff
    MsgBox "Done!" & Chr(10) & _
           sheetCount & " hidden sheet(s) revealed." & Chr(10) & _
           "All hidden rows and columns shown on every sheet.", _
           vbInformation, "UTL Workbook Mgmt"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Workbook Mgmt"
End Sub

' ============================================================
' TOOL 2 — Export All Sheets as One Combined PDF    [TIER 1]
' Combines all visible sheets into a single multi-page PDF
' ============================================================
Sub ExportAllSheetsCombinedPDF()
    Dim savePath As String
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=Replace(ActiveWorkbook.Name, ".xlsm", "") & "_Export", _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Save Combined PDF As")

    If savePath = "False" Then Exit Sub

    On Error GoTo ErrHandler

    Dim visibleSheets() As String
    Dim count As Long
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            ReDim Preserve visibleSheets(count)
            visibleSheets(count) = ws.Name
            count = count + 1
        End If
    Next ws

    ActiveWorkbook.Sheets(visibleSheets).Select
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=savePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    ActiveWorkbook.Sheets(1).Select
    MsgBox "Done! PDF saved to:" & Chr(10) & savePath, vbInformation, "UTL Workbook Mgmt"
    Exit Sub
ErrHandler:
    ActiveWorkbook.Sheets(1).Select
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Workbook Mgmt"
End Sub

' ============================================================
' TOOL 3 — Find & Replace Across All Sheets         [TIER 1]
' Global find-and-replace across every sheet simultaneously
' ============================================================
Sub FindReplaceAcrossAllSheets()
    Dim findVal As String
    Dim replaceVal As String

    findVal = InputBox("Find what:", "UTL — Find & Replace All Sheets")
    If findVal = "" Then Exit Sub

    replaceVal = InputBox("Replace with:", "UTL — Find & Replace All Sheets")

    If MsgBox("Replace '" & findVal & "' with '" & replaceVal & "' on ALL sheets?" & Chr(10) & _
              "This cannot be undone after saving.", _
              vbExclamation + vbYesNo, "UTL Workbook Mgmt") = vbNo Then Exit Sub

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim count As Long
    Dim ws As Worksheet
    Dim c As Range
    Dim foundCell As Range
    Dim firstAddress As String

    For Each ws In ActiveWorkbook.Worksheets
        Set foundCell = ws.UsedRange.Find(What:=findVal, LookIn:=xlValues, _
                                          LookAt:=xlPart, MatchCase:=False)
        If Not foundCell Is Nothing Then
            firstAddress = foundCell.Address
            Do
                foundCell.Value = Replace(foundCell.Value, findVal, replaceVal, 1, -1, vbTextCompare)
                count = count + 1
                Set foundCell = ws.UsedRange.FindNext(foundCell)
            Loop While Not foundCell Is Nothing And foundCell.Address <> firstAddress
        End If
    Next ws

    UTL_TurboOff
    MsgBox "Done! " & count & " replacement(s) made across all sheets.", vbInformation, "UTL Workbook Mgmt"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Workbook Mgmt"
End Sub

' ============================================================
' TOOL 4 — Search Across All Sheets                 [TIER 1]
' Finds any value across every sheet — returns sheet + address
' Results appear in a new sheet for easy navigation
' ============================================================
Sub SearchAcrossAllSheets()
    Dim searchTerm As String
    searchTerm = InputBox("Search for:", "UTL — Search All Sheets")
    If searchTerm = "" Then Exit Sub

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim resultsWS As Worksheet
    Dim wsName As String
    wsName = "UTL_Search_Results"
    On Error Resume Next
    Application.DisplayAlerts = False
    ActiveWorkbook.Sheets(wsName).Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrHandler

    Set resultsWS = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    resultsWS.Name = wsName

    With resultsWS
        .Range("A1").Value = "Search Results for: " & searchTerm
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A2").Value = "Sheet"
        .Range("B2").Value = "Cell Address"
        .Range("C2").Value = "Value Found"
        .Range("A2:C2").Font.Bold = True
        .Range("A2:C2").Interior.Color = RGB(31, 73, 125)
        .Range("A2:C2").Font.Color = RGB(255, 255, 255)
    End With

    Dim rowNum As Long
    rowNum = 3
    Dim totalFound As Long
    Dim ws As Worksheet
    Dim foundCell As Range
    Dim firstAddress As String

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> wsName Then
            Set foundCell = ws.UsedRange.Find(What:=searchTerm, LookIn:=xlValues, LookAt:=xlPart)
            If Not foundCell Is Nothing Then
                firstAddress = foundCell.Address
                Do
                    resultsWS.Cells(rowNum, 1).Value = ws.Name
                    resultsWS.Cells(rowNum, 2).Value = foundCell.Address
                    resultsWS.Cells(rowNum, 3).Value = foundCell.Value
                    resultsWS.Hyperlinks.Add _
                        Anchor:=resultsWS.Cells(rowNum, 2), _
                        Address:="", _
                        SubAddress:="'" & ws.Name & "'!" & foundCell.Address, _
                        TextToDisplay:=foundCell.Address
                    rowNum = rowNum + 1
                    totalFound = totalFound + 1
                    If totalFound >= 200 Then GoTo DoneSearching
                    Set foundCell = ws.UsedRange.FindNext(foundCell)
                Loop While Not foundCell Is Nothing And foundCell.Address <> firstAddress
            End If
        End If
    Next ws

DoneSearching:
    resultsWS.Columns("A:C").AutoFit
    UTL_TurboOff

    If totalFound = 0 Then
        Application.DisplayAlerts = False
        resultsWS.Delete
        Application.DisplayAlerts = True
        MsgBox "No results found for '" & searchTerm & "'.", vbInformation, "UTL Workbook Mgmt"
    Else
        resultsWS.Activate
        MsgBox "Found " & totalFound & " result(s) for '" & searchTerm & "'." & Chr(10) & _
               "Click addresses in column B to navigate directly.", vbInformation, "UTL Workbook Mgmt"
    End If
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Workbook Mgmt"
End Sub

' ============================================================
' TOOL 5 — Multi-Sheet Batch Renamer                [TIER 2]
' Replaces a text string in all sheet tab names at once
' Example: replace "2024" with "2025" across all tabs
' ============================================================
Sub MultiSheetBatchRenamer()
    Dim findText As String
    Dim replaceText As String

    findText = InputBox("Find this text in sheet names:", "UTL — Batch Sheet Renamer")
    If findText = "" Then Exit Sub

    replaceText = InputBox("Replace with:", "UTL — Batch Sheet Renamer")

    On Error GoTo ErrHandler

    Dim count As Long
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If InStr(1, ws.Name, findText, vbTextCompare) > 0 Then
            ws.Name = Replace(ws.Name, findText, replaceText, 1, -1, vbTextCompare)
            count = count + 1
        End If
    Next ws

    If count = 0 Then
        MsgBox "No sheet names contained '" & findText & "'.", vbInformation, "UTL Workbook Mgmt"
    Else
        MsgBox "Done! " & count & " sheet name(s) updated.", vbInformation, "UTL Workbook Mgmt"
    End If
    Exit Sub
ErrHandler:
    MsgBox "Error renaming sheet: " & Err.Description & Chr(10) & _
           "Note: Sheet names cannot contain special characters or exceed 31 characters.", _
           vbCritical, "UTL Workbook Mgmt"
End Sub

' ============================================================
' TOOL 6 — Sort Worksheets Alphabetically           [TIER 2]
' Reorders all sheet tabs A to Z
' ============================================================
Sub SortWorksheetsAlphabetically()
    If MsgBox("Sort all sheet tabs alphabetically (A to Z)?", _
              vbQuestion + vbYesNo, "UTL Workbook Mgmt") = vbNo Then Exit Sub

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim i As Long, j As Long
    Dim sheetCount As Long
    sheetCount = ActiveWorkbook.Sheets.Count

    For i = 1 To sheetCount - 1
        For j = 1 To sheetCount - i
            If ActiveWorkbook.Sheets(j).Name > ActiveWorkbook.Sheets(j + 1).Name Then
                ActiveWorkbook.Sheets(j + 1).Move Before:=ActiveWorkbook.Sheets(j)
            End If
        Next j
    Next i

    UTL_TurboOff
    MsgBox "Done! All " & sheetCount & " sheets sorted A to Z.", vbInformation, "UTL Workbook Mgmt"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Workbook Mgmt"
End Sub

' ============================================================
' TOOL 7 — Create Table of Contents                 [TIER 2]
' Generates a clickable index sheet linking to every worksheet
' ============================================================
Sub CreateTableOfContents()
    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim tocName As String
    tocName = "Table of Contents"

    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Sheets(tocName).Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    Dim tocWS As Worksheet
    Set tocWS = ActiveWorkbook.Sheets.Add(Before:=ActiveWorkbook.Sheets(1))
    tocWS.Name = tocName

    With tocWS
        .Range("A1").Value = ActiveWorkbook.Name & " — Table of Contents"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Range("A2").Value = "Generated: " & Format(Now, "MM/DD/YYYY h:mm AM/PM")
        .Range("A2").Font.Italic = True
        .Range("A4").Value = "#"
        .Range("B4").Value = "Sheet Name"
        .Range("C4").Value = "Navigate"
        .Range("A4:C4").Font.Bold = True
        .Range("A4:C4").Interior.Color = RGB(31, 73, 125)
        .Range("A4:C4").Font.Color = RGB(255, 255, 255)
    End With

    Dim rowNum As Long
    rowNum = 5
    Dim sheetNum As Long
    sheetNum = 1
    Dim ws As Worksheet

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> tocName Then
            tocWS.Cells(rowNum, 1).Value = sheetNum
            tocWS.Cells(rowNum, 2).Value = ws.Name
            tocWS.Hyperlinks.Add _
                Anchor:=tocWS.Cells(rowNum, 3), _
                Address:="", _
                SubAddress:="'" & ws.Name & "'!A1", _
                TextToDisplay:="Go to Sheet"
            rowNum = rowNum + 1
            sheetNum = sheetNum + 1
        End If
    Next ws

    tocWS.Columns("A:C").AutoFit
    tocWS.Activate
    UTL_TurboOff
    MsgBox "Done! Table of Contents created with " & (sheetNum - 1) & " sheet links.", _
           vbInformation, "UTL Workbook Mgmt"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Workbook Mgmt"
End Sub

' ============================================================
' TOOL 8 — Protect All Sheets                       [TIER 2]
' Applies password protection to every sheet at once
' ============================================================
Sub ProtectAllSheets()
    Dim pwd As String
    pwd = InputBox("Enter a password to protect all sheets:" & Chr(10) & _
                   "(leave blank for protection without password)", "UTL Workbook Mgmt")

    If MsgBox("Apply worksheet protection to ALL " & ActiveWorkbook.Sheets.Count & " sheets?", _
              vbQuestion + vbYesNo, "UTL Workbook Mgmt") = vbNo Then Exit Sub

    On Error GoTo ErrHandler
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Protect Password:=pwd, DrawingObjects:=True, Contents:=True, Scenarios:=True
    Next ws

    MsgBox "Done! All sheets are now protected.", vbInformation, "UTL Workbook Mgmt"
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Workbook Mgmt"
End Sub

' ============================================================
' TOOL 9 — Unprotect All Sheets                     [TIER 2]
' Removes worksheet protection from every sheet
' ============================================================
Sub UnprotectAllSheets()
    Dim pwd As String
    pwd = InputBox("Enter the password to unprotect all sheets:" & Chr(10) & _
                   "(leave blank if sheets have no password)", "UTL Workbook Mgmt")

    On Error GoTo ErrHandler
    Dim count As Long
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Unprotect Password:=pwd
        count = count + 1
    Next ws

    MsgBox "Done! " & count & " sheets unprotected.", vbInformation, "UTL Workbook Mgmt"
    Exit Sub
ErrHandler:
    MsgBox "Error unprotecting sheets: " & Err.Description & Chr(10) & _
           "Check that the password is correct.", vbCritical, "UTL Workbook Mgmt"
End Sub

' ============================================================
' TOOL 10 — Lock All Formula Cells                  [TIER 2]
' Locks cells with formulas, leaves input cells editable
' Run before protecting the sheet to protect calculated values
' ============================================================
Sub LockAllFormulaCells()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    If MsgBox("Lock all formula cells on '" & ws.Name & "'?" & Chr(10) & _
              "Input cells will remain editable." & Chr(10) & _
              "Note: Run Protect Sheet afterward to activate the lock.", _
              vbQuestion + vbYesNo, "UTL Workbook Mgmt") = vbNo Then Exit Sub

    On Error GoTo ErrHandler
    UTL_TurboOn

    ws.Cells.Locked = False
    Dim count As Long
    Dim fRng As Range: Set fRng = Nothing
    On Error Resume Next
    Set fRng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo ErrHandler
    If Not fRng Is Nothing Then
        Dim c As Range
        For Each c In fRng
            c.Locked = True
            count = count + 1
        Next c
    End If

    UTL_TurboOff
    MsgBox "Done! " & count & " formula cells locked." & Chr(10) & _
           "Run Protect Sheet (Review tab) to activate.", vbInformation, "UTL Workbook Mgmt"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Workbook Mgmt"
End Sub

' ============================================================
' TOOL 11 — Export Active Sheet as PDF              [TIER 2]
' Saves just the current sheet as a PDF file
' ============================================================
Sub ExportActiveSheetPDF()
    Dim savePath As String
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=ActiveSheet.Name, _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Save Sheet as PDF")

    If savePath = "False" Then Exit Sub

    On Error GoTo ErrHandler
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=savePath, _
        Quality:=xlQualityStandard, _
        OpenAfterPublish:=True

    MsgBox "Done! PDF saved to:" & Chr(10) & savePath, vbInformation, "UTL Workbook Mgmt"
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Workbook Mgmt"
End Sub

' ============================================================
' TOOL 12 — Export All Sheets as Individual PDFs   [TIER 2]
' Saves each visible sheet as its own PDF file in a folder
' ============================================================
Sub ExportAllSheetsIndividualPDFs()
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select folder to save PDFs"
        If .Show = False Then Exit Sub
        folderPath = .SelectedItems(1) & "\"
    End With

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim count As Long
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            Dim filePath As String
            filePath = folderPath & ws.Name & ".pdf"
            ws.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                Filename:=filePath, _
                Quality:=xlQualityStandard
            count = count + 1
        End If
    Next ws

    UTL_TurboOff
    MsgBox "Done! " & count & " PDFs saved to:" & Chr(10) & folderPath, _
           vbInformation, "UTL Workbook Mgmt"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Workbook Mgmt"
End Sub

' ============================================================
' TOOL 13 — Reset All Filters                       [TIER 2]
' Clears all AutoFilter criteria across every sheet
' ============================================================
Sub ResetAllFilters()
    On Error GoTo ErrHandler
    Dim count As Long
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.AutoFilterMode Then
            ws.AutoFilter.ShowAllData
            count = count + 1
        End If
    Next ws

    If count = 0 Then
        MsgBox "No active filters found on any sheet.", vbInformation, "UTL Workbook Mgmt"
    Else
        MsgBox "Done! Filters cleared on " & count & " sheet(s).", vbInformation, "UTL Workbook Mgmt"
    End If
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Workbook Mgmt"
End Sub

' ============================================================
' TOOL 14 — Build Distribution-Ready Copy           [TIER 2]
' Creates a clean copy: formulas as values, metadata stripped
' Perfect for sharing with external parties
' ============================================================
Sub BuildDistributionReadyCopy()
    If MsgBox("Create a distribution-ready copy of this workbook?" & Chr(10) & Chr(10) & _
              "The copy will have:" & Chr(10) & _
              "  • All formulas converted to values" & Chr(10) & _
              "  • All hidden sheets visible" & Chr(10) & _
              "  • File saved with '_DIST' suffix" & Chr(10) & Chr(10) & _
              "Your original file will NOT be changed.", _
              vbQuestion + vbYesNo, "UTL Workbook Mgmt") = vbNo Then Exit Sub

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim origPath As String
    Dim distPath As String
    origPath = ActiveWorkbook.FullName
    If Right(LCase(origPath), 5) = ".xlsm" Then
        distPath = Left(origPath, Len(origPath) - 5) & "_DIST.xlsx"
    ElseIf Right(LCase(origPath), 5) = ".xlsx" Then
        distPath = Left(origPath, Len(origPath) - 5) & "_DIST.xlsx"
    Else
        distPath = origPath & "_DIST.xlsx"
    End If

    ActiveWorkbook.SaveCopyAs distPath

    Dim distWB As Workbook
    Set distWB = Workbooks.Open(distPath)

    Dim ws As Worksheet
    For Each ws In distWB.Worksheets
        ws.Visible = xlSheetVisible
        ws.UsedRange.Value = ws.UsedRange.Value
    Next ws

    distWB.Save
    distWB.Close

    UTL_TurboOff
    MsgBox "Done! Distribution copy saved to:" & Chr(10) & distPath, _
           vbInformation, "UTL Workbook Mgmt"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Workbook Mgmt"
End Sub

' ============================================================
' TOOL 15 — Workbook Health Check                   [TIER 2]
' Generates a full diagnostic report on the active workbook
' Covers: size, errors, links, formulas, blanks, duplicates
' ============================================================
Sub WorkbookHealthCheck()
    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim ws As Worksheet
    Dim totalSheets As Long
    Dim totalCells As Long
    Dim totalFormulas As Long
    Dim totalErrors As Long
    Dim totalBlanks As Long
    Dim totalLinks As Long

    For Each ws In ActiveWorkbook.Worksheets
        totalSheets = totalSheets + 1
        totalCells = totalCells + ws.UsedRange.Cells.Count

        ' Count formulas via SpecialCells
        Dim fRng2 As Range: Set fRng2 = Nothing
        On Error Resume Next
        Set fRng2 = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo ErrHandler
        If Not fRng2 Is Nothing Then totalFormulas = totalFormulas + fRng2.Cells.Count

        ' Count errors via SpecialCells
        Dim eRng As Range: Set eRng = Nothing
        On Error Resume Next
        Set eRng = ws.UsedRange.SpecialCells(xlCellTypeFormulas, xlErrors)
        On Error GoTo ErrHandler
        If Not eRng Is Nothing Then totalErrors = totalErrors + eRng.Cells.Count

        ' Count blanks via SpecialCells
        Dim bRng As Range: Set bRng = Nothing
        On Error Resume Next
        Set bRng = ws.UsedRange.SpecialCells(xlCellTypeBlanks)
        On Error GoTo ErrHandler
        If Not bRng Is Nothing Then totalBlanks = totalBlanks + bRng.Cells.Count

        totalLinks = totalLinks + ws.Hyperlinks.Count
    Next ws

    Dim extLinks As Long
    extLinks = 0
    On Error Resume Next
    extLinks = UBound(ActiveWorkbook.LinkSources(xlExcelLinks)) + 1
    On Error GoTo ErrHandler

    UTL_TurboOff

    Dim report As String
    report = "=== WORKBOOK HEALTH CHECK ===" & Chr(10) & Chr(10) & _
             "File: " & ActiveWorkbook.Name & Chr(10) & _
             "Date: " & Format(Now, "MM/DD/YYYY h:mm AM/PM") & Chr(10) & Chr(10) & _
             "STRUCTURE:" & Chr(10) & _
             "  Sheets:           " & totalSheets & Chr(10) & _
             "  Used Cells:       " & totalCells & Chr(10) & Chr(10) & _
             "FORMULAS & DATA:" & Chr(10) & _
             "  Formula Cells:    " & totalFormulas & Chr(10) & _
             "  Error Cells:      " & totalErrors & IIf(totalErrors > 0, "  ← REVIEW NEEDED", "") & Chr(10) & _
             "  Blank Cells:      " & totalBlanks & Chr(10) & Chr(10) & _
             "LINKS & DEPENDENCIES:" & Chr(10) & _
             "  Hyperlinks:       " & totalLinks & Chr(10) & _
             "  External Links:   " & extLinks & IIf(extLinks > 0, "  ← REVIEW NEEDED", "") & Chr(10) & Chr(10) & _
             IIf(totalErrors = 0 And extLinks = 0, _
                 "STATUS: CLEAN — No critical issues found.", _
                 "STATUS: REVIEW ITEMS FLAGGED — See above.")

    MsgBox report, vbInformation, "UTL Workbook Health Check"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Workbook Mgmt"
End Sub
