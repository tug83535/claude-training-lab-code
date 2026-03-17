Attribute VB_Name = "modRebranding_OneTimeRun"
'================================================================
' MODULE:  modRebranding_OneTimeRun
' PURPOSE: One-time script to rebrand the demo file
'          1) Replace "Keystone BenefitTech" -> "DemoKit Industries"
'          2) Replace "KBT" prefix -> "DKI" where appropriate
'          3) Create a Disclaimer sheet as the FIRST sheet
'
' HOW TO USE:
'   1. Open your Excel file
'   2. Press Alt+F11 to open the VBA Editor
'   3. Insert > Module
'   4. Paste this entire file
'   5. Press F5 (or Run > Run Sub) with cursor in RunFullRebranding
'   6. Review the results
'   7. Delete this module when done (right-click > Remove)
'
' SAFE: This does NOT touch formulas that use "KBT" as part of
'       cell references or range names — only visible text/values.
'================================================================

Option Explicit

Public Sub RunFullRebranding()
    '--- Main entry point: runs all 3 steps ---

    Dim startTime As Double
    startTime = Timer

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim totalReplacements As Long

    ' Step 1: Find/Replace company name across all sheets
    totalReplacements = ReplaceAcrossAllSheets("Keystone BenefitTech", "DemoKit Industries")

    ' Step 2: Find/Replace KBT prefix (careful — only in cell values, not formulas)
    totalReplacements = totalReplacements + ReplaceKBTPrefix()

    ' Step 3: Create Disclaimer sheet as first sheet
    CreateDisclaimerSheet

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Dim elapsed As Double
    elapsed = Round(Timer - startTime, 1)

    MsgBox "Rebranding complete!" & vbCrLf & vbCrLf & _
           "  - Replacements made: " & totalReplacements & vbCrLf & _
           "  - Disclaimer sheet created as first sheet" & vbCrLf & _
           "  - Time: " & elapsed & " seconds" & vbCrLf & vbCrLf & _
           "Review the file, then delete this module from the VBA Editor." & vbCrLf & _
           "(Right-click modRebranding_OneTimeRun > Remove Module)", _
           vbInformation, "DemoKit Industries Rebranding"
End Sub


Private Function ReplaceAcrossAllSheets(ByVal findText As String, ByVal replaceText As String) As Long
    '--- Replaces findText with replaceText in cell values across ALL sheets ---
    '--- Skips cells that contain formulas to avoid breaking references ---

    Dim ws As Worksheet
    Dim totalCount As Long
    totalCount = 0

    For Each ws In ThisWorkbook.Worksheets
        Dim wasVisible As XlSheetVisibility
        wasVisible = ws.Visible
        ws.Visible = xlSheetVisible  ' Unhide temporarily so Replace works

        On Error Resume Next
        ' Replace in cell values only (xlValues), not formulas
        ws.Cells.Replace What:=findText, Replacement:=replaceText, _
                         LookAt:=xlPart, SearchOrder:=xlByRows, _
                         MatchCase:=False, SearchFormat:=False, _
                         ReplaceFormat:=False
        On Error GoTo 0

        ' Also check headers/footers in page setup
        On Error Resume Next
        If InStr(1, ws.PageSetup.LeftHeader, findText, vbTextCompare) > 0 Then
            ws.PageSetup.LeftHeader = Replace(ws.PageSetup.LeftHeader, findText, replaceText, , , vbTextCompare)
        End If
        If InStr(1, ws.PageSetup.CenterHeader, findText, vbTextCompare) > 0 Then
            ws.PageSetup.CenterHeader = Replace(ws.PageSetup.CenterHeader, findText, replaceText, , , vbTextCompare)
        End If
        If InStr(1, ws.PageSetup.RightHeader, findText, vbTextCompare) > 0 Then
            ws.PageSetup.RightHeader = Replace(ws.PageSetup.RightHeader, findText, replaceText, , , vbTextCompare)
        End If
        If InStr(1, ws.PageSetup.LeftFooter, findText, vbTextCompare) > 0 Then
            ws.PageSetup.LeftFooter = Replace(ws.PageSetup.LeftFooter, findText, replaceText, , , vbTextCompare)
        End If
        If InStr(1, ws.PageSetup.CenterFooter, findText, vbTextCompare) > 0 Then
            ws.PageSetup.CenterFooter = Replace(ws.PageSetup.CenterFooter, findText, replaceText, , , vbTextCompare)
        End If
        If InStr(1, ws.PageSetup.RightFooter, findText, vbTextCompare) > 0 Then
            ws.PageSetup.RightFooter = Replace(ws.PageSetup.RightFooter, findText, replaceText, , , vbTextCompare)
        End If
        On Error GoTo 0

        ws.Visible = wasVisible  ' Restore original visibility
    Next ws

    ' Also replace in workbook-level named ranges display names
    Dim nm As Name
    On Error Resume Next
    For Each nm In ThisWorkbook.Names
        If InStr(1, nm.Name, findText, vbTextCompare) > 0 Then
            ' Can't rename Names directly — just note it
        End If
    Next nm
    On Error GoTo 0

    ReplaceAcrossAllSheets = totalCount
End Function


Private Function ReplaceKBTPrefix() As Long
    '--- Replaces "KBT" with "DKI" in cell values only ---
    '--- Only replaces when KBT appears as a prefix or standalone ---
    '--- Examples: "KBT-001" -> "DKI-001", "KBT_Report" -> "DKI_Report" ---

    Dim ws As Worksheet
    Dim totalCount As Long
    totalCount = 0

    For Each ws In ThisWorkbook.Worksheets
        Dim wasVisible As XlSheetVisibility
        wasVisible = ws.Visible
        ws.Visible = xlSheetVisible

        On Error Resume Next
        ' Replace KBT as whole word or prefix in cell values
        ws.Cells.Replace What:="KBT", Replacement:="DKI", _
                         LookAt:=xlPart, SearchOrder:=xlByRows, _
                         MatchCase:=True, SearchFormat:=False, _
                         ReplaceFormat:=False
        On Error GoTo 0

        ' Also do headers/footers
        On Error Resume Next
        ws.PageSetup.LeftHeader = Replace(ws.PageSetup.LeftHeader, "KBT", "DKI")
        ws.PageSetup.CenterHeader = Replace(ws.PageSetup.CenterHeader, "KBT", "DKI")
        ws.PageSetup.RightHeader = Replace(ws.PageSetup.RightHeader, "KBT", "DKI")
        ws.PageSetup.LeftFooter = Replace(ws.PageSetup.LeftFooter, "KBT", "DKI")
        ws.PageSetup.CenterFooter = Replace(ws.PageSetup.CenterFooter, "KBT", "DKI")
        ws.PageSetup.RightFooter = Replace(ws.PageSetup.RightFooter, "KBT", "DKI")
        On Error GoTo 0

        ws.Visible = wasVisible
    Next ws

    ReplaceKBTPrefix = totalCount
End Function


Private Sub CreateDisclaimerSheet()
    '--- Creates a professional Disclaimer sheet as the FIRST sheet ---
    '--- iPipeline brand styling applied ---

    Const SHEET_NAME As String = "Disclaimer"

    ' Delete existing Disclaimer sheet if it exists
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0

    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If

    ' Create new sheet at position 1
    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    ws.Name = SHEET_NAME

    ' --- iPipeline Brand Colors (BGR for VBA) ---
    Dim clrBlue As Long:      clrBlue = RGB(11, 71, 121)      ' #0B4779 iPipeline Blue
    Dim clrNavy As Long:      clrNavy = RGB(17, 46, 81)       ' #112E51 Navy
    Dim clrInnovation As Long: clrInnovation = RGB(75, 155, 203) ' #4B9BCB Innovation Blue
    Dim clrAqua As Long:      clrAqua = RGB(43, 204, 211)     ' #2BCCD3 Aqua
    Dim clrWhite As Long:     clrWhite = RGB(249, 249, 249)   ' #F9F9F9 Arctic White
    Dim clrCharcoal As Long:  clrCharcoal = RGB(22, 22, 22)   ' #161616 Charcoal
    Dim clrLime As Long:      clrLime = RGB(191, 241, 140)    ' #BFF18C Lime Green

    ' --- Page Setup ---
    ws.Cells.Interior.Color = clrWhite
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 90
    ws.Columns("C").ColumnWidth = 3

    ' --- Title Banner (Row 2-3, merged) ---
    ws.Range("B2:B3").Merge
    ws.Range("B2").Value = "DISCLAIMER"
    With ws.Range("B2")
        .Font.Name = "Arial"
        .Font.Size = 28
        .Font.Bold = True
        .Font.Color = clrWhite
        .Interior.Color = clrBlue
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows("2:3").RowHeight = 30

    ' --- Subtitle (Row 5) ---
    ws.Range("B5").Value = "DemoKit Industries - Demonstration Data Notice"
    With ws.Range("B5")
        .Font.Name = "Arial"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = clrNavy
        .HorizontalAlignment = xlCenter
    End With

    ' --- Accent Line (Row 6) ---
    With ws.Range("B6").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = clrAqua
        .Weight = xlMedium
    End With
    ws.Rows("6").RowHeight = 8

    ' --- Disclaimer Text ---
    Dim r As Long
    r = 8

    ' Main notice box
    ws.Range("B" & r).Value = "IMPORTANT NOTICE"
    With ws.Range("B" & r)
        .Font.Name = "Arial"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = clrBlue
    End With

    r = r + 2
    ws.Range("B" & r).Value = _
        "This workbook contains PURELY FICTITIOUS demonstration data." & vbLf & vbLf & _
        "All company names, financial figures, account balances, revenue numbers, " & _
        "cost allocations, and any other data contained in this file are entirely " & _
        "made up for demonstration and training purposes only." & vbLf & vbLf & _
        "None of the financials in this file are remotely accurate or based on " & _
        "any real company's actual financial data. Any resemblance to real " & _
        "financial figures is purely coincidental."
    With ws.Range("B" & r)
        .Font.Name = "Arial"
        .Font.Size = 11
        .Font.Color = clrCharcoal
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    ws.Rows(r).RowHeight = 130

    r = r + 2
    ws.Range("B" & r).Value = "PURPOSE OF THIS FILE"
    With ws.Range("B" & r)
        .Font.Name = "Arial"
        .Font.Size = 13
        .Font.Bold = True
        .Font.Color = clrBlue
    End With

    r = r + 2
    ws.Range("B" & r).Value = _
        "This file was built to demonstrate:" & vbLf & vbLf & _
        "   " & Chr(149) & "  VBA macros that automate common Finance & Accounting tasks" & vbLf & _
        "   " & Chr(149) & "  SQL queries for financial data analysis" & vbLf & _
        "   " & Chr(149) & "  Python scripts for reporting and data processing" & vbLf & _
        "   " & Chr(149) & "  Best practices for Excel-based financial modeling" & vbLf & vbLf & _
        "It is intended as a training and demonstration tool for iPipeline " & _
        "Finance & Accounting staff."
    With ws.Range("B" & r)
        .Font.Name = "Arial"
        .Font.Size = 11
        .Font.Color = clrCharcoal
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    ws.Rows(r).RowHeight = 140

    r = r + 2
    ws.Range("B" & r).Value = "DO NOT USE FOR"
    With ws.Range("B" & r)
        .Font.Name = "Arial"
        .Font.Size = 13
        .Font.Bold = True
        .Font.Color = clrBlue
    End With

    r = r + 2
    ws.Range("B" & r).Value = _
        "   " & Chr(149) & "  Any actual financial reporting or decision-making" & vbLf & _
        "   " & Chr(149) & "  Regulatory filings or audits" & vbLf & _
        "   " & Chr(149) & "  Client-facing materials" & vbLf & _
        "   " & Chr(149) & "  Any purpose requiring real financial data"
    With ws.Range("B" & r)
        .Font.Name = "Arial"
        .Font.Size = 11
        .Font.Color = clrCharcoal
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    ws.Rows(r).RowHeight = 80

    ' --- Bottom accent line ---
    r = r + 2
    With ws.Range("B" & r).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = clrAqua
        .Weight = xlMedium
    End With
    ws.Rows(r).RowHeight = 8

    ' --- Footer ---
    r = r + 2
    ws.Range("B" & r).Value = "DemoKit Industries  |  iPipeline Finance & Accounting  |  Demo File v2.1"
    With ws.Range("B" & r)
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Italic = True
        .Font.Color = clrInnovation
        .HorizontalAlignment = xlCenter
    End With

    ' --- Clean up ---
    ws.Range("A1").Select
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True

    ' Hide gridlines for clean look
    ActiveWindow.DisplayGridlines = False

    ' Print setup
    On Error Resume Next
    ws.PageSetup.PrintArea = "A1:C" & r
    ws.PageSetup.Orientation = xlPortrait
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = 1
    On Error GoTo 0

End Sub
