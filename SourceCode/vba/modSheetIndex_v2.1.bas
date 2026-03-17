Attribute VB_Name = "modSheetIndex"
Option Explicit

'===============================================================================
' modSheetIndex - Home Sheet & Sheet Index Builder
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Creates a "Home" sheet with a one-click Command Center launch
'           button, and a "Sheet Index" sheet listing every tab with
'           clickable hyperlinks.
'
' PUBLIC SUBS:
'   CreateHomeSheet         - Builds the Home sheet with Command Center button
'   ListAllSheetsWithLinks  - Lists all sheets in column A with links in B
'
' VERSION:  2.1.0 (New module - 2026-03-04)
'===============================================================================

'===============================================================================
' CreateHomeSheet - Build a Home sheet with Command Center launch button
' Places the sheet at position 1 so it's the first tab the user sees.
' Adds a styled button that calls LaunchCommandCenter from modFormBuilder.
'===============================================================================
Public Sub CreateHomeSheet()
    On Error GoTo ErrHandler

    Dim wsName As String
    wsName = "Home"

    ' If Home already exists, just activate it
    If modConfig.SheetExists(wsName) Then
        ThisWorkbook.Worksheets(wsName).Activate
        MsgBox "Home sheet already exists. Activated.", vbInformation, APP_NAME
        Exit Sub
    End If

    modPerformance.TurboOn

    ' Create the Home sheet at position 1
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    ws.Name = wsName
    ws.Tab.Color = CLR_NAVY

    ' Title
    With ws.Range("A1")
        .Value = APP_NAME
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = CLR_NAVY
        .Font.Name = "Arial"
    End With

    ' Subtitle
    With ws.Range("A2")
        .Value = "Command Center Home"
        .Font.Size = 14
        .Font.Italic = True
        .Font.Color = RGB(75, 155, 203)
        .Font.Name = "Arial"
    End With

    ' Version info
    ws.Range("A3").Value = "Version " & APP_VERSION & " | " & Format(Now, "MMMM D, YYYY")
    ws.Range("A3").Font.Size = 10
    ws.Range("A3").Font.Color = RGB(128, 128, 128)

    ' --- Command Center launch button ---
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, 30, 90, 280, 60)
    With btn
        .Name = "btnLaunchCommandCenter"
        .Fill.ForeColor.RGB = CLR_NAVY
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = "Open Command Center"
        .TextFrame2.TextRange.Font.Size = 16
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = CLR_WHITE
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .OnAction = "LaunchCommandCenter"
    End With

    ' --- List All Sheets button ---
    Dim btn2 As Shape
    Set btn2 = ws.Shapes.AddShape(msoShapeRoundedRectangle, 30, 170, 280, 50)
    With btn2
        .Name = "btnListAllSheets"
        .Fill.ForeColor.RGB = RGB(75, 155, 203)
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = "View Sheet Index"
        .TextFrame2.TextRange.Font.Size = 14
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = CLR_WHITE
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .OnAction = "ListAllSheetsWithLinks"
    End With

    ' Instructions
    ws.Range("A16").Value = "Click 'Open Command Center' to access all 62 actions."
    ws.Range("A16").Font.Size = 11
    ws.Range("A16").Font.Italic = True
    ws.Range("A17").Value = "Click 'View Sheet Index' to see a clickable list of every tab."
    ws.Range("A17").Font.Size = 11
    ws.Range("A17").Font.Italic = True

    ws.Columns("A:A").ColumnWidth = 55
    ws.Activate

    modPerformance.TurboOff
    modLogger.LogAction "modSheetIndex", "CreateHomeSheet", _
        "Home sheet created with Command Center button"

    MsgBox "Home sheet created!" & vbCrLf & vbCrLf & _
           "Click the blue button to open the Command Center." & vbCrLf & _
           "Click 'View Sheet Index' to see a clickable list of every tab.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    MsgBox "CreateHomeSheet error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' ListAllSheetsWithLinks - Build/update a Sheet Index with hyperlinks
' Creates a "Sheet Index" sheet listing every tab in the workbook.
' Column A = Sheet name, Column B = Clickable hyperlink, Column C = Visibility.
' If run again, only adds NEW sheets that aren't already listed (no duplicates).
'===============================================================================
Public Sub ListAllSheetsWithLinks()
    On Error GoTo ErrHandler
    modPerformance.TurboOn

    Dim indexName As String
    indexName = "Sheet Index"

    Dim ws As Worksheet

    ' Create or reuse the Sheet Index tab
    If modConfig.SheetExists(indexName) Then
        Set ws = ThisWorkbook.Worksheets(indexName)
    Else
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = indexName
        ws.Tab.Color = RGB(75, 155, 203)
    End If

    ' Build a dictionary of sheets already in the index to avoid duplicates
    Dim existingLinks As Object
    Set existingLinks = CreateObject("Scripting.Dictionary")

    Dim lastExisting As Long
    lastExisting = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastExisting >= 3 Then
        Dim chk As Long
        For chk = 3 To lastExisting
            Dim existName As String
            existName = Trim(CStr(ws.Cells(chk, 1).Value))
            If Len(existName) > 0 Then
                existingLinks(existName) = True
            End If
        Next chk
    End If

    ' Write title and headers if sheet is new/empty
    If Trim(CStr(ws.Cells(1, 1).Value)) = "" Then
        ws.Range("A1").Value = "Sheet Index - " & ThisWorkbook.Name
        ws.Range("A1").Font.Size = 14
        ws.Range("A1").Font.Bold = True
        ws.Range("A1").Font.Color = CLR_NAVY

        modConfig.StyleHeader ws, 2, Array("Sheet Name", "Navigate", "Status")
    End If

    ' Find the next available row (below existing entries)
    Dim outRow As Long
    outRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If outRow < 3 Then outRow = 3

    Dim addedCount As Long
    Dim skippedCount As Long
    Dim sheetWS As Worksheet

    For Each sheetWS In ThisWorkbook.Worksheets
        ' Skip the index sheet itself
        If sheetWS.Name = indexName Then GoTo NextSheet

        ' Skip if already in the index
        If existingLinks.exists(sheetWS.Name) Then
            skippedCount = skippedCount + 1
            GoTo NextSheet
        End If

        ' Write sheet name
        ws.Cells(outRow, 1).Value = sheetWS.Name

        ' Add clickable hyperlink
        ws.Hyperlinks.Add _
            Anchor:=ws.Cells(outRow, 2), _
            Address:="", _
            SubAddress:="'" & Replace(sheetWS.Name, "'", "''") & "'!A1", _
            TextToDisplay:="Go to Sheet"
        ws.Cells(outRow, 2).Font.Color = RGB(31, 78, 121)

        ' Visibility status
        Select Case sheetWS.Visible
            Case xlSheetVisible
                ws.Cells(outRow, 3).Value = "Visible"
            Case xlSheetHidden
                ws.Cells(outRow, 3).Value = "Hidden"
                ws.Cells(outRow, 3).Font.Color = RGB(192, 0, 0)
            Case xlSheetVeryHidden
                ws.Cells(outRow, 3).Value = "Very Hidden"
                ws.Cells(outRow, 3).Font.Color = RGB(192, 0, 0)
        End Select

        ' Alternating row shading
        If outRow Mod 2 = 1 Then
            ws.Range(ws.Cells(outRow, 1), ws.Cells(outRow, 3)).Interior.Color = CLR_ALT_ROW
        End If

        addedCount = addedCount + 1
        outRow = outRow + 1
NextSheet:
    Next sheetWS

    ws.Columns("A:C").AutoFit
    ws.Activate

    modPerformance.TurboOff
    modLogger.LogAction "modSheetIndex", "ListAllSheetsWithLinks", _
        addedCount & " added, " & skippedCount & " already in index"

    MsgBox "Sheet index updated!" & vbCrLf & vbCrLf & _
           addedCount & " new sheet(s) added." & vbCrLf & _
           skippedCount & " sheet(s) already listed (skipped)." & vbCrLf & _
           "Click links in column B to navigate.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    MsgBox "ListAllSheetsWithLinks error: " & Err.Description, vbCritical, APP_NAME
End Sub
