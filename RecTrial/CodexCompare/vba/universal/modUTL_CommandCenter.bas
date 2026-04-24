Attribute VB_Name = "modUTL_CommandCenter"
Option Explicit

Private Const COMMAND_CENTER_SHEET As String = "UTL_CommandCenter"

Public Sub BuildCommandCenter()
    Dim ws As Worksheet

    Application.ScreenUpdating = False

    Set ws = GetOrCreateCommandCenterSheet()
    ws.Cells.Clear

    ApplyBrandHeader ws
    AddActionButton ws, "Run Full Workbook Sanitizer", "Run_CommandCenter_Sanitize", 6
    AddActionButton ws, "Preview Sanitizer Impact", "Run_CommandCenter_Preview", 10
    AddActionButton ws, "Create Workbook Profile", "Run_CommandCenter_Profile", 14
    AddActionButton ws, "Consolidate Visible Sheets", "Run_CommandCenter_Consolidate", 18
    AddActionButton ws, "Classify Materiality", "Run_CommandCenter_Materiality", 22
    AddActionButton ws, "Generate Exception Narratives", "Run_CommandCenter_Narratives", 26
    AddActionButton ws, "Build Executive One-Pager", "Run_CommandCenter_OnePager", 30

    ws.Range("B35").Value = "Status"
    ws.Range("C35").Value = "Ready"
    ws.Range("B35:C35").Font.Bold = True

    ws.Columns("B:F").ColumnWidth = 34

    Application.ScreenUpdating = True

    UTL_LogAction "modUTL_CommandCenter", "BuildCommandCenter", "PASS", "Command Center rebuilt"
    UTL_ShowCompletion "Universal Command Center", "Command Center is ready on sheet 'UTL_CommandCenter'."
End Sub

Public Sub Run_CommandCenter_Sanitize()
    UpdateStatus "Running full sanitizer..."
    RunFullSanitize False
    UpdateStatus "Sanitizer finished"
End Sub

Public Sub Run_CommandCenter_Preview()
    UpdateStatus "Previewing sanitizer impact..."
    PreviewSanitizeChanges False
    UpdateStatus "Preview finished"
End Sub

Public Sub Run_CommandCenter_Profile()
    Dim ws As Worksheet
    Dim reportRow As Long
    Dim targets As Collection
    Dim item As Variant

    Set ws = GetOrCreateCommandCenterSheet()
    Set targets = UTL_GetTargetSheets(False)

    reportRow = 39
    ws.Range("B39:F1000").ClearContents
    ws.Range("B38:F38").Value = Array("Sheet", "Header Row", "Rows", "Columns", "Data Range")
    ws.Range("B38:F38").Font.Bold = True

    For Each item In targets
        If TypeName(item) = "Worksheet" Then
            ws.Cells(reportRow, 2).Value = item.Name
            ws.Cells(reportRow, 3).Value = UTL_DetectHeaderRow(item)
            ws.Cells(reportRow, 4).Value = UTL_LastUsedRow(item)
            ws.Cells(reportRow, 5).Value = UTL_LastUsedColumn(item)
            ws.Cells(reportRow, 6).Value = UTL_DetectDataRange(item).Address(False, False)
            reportRow = reportRow + 1
        End If
    Next item

    UpdateStatus "Workbook profile refreshed"
    UTL_LogAction "modUTL_CommandCenter", "Run_CommandCenter_Profile", "PASS", "Profile created", reportRow - 39, 0
    UTL_ShowCompletion "Workbook Profile", "Profile rows written: " & (reportRow - 39)
End Sub

Public Sub Run_CommandCenter_Consolidate()
    UpdateStatus "Consolidating visible sheets..."
    ConsolidateVisibleSheetsByHeader
    UpdateStatus "Consolidation finished"
End Sub

Public Sub Run_CommandCenter_Materiality()
    UpdateStatus "Classifying materiality..."
    MaterialityClassifierActiveSheet
    UpdateStatus "Materiality classification finished"
End Sub

Public Sub Run_CommandCenter_Narratives()
    UpdateStatus "Generating narratives..."
    GenerateExceptionNarrativesActiveSheet
    UpdateStatus "Narratives finished"
End Sub

Public Sub Run_CommandCenter_OnePager()
    UpdateStatus "Building executive one-pager..."
    BuildExecutiveOnePagerFromActiveSheet
    UpdateStatus "One-pager finished"
End Sub

Private Function GetOrCreateCommandCenterSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(COMMAND_CENTER_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = COMMAND_CENTER_SHEET
    End If

    Set GetOrCreateCommandCenterSheet = ws
End Function

Private Sub ApplyBrandHeader(ByVal ws As Worksheet)
    With ws.Range("B2:F2")
        .Merge
        .Value = "iPipeline"
        .Font.Name = "Arial"
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = RGB(11, 71, 121)
        .HorizontalAlignment = xlLeft
    End With

    With ws.Range("B3:F3")
        .Merge
        .Value = "Finance & Accounting"
        .Font.Name = "Arial"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(17, 46, 81)
        .HorizontalAlignment = xlLeft
    End With

    With ws.Range("B4:F4")
        .Merge
        .Value = "Universal Toolkit Command Center"
        .Font.Name = "Arial"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = RGB(11, 71, 121)
        .Font.Color = RGB(249, 249, 249)
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub AddActionButton(ByVal ws As Worksheet, ByVal LabelText As String, ByVal MacroName As String, ByVal TopRow As Long)
    Dim btn As Shape
    Dim buttonName As String

    buttonName = "btn_" & Replace(MacroName, " ", "_")

    On Error Resume Next
    ws.Shapes(buttonName).Delete
    On Error GoTo 0

    Set btn = ws.Shapes.AddShape(Type:=msoShapeRoundedRectangle, _
                                 Left:=ws.Range("B" & TopRow).Left, _
                                 Top:=ws.Range("B" & TopRow).Top, _
                                 Width:=360, Height:=34)

    btn.Name = buttonName
    btn.Fill.ForeColor.RGB = RGB(75, 155, 203)
    btn.Line.ForeColor.RGB = RGB(17, 46, 81)
    btn.TextFrame2.TextRange.Characters.Text = LabelText
    btn.TextFrame2.TextRange.Font.Name = "Arial"
    btn.TextFrame2.TextRange.Font.Size = 11
    btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(22, 22, 22)
    btn.OnAction = MacroName
End Sub

Private Sub UpdateStatus(ByVal StatusText As String)
    Dim ws As Worksheet

    Set ws = GetOrCreateCommandCenterSheet()
    ws.Range("C35").Value = StatusText
End Sub
