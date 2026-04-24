Attribute VB_Name = "modDemo_CommandCenter"
Option Explicit

Private Const DEMO_CENTER_SHEET As String = "Demo_CommandCenter"

Public Sub BuildDemoCommandCenter()
    Dim ws As Worksheet

    On Error GoTo CenterFail

    Set ws = DemoGetOrCreateCenterSheet()
    ws.Cells.Clear

    BuildDemoHeader ws
    AddDemoButton ws, "Run Reconciliation", "RunDemoReconciliation", 6
    AddDemoButton ws, "Generate Variance Narrative", "GenerateDemoVarianceNarrative", 10
    AddDemoButton ws, "Build Executive Brief Pack", "BuildDemoExecutiveBriefPack", 14
    AddDemoButton ws, "Run Scenario Comparison", "RunDemoWhatIfScenarios", 18

    ws.Range("B23").Value = "Status"
    ws.Range("C23").Value = "Ready"
    ws.Range("B23:C23").Font.Bold = True

    ws.Columns("B:F").ColumnWidth = 34

    DemoLog "BuildDemoCommandCenter", "PASS", "Demo command center rebuilt"
    UTL_ShowCompletion "Demo Command Center", "Demo command center is ready on 'Demo_CommandCenter'."
    Exit Sub

CenterFail:
    DemoLog "BuildDemoCommandCenter", "FAIL", Err.Description
    MsgBox "Demo command center failed: " & Err.Description, vbExclamation, "Demo Command Center"
End Sub

Private Function DemoGetOrCreateCenterSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DEMO_CENTER_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = DEMO_CENTER_SHEET
    End If

    Set DemoGetOrCreateCenterSheet = ws
End Function

Private Sub BuildDemoHeader(ByVal ws As Worksheet)
    ws.Range("B2:F2").Merge
    ws.Range("B2").Value = "iPipeline"
    ws.Range("B2").Font.Name = "Arial"
    ws.Range("B2").Font.Size = 20
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Color = RGB(11, 71, 121)

    ws.Range("B3:F3").Merge
    ws.Range("B3").Value = "Finance & Accounting"
    ws.Range("B3").Font.Name = "Arial"
    ws.Range("B3").Font.Size = 10
    ws.Range("B3").Font.Color = RGB(17, 46, 81)

    ws.Range("B4:F4").Merge
    ws.Range("B4").Value = "Demo Command Center"
    ws.Range("B4").Font.Name = "Arial"
    ws.Range("B4").Font.Size = 14
    ws.Range("B4").Font.Bold = True
    ws.Range("B4").Interior.Color = RGB(11, 71, 121)
    ws.Range("B4").Font.Color = RGB(249, 249, 249)
End Sub

Private Sub AddDemoButton(ByVal ws As Worksheet, ByVal buttonText As String, ByVal macroName As String, ByVal topRow As Long)
    Dim shp As Shape
    Dim nameTag As String

    nameTag = "btnDemo_" & Replace(macroName, " ", "_")

    On Error Resume Next
    ws.Shapes(nameTag).Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("B" & topRow).Left, ws.Range("B" & topRow).Top, 360, 34)
    shp.Name = nameTag
    shp.Fill.ForeColor.RGB = RGB(75, 155, 203)
    shp.Line.ForeColor.RGB = RGB(17, 46, 81)
    shp.TextFrame2.TextRange.Text = buttonText
    shp.TextFrame2.TextRange.Font.Name = "Arial"
    shp.TextFrame2.TextRange.Font.Size = 11
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(22, 22, 22)
    shp.OnAction = macroName
End Sub
