Attribute VB_Name = "modDemo_Config"
Option Explicit

Public Const DEMO_SHEET_ASSUMPTIONS As String = "Assumptions"
Public Const DEMO_SHEET_GL As String = "CrossfireHiddenWorksheet"
Public Const DEMO_SHEET_CHECKS As String = "Checks"
Public Const DEMO_SHEET_PNL_TREND As String = "P&L - Monthly Trend"
Public Const DEMO_SHEET_REPORT As String = "Report-->"
Public Const DEMO_SHEET_AUDIT As String = "VBA_AuditLog"

Public Const DEMO_MATERIALITY_ABS As Double = 10000
Public Const DEMO_MATERIALITY_PCT As Double = 0.15

Public Function DemoRequiredSheets() As Variant
    DemoRequiredSheets = Array( _
        DEMO_SHEET_ASSUMPTIONS, _
        DEMO_SHEET_GL, _
        DEMO_SHEET_CHECKS, _
        DEMO_SHEET_PNL_TREND, _
        DEMO_SHEET_REPORT)
End Function

Public Function DemoWorkbookReady() As Boolean
    Dim item As Variant

    For Each item In DemoRequiredSheets()
        If Not DemoSheetExists(CStr(item)) Then
            DemoWorkbookReady = False
            Exit Function
        End If
    Next item

    DemoWorkbookReady = True
End Function

Public Function DemoSheetExists(ByVal SheetName As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SheetName)
    On Error GoTo 0

    DemoSheetExists = Not ws Is Nothing
End Function

Public Function DemoGetSheet(ByVal SheetName As String) As Worksheet
    If Not DemoSheetExists(SheetName) Then
        Err.Raise vbObjectError + 801, "DemoGetSheet", "Required sheet missing: " & SheetName
    End If

    Set DemoGetSheet = ThisWorkbook.Worksheets(SheetName)
End Function

Public Sub DemoValidateWorkbookOrStop()
    If Not DemoWorkbookReady() Then
        Err.Raise vbObjectError + 802, "DemoValidateWorkbookOrStop", "Demo workbook is missing one or more required sheets."
    End If
End Sub
