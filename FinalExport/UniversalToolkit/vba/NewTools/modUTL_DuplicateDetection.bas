Attribute VB_Name = "modUTL_DuplicateDetection"
Option Explicit

' ============================================================
' KBT Universal Tools — Duplicate Detection Module
' Works on ANY Excel file — no project-specific setup required
' Tools: 1 | Small effort
' Date: 2026-03-05
' ============================================================
' Tool 09 — Exact Duplicate Finder (Key Column Based)
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
' TOOL 09 — Exact Duplicate Finder                     [SMALL]
' User picks a key column (e.g. "A" for Invoice #). Scans
' the active sheet and highlights all duplicate values in that
' column with yellow fill. Creates a summary report sheet
' listing each duplicate value and how many times it appears.
' ============================================================
Sub ExactDuplicateFinder()
    On Error GoTo ErrHandler

    Dim colInput As String
    colInput = InputBox("Enter the key column letter to check for duplicates:" & vbCrLf & _
                        "Example: A (for Invoice #), B (for Customer ID), etc.", _
                        "Exact Duplicate Finder", "A")
    If colInput = "" Then Exit Sub

    Dim headerRowStr As String
    headerRowStr = InputBox("Which row contains headers?" & vbCrLf & _
                            "(Data starts on the next row)", _
                            "Exact Duplicate Finder", "1")
    If headerRowStr = "" Then Exit Sub
    Dim hRow As Long: hRow = CLng(headerRowStr)

    UTL_TurboOn

    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim keyCol As Long: keyCol = Range(colInput & "1").Column
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row
    Dim dataStart As Long: dataStart = hRow + 1

    If lastRow < dataStart Then
        UTL_TurboOff
        MsgBox "No data found below row " & hRow & " in column " & colInput & ".", vbInformation
        Exit Sub
    End If

    ' Count occurrences using a dictionary
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = dataStart To lastRow
        Dim val As String: val = Trim(CStr(ws.Cells(r, keyCol).Value))
        If Len(val) > 0 Then
            If dict.Exists(val) Then
                dict(val) = dict(val) + 1
            Else
                dict.Add val, 1
            End If
        End If
    Next r

    ' Identify duplicates (count > 1)
    Dim dupCount As Long: dupCount = 0
    Dim dupValues As Object: Set dupValues = CreateObject("Scripting.Dictionary")
    Dim key As Variant
    For Each key In dict.Keys
        If dict(key) > 1 Then
            dupValues.Add key, dict(key)
            dupCount = dupCount + 1
        End If
    Next key

    ' Highlight duplicate rows on the source sheet
    Dim highlightCount As Long: highlightCount = 0
    For r = dataStart To lastRow
        Dim cellVal As String: cellVal = Trim(CStr(ws.Cells(r, keyCol).Value))
        If dupValues.Exists(cellVal) Then
            ws.Cells(r, keyCol).Interior.Color = RGB(255, 255, 150)  ' Yellow highlight
            highlightCount = highlightCount + 1
        End If
    Next r

    ' Create report sheet if duplicates found
    If dupCount > 0 Then
        Dim rptName As String: rptName = "UTL_DuplicateReport"
        Dim wsOld As Worksheet
        On Error Resume Next
        Set wsOld = ActiveWorkbook.Worksheets(rptName)
        On Error GoTo ErrHandler
        If Not wsOld Is Nothing Then
            Application.DisplayAlerts = False
            wsOld.Delete
            Application.DisplayAlerts = True
        End If

        Dim wsRpt As Worksheet
        Set wsRpt = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        wsRpt.Name = rptName

        ' Headers
        wsRpt.Cells(1, 1).Value = "Duplicate Value"
        wsRpt.Cells(1, 2).Value = "Occurrences"
        wsRpt.Cells(1, 3).Value = "Source Column"
        Dim hdrC As Long
        For hdrC = 1 To 3
            wsRpt.Cells(1, hdrC).Font.Bold = True
            wsRpt.Cells(1, hdrC).Interior.Color = RGB(11, 71, 121)
            wsRpt.Cells(1, hdrC).Font.Color = RGB(255, 255, 255)
        Next hdrC

        Dim outRow As Long: outRow = 2
        For Each key In dupValues.Keys
            wsRpt.Cells(outRow, 1).Value = key
            wsRpt.Cells(outRow, 2).Value = dupValues(key)
            wsRpt.Cells(outRow, 3).Value = colInput
            outRow = outRow + 1
        Next key

        wsRpt.Columns("A:C").AutoFit
        ws.Activate  ' Go back to source sheet
    End If

    UTL_TurboOff

    If dupCount = 0 Then
        MsgBox "No duplicates found in column " & colInput & " on '" & ws.Name & "'." & vbCrLf & _
               "All " & dict.Count & " values are unique.", vbInformation, "Exact Duplicate Finder"
    Else
        MsgBox "Duplicates Found!" & vbCrLf & vbCrLf & _
               dupCount & " value(s) appear more than once." & vbCrLf & _
               highlightCount & " cell(s) highlighted yellow on '" & ws.Name & "'." & vbCrLf & vbCrLf & _
               "Full report on '" & rptName & "'.", vbExclamation, "Exact Duplicate Finder"
    End If
    Exit Sub

ErrHandler:
    UTL_TurboOff
    MsgBox "Exact Duplicate Finder error: " & Err.Description, vbCritical
End Sub
