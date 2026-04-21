Attribute VB_Name = "modFolderOrganizer"
'===============================================================================
' modFolderOrganizer
' PURPOSE: Scan a folder tree, read every file's metadata (name, size, modified,
'          type, owner, path) into Excel, then apply rename / move / archive
'          rules you write in Excel formulas. Executes the rules in one click.
'
' WHY THIS IS NOT NATIVE: OneDrive/Windows cannot rename or relocate 10,000
'          files based on a spreadsheet-driven rule table. File Explorer has
'          no batch rename engine with business logic. PowerShell can, but
'          Finance & Accounting usually doesn't know PowerShell - they know Excel.
'
' USE CASE (software business):
'   - 12 years of quarterly close folders, 40,000 files, inconsistent naming.
'     Scan them, write rename rules using Excel logic (e.g. "=B2 & '_' &
'     TEXT(D2,'yyyy-mm')"), and apply.
'   - Vendor onboarding packet arrives as a zip of 50 PDFs. Batch-rename them
'     to Vendor_<Name>_<DocType>_<Date>.pdf by typing 50 new names in column G.
'
' SHEETS:
'   "FolderScan"     - filled by ScanFolderToSheet
'   "FolderRules"    - you fill: column G = new name, column H = new folder
'===============================================================================
Option Explicit

'-------------------------------------------------------------------------------
' ScanFolderToSheet - Recursive file inventory.
'-------------------------------------------------------------------------------
Public Sub ScanFolderToSheet()
    Dim fd As FileDialog, root As String
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Choose folder to scan"
    If fd.Show <> -1 Then Exit Sub
    root = fd.SelectedItems(1)
    If Right(root, 1) <> "\" Then root = root & "\"

    Dim ws As Worksheet
    Set ws = EnsureSheet("FolderScan")
    ws.Cells.Clear
    ws.Range("A1:G1").Value = Array("Path", "FileName", "Ext", "SizeKB", _
                                     "ModifiedAt", "RenameTo", "MoveTo")
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.Color = RGB(11, 71, 121)
    ws.Rows(1).Font.Color = vbWhite

    Dim row As Long: row = 2
    WalkFolder root, ws, row
    ws.Columns("A:E").AutoFit
    MsgBox "Scanned " & (row - 2) & " files.", vbInformation
End Sub

Private Sub WalkFolder(ByVal folder As String, ws As Worksheet, ByRef row As Long)
    Dim fso As Object, f As Object, sf As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    For Each f In fso.GetFolder(folder).Files
        ws.Cells(row, 1).Value = f.ParentFolder.path
        ws.Cells(row, 2).Value = f.Name
        ws.Cells(row, 3).Value = fso.GetExtensionName(f.Name)
        ws.Cells(row, 4).Value = Round(f.Size / 1024, 2)
        ws.Cells(row, 5).Value = f.DateLastModified
        row = row + 1
        If row Mod 250 = 0 Then
            Application.StatusBar = "Scanned " & (row - 2) & " files..."
            DoEvents
        End If
    Next f
    For Each sf In fso.GetFolder(folder).Subfolders
        WalkFolder sf.path, ws, row
    Next sf
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' ApplyRenameAndMove - Reads column F (RenameTo) and G (MoveTo). Non-blank rows
'                      are executed. Writes outcome to column H.
'-------------------------------------------------------------------------------
Public Sub ApplyRenameAndMove()
    Dim ws As Worksheet, r As Long, lastRow As Long
    Dim origPath As String, newName As String, newFolder As String
    Dim renamed As Long, moved As Long, failed As Long
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Set ws = ThisWorkbook.Worksheets("FolderScan")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    If MsgBox("About to apply rename/move to " & (lastRow - 1) & " rows." & _
             vbCrLf & "Recommend closing any open copies first. Continue?", _
             vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    ws.Cells(1, 8).Value = "Result"
    ws.Cells(1, 8).Font.Bold = True

    For r = 2 To lastRow
        origPath = ws.Cells(r, "A").Value & "\" & ws.Cells(r, "B").Value
        newName = CStr(ws.Cells(r, "F").Value)
        newFolder = CStr(ws.Cells(r, "G").Value)

        If Len(newName) = 0 And Len(newFolder) = 0 Then
            ws.Cells(r, 8).Value = ""
            GoTo NextRow
        End If

        On Error Resume Next
        If Len(newName) > 0 Then
            fso.GetFile(origPath).Name = newName
            If Err.Number = 0 Then
                renamed = renamed + 1
                ws.Cells(r, "B").Value = newName
                origPath = ws.Cells(r, "A").Value & "\" & newName
            Else
                failed = failed + 1
                ws.Cells(r, 8).Value = "Rename failed: " & Err.Description
                Err.Clear
                GoTo NextRow
            End If
        End If
        If Len(newFolder) > 0 Then
            If Not fso.FolderExists(newFolder) Then fso.CreateFolder newFolder
            fso.MoveFile origPath, newFolder & "\" & ws.Cells(r, "B").Value
            If Err.Number = 0 Then
                moved = moved + 1
                ws.Cells(r, "A").Value = newFolder
            Else
                failed = failed + 1
                ws.Cells(r, 8).Value = "Move failed: " & Err.Description
                Err.Clear
                GoTo NextRow
            End If
        End If
        ws.Cells(r, 8).Value = "OK"
        ws.Cells(r, 8).Interior.Color = RGB(220, 245, 220)
        On Error GoTo 0
NextRow:
    Next r

    MsgBox "Done." & vbCrLf & "Renamed: " & renamed & vbCrLf & _
           "Moved: " & moved & vbCrLf & "Failed: " & failed, vbInformation
End Sub

Private Function EnsureSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        EnsureSheet.Name = name
    End If
End Function
