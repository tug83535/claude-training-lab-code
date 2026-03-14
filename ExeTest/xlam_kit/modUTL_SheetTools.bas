Attribute VB_Name = "modUTL_SheetTools"
Option Explicit

' ============================================================
' KBT Universal Tools — Sheet Tools Module
' Works on ANY Excel file — no project-specific setup required
' Install in Personal.xlsb to use across all Excel sessions
' Tools: 4 | Tier 1: 4
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
' TOOL 1 — List All Sheets with Hyperlinks         [TIER 1]
' Creates a "Sheet Index" tab listing every worksheet
' Column A = Sheet Name, Column B = Clickable Link, Column C = Status
' Safe to re-run: only adds sheets not already listed
' ============================================================
Sub ListAllSheetsWithLinks()
    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim indexName As String
    indexName = "UTL_SheetIndex"

    Dim ws As Worksheet

    ' Create or reuse the index sheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets(indexName)
    On Error GoTo ErrHandler

    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Sheets.Add( _
            After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
        ws.Name = indexName
    End If

    ' Build dictionary of sheets already listed
    Dim existing As Object
    Set existing = CreateObject("Scripting.Dictionary")

    Dim lastExisting As Long
    lastExisting = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastExisting >= 4 Then
        Dim chk As Long
        For chk = 4 To lastExisting
            Dim existName As String
            existName = Trim(CStr(ws.Cells(chk, 1).Value))
            If Len(existName) > 0 Then existing(existName) = True
        Next chk
    End If

    ' Write title and headers if sheet is new
    If Trim(CStr(ws.Cells(1, 1).Value)) = "" Then
        ws.Range("A1").Value = "Sheet Index — " & ActiveWorkbook.Name
        ws.Range("A1").Font.Bold = True
        ws.Range("A1").Font.Size = 14
        ws.Range("A2").Value = "Generated: " & Format(Now, "MM/DD/YYYY h:mm AM/PM")
        ws.Range("A2").Font.Italic = True

        ws.Range("A3").Value = "Sheet Name"
        ws.Range("B3").Value = "Navigate"
        ws.Range("C3").Value = "Status"
        ws.Range("A3:C3").Font.Bold = True
        ws.Range("A3:C3").Interior.Color = RGB(11, 71, 121)
        ws.Range("A3:C3").Font.Color = RGB(255, 255, 255)
    End If

    ' Find next available row
    Dim outRow As Long
    outRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If outRow < 4 Then outRow = 4

    Dim addedCount As Long
    Dim skippedCount As Long
    Dim sheetWS As Worksheet

    For Each sheetWS In ActiveWorkbook.Worksheets
        If sheetWS.Name = indexName Then GoTo NextSheet

        If existing.exists(sheetWS.Name) Then
            skippedCount = skippedCount + 1
            GoTo NextSheet
        End If

        ws.Cells(outRow, 1).Value = sheetWS.Name

        ws.Hyperlinks.Add _
            Anchor:=ws.Cells(outRow, 2), _
            Address:="", _
            SubAddress:="'" & sheetWS.Name & "'!A1", _
            TextToDisplay:="Go to Sheet"

        Select Case sheetWS.Visible
            Case xlSheetVisible:    ws.Cells(outRow, 3).Value = "Visible"
            Case xlSheetHidden:     ws.Cells(outRow, 3).Value = "Hidden"
            Case xlSheetVeryHidden: ws.Cells(outRow, 3).Value = "Very Hidden"
        End Select

        ' Alternating rows
        If outRow Mod 2 = 0 Then
            ws.Range(ws.Cells(outRow, 1), ws.Cells(outRow, 3)).Interior.Color = RGB(237, 242, 249)
        End If

        addedCount = addedCount + 1
        outRow = outRow + 1
NextSheet:
    Next sheetWS

    ws.Columns("A:C").AutoFit
    ws.Activate
    UTL_TurboOff

    MsgBox "Sheet index updated!" & Chr(10) & Chr(10) & _
           addedCount & " new sheet(s) added." & Chr(10) & _
           skippedCount & " sheet(s) already listed (skipped)." & Chr(10) & _
           "Click links in column B to navigate.", _
           vbInformation, "UTL Sheet Tools"
    Exit Sub

ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Sheet Tools"
End Sub

' ============================================================
' TOOL 2 — Template Cloner                         [TIER 1]
' Pick any sheet, type how many copies, and get instant clones.
' Names each clone "SheetName (1)", "SheetName (2)", etc.
' Handles name conflicts automatically.
' ============================================================
Sub TemplateCloner()
    ' Pick the sheet to clone
    Dim sheetList As String
    Dim ws As Worksheet
    Dim sheetIdx As Long
    sheetIdx = 1

    For Each ws In ActiveWorkbook.Worksheets
        sheetList = sheetList & sheetIdx & ". " & ws.Name & Chr(10)
        sheetIdx = sheetIdx + 1
    Next ws

    Dim choice As String
    choice = InputBox("Which sheet do you want to clone?" & Chr(10) & Chr(10) & _
                      sheetList & Chr(10) & _
                      "Enter the number:", "UTL — Template Cloner")
    If choice = "" Then Exit Sub
    If Not IsNumeric(choice) Then
        MsgBox "Please enter a number.", vbExclamation, "UTL Sheet Tools"
        Exit Sub
    End If

    Dim sheetNum As Long
    sheetNum = CLng(choice)
    If sheetNum < 1 Or sheetNum > ActiveWorkbook.Worksheets.Count Then
        MsgBox "Invalid sheet number. Choose between 1 and " & _
               ActiveWorkbook.Worksheets.Count & ".", vbExclamation, "UTL Sheet Tools"
        Exit Sub
    End If

    Dim sourceWS As Worksheet
    Set sourceWS = ActiveWorkbook.Worksheets(sheetNum)

    ' How many clones?
    Dim countStr As String
    countStr = InputBox("How many copies of '" & sourceWS.Name & "'?", _
                        "UTL — Template Cloner", "3")
    If countStr = "" Then Exit Sub
    If Not IsNumeric(countStr) Then
        MsgBox "Please enter a number.", vbExclamation, "UTL Sheet Tools"
        Exit Sub
    End If

    Dim copyCount As Long
    copyCount = CLng(countStr)
    If copyCount < 1 Or copyCount > 50 Then
        MsgBox "Please enter a number between 1 and 50.", vbExclamation, "UTL Sheet Tools"
        Exit Sub
    End If

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim baseName As String
    baseName = sourceWS.Name

    Dim i As Long
    Dim createdCount As Long
    For i = 1 To copyCount
        sourceWS.Copy After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)

        ' Name the new sheet safely
        Dim newWS As Worksheet
        Set newWS = ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)

        Dim newName As String
        newName = baseName & " (" & i & ")"

        ' Handle name conflicts (truncate if > 31 chars)
        If Len(newName) > 31 Then
            newName = Left(baseName, 31 - Len(" (" & i & ")")) & " (" & i & ")"
        End If

        ' Try to rename; if conflict, add a suffix
        On Error Resume Next
        newWS.Name = newName
        If Err.Number <> 0 Then
            Err.Clear
            newWS.Name = Left(newName, 27) & "_" & i
        End If
        On Error GoTo ErrHandler

        createdCount = createdCount + 1
    Next i

    sourceWS.Activate
    UTL_TurboOff

    MsgBox "Done! " & createdCount & " clone(s) of '" & baseName & "' created." & Chr(10) & _
           "New sheets are at the end of the workbook.", _
           vbInformation, "UTL Sheet Tools"
    Exit Sub

ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Sheet Tools"
End Sub

' ============================================================
' TOOL 3 — Unique Customer ID Generator            [TIER 1]
' Assigns unique sequential IDs to customers.
' Scans the ID column to find the highest existing ID, then
' fills blank ID cells with the next available number.
' Format: CUST-00001, CUST-00002, etc. (or custom prefix)
' Never duplicates — always scans for the max first.
' ============================================================
Sub GenerateUniqueCustomerIDs()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Ask which column has the IDs
    Dim idColStr As String
    idColStr = InputBox("Which column should have the Customer IDs? (e.g. A)" & Chr(10) & Chr(10) & _
                        "The tool will scan this column for existing IDs," & Chr(10) & _
                        "then fill in new IDs for any blank cells." & Chr(10) & _
                        "Row 1 is assumed to be a header row.", _
                        "UTL — Customer ID Generator", "A")
    If idColStr = "" Then Exit Sub

    ' Convert column letter to number
    Dim idCol As Long
    On Error Resume Next
    idCol = ws.Range(idColStr & "1").Column
    On Error GoTo ErrHandler
    If idCol = 0 Then
        MsgBox "Invalid column letter.", vbExclamation, "UTL Sheet Tools"
        Exit Sub
    End If

    ' Ask for prefix
    Dim prefix As String
    prefix = InputBox("Enter the ID prefix:" & Chr(10) & _
                      "(e.g. CUST-, CLI-, ID-)" & Chr(10) & Chr(10) & _
                      "IDs will look like: CUST-00001, CUST-00002, etc.", _
                      "UTL — Customer ID Generator", "CUST-")
    If prefix = "" Then prefix = "CUST-"

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim lastRow As Long
    ' Find last row across all columns (not just ID column)
    Dim testCol As Long
    lastRow = 1
    For testCol = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Dim colLast As Long
        colLast = ws.Cells(ws.Rows.Count, testCol).End(xlUp).Row
        If colLast > lastRow Then lastRow = colLast
    Next testCol

    If lastRow < 2 Then
        UTL_TurboOff
        MsgBox "No data found below the header row.", vbInformation, "UTL Sheet Tools"
        Exit Sub
    End If

    ' Scan existing IDs to find the max number
    Dim maxNum As Long
    maxNum = 0
    Dim r As Long
    For r = 2 To lastRow
        Dim cellVal As String
        cellVal = Trim(CStr(ws.Cells(r, idCol).Value))
        If Len(cellVal) > 0 And Left(cellVal, Len(prefix)) = prefix Then
            Dim numPart As String
            numPart = Mid(cellVal, Len(prefix) + 1)
            If IsNumeric(numPart) Then
                Dim thisNum As Long
                thisNum = CLng(numPart)
                If thisNum > maxNum Then maxNum = thisNum
            End If
        End If
    Next r

    ' Fill blank ID cells with next sequential IDs
    Dim nextNum As Long
    nextNum = maxNum + 1
    Dim filledCount As Long

    For r = 2 To lastRow
        Dim existingVal As String
        existingVal = Trim(CStr(ws.Cells(r, idCol).Value))

        ' Only fill if the ID cell is blank AND the row has data
        If Len(existingVal) = 0 Then
            ' Check if the row has any other data (not a completely empty row)
            Dim hasData As Boolean
            hasData = False
            Dim checkCol As Long
            For checkCol = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
                If checkCol <> idCol Then
                    If Trim(CStr(ws.Cells(r, checkCol).Value)) <> "" Then
                        hasData = True
                        Exit For
                    End If
                End If
            Next checkCol

            If hasData Then
                ws.Cells(r, idCol).Value = prefix & Format(nextNum, "00000")
                ws.Cells(r, idCol).NumberFormat = "@"  ' Set as text format
                nextNum = nextNum + 1
                filledCount = filledCount + 1
            End If
        End If
    Next r

    UTL_TurboOff

    If filledCount = 0 Then
        MsgBox "No blank ID cells found. All rows already have IDs." & Chr(10) & _
               "Highest existing ID: " & prefix & Format(maxNum, "00000"), _
               vbInformation, "UTL Sheet Tools"
    Else
        MsgBox "Done! " & filledCount & " new ID(s) assigned." & Chr(10) & Chr(10) & _
               "Range: " & prefix & Format(maxNum + 1, "00000") & _
               " through " & prefix & Format(nextNum - 1, "00000") & Chr(10) & _
               "Existing IDs were NOT changed." & Chr(10) & _
               "No duplicates — guaranteed unique.", _
               vbInformation, "UTL Sheet Tools"
    End If
    Exit Sub

ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Sheet Tools"
End Sub

' ============================================================
' TOOL 4 — Create Folders from Selection              [TIER 1]
' Highlight a column of cell values (names, projects, etc.)
' and this tool creates a Windows folder for each value.
' Asks you where to create the folders first.
' Skips blanks, duplicates, and illegal filename characters.
' ============================================================
Sub CreateFoldersFromSelection()
    ' Validate selection
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells first." & Chr(10) & Chr(10) & _
               "How to use:" & Chr(10) & _
               "1. Highlight a column of names/values" & Chr(10) & _
               "2. Run this tool" & Chr(10) & _
               "3. Pick where to create the folders", _
               vbExclamation, "UTL — Create Folders"
        Exit Sub
    End If

    Dim sel As Range
    Set sel = Selection

    ' Collect unique, non-blank folder names from selection
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim cell As Range
    Dim rawName As String
    Dim cleanName As String
    Dim skippedBlanks As Long
    Dim skippedDupes As Long
    Dim skippedBadChars As Long

    For Each cell In sel.Cells
        rawName = Trim(CStr(cell.Value))

        ' Skip blanks
        If Len(rawName) = 0 Then
            skippedBlanks = skippedBlanks + 1
            GoTo NextCell
        End If

        ' Clean illegal Windows folder characters: \ / : * ? " < > |
        cleanName = rawName
        cleanName = Replace(cleanName, "\", "-")
        cleanName = Replace(cleanName, "/", "-")
        cleanName = Replace(cleanName, ":", "-")
        cleanName = Replace(cleanName, "*", "-")
        cleanName = Replace(cleanName, "?", "")
        cleanName = Replace(cleanName, """", "")
        cleanName = Replace(cleanName, "<", "")
        cleanName = Replace(cleanName, ">", "")
        cleanName = Replace(cleanName, "|", "-")

        ' Trim trailing dots and spaces (Windows doesn't allow them at end of folder names)
        Do While Len(cleanName) > 0 And (Right(cleanName, 1) = "." Or Right(cleanName, 1) = " ")
            cleanName = Left(cleanName, Len(cleanName) - 1)
        Loop

        ' Skip if cleaning emptied the name
        If Len(cleanName) = 0 Then
            skippedBadChars = skippedBadChars + 1
            GoTo NextCell
        End If

        ' Truncate to 255 chars (Windows folder name limit)
        If Len(cleanName) > 255 Then cleanName = Left(cleanName, 255)

        ' Skip duplicates
        If dict.exists(cleanName) Then
            skippedDupes = skippedDupes + 1
            GoTo NextCell
        End If

        dict(cleanName) = rawName  ' Store clean -> original mapping
NextCell:
    Next cell

    If dict.Count = 0 Then
        MsgBox "No valid folder names found in the selected cells." & Chr(10) & Chr(10) & _
               "Blanks skipped: " & skippedBlanks & Chr(10) & _
               "Duplicates skipped: " & skippedDupes & Chr(10) & _
               "Invalid names skipped: " & skippedBadChars, _
               vbInformation, "UTL — Create Folders"
        Exit Sub
    End If

    ' Build preview list (show first 20, then "...and X more")
    Dim previewList As String
    Dim previewCount As Long
    Dim folderName As Variant
    previewCount = 0
    For Each folderName In dict.Keys
        previewCount = previewCount + 1
        If previewCount <= 20 Then
            previewList = previewList & "  " & folderName & Chr(10)
        End If
    Next folderName
    If dict.Count > 20 Then
        previewList = previewList & "  ...and " & (dict.Count - 20) & " more" & Chr(10)
    End If

    ' Show preview and ask for confirmation
    Dim confirmMsg As String
    confirmMsg = "Ready to create " & dict.Count & " folder(s):" & Chr(10) & Chr(10) & _
                 previewList & Chr(10)
    If skippedBlanks > 0 Then confirmMsg = confirmMsg & "Blanks skipped: " & skippedBlanks & Chr(10)
    If skippedDupes > 0 Then confirmMsg = confirmMsg & "Duplicates skipped: " & skippedDupes & Chr(10)
    If skippedBadChars > 0 Then confirmMsg = confirmMsg & "Invalid names cleaned: " & skippedBadChars & Chr(10)
    confirmMsg = confirmMsg & Chr(10) & "Continue? (You will pick the location next)"

    If MsgBox(confirmMsg, vbOKCancel + vbQuestion, "UTL — Create Folders — Preview") = vbCancel Then
        Exit Sub
    End If

    ' Ask where to create the folders using folder picker
    Dim parentFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select the parent folder — new folders will be created inside this location"
        .ButtonName = "Select This Folder"
        If .Show = 0 Then Exit Sub  ' User cancelled
        parentFolder = .SelectedItems(1)
    End With

    ' Make sure path ends with backslash
    If Right(parentFolder, 1) <> "\" Then parentFolder = parentFolder & "\"

    ' Create the folders
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim createdCount As Long
    Dim skippedExist As Long
    Dim failedCount As Long
    Dim failedNames As String

    For Each folderName In dict.Keys
        Dim fullPath As String
        fullPath = parentFolder & CStr(folderName)

        If fso.FolderExists(fullPath) Then
            skippedExist = skippedExist + 1
        Else
            On Error Resume Next
            fso.CreateFolder fullPath
            If Err.Number <> 0 Then
                failedCount = failedCount + 1
                If failedCount <= 5 Then
                    failedNames = failedNames & "  " & CStr(folderName) & " (" & Err.Description & ")" & Chr(10)
                End If
                Err.Clear
            Else
                createdCount = createdCount + 1
            End If
            On Error GoTo 0
        End If
    Next folderName

    ' Summary message
    Dim summary As String
    summary = "Folder creation complete!" & Chr(10) & Chr(10) & _
              "Location: " & parentFolder & Chr(10) & Chr(10) & _
              "Created: " & createdCount & " new folder(s)" & Chr(10)
    If skippedExist > 0 Then summary = summary & "Already existed (skipped): " & skippedExist & Chr(10)
    If failedCount > 0 Then
        summary = summary & "Failed: " & failedCount & Chr(10) & failedNames
    End If

    MsgBox summary, vbInformation, "UTL — Create Folders"
End Sub
