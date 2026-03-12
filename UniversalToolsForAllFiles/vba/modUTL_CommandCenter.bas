Attribute VB_Name = "modUTL_CommandCenter"
'==============================================================================
' modUTL_CommandCenter — Universal Command Center
'==============================================================================
' PURPOSE:  One-click launcher for all Universal Toolkit tools.
'           Auto-discovers modUTL_* modules (if Trust Access is enabled),
'           falls back to built-in registry, and lets users register
'           their own custom macros.
'
' PUBLIC SUBS:
'   LaunchCommandCenter    — Main menu (categories → tools → run)
'   SearchTools            — Search all tools by keyword
'   ListAllTools           — Print full tool inventory to a sheet
'   RegisterCustomTool     — Add a user's own macro to the menu
'   RemoveCustomTool       — Remove a custom-registered macro
'   ViewCustomTools        — Show all custom-registered tools
'
' DEPENDENCIES: None (standalone). Works in any Excel workbook.
' TRUST ACCESS: Optional. If enabled, auto-discovers new modUTL_* modules.
'               If disabled, uses the built-in registry (still fully functional).
'
' AUTHOR:   KBT P&L Toolkit — Universal Tools
' VERSION:  1.0.0
' DATE:     2026-03-12
'==============================================================================
Option Explicit

'--- Constants ----------------------------------------------------------------
Private Const MAX_TOOLS As Long = 200
Private Const MAX_CUSTOM As Long = 50
Private Const CUSTOM_SHEET As String = "UTL_CustomTools"
Private Const INVENTORY_SHEET As String = "UTL_ToolInventory"
Private Const VERSION_LABEL As String = "Universal Command Center v1.1"

'--- Tool record type ---------------------------------------------------------
Private Type ToolRecord
    Category As String
    ToolName As String
    MacroName As String
    Description As String
    Source As String  ' "Built-in" or "Custom" or "Auto-discovered"
End Type

'--- Module-level registry ----------------------------------------------------
Private m_Tools() As ToolRecord
Private m_ToolCount As Long
Private m_Loaded As Boolean

'==============================================================================
' PUBLIC: LaunchCommandCenter
' Main entry point. Shows categories, then tools in chosen category, then runs.
'==============================================================================
Public Sub LaunchCommandCenter()
    On Error GoTo ErrHandler

    LoadRegistry

    '--- Build category list ---
    Dim cats() As String
    Dim catCount As Long
    catCount = 0
    ReDim cats(1 To 20)

    Dim i As Long, j As Long
    Dim found As Boolean

    For i = 1 To m_ToolCount
        found = False
        For j = 1 To catCount
            If cats(j) = m_Tools(i).Category Then
                found = True
                Exit For
            End If
        Next j
        If Not found Then
            catCount = catCount + 1
            If catCount > UBound(cats) Then ReDim Preserve cats(1 To catCount + 10)
            cats(catCount) = m_Tools(i).Category
        End If
    Next i

    If catCount = 0 Then
        MsgBox "No tools found in registry.", vbExclamation, VERSION_LABEL
        Exit Sub
    End If

    '--- Show category menu ---
    Dim menuText As String
    menuText = VERSION_LABEL & " (" & m_ToolCount & " tools)" & vbCrLf & vbCrLf
    menuText = menuText & "CATEGORIES:" & vbCrLf
    menuText = menuText & String(40, "-") & vbCrLf

    For i = 1 To catCount
        '--- Count tools in category ---
        Dim toolsInCat As Long
        toolsInCat = 0
        For j = 1 To m_ToolCount
            If m_Tools(j).Category = cats(i) Then toolsInCat = toolsInCat + 1
        Next j
        menuText = menuText & "  " & i & ". " & cats(i) & " (" & toolsInCat & " tools)" & vbCrLf
    Next i

    menuText = menuText & vbCrLf & "  S. Search all tools by keyword" & vbCrLf
    menuText = menuText & "  L. List all tools to a sheet" & vbCrLf
    menuText = menuText & vbCrLf & "Enter a number (or S/L):"

    Dim choice As String
    choice = InputBox(menuText, VERSION_LABEL)

    If Len(Trim(choice)) = 0 Then Exit Sub

    '--- Handle S = Search ---
    If UCase(Trim(choice)) = "S" Then
        SearchTools
        Exit Sub
    End If

    '--- Handle L = List ---
    If UCase(Trim(choice)) = "L" Then
        ListAllTools
        Exit Sub
    End If

    '--- Validate category number ---
    Dim catNum As Long
    If Not IsNumeric(choice) Then
        MsgBox "Invalid selection.", vbExclamation, VERSION_LABEL
        Exit Sub
    End If
    catNum = CLng(choice)
    If catNum < 1 Or catNum > catCount Then
        MsgBox "Invalid category number. Enter 1-" & catCount & ".", vbExclamation, VERSION_LABEL
        Exit Sub
    End If

    '--- Show tools in selected category ---
    Dim selectedCat As String
    selectedCat = cats(catNum)

    Dim toolIdxs() As Long
    Dim toolIdxCount As Long
    toolIdxCount = 0
    ReDim toolIdxs(1 To MAX_TOOLS)

    For i = 1 To m_ToolCount
        If m_Tools(i).Category = selectedCat Then
            toolIdxCount = toolIdxCount + 1
            toolIdxs(toolIdxCount) = i
        End If
    Next i

    Dim toolMenu As String
    toolMenu = selectedCat & " (" & toolIdxCount & " tools)" & vbCrLf
    toolMenu = toolMenu & String(40, "-") & vbCrLf & vbCrLf

    For i = 1 To toolIdxCount
        Dim idx As Long
        idx = toolIdxs(i)
        toolMenu = toolMenu & "  " & i & ". " & m_Tools(idx).ToolName & vbCrLf
        If Len(m_Tools(idx).Description) > 0 Then
            toolMenu = toolMenu & "     " & m_Tools(idx).Description & vbCrLf
        End If
    Next i

    toolMenu = toolMenu & vbCrLf & "Enter tool number to run (or 0 to go back):"

    Dim toolChoice As String
    toolChoice = InputBox(toolMenu, selectedCat & " — " & VERSION_LABEL)

    If Len(Trim(toolChoice)) = 0 Then Exit Sub
    If Not IsNumeric(toolChoice) Then
        MsgBox "Invalid selection.", vbExclamation, VERSION_LABEL
        Exit Sub
    End If

    Dim toolNum As Long
    toolNum = CLng(toolChoice)
    If toolNum = 0 Then
        LaunchCommandCenter  ' Go back to main menu
        Exit Sub
    End If
    If toolNum < 1 Or toolNum > toolIdxCount Then
        MsgBox "Invalid tool number. Enter 1-" & toolIdxCount & ".", vbExclamation, VERSION_LABEL
        Exit Sub
    End If

    '--- Run the selected tool ---
    Dim macroToRun As String
    macroToRun = m_Tools(toolIdxs(toolNum)).MacroName

    Application.StatusBar = "Running: " & m_Tools(toolIdxs(toolNum)).ToolName & "..."
    Application.Run macroToRun
    Application.StatusBar = False

    Exit Sub

ErrHandler:
    Application.StatusBar = False
    If Err.Number = 1004 Then
        MsgBox "Could not run macro: " & macroToRun & vbCrLf & vbCrLf & _
               "Make sure the module is imported into this workbook.", _
               vbExclamation, VERSION_LABEL
    Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, VERSION_LABEL
    End If
End Sub

'==============================================================================
' PUBLIC: SearchTools
' Keyword search across all tool names and descriptions.
'==============================================================================
Public Sub SearchTools()
    On Error GoTo ErrHandler

    LoadRegistry

    Dim keyword As String
    keyword = InputBox("Enter a keyword to search for:" & vbCrLf & vbCrLf & _
                       "Examples: duplicate, format, PDF, audit, clean", _
                       "Search Tools — " & VERSION_LABEL)

    If Len(Trim(keyword)) = 0 Then Exit Sub

    keyword = LCase(Trim(keyword))

    '--- Find matches ---
    Dim matchIdxs() As Long
    Dim matchCount As Long
    matchCount = 0
    ReDim matchIdxs(1 To MAX_TOOLS)

    Dim i As Long
    For i = 1 To m_ToolCount
        If InStr(1, LCase(m_Tools(i).ToolName), keyword) > 0 Or _
           InStr(1, LCase(m_Tools(i).Description), keyword) > 0 Or _
           InStr(1, LCase(m_Tools(i).Category), keyword) > 0 Or _
           InStr(1, LCase(m_Tools(i).MacroName), keyword) > 0 Then
            matchCount = matchCount + 1
            matchIdxs(matchCount) = i
        End If
    Next i

    If matchCount = 0 Then
        MsgBox "No tools found matching '" & keyword & "'.", vbInformation, VERSION_LABEL
        Exit Sub
    End If

    '--- Show results ---
    Dim results As String
    results = "Found " & matchCount & " tool(s) matching '" & keyword & "':" & vbCrLf
    results = results & String(40, "-") & vbCrLf & vbCrLf

    For i = 1 To matchCount
        Dim mi As Long
        mi = matchIdxs(i)
        results = results & "  " & i & ". [" & m_Tools(mi).Category & "] " & m_Tools(mi).ToolName & vbCrLf
        If Len(m_Tools(mi).Description) > 0 Then
            results = results & "     " & m_Tools(mi).Description & vbCrLf
        End If
    Next i

    results = results & vbCrLf & "Enter tool number to run (or 0 to cancel):"

    Dim choice As String
    choice = InputBox(results, "Search Results — " & VERSION_LABEL)

    If Len(Trim(choice)) = 0 Then Exit Sub
    If Not IsNumeric(choice) Then Exit Sub

    Dim choiceNum As Long
    choiceNum = CLng(choice)
    If choiceNum < 1 Or choiceNum > matchCount Then Exit Sub

    Dim macroToRun As String
    macroToRun = m_Tools(matchIdxs(choiceNum)).MacroName

    Application.StatusBar = "Running: " & m_Tools(matchIdxs(choiceNum)).ToolName & "..."
    Application.Run macroToRun
    Application.StatusBar = False

    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, VERSION_LABEL
End Sub

'==============================================================================
' PUBLIC: ListAllTools
' Prints the full tool inventory to a styled sheet.
'==============================================================================
Public Sub ListAllTools()
    On Error GoTo ErrHandler

    LoadRegistry

    Application.ScreenUpdating = False

    '--- Create or clear inventory sheet ---
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(INVENTORY_SHEET)
    On Error GoTo ErrHandler

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = INVENTORY_SHEET
    Else
        ws.Cells.Clear
    End If

    '--- Header row ---
    ws.Range("A1").Value = VERSION_LABEL & " — Tool Inventory"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    ws.Range("A2").Value = "Generated: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Range("A2").Font.Italic = True
    ws.Range("A2").Font.Color = RGB(100, 100, 100)

    ws.Range("A3").Value = "Total Tools: " & m_ToolCount
    ws.Range("A3").Font.Bold = True

    '--- Column headers ---
    Dim hdrRow As Long
    hdrRow = 5
    ws.Cells(hdrRow, 1).Value = "#"
    ws.Cells(hdrRow, 2).Value = "Category"
    ws.Cells(hdrRow, 3).Value = "Tool Name"
    ws.Cells(hdrRow, 4).Value = "Macro Name"
    ws.Cells(hdrRow, 5).Value = "Description"
    ws.Cells(hdrRow, 6).Value = "Source"

    Dim hdrRng As Range
    Set hdrRng = ws.Range(ws.Cells(hdrRow, 1), ws.Cells(hdrRow, 6))
    hdrRng.Font.Bold = True
    hdrRng.Font.Color = RGB(255, 255, 255)
    hdrRng.Interior.Color = RGB(11, 71, 121)

    '--- Data rows ---
    Dim r As Long
    Dim i As Long
    For i = 1 To m_ToolCount
        r = hdrRow + i
        ws.Cells(r, 1).Value = i
        ws.Cells(r, 2).Value = m_Tools(i).Category
        ws.Cells(r, 3).Value = m_Tools(i).ToolName
        ws.Cells(r, 4).Value = m_Tools(i).MacroName
        ws.Cells(r, 5).Value = m_Tools(i).Description
        ws.Cells(r, 6).Value = m_Tools(i).Source

        '--- Alternating row color ---
        If i Mod 2 = 0 Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).Interior.Color = RGB(235, 241, 250)
        End If
    Next i

    '--- Auto-fit ---
    ws.Columns("A:F").AutoFit
    If ws.Columns("E").ColumnWidth > 60 Then ws.Columns("E").ColumnWidth = 60

    ws.Activate
    ws.Range("A1").Select

    Application.ScreenUpdating = True

    MsgBox m_ToolCount & " tools listed on '" & INVENTORY_SHEET & "' sheet.", _
           vbInformation, VERSION_LABEL

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, VERSION_LABEL
End Sub

'==============================================================================
' PUBLIC: RegisterCustomTool
' Lets users add their own macros to the Command Center.
'==============================================================================
Public Sub RegisterCustomTool()
    On Error GoTo ErrHandler

    Dim macroName As String
    macroName = InputBox("Enter the macro name to register:" & vbCrLf & vbCrLf & _
                         "This must be a Public Sub in any module in this workbook." & vbCrLf & _
                         "Example: MyModule.MyMacro  or just  MyMacro", _
                         "Register Custom Tool — " & VERSION_LABEL)
    If Len(Trim(macroName)) = 0 Then Exit Sub
    macroName = Trim(macroName)

    Dim displayName As String
    displayName = InputBox("Enter a display name for this tool:" & vbCrLf & vbCrLf & _
                           "Example: My Custom Report Builder", _
                           "Register Custom Tool — " & VERSION_LABEL)
    If Len(Trim(displayName)) = 0 Then Exit Sub
    displayName = Trim(displayName)

    Dim category As String
    category = InputBox("Enter a category for this tool:" & vbCrLf & vbCrLf & _
                        "Example: My Tools" & vbCrLf & _
                        "(Leave blank to use 'Custom Tools')", _
                        "Register Custom Tool — " & VERSION_LABEL)
    If Len(Trim(category)) = 0 Then category = "Custom Tools"
    category = Trim(category)

    Dim description As String
    description = InputBox("Enter a short description (optional):" & vbCrLf & vbCrLf & _
                           "Example: Builds my weekly status report", _
                           "Register Custom Tool — " & VERSION_LABEL)
    description = Trim(description)

    '--- Save to hidden sheet ---
    Dim ws As Worksheet
    Set ws = GetOrCreateCustomSheet()

    '--- Check for duplicate ---
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        If LCase(ws.Cells(i, 2).Value) = LCase(macroName) Then
            MsgBox "A tool with macro name '" & macroName & "' is already registered." & vbCrLf & _
                   "Use RemoveCustomTool first if you want to replace it.", _
                   vbExclamation, VERSION_LABEL
            Exit Sub
        End If
    Next i

    '--- Add new row ---
    Dim newRow As Long
    newRow = lastRow + 1
    ws.Cells(newRow, 1).Value = category
    ws.Cells(newRow, 2).Value = macroName
    ws.Cells(newRow, 3).Value = displayName
    ws.Cells(newRow, 4).Value = description
    ws.Cells(newRow, 5).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")

    '--- Force reload on next launch ---
    m_Loaded = False

    MsgBox "Tool registered successfully!" & vbCrLf & vbCrLf & _
           "Name: " & displayName & vbCrLf & _
           "Macro: " & macroName & vbCrLf & _
           "Category: " & category & vbCrLf & vbCrLf & _
           "It will appear in the Command Center next time you open it.", _
           vbInformation, VERSION_LABEL

    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, VERSION_LABEL
End Sub

'==============================================================================
' PUBLIC: RemoveCustomTool
' Removes a custom-registered tool from the Command Center.
'==============================================================================
Public Sub RemoveCustomTool()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(CUSTOM_SHEET)
    On Error GoTo ErrHandler

    If ws Is Nothing Then
        MsgBox "No custom tools have been registered yet.", vbInformation, VERSION_LABEL
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No custom tools have been registered yet.", vbInformation, VERSION_LABEL
        Exit Sub
    End If

    '--- Build list ---
    Dim menuText As String
    menuText = "Custom Tools Registered:" & vbCrLf & String(40, "-") & vbCrLf & vbCrLf

    Dim i As Long
    Dim count As Long
    count = 0
    For i = 2 To lastRow
        If Len(ws.Cells(i, 2).Value) > 0 Then
            count = count + 1
            menuText = menuText & "  " & count & ". " & ws.Cells(i, 3).Value & _
                       " (" & ws.Cells(i, 2).Value & ")" & vbCrLf
        End If
    Next i

    If count = 0 Then
        MsgBox "No custom tools have been registered yet.", vbInformation, VERSION_LABEL
        Exit Sub
    End If

    menuText = menuText & vbCrLf & "Enter number to remove (or 0 to cancel):"

    Dim choice As String
    choice = InputBox(menuText, "Remove Custom Tool — " & VERSION_LABEL)

    If Len(Trim(choice)) = 0 Then Exit Sub
    If Not IsNumeric(choice) Then Exit Sub

    Dim choiceNum As Long
    choiceNum = CLng(choice)
    If choiceNum < 1 Or choiceNum > count Then Exit Sub

    '--- Find and delete the row ---
    Dim targetRow As Long
    count = 0
    For i = 2 To lastRow
        If Len(ws.Cells(i, 2).Value) > 0 Then
            count = count + 1
            If count = choiceNum Then
                targetRow = i
                Exit For
            End If
        End If
    Next i

    Dim removedName As String
    removedName = ws.Cells(targetRow, 3).Value
    ws.Rows(targetRow).Delete

    m_Loaded = False

    MsgBox "Removed: " & removedName, vbInformation, VERSION_LABEL

    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, VERSION_LABEL
End Sub

'==============================================================================
' PUBLIC: ViewCustomTools
' Shows all custom-registered tools.
'==============================================================================
Public Sub ViewCustomTools()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(CUSTOM_SHEET)
    On Error GoTo ErrHandler

    If ws Is Nothing Then
        MsgBox "No custom tools have been registered yet." & vbCrLf & vbCrLf & _
               "Use RegisterCustomTool to add your own macros.", _
               vbInformation, VERSION_LABEL
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No custom tools have been registered yet." & vbCrLf & vbCrLf & _
               "Use RegisterCustomTool to add your own macros.", _
               vbInformation, VERSION_LABEL
        Exit Sub
    End If

    Dim msg As String
    msg = "Custom Tools Registered:" & vbCrLf & String(40, "-") & vbCrLf & vbCrLf

    Dim count As Long
    count = 0
    Dim i As Long
    For i = 2 To lastRow
        If Len(ws.Cells(i, 2).Value) > 0 Then
            count = count + 1
            msg = msg & count & ". " & ws.Cells(i, 3).Value & vbCrLf
            msg = msg & "   Macro: " & ws.Cells(i, 2).Value & vbCrLf
            msg = msg & "   Category: " & ws.Cells(i, 1).Value & vbCrLf
            If Len(ws.Cells(i, 4).Value) > 0 Then
                msg = msg & "   Description: " & ws.Cells(i, 4).Value & vbCrLf
            End If
            msg = msg & "   Added: " & ws.Cells(i, 5).Value & vbCrLf & vbCrLf
        End If
    Next i

    If count = 0 Then
        msg = "No custom tools registered."
    Else
        msg = msg & "Total: " & count & " custom tool(s)"
    End If

    MsgBox msg, vbInformation, VERSION_LABEL

    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, VERSION_LABEL
End Sub

'==============================================================================
' PRIVATE: LoadRegistry
' Loads built-in tools, auto-discovered tools, and custom tools.
'==============================================================================
Private Sub LoadRegistry()
    If m_Loaded Then Exit Sub

    m_ToolCount = 0
    ReDim m_Tools(1 To MAX_TOOLS)

    '--- 1. Load built-in registry ---
    LoadBuiltInTools

    '--- 2. Try auto-discovery (requires Trust Access) ---
    AutoDiscoverTools

    '--- 3. Load custom tools from hidden sheet ---
    LoadCustomTools

    m_Loaded = True
End Sub

'==============================================================================
' PRIVATE: LoadBuiltInTools
' Hard-coded registry of all known Universal Toolkit tools.
' This ensures the Command Center works even without Trust Access.
'==============================================================================
Private Sub LoadBuiltInTools()

    '=== AUDIT (8 tools) ===
    AddTool "Audit", "External Link Finder", "modUTL_Audit.ExternalLinkFinder", _
            "Find all external links in the workbook", "Built-in"
    AddTool "Audit", "Circular Reference Detector", "modUTL_Audit.CircularReferenceDetector", _
            "Detect circular references across all sheets", "Built-in"
    AddTool "Audit", "Workbook Error Scanner", "modUTL_Audit.WorkbookErrorScanner", _
            "Scan for #REF, #VALUE, #N/A and other errors", "Built-in"
    AddTool "Audit", "Data Quality Scorecard", "modUTL_Audit.DataQualityScorecard", _
            "Generate a data quality score for the workbook", "Built-in"
    AddTool "Audit", "Named Range Auditor", "modUTL_Audit.NamedRangeAuditor", _
            "Audit all named ranges for validity", "Built-in"
    AddTool "Audit", "Data Validation Checker", "modUTL_Audit.DataValidationChecker", _
            "Check data validation rules across all sheets", "Built-in"
    AddTool "Audit", "Inconsistent Formulas Auditor", "modUTL_Audit.InconsistentFormulasAuditor", _
            "Find formulas that break row/column patterns", "Built-in"
    AddTool "Audit", "External Link Severance", "modUTL_Audit.ExternalLinkSeveranceProtocol", _
            "Break and remove all external links safely", "Built-in"

    '=== BRANDING (2 tools) ===
    AddTool "Branding", "Apply iPipeline Branding", "modUTL_Branding.ApplyiPipelineBranding", _
            "Apply iPipeline brand colors and fonts to active sheet", "Built-in"
    AddTool "Branding", "Set iPipeline Theme Colors", "modUTL_Branding.SetiPipelineThemeColors", _
            "Set workbook theme to iPipeline brand palette", "Built-in"

    '=== DATA CLEANING (12 tools) ===
    AddTool "Data Cleaning", "Unmerge and Fill Down", "modUTL_DataCleaning.UnmergeAndFillDown", _
            "Unmerge cells and fill down with the merged value", "Built-in"
    AddTool "Data Cleaning", "Fill Blanks Down", "modUTL_DataCleaning.FillBlanksDown", _
            "Fill blank cells with the value above", "Built-in"
    AddTool "Data Cleaning", "Convert Text to Numbers", "modUTL_DataCleaning.ConvertTextToNumbers", _
            "Convert text-stored numbers to real numbers", "Built-in"
    AddTool "Data Cleaning", "Remove Spaces", "modUTL_DataCleaning.RemoveLeadingTrailingSpaces", _
            "Remove leading and trailing spaces from all cells", "Built-in"
    AddTool "Data Cleaning", "Delete Blank Rows", "modUTL_DataCleaning.DeleteBlankRows", _
            "Delete all completely blank rows", "Built-in"
    AddTool "Data Cleaning", "Replace Error Values", "modUTL_DataCleaning.ReplaceErrorValues", _
            "Replace #REF, #N/A, etc. with zero or blank", "Built-in"
    AddTool "Data Cleaning", "Highlight Duplicate Rows", "modUTL_DataCleaning.HighlightDuplicateRows", _
            "Highlight duplicate rows in yellow", "Built-in"
    AddTool "Data Cleaning", "Remove Duplicate Rows", "modUTL_DataCleaning.RemoveDuplicateRows", _
            "Remove duplicate rows (keeps first occurrence)", "Built-in"
    AddTool "Data Cleaning", "Multi-Replace Data Cleaner", "modUTL_DataCleaning.MultiReplaceDataCleaner", _
            "Find and replace multiple values at once", "Built-in"
    AddTool "Data Cleaning", "Formula to Value", "modUTL_DataCleaning.FormulaToValueHardcoder", _
            "Convert formulas to their current values", "Built-in"
    AddTool "Data Cleaning", "Phantom Hyperlink Purger", "modUTL_DataCleaning.PhantomHyperlinkPurger", _
            "Remove invisible/phantom hyperlinks", "Built-in"
    AddTool "Data Cleaning", "Numbers to Words", "modUTL_DataCleaning.ConvertNumbersToWords", _
            "Convert numbers to written words (e.g., 100 = One Hundred)", "Built-in"

    '=== DATA SANITIZER (4 tools) ===
    AddTool "Data Sanitizer", "Run Full Sanitize", "modUTL_DataSanitizer.RunFullSanitize", _
            "Fix text-numbers, floating-point tails, and integer formats", "Built-in"
    AddTool "Data Sanitizer", "Preview Sanitize Changes", "modUTL_DataSanitizer.PreviewSanitizeChanges", _
            "Dry-run report showing what would change (no edits)", "Built-in"
    AddTool "Data Sanitizer", "Fix Floating-Point Tails", "modUTL_DataSanitizer.FixFloatingPointTails", _
            "Fix floating-point noise (e.g., 10.000000001)", "Built-in"
    AddTool "Data Sanitizer", "Convert Text-Stored Numbers", "modUTL_DataSanitizer.ConvertTextStoredNumbers", _
            "Convert text-stored numbers to real numbers (smart detection)", "Built-in"

    '=== EXECUTIVE BRIEF (1 tool) ===
    AddTool "Executive Brief", "Generate Exec Brief", "modUTL_ExecBrief.GenerateExecBrief", _
            "Scan any workbook and build an executive brief report", "Built-in"

    '=== FINANCE (14 tools) ===
    AddTool "Finance", "Duplicate Invoice Detector", "modUTL_Finance.DuplicateInvoiceDetector", _
            "Find duplicate invoices by amount, date, vendor", "Built-in"
    AddTool "Finance", "Auto-Balancing GL Validator", "modUTL_Finance.AutoBalancingGLValidator", _
            "Validate that debits equal credits in GL data", "Built-in"
    AddTool "Finance", "Trial Balance Checker", "modUTL_Finance.TrialBalanceChecker", _
            "Verify trial balance totals to zero", "Built-in"
    AddTool "Finance", "Journal Entry Validator", "modUTL_Finance.JournalEntryValidator", _
            "Validate journal entries for completeness and balance", "Built-in"
    AddTool "Finance", "Flux Analysis", "modUTL_Finance.FluxAnalysis", _
            "Period-over-period flux analysis with thresholds", "Built-in"
    AddTool "Finance", "AP Aging Summary", "modUTL_Finance.APAgingSummaryGenerator", _
            "Generate accounts payable aging summary", "Built-in"
    AddTool "Finance", "AR Aging Summary", "modUTL_Finance.ARAgingSummaryGenerator", _
            "Generate accounts receivable aging summary", "Built-in"
    AddTool "Finance", "Aging Bucket Calculator", "modUTL_Finance.AgingBucketCalculator", _
            "Calculate aging buckets (Current, 30, 60, 90+)", "Built-in"
    AddTool "Finance", "Variance Analysis Template", "modUTL_Finance.VarianceAnalysisTemplate", _
            "Build a variance analysis template", "Built-in"
    AddTool "Finance", "Quick Corkscrew Builder", "modUTL_Finance.QuickCorkscrewBuilder", _
            "Build opening-additions-disposals-closing rollforward", "Built-in"
    AddTool "Finance", "Period Roll Forward", "modUTL_Finance.FinancialPeriodRollForward", _
            "Roll financial periods forward", "Built-in"
    AddTool "Finance", "Multi-Currency Consolidation", "modUTL_Finance.MultiCurrencyConsolidationAggregator", _
            "Consolidate multi-currency data", "Built-in"
    AddTool "Finance", "Ratio Analysis Dashboard", "modUTL_Finance.RatioAnalysisDashboard", _
            "Calculate and display key financial ratios", "Built-in"
    AddTool "Finance", "GL Journal Mapper", "modUTL_Finance.GeneralLedgerJournalMapper", _
            "Map GL entries to journal categories", "Built-in"

    '=== FORMATTING (9 tools) ===
    AddTool "Formatting", "Auto-Fit All Columns/Rows", "modUTL_Formatting.AutoFitAllColumnsRows", _
            "Auto-fit column widths and row heights", "Built-in"
    AddTool "Formatting", "Freeze Top Row (All Sheets)", "modUTL_Formatting.FreezeTopRowAllSheets", _
            "Freeze the top row on every sheet", "Built-in"
    AddTool "Formatting", "Number Format Standardizer", "modUTL_Formatting.NumberFormatStandardizer", _
            "Standardize number formats across selection", "Built-in"
    AddTool "Formatting", "Currency Format Standardizer", "modUTL_Formatting.CurrencyFormatStandardizer", _
            "Standardize currency formats", "Built-in"
    AddTool "Formatting", "Date Format Standardizer", "modUTL_Formatting.DateFormatStandardizer", _
            "Standardize date formats across selection", "Built-in"
    AddTool "Formatting", "Highlight Negatives Red", "modUTL_Formatting.HighlightNegativesRed", _
            "Highlight negative numbers in red", "Built-in"
    AddTool "Formatting", "Financial Number Formatting", "modUTL_Formatting.FinancialNumberFormattingSuite", _
            "Apply financial number formatting (parentheses for negatives)", "Built-in"
    AddTool "Formatting", "Conditional Format Purger", "modUTL_Formatting.ConditionalFormatPurger", _
            "Remove all conditional formatting rules", "Built-in"
    AddTool "Formatting", "Print Header/Footer Standardizer", "modUTL_Formatting.PrintHeaderFooterStandardizer", _
            "Standardize print headers and footers", "Built-in"

    '=== PROGRESS BAR (3 tools) ===
    AddTool "Progress Bar", "Start Progress", "modUTL_ProgressBar.StartProgress", _
            "Initialize progress bar (call from your own macros)", "Built-in"
    AddTool "Progress Bar", "Update Progress", "modUTL_ProgressBar.UpdateProgress", _
            "Update progress bar (call from your own macros)", "Built-in"
    AddTool "Progress Bar", "End Progress", "modUTL_ProgressBar.EndProgress", _
            "Close progress bar (call from your own macros)", "Built-in"

    '=== SHEET TOOLS (4 tools) ===
    AddTool "Sheet Tools", "List All Sheets with Links", "modUTL_SheetTools.ListAllSheetsWithLinks", _
            "Create a sheet index with clickable hyperlinks", "Built-in"
    AddTool "Sheet Tools", "Template Cloner", "modUTL_SheetTools.TemplateCloner", _
            "Clone any sheet multiple times (1-50 copies)", "Built-in"
    AddTool "Sheet Tools", "Generate Unique Customer IDs", "modUTL_SheetTools.GenerateUniqueCustomerIDs", _
            "Fill blank cells with sequential unique IDs", "Built-in"
    AddTool "Sheet Tools", "Create Folders from Selection", "modUTL_SheetTools.CreateFoldersFromSelection", _
            "Create Windows folders from selected cell values", "Built-in"

    '=== SPLASH SCREEN (2 tools) ===
    AddTool "Splash Screen", "Show Splash", "modUTL_SplashScreen.ShowSplash", _
            "Show a branded welcome splash screen", "Built-in"
    AddTool "Splash Screen", "Show Custom Splash", "modUTL_SplashScreen.ShowSplashCustom", _
            "Show a customizable splash screen", "Built-in"

    '=== WHAT-IF (4 tools) ===
    AddTool "What-If Analysis", "Run What-If Presets", "modUTL_WhatIf.RunWhatIfPresets", _
            "Apply preset percentage changes to selected cells", "Built-in"
    AddTool "What-If Analysis", "Run Custom What-If", "modUTL_WhatIf.RunWhatIf", _
            "Apply a custom percentage change to selected cells", "Built-in"
    AddTool "What-If Analysis", "Restore Baseline", "modUTL_WhatIf.RestoreBaseline", _
            "Undo all What-If changes and restore original values", "Built-in"
    AddTool "What-If Analysis", "View Baseline", "modUTL_WhatIf.ViewBaseline", _
            "View the saved baseline values", "Built-in"

    '=== WORKBOOK MANAGEMENT (15 tools) ===
    AddTool "Workbook Management", "Unhide All Sheets/Rows/Columns", "modUTL_WorkbookMgmt.UnhideAllSheetsRowsColumns", _
            "Unhide all hidden sheets, rows, and columns", "Built-in"
    AddTool "Workbook Management", "Export All Sheets (Combined PDF)", "modUTL_WorkbookMgmt.ExportAllSheetsCombinedPDF", _
            "Export all sheets to a single PDF", "Built-in"
    AddTool "Workbook Management", "Find & Replace (All Sheets)", "modUTL_WorkbookMgmt.FindReplaceAcrossAllSheets", _
            "Find and replace text across all sheets", "Built-in"
    AddTool "Workbook Management", "Search All Sheets", "modUTL_WorkbookMgmt.SearchAcrossAllSheets", _
            "Search for text across all sheets", "Built-in"
    AddTool "Workbook Management", "Batch Sheet Renamer", "modUTL_WorkbookMgmt.MultiSheetBatchRenamer", _
            "Rename multiple sheets at once using a pattern", "Built-in"
    AddTool "Workbook Management", "Sort Sheets Alphabetically", "modUTL_WorkbookMgmt.SortWorksheetsAlphabetically", _
            "Sort all worksheet tabs alphabetically", "Built-in"
    AddTool "Workbook Management", "Create Table of Contents", "modUTL_WorkbookMgmt.CreateTableOfContents", _
            "Create a clickable table of contents sheet", "Built-in"
    AddTool "Workbook Management", "Protect All Sheets", "modUTL_WorkbookMgmt.ProtectAllSheets", _
            "Protect all sheets with a password", "Built-in"
    AddTool "Workbook Management", "Unprotect All Sheets", "modUTL_WorkbookMgmt.UnprotectAllSheets", _
            "Unprotect all sheets", "Built-in"
    AddTool "Workbook Management", "Lock All Formula Cells", "modUTL_WorkbookMgmt.LockAllFormulaCells", _
            "Lock cells containing formulas", "Built-in"
    AddTool "Workbook Management", "Export Active Sheet PDF", "modUTL_WorkbookMgmt.ExportActiveSheetPDF", _
            "Export the active sheet to PDF", "Built-in"
    AddTool "Workbook Management", "Export All Sheets (Individual PDFs)", "modUTL_WorkbookMgmt.ExportAllSheetsIndividualPDFs", _
            "Export each sheet to its own PDF file", "Built-in"
    AddTool "Workbook Management", "Reset All Filters", "modUTL_WorkbookMgmt.ResetAllFilters", _
            "Clear all AutoFilter and Advanced Filter settings", "Built-in"
    AddTool "Workbook Management", "Build Distribution-Ready Copy", "modUTL_WorkbookMgmt.BuildDistributionReadyCopy", _
            "Create a clean copy for distribution (values only, no macros)", "Built-in"
    AddTool "Workbook Management", "Workbook Health Check", "modUTL_WorkbookMgmt.WorkbookHealthCheckagentId", _
            "Run a comprehensive health check on the workbook", "Built-in"

    '=== COLUMN OPERATIONS (4 tools) ===
    AddTool "Column Operations", "Split Column", "modUTL_ColumnOps.SplitColumn", _
            "Split a column by delimiter into multiple columns", "Built-in"
    AddTool "Column Operations", "Combine Columns", "modUTL_ColumnOps.CombineColumns", _
            "Merge multiple columns into one with a separator", "Built-in"
    AddTool "Column Operations", "Extract Pattern", "modUTL_ColumnOps.ExtractPattern", _
            "Extract numbers, text before/after delimiter, first/last N chars", "Built-in"
    AddTool "Column Operations", "Swap Columns", "modUTL_ColumnOps.SwapColumns", _
            "Swap the contents of two columns", "Built-in"

    '=== COMMENTS (4 tools) ===
    AddTool "Comments", "Extract All Comments", "modUTL_Comments.ExtractAllComments", _
            "Export all comments to a summary sheet", "Built-in"
    AddTool "Comments", "Delete Sheet Comments", "modUTL_Comments.DeleteSheetComments", _
            "Delete comments from user-selected sheets", "Built-in"
    AddTool "Comments", "Delete All Comments", "modUTL_Comments.DeleteAllComments", _
            "Delete all comments in the workbook", "Built-in"
    AddTool "Comments", "Count Comments", "modUTL_Comments.CountComments", _
            "Quick count of comments per sheet", "Built-in"

    '=== COMPARISON (3 tools) ===
    AddTool "Comparison", "Compare Sheets", "modUTL_Compare.CompareSheets", _
            "Compare two sheets cell-by-cell and highlight differences", "Built-in"
    AddTool "Comparison", "Compare Ranges", "modUTL_Compare.CompareRanges", _
            "Compare two selected ranges cell-by-cell", "Built-in"
    AddTool "Comparison", "Clear Compare Highlights", "modUTL_Compare.ClearCompareHighlights", _
            "Remove comparison highlighting from sheets", "Built-in"

    '=== CONSOLIDATION (2 tools) ===
    AddTool "Consolidation", "Consolidate Sheets", "modUTL_Consolidate.ConsolidateSheets", _
            "Combine selected sheets into one master sheet", "Built-in"
    AddTool "Consolidation", "Consolidate by Pattern", "modUTL_Consolidate.ConsolidateByPattern", _
            "Combine sheets matching a keyword pattern", "Built-in"

    '=== HIGHLIGHTS (5 tools) ===
    AddTool "Highlights", "Highlight by Threshold", "modUTL_Highlights.HighlightByThreshold", _
            "Highlight cells above/below a value you type", "Built-in"
    AddTool "Highlights", "Highlight Top/Bottom N", "modUTL_Highlights.HighlightTopBottom", _
            "Highlight the top N or bottom N values", "Built-in"
    AddTool "Highlights", "Highlight Duplicate Values", "modUTL_Highlights.HighlightDuplicateValues", _
            "Highlight cells with duplicate values", "Built-in"
    AddTool "Highlights", "Apply Color Scale", "modUTL_Highlights.ApplyColorScale", _
            "Red-Yellow-Green gradient based on values", "Built-in"
    AddTool "Highlights", "Clear Highlights", "modUTL_Highlights.ClearHighlights", _
            "Remove highlighting from selection, sheet, or all sheets", "Built-in"

    '=== LOOKUP BUILDER (4 tools) ===
    AddTool "Lookup Builder", "Build VLOOKUP", "modUTL_LookupBuilder.BuildVLOOKUP", _
            "Build VLOOKUP formulas step by step", "Built-in"
    AddTool "Lookup Builder", "Build INDEX-MATCH", "modUTL_LookupBuilder.BuildINDEXMATCH", _
            "Build INDEX-MATCH formulas (more flexible than VLOOKUP)", "Built-in"
    AddTool "Lookup Builder", "Match and Pull", "modUTL_LookupBuilder.MatchAndPull", _
            "Match two lists and pull values across", "Built-in"
    AddTool "Lookup Builder", "Find Unmatched", "modUTL_LookupBuilder.FindUnmatched", _
            "Find values in list A that don't exist in list B", "Built-in"

    '=== PIVOT TOOLS (4 tools) ===
    AddTool "Pivot Tools", "Refresh All Pivots", "modUTL_PivotTools.RefreshAllPivots", _
            "Refresh every pivot table (skips external connections)", "Built-in"
    AddTool "Pivot Tools", "Refresh Selected Pivots", "modUTL_PivotTools.RefreshSelectedPivots", _
            "Pick which pivot tables to refresh", "Built-in"
    AddTool "Pivot Tools", "List All Pivots", "modUTL_PivotTools.ListAllPivots", _
            "Build an inventory sheet of all pivot tables", "Built-in"
    AddTool "Pivot Tools", "Clear Old Pivot Cache", "modUTL_PivotTools.ClearOldPivotCache", _
            "Report on orphaned pivot caches (file size bloat)", "Built-in"

    '=== TAB ORGANIZER (6 tools) ===
    AddTool "Tab Organizer", "Color Tabs by Keyword", "modUTL_TabOrganizer.ColorTabsByKeyword", _
            "Color tabs matching a keyword", "Built-in"
    AddTool "Tab Organizer", "Color Tabs Interactive", "modUTL_TabOrganizer.ColorTabsInteractive", _
            "Pick specific sheets and assign a tab color", "Built-in"
    AddTool "Tab Organizer", "Bulk Hide Sheets", "modUTL_TabOrganizer.BulkHideSheets", _
            "Hide multiple sheets at once", "Built-in"
    AddTool "Tab Organizer", "Bulk Unhide Sheets", "modUTL_TabOrganizer.BulkUnhideSheets", _
            "Unhide multiple sheets at once", "Built-in"
    AddTool "Tab Organizer", "Reorder Tabs", "modUTL_TabOrganizer.ReorderTabs", _
            "Move sheets to front, back, or after a specific sheet", "Built-in"
    AddTool "Tab Organizer", "Bulk Rename Tabs", "modUTL_TabOrganizer.BulkRenameTabs", _
            "Find/replace text in sheet tab names", "Built-in"

    '=== VALIDATION BUILDER (6 tools) ===
    AddTool "Validation Builder", "Create Dropdown List", "modUTL_ValidationBuilder.CreateDropdownList", _
            "Create a dropdown from a range or typed list", "Built-in"
    AddTool "Validation Builder", "Number Validation", "modUTL_ValidationBuilder.ApplyNumberValidation", _
            "Restrict cells to numbers (with min/max options)", "Built-in"
    AddTool "Validation Builder", "Date Validation", "modUTL_ValidationBuilder.ApplyDateValidation", _
            "Restrict cells to dates (with range options)", "Built-in"
    AddTool "Validation Builder", "Copy Validation Rules", "modUTL_ValidationBuilder.CopyValidationRules", _
            "Copy validation from one cell to another range", "Built-in"
    AddTool "Validation Builder", "Find Validation Violations", "modUTL_ValidationBuilder.FindValidationViolations", _
            "Find cells that break their validation rules", "Built-in"
    AddTool "Validation Builder", "Remove All Validation", "modUTL_ValidationBuilder.RemoveAllValidation", _
            "Remove validation from selection, sheet, or workbook", "Built-in"

    '=== COMMAND CENTER SELF (3 tools) ===
    AddTool "Command Center", "Register Custom Tool", "modUTL_CommandCenter.RegisterCustomTool", _
            "Add your own macro to the Command Center", "Built-in"
    AddTool "Command Center", "Remove Custom Tool", "modUTL_CommandCenter.RemoveCustomTool", _
            "Remove a custom-registered tool", "Built-in"
    AddTool "Command Center", "View Custom Tools", "modUTL_CommandCenter.ViewCustomTools", _
            "View all custom-registered tools", "Built-in"

End Sub

'==============================================================================
' PRIVATE: AutoDiscoverTools
' Scans VBProject for modUTL_* modules and adds any Public Subs not already
' in the registry. Requires Trust Access to be enabled.
'==============================================================================
Private Sub AutoDiscoverTools()
    On Error Resume Next

    '--- Test if Trust Access is available ---
    Dim compCount As Long
    compCount = ThisWorkbook.VBProject.VBComponents.Count
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub  ' Trust Access not enabled — skip discovery
    End If
    On Error GoTo 0

    Dim comp As Object  ' VBComponent
    Dim i As Long

    For Each comp In ThisWorkbook.VBProject.VBComponents
        '--- Only scan modUTL_* standard modules ---
        If comp.Type = 1 Then  ' vbext_ct_StdModule
            If Left(comp.Name, 7) = "modUTL_" And comp.Name <> "modUTL_CommandCenter" Then
                '--- Scan code for Public Sub/Function ---
                Dim codeModule As Object
                Set codeModule = comp.CodeModule

                Dim lineCount As Long
                lineCount = codeModule.CountOfLines

                Dim ln As Long
                For ln = 1 To lineCount
                    Dim codeLine As String
                    codeLine = Trim(codeModule.Lines(ln, 1))

                    Dim subName As String
                    subName = ""

                    If Left(codeLine, 11) = "Public Sub " Then
                        subName = Mid(codeLine, 12)
                        Dim parenPos As Long
                        parenPos = InStr(subName, "(")
                        If parenPos > 0 Then subName = Left(subName, parenPos - 1)
                        subName = Trim(subName)
                    ElseIf Left(codeLine, 16) = "Public Function " Then
                        subName = Mid(codeLine, 17)
                        parenPos = InStr(subName, "(")
                        If parenPos > 0 Then subName = Left(subName, parenPos - 1)
                        subName = Trim(subName)
                    End If

                    If Len(subName) > 0 Then
                        '--- Check if already in registry ---
                        Dim macroFull As String
                        macroFull = comp.Name & "." & subName

                        Dim alreadyExists As Boolean
                        alreadyExists = False
                        Dim j As Long
                        For j = 1 To m_ToolCount
                            If LCase(m_Tools(j).MacroName) = LCase(macroFull) Then
                                alreadyExists = True
                                Exit For
                            End If
                        Next j

                        If Not alreadyExists Then
                            '--- Derive category from module name ---
                            Dim catName As String
                            catName = Mid(comp.Name, 8)  ' Remove "modUTL_"
                            catName = catName & " (Discovered)"

                            AddTool catName, subName, macroFull, _
                                    "(Auto-discovered from " & comp.Name & ")", "Auto-discovered"
                        End If
                    End If
                Next ln
            End If
        End If
    Next comp
End Sub

'==============================================================================
' PRIVATE: LoadCustomTools
' Loads user-registered tools from the hidden custom tools sheet.
'==============================================================================
Private Sub LoadCustomTools()
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(CUSTOM_SHEET)
    If ws Is Nothing Then Exit Sub

    On Error GoTo 0

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        If Len(ws.Cells(i, 2).Value) > 0 Then
            AddTool CStr(ws.Cells(i, 1).Value), _
                    CStr(ws.Cells(i, 3).Value), _
                    CStr(ws.Cells(i, 2).Value), _
                    CStr(ws.Cells(i, 4).Value), _
                    "Custom"
        End If
    Next i
End Sub

'==============================================================================
' PRIVATE: AddTool
' Adds a tool to the in-memory registry.
'==============================================================================
Private Sub AddTool(ByVal cat As String, ByVal toolName As String, _
                    ByVal macroName As String, ByVal desc As String, _
                    ByVal src As String)
    If m_ToolCount >= MAX_TOOLS Then Exit Sub

    m_ToolCount = m_ToolCount + 1
    m_Tools(m_ToolCount).Category = cat
    m_Tools(m_ToolCount).ToolName = toolName
    m_Tools(m_ToolCount).MacroName = macroName
    m_Tools(m_ToolCount).Description = desc
    m_Tools(m_ToolCount).Source = src
End Sub

'==============================================================================
' PRIVATE: GetOrCreateCustomSheet
' Gets or creates the hidden sheet for custom tool storage.
'==============================================================================
Private Function GetOrCreateCustomSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(CUSTOM_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Application.ScreenUpdating = False
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = CUSTOM_SHEET

        '--- Header row ---
        ws.Cells(1, 1).Value = "Category"
        ws.Cells(1, 2).Value = "MacroName"
        ws.Cells(1, 3).Value = "DisplayName"
        ws.Cells(1, 4).Value = "Description"
        ws.Cells(1, 5).Value = "DateAdded"
        ws.Rows(1).Font.Bold = True

        '--- Hide the sheet ---
        ws.Visible = xlSheetHidden

        Application.ScreenUpdating = True
    End If

    Set GetOrCreateCustomSheet = ws
End Function
