Attribute VB_Name = "modSearch"
Option Explicit

'===============================================================================
' modSearch - Cross-Sheet Search Engine
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Search for any value, label, or keyword across all visible sheets.
'           Returns a clickable results list with sheet name, cell address, and
'           context. Replaces the need to manually Ctrl+F each sheet.
'
' PUBLIC SUBS:
'   SearchAll          - Prompt for keyword and search all sheets
'   SearchAndNavigate  - Search and jump to selected result
'   SearchCurrentSheet - Search active sheet only, highlight matches
'
' VERSION:  2.1.0
' CHANGES:  v2.0 -> v2.1:
'           + ISSUE-012: When MAX_RESULTS cap is hit, continues counting
'             total matches and reports "Showing first 200 of N total"
'           + Uses SH_SEARCH constant from modConfig (was Private Const)
'           + Uses SafeDeleteSheet (was inline delete)
'           + Added m_TotalMatches counter for uncapped count
'===============================================================================

Private Const MAX_RESULTS  As Long = 200

Private Type SearchHit
    SheetName As String
    CellRef   As String
    CellValue As String
    Context   As String  ' Value from column A of the same row for context
End Type

Private m_Hits() As SearchHit
Private m_HitCount As Long
Private m_TotalMatches As Long   ' v2.1: tracks ALL matches even past MAX_RESULTS

'===============================================================================
' SearchAll - Search all visible sheets for a keyword
'===============================================================================
Public Sub SearchAll()
    On Error GoTo ErrHandler
    
    Dim keyword As String
    keyword = InputBox("Enter search term:" & vbCrLf & vbCrLf & _
                       "Searches all visible sheets for matching cell values.", _
                       APP_NAME & " - Search Workbook")
    If keyword = "" Then Exit Sub
    
    m_HitCount = 0
    m_TotalMatches = 0
    Erase m_Hits
    
    modPerformance.TurboOn
    modPerformance.UpdateStatus "Searching for '" & keyword & "'...", 0
    
    Dim ws As Worksheet
    Dim sheetIdx As Long: sheetIdx = 0
    Dim totalSheets As Long: totalSheets = ThisWorkbook.Worksheets.Count
    
    For Each ws In ThisWorkbook.Worksheets
        sheetIdx = sheetIdx + 1
        If ws.Visible = xlSheetVisible And ws.Name <> SH_SEARCH Then
            modPerformance.UpdateStatus "Searching " & ws.Name & "...", sheetIdx / totalSheets
            SearchSheet ws, keyword
        End If
    Next ws
    
    ' Write results
    WriteSearchResults keyword
    
    modPerformance.TurboOff
    
    modLogger.LogAction "modSearch", "SearchAll", _
                        "'" & keyword & "' -> " & m_TotalMatches & " total, " & _
                        m_HitCount & " displayed"
    
    If m_TotalMatches = 0 Then
        MsgBox "No results found for '" & keyword & "'.", vbInformation, APP_NAME
    Else
        ' ISSUE-012 FIX: Report total vs displayed when cap is hit
        Dim msg As String
        If m_TotalMatches > MAX_RESULTS Then
            msg = "Showing first " & MAX_RESULTS & " of " & m_TotalMatches & _
                  " total matches for '" & keyword & "'." & vbCrLf & vbCrLf & _
                  "Tip: Use a more specific search term to narrow results."
        Else
            msg = m_TotalMatches & " results for '" & keyword & "'."
        End If
        msg = msg & vbCrLf & "See '" & SH_SEARCH & "' sheet."
        MsgBox msg, vbInformation, APP_NAME
    End If
    Exit Sub
    
ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modSearch", "ERROR-SearchAll", Err.Description
    MsgBox "Search error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' SearchAndNavigate - Search and jump to a selected result
'===============================================================================
Public Sub SearchAndNavigate()
    On Error GoTo ErrHandler
    
    Dim keyword As String
    keyword = InputBox("Enter search term to find and navigate to:", _
                       APP_NAME & " - Find & Go")
    If keyword = "" Then Exit Sub
    
    m_HitCount = 0
    m_TotalMatches = 0
    Erase m_Hits
    
    modPerformance.TurboOn
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            SearchSheet ws, keyword
            If m_HitCount >= 20 Then Exit For  ' Limit for navigation list
        End If
    Next ws
    
    modPerformance.TurboOff
    
    If m_HitCount = 0 Then
        MsgBox "No results found for '" & keyword & "'.", vbInformation, APP_NAME
        Exit Sub
    End If
    
    ' Build navigation list
    Dim navList As String: navList = ""
    Dim i As Long
    For i = 0 To m_HitCount - 1
        navList = navList & (i + 1) & ". [" & m_Hits(i).SheetName & "] " & _
                  m_Hits(i).CellRef & " = " & Left(m_Hits(i).CellValue, 40) & vbCrLf
    Next i
    
    Dim choice As String
    choice = InputBox(m_HitCount & " results found. Enter number to navigate:" & vbCrLf & vbCrLf & _
                      navList, APP_NAME & " - Navigate")
    If choice = "" Then Exit Sub
    If Not IsNumeric(choice) Then Exit Sub
    
    Dim idx As Long: idx = CLng(choice) - 1
    If idx < 0 Or idx > m_HitCount - 1 Then
        MsgBox "Invalid selection.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' Navigate to the result (final navigation only — permitted)
    ThisWorkbook.Worksheets(m_Hits(idx).SheetName).Activate
    ThisWorkbook.Worksheets(m_Hits(idx).SheetName).Range(m_Hits(idx).CellRef).Select
    
    modLogger.LogAction "modSearch", "SearchAndNavigate", _
        "'" & keyword & "' -> navigated to " & m_Hits(idx).SheetName & "!" & m_Hits(idx).CellRef
    Exit Sub
    
ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modSearch", "ERROR-Navigate", Err.Description
    MsgBox "Navigation error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' SearchCurrentSheet - Search active sheet only, highlight matches
'===============================================================================
Public Sub SearchCurrentSheet()
    On Error GoTo ErrHandler
    
    Dim keyword As String
    keyword = InputBox("Search '" & ActiveSheet.Name & "' for:", _
                       APP_NAME & " - Search Sheet")
    If keyword = "" Then Exit Sub
    
    m_HitCount = 0
    m_TotalMatches = 0
    Erase m_Hits
    
    SearchSheet ActiveSheet, keyword
    
    If m_HitCount = 0 Then
        MsgBox "No results for '" & keyword & "' on this sheet.", vbInformation, APP_NAME
        Exit Sub
    End If
    
    ' Highlight results on the active sheet
    Dim i As Long
    For i = 0 To m_HitCount - 1
        ActiveSheet.Range(m_Hits(i).CellRef).Interior.Color = RGB(255, 255, 0)
    Next i
    
    ' Select first result (final navigation — permitted)
    ActiveSheet.Range(m_Hits(0).CellRef).Select
    
    MsgBox m_HitCount & " matches found and highlighted in yellow." & vbCrLf & _
           "First match selected.", vbInformation, APP_NAME
    
    modLogger.LogAction "modSearch", "SearchCurrentSheet", _
        "'" & keyword & "' on " & ActiveSheet.Name & " -> " & m_HitCount & " matches"
    Exit Sub
    
ErrHandler:
    modLogger.LogAction "modSearch", "ERROR-SearchCurrent", Err.Description
    MsgBox "Search error: " & Err.Description, vbCritical, APP_NAME
End Sub


'===============================================================================
'
' ===  PRIVATE HELPERS  ========================================================
'
'===============================================================================

'===============================================================================
' SearchSheet - Scan a single sheet for the keyword
' ISSUE-012 FIX: Always counts ALL matches in m_TotalMatches.
' Only stores the first MAX_RESULTS hits in m_Hits() for display.
'===============================================================================
Private Sub SearchSheet(ByVal ws As Worksheet, ByVal keyword As String)
    Dim firstAddr As String
    Dim found As Range
    
    Set found = ws.UsedRange.Find( _
        What:=keyword, _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        MatchCase:=False)
    
    If found Is Nothing Then Exit Sub
    firstAddr = found.Address
    
    Do
        ' Always count total matches (even past cap)
        m_TotalMatches = m_TotalMatches + 1
        
        ' Only store details if under the cap
        If m_HitCount < MAX_RESULTS Then
            ReDim Preserve m_Hits(m_HitCount)
            With m_Hits(m_HitCount)
                .SheetName = ws.Name
                .CellRef = found.Address(False, False)
                .CellValue = Left(CStr(found.Value), 100)
                ' Grab context: col A value from the same row
                If found.Column > 1 Then
                    .Context = Left(modConfig.SafeStr(ws.Cells(found.row, 1).Value), 60)
                Else
                    .Context = ""
                End If
            End With
            m_HitCount = m_HitCount + 1
        End If
        
        Set found = ws.UsedRange.FindNext(found)
        If found Is Nothing Then Exit Do
    Loop While found.Address <> firstAddr
End Sub

'===============================================================================
' WriteSearchResults - Create a results sheet with clickable hyperlinks
'===============================================================================
Private Sub WriteSearchResults(ByVal keyword As String)
    modConfig.SafeDeleteSheet SH_SEARCH
    
    If m_HitCount = 0 Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = SH_SEARCH
    ws.Tab.Color = RGB(0, 176, 240)
    
    ' Title
    ws.Range("A1").Value = "Search Results - '" & keyword & "'"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    
    ' Subtitle with cap warning if applicable
    Dim subtitle As String
    If m_TotalMatches > MAX_RESULTS Then
        subtitle = "Showing " & m_HitCount & " of " & m_TotalMatches & " total matches"
    Else
        subtitle = m_HitCount & " results"
    End If
    ws.Range("A2").Value = subtitle & " | " & Format(Now, "yyyy-mm-dd hh:mm")
    ws.Range("A2").Font.Italic = True
    If m_TotalMatches > MAX_RESULTS Then
        ws.Range("A2").Font.Color = RGB(192, 0, 0)  ' Red warning when capped
    End If
    
    ' Headers
    Dim headers As Variant
    headers = Array("Sheet", "Cell", "Value", "Row Context (Col A)")
    Dim c As Long
    For c = 0 To UBound(headers)
        ws.Cells(4, c + 1).Value = headers(c)
    Next c
    With ws.Range("A4:D4")
        .Font.Bold = True
        .Interior.Color = CLR_NAVY
        .Font.Color = CLR_WHITE
    End With
    
    ' Data rows with hyperlinks to source cells
    Dim r As Long: r = 5
    Dim i As Long
    For i = 0 To m_HitCount - 1
        ws.Cells(r, 1).Value = m_Hits(i).SheetName
        
        ' Cell reference as hyperlink
        ws.Hyperlinks.Add _
            Anchor:=ws.Cells(r, 2), _
            Address:="", _
            SubAddress:="'" & m_Hits(i).SheetName & "'!" & m_Hits(i).CellRef, _
            TextToDisplay:=m_Hits(i).CellRef
        ws.Cells(r, 2).Font.Color = RGB(31, 78, 121)
        
        ws.Cells(r, 3).Value = m_Hits(i).CellValue
        ws.Cells(r, 4).Value = m_Hits(i).Context
        
        ' Alternating row
        If r Mod 2 = 1 Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).Interior.Color = CLR_ALT_ROW
        End If
        
        r = r + 1
    Next i
    
    ' Cap warning row at bottom
    If m_TotalMatches > MAX_RESULTS Then
        ws.Cells(r + 1, 1).Value = "Results capped at " & MAX_RESULTS & _
            ". " & (m_TotalMatches - MAX_RESULTS) & " additional matches not shown."
        ws.Cells(r + 1, 1).Font.Italic = True
        ws.Cells(r + 1, 1).Font.Color = RGB(192, 0, 0)
    End If
    
    ws.Columns("A:D").AutoFit
    ws.Activate
End Sub
