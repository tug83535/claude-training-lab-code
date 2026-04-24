Attribute VB_Name = "modUTL_Branding"
Option Explicit

' ============================================================
' KBT Universal Tools — iPipeline Branding Module
' Works on ANY Excel file — no project-specific setup required
' Install in Personal.xlsb to use across all Excel sessions
' Tools: 2 | Tier 1: 2
' ============================================================
' iPipeline Brand Colors (from docs/ipipeline-brand-styling.md):
'   iPipeline Blue:   #0B4779  RGB(11,  71,  121) — Primary
'   Navy Blue:        #112E51  RGB(17,  46,  81)  — Secondary
'   Innovation Blue:  #4B9BCB  RGB(75,  155, 203) — Secondary
'   Lime Green:       #BFF18C  RGB(191, 241, 140) — Accent
'   Aqua:             #2BCCD3  RGB(43,  204, 211) — Accent
'   Arctic White:     #F9F9F9  RGB(249, 249, 249) — Neutral
'   Charcoal:         #161616  RGB(22,  22,  22)  — Neutral
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
' TOOL 1 — Apply iPipeline Branding                [TIER 1]
' Styles headers, alternating rows, and totals rows on the
' active sheet using official iPipeline brand colors and fonts.
' Detects header row automatically (first non-empty row with
' text in 3+ columns).
' ============================================================
Sub ApplyiPipelineBranding()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    If MsgBox("Apply iPipeline brand formatting to sheet '" & ws.Name & "'?" & Chr(10) & Chr(10) & _
              "This will style:" & Chr(10) & _
              "  - Header row: iPipeline Blue background, white text, Arial Bold" & Chr(10) & _
              "  - Data rows: alternating white/light gray" & Chr(10) & _
              "  - Total/Summary rows: Navy Blue background, white text" & Chr(10) & Chr(10) & _
              "Existing cell values will NOT be changed.", _
              vbQuestion + vbYesNo, "UTL iPipeline Branding") = vbNo Then Exit Sub

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.UsedRange.Columns.Count

    If lastRow < 2 Or lastCol < 1 Then
        UTL_TurboOff
        MsgBox "Sheet appears empty. No formatting applied.", vbInformation, "UTL iPipeline Branding"
        Exit Sub
    End If

    ' Find the last used column more reliably
    Dim r As Long
    For r = 1 To Application.Min(lastRow, 10)
        Dim rowLastCol As Long
        rowLastCol = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        If rowLastCol > lastCol Then lastCol = rowLastCol
    Next r

    ' --- Detect header row (first row with 3+ non-empty cells) ---
    Dim headerRow As Long
    headerRow = 0
    For r = 1 To Application.Min(lastRow, 10)
        Dim filledCells As Long
        filledCells = 0
        Dim c As Long
        For c = 1 To lastCol
            If Trim(CStr(ws.Cells(r, c).Value)) <> "" Then
                filledCells = filledCells + 1
            End If
        Next c
        If filledCells >= 3 Then
            headerRow = r
            Exit For
        End If
    Next r

    If headerRow = 0 Then headerRow = 1

    ' --- Style header row: iPipeline Blue bg, Arctic White text, Arial Bold ---
    Dim hdrRange As Range
    Set hdrRange = ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, lastCol))
    With hdrRange
        .Interior.Color = RGB(11, 71, 121)
        .Font.Color = RGB(249, 249, 249)
        .Font.Bold = True
        .Font.Name = "Arial"
        .Font.Size = 11
    End With

    ' --- Style title rows above header (if any) ---
    If headerRow > 1 Then
        Dim titleR As Long
        For titleR = 1 To headerRow - 1
            With ws.Range(ws.Cells(titleR, 1), ws.Cells(titleR, lastCol))
                .Font.Name = "Arial"
                .Font.Bold = True
                .Font.Color = RGB(17, 46, 81)
            End With
        Next titleR
    End If

    ' --- Style data rows: alternating Arctic White / light gray ---
    Dim dataStartRow As Long
    dataStartRow = headerRow + 1
    Dim totalRowsStyled As Long
    Dim totalRowsFound As Long

    For r = dataStartRow To lastRow
        Dim rowRange As Range
        Set rowRange = ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))

        ' Check if this is a totals/summary row
        Dim cellA As String
        cellA = LCase(Trim(CStr(ws.Cells(r, 1).Value)))
        Dim isTotalRow As Boolean
        isTotalRow = (InStr(cellA, "total") > 0 Or _
                      InStr(cellA, "grand total") > 0 Or _
                      InStr(cellA, "net income") > 0 Or _
                      InStr(cellA, "net revenue") > 0 Or _
                      InStr(cellA, "summary") > 0)

        If isTotalRow Then
            ' Total rows: Navy Blue bg, white text, bold
            With rowRange
                .Interior.Color = RGB(17, 46, 81)
                .Font.Color = RGB(249, 249, 249)
                .Font.Bold = True
                .Font.Name = "Arial"
            End With
            totalRowsFound = totalRowsFound + 1
        Else
            ' Alternating rows
            If (r - dataStartRow) Mod 2 = 0 Then
                rowRange.Interior.Color = RGB(249, 249, 249)
            Else
                rowRange.Interior.Color = RGB(240, 240, 238)
            End If
            rowRange.Font.Name = "Arial"
            rowRange.Font.Color = RGB(22, 22, 22)
        End If
        totalRowsStyled = totalRowsStyled + 1
    Next r

    ' Auto-fit columns
    ws.UsedRange.Columns.AutoFit

    UTL_TurboOff
    MsgBox "iPipeline branding applied!" & Chr(10) & Chr(10) & _
           "Header row " & headerRow & " styled (iPipeline Blue)." & Chr(10) & _
           totalRowsStyled & " data rows formatted." & Chr(10) & _
           totalRowsFound & " total/summary row(s) styled (Navy Blue)." & Chr(10) & Chr(10) & _
           "Font: Arial | Colors: Official iPipeline brand palette.", _
           vbInformation, "UTL iPipeline Branding"
    Exit Sub

ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL iPipeline Branding"
End Sub

' ============================================================
' TOOL 2 — Set iPipeline Theme Colors              [TIER 1]
' Overwrites the workbook's theme color palette so iPipeline
' brand colors appear in the standard Excel color picker.
' After running this, every color dropdown (Font Color, Fill,
' Shape Fill, etc.) will show iPipeline brand colors in the
' "Theme Colors" section at the top.
' ============================================================
Sub SetiPipelineThemeColors()
    If MsgBox("Set this workbook's theme colors to the iPipeline brand palette?" & Chr(10) & Chr(10) & _
              "After this, the color picker will show:" & Chr(10) & _
              "  iPipeline Blue, Navy, Innovation Blue," & Chr(10) & _
              "  Lime Green, Aqua, Arctic White, Charcoal" & Chr(10) & Chr(10) & _
              "This only affects THIS workbook (not other files).", _
              vbQuestion + vbYesNo, "UTL iPipeline Branding") = vbNo Then Exit Sub

    On Error GoTo FallbackMethod

    ' Theme color slots (msoThemeColor enum values):
    '   1 = Dark 1,  2 = Light 1,  3 = Dark 2,   4 = Light 2
    '   5 = Accent1, 6 = Accent2,  7 = Accent3,  8 = Accent4
    '   9 = Accent5, 10 = Accent6, 11 = Hyperlink, 12 = Followed Hyperlink
    With ActiveWorkbook.Theme.ThemeColorScheme
        .Colors(1).RGB = RGB(22, 22, 22)         ' Dark 1:   Charcoal
        .Colors(2).RGB = RGB(249, 249, 249)      ' Light 1:  Arctic White
        .Colors(3).RGB = RGB(17, 46, 81)         ' Dark 2:   Navy Blue
        .Colors(4).RGB = RGB(75, 155, 203)       ' Light 2:  Innovation Blue
        .Colors(5).RGB = RGB(11, 71, 121)        ' Accent 1: iPipeline Blue (Primary)
        .Colors(6).RGB = RGB(43, 204, 211)       ' Accent 2: Aqua
        .Colors(7).RGB = RGB(191, 241, 140)      ' Accent 3: Lime Green
        .Colors(8).RGB = RGB(75, 155, 203)       ' Accent 4: Innovation Blue (repeat)
        .Colors(9).RGB = RGB(17, 46, 81)         ' Accent 5: Navy Blue (repeat)
        .Colors(10).RGB = RGB(11, 71, 121)       ' Accent 6: iPipeline Blue (repeat)
        .Colors(11).RGB = RGB(75, 155, 203)      ' Hyperlink: Innovation Blue
        .Colors(12).RGB = RGB(43, 204, 211)      ' Followed:  Aqua
    End With

    MsgBox "iPipeline theme colors applied!" & Chr(10) & Chr(10) & _
           "Open any color picker (Font Color, Fill, etc.)" & Chr(10) & _
           "and you will see iPipeline brand colors in the" & Chr(10) & _
           "'Theme Colors' section at the top." & Chr(10) & Chr(10) & _
           "This only affects this workbook.", _
           vbInformation, "UTL iPipeline Branding"
    Exit Sub

FallbackMethod:
    ' If the theme method is not supported, use the legacy 56-color palette
    On Error GoTo ErrHandler
    With ActiveWorkbook
        .Colors(17) = RGB(11, 71, 121)    ' iPipeline Blue
        .Colors(18) = RGB(17, 46, 81)     ' Navy Blue
        .Colors(19) = RGB(75, 155, 203)   ' Innovation Blue
        .Colors(20) = RGB(191, 241, 140)  ' Lime Green
        .Colors(21) = RGB(43, 204, 211)   ' Aqua
        .Colors(22) = RGB(249, 249, 249)  ' Arctic White
        .Colors(23) = RGB(22, 22, 22)     ' Charcoal
    End With

    MsgBox "iPipeline colors added to the workbook palette!" & Chr(10) & Chr(10) & _
           "Your Excel version does not support theme color changes," & Chr(10) & _
           "but the iPipeline brand colors have been added to the" & Chr(10) & _
           "workbook's custom color palette. You can access them" & Chr(10) & _
           "through Format Cells > Custom Colors.", _
           vbInformation, "UTL iPipeline Branding"
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "UTL iPipeline Branding"
End Sub
