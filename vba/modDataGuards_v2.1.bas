Attribute VB_Name = "modDataGuards"
Option Explicit

'===============================================================================
' modDataGuards - Data Validation Safety Checks
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Pre-run guards that catch data problems BEFORE they silently
'           corrupt calculations. Run these before any refresh cycle.
'
' PUBLIC SUBS / FUNCTIONS:
'   ValidateAssumptionsPresence  - Block refresh if key drivers are blank (#48)
'   CheckSumOfDrivers            - Validate rev share %s sum to 100% (#49)
'   FindNegativeAmounts          - Flag negative amounts in GL data (#150)
'   FindZeroAmounts              - Flag zero amounts in GL data (#151)
'   FindSuspiciousRoundNumbers   - Flag amounts that are multiples of $1,000 (#155)
'
' VERSION:  2.1.0 (New module — 2026-03-01)
' SOURCE:   Ideas from NewTesting/VBA Examples (200) — items #48, #49, #150, #151, #155
'===============================================================================

'===============================================================================
' ValidateAssumptionsPresence - Block if any key driver cell is blank (#48)
' Returns True if all drivers are present (safe to proceed), False if not.
' Call this at the top of any refresh macro before doing any work.
' Shows a blocking MsgBox listing every blank driver found.
'===============================================================================
Public Function ValidateAssumptionsPresence() As Boolean
    ValidateAssumptionsPresence = True

    If Not modConfig.SheetExists(SH_ASSUMPTIONS) Then
        MsgBox "Assumptions sheet '" & SH_ASSUMPTIONS & "' not found.", _
               vbCritical, APP_NAME
        ValidateAssumptionsPresence = False
        Exit Function
    End If

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_ASSUMPTIONS)
    Dim lastRow As Long: lastRow = modConfig.LastRow(ws, 1)

    Dim missing As Long: missing = 0
    Dim missingList As String
    Dim r As Long
    For r = DATA_ROW_ASSUME To lastRow
        ' Skip section headers and label-only rows (they have no Type in column C)
        If Len(modConfig.SafeStr(ws.Cells(r, 3).Value)) = 0 Then GoTo NextAssRow
        Dim driverName As String: driverName = modConfig.SafeStr(ws.Cells(r, 1).Value)
        If Len(driverName) > 0 Then
            If IsEmpty(ws.Cells(r, 2).Value) Or ws.Cells(r, 2).Value = "" Then
                missing = missing + 1
                missingList = missingList & vbCrLf & "  Row " & r & ": " & driverName
            End If
        End If
NextAssRow:
    Next r

    If missing > 0 Then
        MsgBox "BLOCKED: " & missing & " assumption driver(s) are blank." & _
               vbCrLf & missingList & vbCrLf & vbCrLf & _
               "Fill in all values before running a refresh.", _
               vbCritical, APP_NAME
        ValidateAssumptionsPresence = False
    Else
        MsgBox "All assumption drivers are present. Safe to proceed.", _
               vbInformation, APP_NAME
    End If
End Function

'===============================================================================
' CheckSumOfDrivers - Validate revenue share percentages sum to 100% (#49)
' Reads rows whose label contains "rev share" or "revenue share" from the
' Assumptions sheet and confirms they total exactly 1.00 (within 0.1%).
' Falls back to searching for product names if no labeled rows are found.
'===============================================================================
Public Sub CheckSumOfDrivers()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_ASSUMPTIONS) Then
        MsgBox "Assumptions sheet not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_ASSUMPTIONS)
    Dim lastRow As Long: lastRow = modConfig.LastRow(ws, 1)

    Dim revShareSum   As Double: revShareSum   = 0
    Dim revShareCount As Long:   revShareCount = 0
    Dim itemList      As String

    Dim r As Long
    For r = DATA_ROW_ASSUME To lastRow
        Dim lbl As String: lbl = LCase(modConfig.SafeStr(ws.Cells(r, 1).Value))
        If InStr(lbl, "rev share") > 0 Or InStr(lbl, "revenue share") > 0 Then
            Dim val As Double: val = modConfig.SafeNum(ws.Cells(r, 2).Value)
            revShareSum   = revShareSum + val
            revShareCount = revShareCount + 1
            itemList = itemList & vbCrLf & "  " & ws.Cells(r, 1).Value & _
                       ": " & Format(val, "0.0%")
        End If
    Next r

    ' Fall back: look for product name rows if no "rev share" labels found
    If revShareCount = 0 Then
        Dim products As Variant: products = modConfig.GetProducts()
        Dim p As Long
        For p = 0 To UBound(products)
            Dim pRow As Long
            pRow = modConfig.FindRowByLabel(ws, CStr(products(p)), DATA_ROW_ASSUME)
            If pRow > 0 Then
                val           = modConfig.SafeNum(ws.Cells(pRow, 2).Value)
                revShareSum   = revShareSum + val
                revShareCount = revShareCount + 1
                itemList      = itemList & vbCrLf & "  " & CStr(products(p)) & _
                                ": " & Format(val, "0.0%")
            End If
        Next p
    End If

    If revShareCount = 0 Then
        MsgBox "No revenue share drivers found on Assumptions sheet." & vbCrLf & _
               "Label driver rows with 'Rev Share' in column A.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim diff   As Double: diff   = Abs(revShareSum - 1)
    Dim status As String
    If diff < 0.001 Then
        status = "PASS — Revenue shares sum to 100%."
    Else
        status = "FAIL — Revenue shares sum to " & Format(revShareSum, "0.00%") & _
                 " (off by " & Format(diff, "0.000%") & ")."
    End If

    modLogger.LogAction "modDataGuards", "CheckSumOfDrivers", _
        status & " | " & revShareCount & " drivers | Sum=" & Format(revShareSum, "0.0000")
    MsgBox status & vbCrLf & vbCrLf & _
           revShareCount & " driver(s) checked:" & itemList, _
           IIf(diff < 0.001, vbInformation, vbCritical), APP_NAME
    Exit Sub

ErrHandler:
    MsgBox "CheckSumOfDrivers error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' FindNegativeAmounts - Flag GL rows where Amount < 0 (#150)
' Highlights matching Amount cells red and shows a count.
' Negative amounts in an expense GL are almost always data entry errors.
'===============================================================================
Public Sub FindNegativeAmounts()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_HIDDEN) Then
        MsgBox "GL sheet '" & SH_HIDDEN & "' not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_HIDDEN)
    ws.Visible = xlSheetVisible
    Dim lastRow As Long: lastRow = modConfig.LastRow(ws, COL_GL_ID)

    ' Clear prior highlight on Amount column only
    ws.Range(ws.Cells(DATA_ROW_GL, COL_GL_AMOUNT), _
             ws.Cells(lastRow, COL_GL_AMOUNT)).Interior.ColorIndex = xlNone

    Dim negCount As Long: negCount = 0
    Dim r As Long
    For r = DATA_ROW_GL To lastRow
        If modConfig.SafeNum(ws.Cells(r, COL_GL_AMOUNT).Value) < 0 Then
            ws.Cells(r, COL_GL_AMOUNT).Interior.Color = RGB(255, 200, 200)
            negCount = negCount + 1
        End If
    Next r

    modLogger.LogAction "modDataGuards", "FindNegativeAmounts", _
        negCount & " negative amount(s) in " & SH_HIDDEN
    If negCount > 0 Then
        MsgBox negCount & " row(s) with negative amounts found and highlighted red." & vbCrLf & _
               "Review these rows — negative amounts in expense data are usually data errors.", _
               vbExclamation, APP_NAME
    Else
        MsgBox "No negative amounts found. Data looks clean.", vbInformation, APP_NAME
    End If
    Exit Sub

ErrHandler:
    MsgBox "FindNegativeAmounts error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' FindZeroAmounts - Flag GL rows where Amount = 0 (#151)
' Highlights matching Amount cells yellow. Zero-amount rows usually indicate
' a missing value or a formula that returned nothing.
'===============================================================================
Public Sub FindZeroAmounts()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_HIDDEN) Then
        MsgBox "GL sheet '" & SH_HIDDEN & "' not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_HIDDEN)
    ws.Visible = xlSheetVisible
    Dim lastRow As Long: lastRow = modConfig.LastRow(ws, COL_GL_ID)

    ws.Range(ws.Cells(DATA_ROW_GL, COL_GL_AMOUNT), _
             ws.Cells(lastRow, COL_GL_AMOUNT)).Interior.ColorIndex = xlNone

    Dim zeroCount As Long: zeroCount = 0
    Dim r As Long
    For r = DATA_ROW_GL To lastRow
        Dim rawVal As Variant: rawVal = ws.Cells(r, COL_GL_AMOUNT).Value
        If Not IsEmpty(rawVal) And rawVal <> "" Then
            If modConfig.SafeNum(rawVal) = 0 Then
                ws.Cells(r, COL_GL_AMOUNT).Interior.Color = RGB(255, 255, 180)
                zeroCount = zeroCount + 1
            End If
        End If
    Next r

    modLogger.LogAction "modDataGuards", "FindZeroAmounts", _
        zeroCount & " zero-amount row(s) in " & SH_HIDDEN
    If zeroCount > 0 Then
        MsgBox zeroCount & " row(s) with zero amounts found and highlighted yellow." & vbCrLf & _
               "Zero amounts in expense data usually indicate missing or failed values.", _
               vbExclamation, APP_NAME
    Else
        MsgBox "No zero amounts found.", vbInformation, APP_NAME
    End If
    Exit Sub

ErrHandler:
    MsgBox "FindZeroAmounts error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' FindSuspiciousRoundNumbers - Flag GL amounts that are exact multiples of 1000 (#155)
' Large round numbers (e.g., $5,000, $12,000) often indicate estimates or
' placeholders rather than real transaction data. Highlights them orange.
' Only flags amounts >= $1,000 to avoid noise from small round values.
'===============================================================================
Public Sub FindSuspiciousRoundNumbers()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_HIDDEN) Then
        MsgBox "GL sheet '" & SH_HIDDEN & "' not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_HIDDEN)
    ws.Visible = xlSheetVisible
    Dim lastRow As Long: lastRow = modConfig.LastRow(ws, COL_GL_ID)

    ws.Range(ws.Cells(DATA_ROW_GL, COL_GL_AMOUNT), _
             ws.Cells(lastRow, COL_GL_AMOUNT)).Interior.ColorIndex = xlNone

    Dim roundCount As Long: roundCount = 0
    Dim r As Long
    For r = DATA_ROW_GL To lastRow
        Dim amt As Double: amt = modConfig.SafeNum(ws.Cells(r, COL_GL_AMOUNT).Value)
        If amt >= 1000 Then
            If (amt Mod 1000) = 0 Then
                ws.Cells(r, COL_GL_AMOUNT).Interior.Color = RGB(255, 235, 200)
                roundCount = roundCount + 1
            End If
        End If
    Next r

    modLogger.LogAction "modDataGuards", "FindSuspiciousRoundNumbers", _
        roundCount & " suspicious round amount(s) in " & SH_HIDDEN
    If roundCount > 0 Then
        MsgBox roundCount & " row(s) with suspicious round amounts (multiples of $1,000) highlighted orange." & vbCrLf & vbCrLf & _
               "These may be estimates or placeholders. Review before presenting to the CFO.", _
               vbExclamation, APP_NAME
    Else
        MsgBox "No suspicious round numbers found above $1,000.", vbInformation, APP_NAME
    End If
    Exit Sub

ErrHandler:
    MsgBox "FindSuspiciousRoundNumbers error: " & Err.Description, vbCritical, APP_NAME
End Sub
