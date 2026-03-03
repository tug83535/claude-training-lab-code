Attribute VB_Name = "modUTL_Finance"
Option Explicit

' ============================================================
' KBT Universal Tools — Finance Module
' Works on ANY Excel file — no project-specific setup required
' Install in Personal.xlsb to use across all Excel sessions
' Tools: 14 | Tier 1: 2 | Tier 2: 12
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
' TOOL 1 — Duplicate Invoice Detector                [TIER 1]
' Scans for potential duplicate invoices
' Flags matches on: Vendor + Amount + Date (within 3 days)
' Run: active sheet must have headers — tool asks which columns
' ============================================================
Sub DuplicateInvoiceDetector()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim vendorCol As String
    Dim amtCol As String
    Dim dateCol As String
    Dim invCol As String

    vendorCol = InputBox("Which column has Vendor Name? (e.g. A)", "UTL Finance")
    If vendorCol = "" Then Exit Sub
    amtCol = InputBox("Which column has Amount? (e.g. B)", "UTL Finance")
    If amtCol = "" Then Exit Sub
    dateCol = InputBox("Which column has Date? (e.g. C)", "UTL Finance")
    If dateCol = "" Then Exit Sub
    invCol = InputBox("Which column has Invoice Number? (e.g. D)" & Chr(10) & _
                      "(leave blank to skip invoice number matching)", "UTL Finance")

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, vendorCol).End(xlUp).Row

    Dim flagged As Long
    Dim i As Long, j As Long

    For i = 2 To lastRow
        For j = i + 1 To lastRow
            Dim vendorMatch As Boolean
            Dim amtMatch As Boolean
            Dim dateMatch As Boolean
            Dim invMatch As Boolean

            vendorMatch = (LCase(Trim(ws.Cells(i, vendorCol).Value)) = _
                           LCase(Trim(ws.Cells(j, vendorCol).Value)))
            amtMatch = (ws.Cells(i, amtCol).Value = ws.Cells(j, amtCol).Value)

            Dim d1 As Date, d2 As Date
            If IsDate(ws.Cells(i, dateCol).Value) And IsDate(ws.Cells(j, dateCol).Value) Then
                d1 = CDate(ws.Cells(i, dateCol).Value)
                d2 = CDate(ws.Cells(j, dateCol).Value)
                dateMatch = (Abs(d1 - d2) <= 3)
            Else
                dateMatch = False
            End If

            If invCol <> "" Then
                invMatch = (Trim(ws.Cells(i, invCol).Value) = Trim(ws.Cells(j, invCol).Value)) _
                           And ws.Cells(i, invCol).Value <> ""
            Else
                invMatch = True
            End If

            If vendorMatch And amtMatch And (dateMatch Or invMatch) Then
                ws.Rows(i).Interior.Color = RGB(255, 200, 100)
                ws.Rows(j).Interior.Color = RGB(255, 200, 100)
                flagged = flagged + 1
            End If
        Next j
    Next i

    UTL_TurboOff
    If flagged = 0 Then
        MsgBox "No duplicate invoices detected.", vbInformation, "UTL Finance"
    Else
        MsgBox flagged & " potential duplicate pair(s) flagged in orange." & Chr(10) & _
               "Review highlighted rows before processing payment.", vbExclamation, "UTL Finance"
    End If
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Finance"
End Sub

' ============================================================
' TOOL 2 — Auto-Balancing GL Validator               [TIER 1]
' Sums debit and credit columns, flags imbalance
' Optionally inserts a balancing plug entry
' ============================================================
Sub AutoBalancingGLValidator()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim debitCol As String
    Dim creditCol As String

    debitCol = InputBox("Which column has Debits? (e.g. C)", "UTL Finance")
    If debitCol = "" Then Exit Sub
    creditCol = InputBox("Which column has Credits? (e.g. D)", "UTL Finance")
    If creditCol = "" Then Exit Sub

    On Error GoTo ErrHandler

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, debitCol).End(xlUp).Row

    Dim totalDebits As Double
    Dim totalCredits As Double
    Dim i As Long

    For i = 2 To lastRow
        If IsNumeric(ws.Cells(i, debitCol).Value) Then
            totalDebits = totalDebits + CDbl(ws.Cells(i, debitCol).Value)
        End If
        If IsNumeric(ws.Cells(i, creditCol).Value) Then
            totalCredits = totalCredits + CDbl(ws.Cells(i, creditCol).Value)
        End If
    Next i

    Dim diff As Double
    diff = Round(totalDebits - totalCredits, 2)

    If Abs(diff) < 0.01 Then
        MsgBox "BALANCED — Debits and Credits match." & Chr(10) & Chr(10) & _
               "Total Debits:  $" & Format(totalDebits, "#,##0.00") & Chr(10) & _
               "Total Credits: $" & Format(totalCredits, "#,##0.00"), _
               vbInformation, "UTL Finance — GL Validator"
    Else
        Dim plugChoice As Integer
        plugChoice = MsgBox("OUT OF BALANCE — Difference: $" & Format(Abs(diff), "#,##0.00") & Chr(10) & Chr(10) & _
                            "Total Debits:  $" & Format(totalDebits, "#,##0.00") & Chr(10) & _
                            "Total Credits: $" & Format(totalCredits, "#,##0.00") & Chr(10) & Chr(10) & _
                            "Click YES to insert a balancing plug entry." & Chr(10) & _
                            "Click NO to just flag the imbalance.", _
                            vbExclamation + vbYesNo, "UTL Finance — GL Validator")

        If plugChoice = vbYes Then
            Dim plugRow As Long
            plugRow = lastRow + 1
            ws.Cells(plugRow, debitCol).Offset(0, -1).Value = "*** PLUG ENTRY — REVIEW REQUIRED ***"
            If diff > 0 Then
                ws.Cells(plugRow, creditCol).Value = diff
            Else
                ws.Cells(plugRow, debitCol).Value = Abs(diff)
            End If
            ws.Rows(plugRow).Interior.Color = RGB(255, 100, 100)
            MsgBox "Plug entry inserted in row " & plugRow & " — highlighted red." & Chr(10) & _
                   "Investigate and replace with the correct entry.", vbExclamation, "UTL Finance"
        End If
    End If
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Finance"
End Sub

' ============================================================
' TOOL 3 — Trial Balance Checker                     [TIER 2]
' Verifies total debits equal total credits
' Highlights the imbalance amount and affected rows
' ============================================================
Sub TrialBalanceChecker()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim debitCol As String
    Dim creditCol As String
    debitCol = InputBox("Which column has Debit balances? (e.g. B)", "UTL Finance")
    If debitCol = "" Then Exit Sub
    creditCol = InputBox("Which column has Credit balances? (e.g. C)", "UTL Finance")
    If creditCol = "" Then Exit Sub

    On Error GoTo ErrHandler

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, debitCol).End(xlUp).Row

    Dim sumDebits As Double
    Dim sumCredits As Double
    Dim i As Long
    For i = 2 To lastRow
        If IsNumeric(ws.Cells(i, debitCol).Value) Then sumDebits = sumDebits + ws.Cells(i, debitCol).Value
        If IsNumeric(ws.Cells(i, creditCol).Value) Then sumCredits = sumCredits + ws.Cells(i, creditCol).Value
    Next i

    Dim diff As Double
    diff = Round(sumDebits - sumCredits, 2)

    Dim msg As String
    msg = "TRIAL BALANCE REPORT" & Chr(10) & String(30, "-") & Chr(10) & _
          "Total Debits:   $" & Format(sumDebits, "#,##0.00") & Chr(10) & _
          "Total Credits:  $" & Format(sumCredits, "#,##0.00") & Chr(10) & _
          "Difference:     $" & Format(Abs(diff), "#,##0.00") & Chr(10) & Chr(10)

    If Abs(diff) < 0.01 Then
        msg = msg & "STATUS: BALANCED"
        MsgBox msg, vbInformation, "UTL Finance"
    Else
        msg = msg & "STATUS: OUT OF BALANCE — Investigate immediately."
        MsgBox msg, vbExclamation, "UTL Finance"
    End If
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Finance"
End Sub

' ============================================================
' TOOL 4 — Journal Entry Validator                   [TIER 2]
' Groups journal entries by entry number, checks each balances
' Flags any entry where debits don't equal credits
' ============================================================
Sub JournalEntryValidator()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim jeCol As String
    Dim debitCol As String
    Dim creditCol As String
    jeCol = InputBox("Which column has the Journal Entry number? (e.g. A)", "UTL Finance")
    If jeCol = "" Then Exit Sub
    debitCol = InputBox("Which column has Debits? (e.g. C)", "UTL Finance")
    If debitCol = "" Then Exit Sub
    creditCol = InputBox("Which column has Credits? (e.g. D)", "UTL Finance")
    If creditCol = "" Then Exit Sub

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, jeCol).End(xlUp).Row

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 2 To lastRow
        Dim jeNum As String
        jeNum = CStr(ws.Cells(i, jeCol).Value)
        If jeNum <> "" Then
            Dim dr As Double
            Dim cr As Double
            If IsNumeric(ws.Cells(i, debitCol).Value) Then dr = CDbl(ws.Cells(i, debitCol).Value)
            If IsNumeric(ws.Cells(i, creditCol).Value) Then cr = CDbl(ws.Cells(i, creditCol).Value)
            If Not dict.exists(jeNum) Then
                dict.Add jeNum, Array(dr - cr, i)
            Else
                Dim arr As Variant
                arr = dict(jeNum)
                arr(0) = arr(0) + dr - cr
                dict(jeNum) = arr
            End If
        End If
    Next i

    Dim unbalanced As Long
    Dim key As Variant
    For Each key In dict.Keys
        Dim balance As Double
        balance = Round(dict(key)(0), 2)
        If Abs(balance) > 0.01 Then
            unbalanced = unbalanced + 1
        End If
    Next key

    UTL_TurboOff
    Dim result As String
    result = "JOURNAL ENTRY VALIDATION" & Chr(10) & String(30, "-") & Chr(10) & _
             "Total Unique JE Numbers: " & dict.Count & Chr(10) & _
             "Unbalanced Entries:      " & unbalanced & Chr(10) & Chr(10)
    If unbalanced = 0 Then
        result = result & "STATUS: ALL ENTRIES BALANCED"
        MsgBox result, vbInformation, "UTL Finance"
    Else
        result = result & "STATUS: " & unbalanced & " UNBALANCED ENTRIES — Review required."
        MsgBox result, vbExclamation, "UTL Finance"
    End If
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Finance"
End Sub

' ============================================================
' TOOL 5 — Flux Analysis / Period Comparison         [TIER 2]
' Compares two columns (Actual vs Prior Period) row by row
' Flags lines where the change exceeds a threshold you set
' ============================================================
Sub FluxAnalysis()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim currCol As String
    Dim priorCol As String
    Dim threshPct As String

    currCol = InputBox("Which column has Current Period values? (e.g. B)", "UTL Finance")
    If currCol = "" Then Exit Sub
    priorCol = InputBox("Which column has Prior Period values? (e.g. C)", "UTL Finance")
    If priorCol = "" Then Exit Sub
    threshPct = InputBox("Flag rows where change exceeds what % (e.g. 10 for 10%):", "UTL Finance", "10")
    If threshPct = "" Then Exit Sub

    Dim threshold As Double
    threshold = CDbl(threshPct) / 100

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, currCol).End(xlUp).Row

    ' Add variance columns
    Dim varDolCol As Long
    Dim varPctCol As Long
    varDolCol = ws.Cells(1, currCol).Column + 1
    varPctCol = varDolCol + 1

    ws.Cells(1, varDolCol).Value = "$ Variance"
    ws.Cells(1, varPctCol).Value = "% Variance"
    ws.Cells(1, varDolCol).Font.Bold = True
    ws.Cells(1, varPctCol).Font.Bold = True

    Dim flagged As Long
    Dim i As Long
    For i = 2 To lastRow
        Dim curr As Double
        Dim prior As Double
        If IsNumeric(ws.Cells(i, currCol).Value) And IsNumeric(ws.Cells(i, priorCol).Value) Then
            curr = CDbl(ws.Cells(i, currCol).Value)
            prior = CDbl(ws.Cells(i, priorCol).Value)
            Dim dolVar As Double
            dolVar = curr - prior
            ws.Cells(i, varDolCol).Value = dolVar
            ws.Cells(i, varDolCol).NumberFormat = "#,##0.00;[Red](#,##0.00)"
            If prior <> 0 Then
                Dim pctVar As Double
                pctVar = dolVar / Abs(prior)
                ws.Cells(i, varPctCol).Value = pctVar
                ws.Cells(i, varPctCol).NumberFormat = "0.0%"
                If Abs(pctVar) > threshold Then
                    ws.Rows(i).Interior.Color = RGB(255, 235, 59)
                    flagged = flagged + 1
                End If
            Else
                ws.Cells(i, varPctCol).Value = "N/A"
            End If
        End If
    Next i

    UTL_TurboOff
    MsgBox "Done! Flux analysis complete." & Chr(10) & _
           flagged & " row(s) flagged (change > " & threshPct & "%) highlighted yellow.", _
           vbInformation, "UTL Finance"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Finance"
End Sub

' ============================================================
' TOOL 6 — AP Aging Summary Generator               [TIER 2]
' Buckets AP invoices by days overdue from a due date column
' Creates a new summary tab with bucket totals
' ============================================================
Sub APAgingSummaryGenerator()
    Call UTL_AgingGenerator("AP")
End Sub

' ============================================================
' TOOL 7 — AR Aging Summary Generator               [TIER 2]
' Buckets AR invoices by days outstanding from invoice date
' Creates a new summary tab with bucket totals
' ============================================================
Sub ARAgingSummaryGenerator()
    Call UTL_AgingGenerator("AR")
End Sub

Private Sub UTL_AgingGenerator(agingType As String)
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim dateCol As String
    Dim amtCol As String
    Dim vendorCol As String

    dateCol = InputBox("Which column has the " & IIf(agingType = "AP", "Due Date", "Invoice Date") & "? (e.g. B)", "UTL Finance")
    If dateCol = "" Then Exit Sub
    amtCol = InputBox("Which column has the Amount? (e.g. C)", "UTL Finance")
    If amtCol = "" Then Exit Sub
    vendorCol = InputBox("Which column has the " & IIf(agingType = "AP", "Vendor", "Customer") & " Name? (e.g. A)", "UTL Finance")
    If vendorCol = "" Then Exit Sub

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim today As Date
    today = Date

    Dim b0_30 As Double, b31_60 As Double, b61_90 As Double, b91plus As Double, bCurrent As Double
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, dateCol).End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        If IsDate(ws.Cells(i, dateCol).Value) And IsNumeric(ws.Cells(i, amtCol).Value) Then
            Dim daysOld As Long
            daysOld = today - CDate(ws.Cells(i, dateCol).Value)
            Dim amt As Double
            amt = CDbl(ws.Cells(i, amtCol).Value)

            Select Case daysOld
                Case Is < 0:      bCurrent = bCurrent + amt
                Case 0 To 30:     b0_30 = b0_30 + amt
                Case 31 To 60:    b31_60 = b31_60 + amt
                Case 61 To 90:    b61_90 = b61_90 + amt
                Case Else:        b91plus = b91plus + amt
            End Select
        End If
    Next i

    ' Write summary to new sheet
    Dim summaryWS As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Sheets(agingType & " Aging").Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    Set summaryWS = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    summaryWS.Name = agingType & " Aging"

    With summaryWS
        .Range("A1").Value = agingType & " Aging Summary — " & Format(today, "MM/DD/YYYY")
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A3").Value = "Bucket"
        .Range("B3").Value = "Total Amount"
        .Range("A3:B3").Font.Bold = True
        .Range("A3:B3").Interior.Color = RGB(31, 73, 125)
        .Range("A3:B3").Font.Color = RGB(255, 255, 255)
        .Range("A4").Value = "Current (Not Yet Due)"
        .Range("A5").Value = "0 - 30 Days"
        .Range("A6").Value = "31 - 60 Days"
        .Range("A7").Value = "61 - 90 Days"
        .Range("A8").Value = "90+ Days"
        .Range("A9").Value = "TOTAL"
        .Range("A9").Font.Bold = True
        .Range("B4").Value = bCurrent
        .Range("B5").Value = b0_30
        .Range("B6").Value = b31_60
        .Range("B7").Value = b61_90
        .Range("B8").Value = b91plus
        .Range("B9").Value = bCurrent + b0_30 + b31_60 + b61_90 + b91plus
        .Range("B4:B9").NumberFormat = "$#,##0.00"
        .Range("B8").Interior.Color = IIf(b91plus > 0, RGB(255, 200, 100), RGB(200, 255, 200))
        .Columns("A:B").AutoFit
    End With

    UTL_TurboOff
    summaryWS.Activate
    MsgBox "Done! " & agingType & " Aging Summary created.", vbInformation, "UTL Finance"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Finance"
End Sub

' ============================================================
' TOOL 8 — Aging Bucket Calculator                   [TIER 2]
' Adds a Bucket column to the active sheet based on any date
' Works for any aging scenario, not just AP/AR
' ============================================================
Sub AgingBucketCalculator()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim dateCol As String
    dateCol = InputBox("Which column has the date to age from? (e.g. B)", "UTL Finance")
    If dateCol = "" Then Exit Sub

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, dateCol).End(xlUp).Row

    Dim bucketCol As Long
    bucketCol = ws.Cells(1, dateCol).Column + 1
    ws.Cells(1, bucketCol).Value = "Aging Bucket"
    ws.Cells(1, bucketCol).Font.Bold = True

    Dim today As Date
    today = Date
    Dim i As Long

    For i = 2 To lastRow
        If IsDate(ws.Cells(i, dateCol).Value) Then
            Dim days As Long
            days = today - CDate(ws.Cells(i, dateCol).Value)
            Dim bucket As String
            Select Case days
                Case Is < 0:   bucket = "Current"
                Case 0 To 30:  bucket = "0-30 Days"
                Case 31 To 60: bucket = "31-60 Days"
                Case 61 To 90: bucket = "61-90 Days"
                Case Else:     bucket = "90+ Days"
            End Select
            ws.Cells(i, bucketCol).Value = bucket
        End If
    Next i

    UTL_TurboOff
    MsgBox "Done! Aging Bucket column added.", vbInformation, "UTL Finance"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Finance"
End Sub

' ============================================================
' TOOL 9 — Variance Analysis Template               [TIER 2]
' Adds $ Variance and % Variance columns next to Actual/Budget
' Run: tell it which columns are Actual and Budget
' ============================================================
Sub VarianceAnalysisTemplate()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim actualCol As String
    Dim budgetCol As String
    actualCol = InputBox("Which column has Actual values? (e.g. B)", "UTL Finance")
    If actualCol = "" Then Exit Sub
    budgetCol = InputBox("Which column has Budget values? (e.g. C)", "UTL Finance")
    If budgetCol = "" Then Exit Sub

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, actualCol).End(xlUp).Row

    Dim varDolCol As Long
    Dim varPctCol As Long
    varDolCol = ws.Cells(1, budgetCol).Column + 1
    varPctCol = varDolCol + 1

    ws.Columns(varDolCol).Insert
    ws.Columns(varPctCol).Insert

    ws.Cells(1, varDolCol).Value = "$ Variance (Act-Bud)"
    ws.Cells(1, varPctCol).Value = "% Variance"
    ws.Cells(1, varDolCol).Font.Bold = True
    ws.Cells(1, varPctCol).Font.Bold = True

    Dim i As Long
    For i = 2 To lastRow
        If IsNumeric(ws.Cells(i, actualCol).Value) And IsNumeric(ws.Cells(i, budgetCol).Value) Then
            Dim act As Double
            Dim bud As Double
            act = CDbl(ws.Cells(i, actualCol).Value)
            bud = CDbl(ws.Cells(i, budgetCol).Value)
            ws.Cells(i, varDolCol).Value = act - bud
            ws.Cells(i, varDolCol).NumberFormat = "#,##0.00;[Red](#,##0.00)"
            If bud <> 0 Then
                ws.Cells(i, varPctCol).Value = (act - bud) / Abs(bud)
                ws.Cells(i, varPctCol).NumberFormat = "0.0%"
            Else
                ws.Cells(i, varPctCol).Value = "N/A"
            End If
        End If
    Next i

    UTL_TurboOff
    MsgBox "Done! Variance columns added.", vbInformation, "UTL Finance"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Finance"
End Sub

' ============================================================
' TOOL 10 — Quick Corkscrew Builder                 [TIER 2]
' Builds a standard roll-forward schedule for any balance
' Beginning Balance + Additions - Deductions = Ending Balance
' ============================================================
Sub QuickCorkscrewBuilder()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim labelName As String
    labelName = InputBox("What is this corkscrew for?" & Chr(10) & _
                         "(e.g. Fixed Assets, AR Reserve, Deferred Revenue)", "UTL Finance", "Balance")
    If labelName = "" Then Exit Sub

    Dim begBal As Double
    Dim begBalStr As String
    begBalStr = InputBox("Enter the beginning balance:", "UTL Finance", "0")
    If begBalStr = "" Then Exit Sub
    If Not IsNumeric(begBalStr) Then
        MsgBox "Please enter a number.", vbExclamation, "UTL Finance"
        Exit Sub
    End If
    begBal = CDbl(begBalStr)

    On Error GoTo ErrHandler

    Dim corkWS As Worksheet
    Set corkWS = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    corkWS.Name = Left(labelName, 28) & " Roll"

    With corkWS
        .Range("A1").Value = labelName & " — Roll-Forward Schedule"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A3").Value = "Beginning Balance"
        .Range("A4").Value = "  + Additions"
        .Range("A5").Value = "  - Deductions / Write-offs"
        .Range("A6").Value = "  +/- Adjustments"
        .Range("A7").Value = "Ending Balance"

        .Range("B3").Value = begBal
        .Range("B4").Value = 0
        .Range("B5").Value = 0
        .Range("B6").Value = 0
        .Range("B7").Formula = "=B3+B4-B5+B6"

        .Range("A3:B7").NumberFormat = "$#,##0.00"
        .Range("A3").Font.Bold = True
        .Range("A7").Font.Bold = True
        .Range("B7").Font.Bold = True
        .Range("A7:B7").Interior.Color = RGB(31, 73, 125)
        .Range("A7:B7").Font.Color = RGB(255, 255, 255)

        .Range("A3:A7").ColumnWidth = 30
        .Columns("B").AutoFit

        .Range("C3").Value = "← Enter beginning balance"
        .Range("C4").Value = "← Enter additions (purchases, accruals)"
        .Range("C5").Value = "← Enter deductions (disposals, write-offs)"
        .Range("C6").Value = "← Enter net adjustments"
        .Range("C7").Value = "← Calculated automatically"
        .Range("C3:C7").Font.Italic = True
        .Range("C3:C7").Font.Color = RGB(128, 128, 128)
        .Columns("C").AutoFit
    End With

    corkWS.Activate
    MsgBox "Done! Corkscrew schedule created on new sheet '" & corkWS.Name & "'.", _
           vbInformation, "UTL Finance"
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Finance"
End Sub

' ============================================================
' TOOL 11 — Financial Period Roll-Forward           [TIER 2]
' Updates month-end header dates across the active sheet
' Clears designated input cells to prepare for new period data
' ============================================================
Sub FinancialPeriodRollForward()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim newPeriod As String
    newPeriod = InputBox("Enter the new period label (e.g. April 2026):", "UTL Finance")
    If newPeriod = "" Then Exit Sub

    Dim headerRow As String
    headerRow = InputBox("Which row has the period headers? (e.g. 1)", "UTL Finance", "1")
    If headerRow = "" Then Exit Sub

    Dim oldPeriod As String
    oldPeriod = InputBox("What is the OLD period label to replace? (e.g. March 2026)" & Chr(10) & _
                         "(leave blank to skip header replacement)", "UTL Finance")

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim hRow As Long
    hRow = CLng(headerRow)
    Dim headerReplaced As Long

    ' Replace header labels
    If oldPeriod <> "" Then
        Dim c As Range
        For Each c In ws.Rows(hRow).Cells
            If InStr(1, CStr(c.Value), oldPeriod, vbTextCompare) > 0 Then
                c.Value = Replace(c.Value, oldPeriod, newPeriod, 1, -1, vbTextCompare)
                headerReplaced = headerReplaced + 1
            End If
        Next c
    End If

    ' Clear input cells (non-formula cells in used range below header)
    Dim clearedCount As Long
    If MsgBox("Clear all input (non-formula) cells below row " & hRow & " to prepare for new period data?", _
              vbQuestion + vbYesNo, "UTL Finance") = vbYes Then
        Dim inputCell As Range
        For Each inputCell In ws.UsedRange
            If inputCell.Row > hRow Then
                If Not inputCell.HasFormula And IsNumeric(inputCell.Value) Then
                    inputCell.ClearContents
                    clearedCount = clearedCount + 1
                End If
            End If
        Next inputCell
    End If

    UTL_TurboOff
    MsgBox "Done! Period roll-forward complete." & Chr(10) & _
           headerReplaced & " header(s) updated to '" & newPeriod & "'." & Chr(10) & _
           clearedCount & " input cell(s) cleared.", vbInformation, "UTL Finance"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Finance"
End Sub

' ============================================================
' TOOL 12 — Multi-Currency Consolidation Aggregator [TIER 2]
' Consolidates currency amounts using FX rates from a rate table
' Rate table: Currency code in col A, Rate to USD in col B
' ============================================================
Sub MultiCurrencyConsolidationAggregator()
    Dim rateSheetName As String
    rateSheetName = InputBox("Enter the name of your FX Rate sheet:" & Chr(10) & _
                             "(Must have: Currency Code in col A, Rate to USD in col B)", _
                             "UTL Finance", "FX Rates")
    If rateSheetName = "" Then Exit Sub

    Dim rateWS As Worksheet
    On Error Resume Next
    Set rateWS = ActiveWorkbook.Sheets(rateSheetName)
    On Error GoTo ErrHandler
    If rateWS Is Nothing Then
        MsgBox "Sheet '" & rateSheetName & "' not found.", vbExclamation, "UTL Finance"
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim currencyCol As String
    Dim amtCol As String
    currencyCol = InputBox("Which column has the Currency code? (e.g. B)", "UTL Finance")
    If currencyCol = "" Then Exit Sub
    amtCol = InputBox("Which column has the Amount? (e.g. C)", "UTL Finance")
    If amtCol = "" Then Exit Sub

    On Error GoTo ErrHandler
    UTL_TurboOn

    ' Build FX rate dictionary
    Dim fxDict As Object
    Set fxDict = CreateObject("Scripting.Dictionary")
    Dim rateLastRow As Long
    rateLastRow = rateWS.Cells(rateWS.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To rateLastRow
        Dim ccy As String
        ccy = UCase(Trim(rateWS.Cells(r, 1).Value))
        If ccy <> "" And IsNumeric(rateWS.Cells(r, 2).Value) Then
            fxDict(ccy) = CDbl(rateWS.Cells(r, 2).Value)
        End If
    Next r

    ' Add USD equivalent column
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, amtCol).End(xlUp).Row
    Dim usdCol As Long
    usdCol = ws.Cells(1, amtCol).Column + 1
    ws.Cells(1, usdCol).Value = "USD Equivalent"
    ws.Cells(1, usdCol).Font.Bold = True

    Dim converted As Long
    Dim notFound As Long
    Dim i As Long
    For i = 2 To lastRow
        Dim ccyCode As String
        ccyCode = UCase(Trim(ws.Cells(i, currencyCol).Value))
        If fxDict.exists(ccyCode) And IsNumeric(ws.Cells(i, amtCol).Value) Then
            ws.Cells(i, usdCol).Value = CDbl(ws.Cells(i, amtCol).Value) * fxDict(ccyCode)
            ws.Cells(i, usdCol).NumberFormat = "$#,##0.00"
            converted = converted + 1
        Else
            ws.Cells(i, usdCol).Value = "Rate not found"
            notFound = notFound + 1
        End If
    Next i

    UTL_TurboOff
    MsgBox "Done! FX conversion complete." & Chr(10) & _
           converted & " amounts converted to USD." & Chr(10) & _
           notFound & " row(s) had no matching rate.", vbInformation, "UTL Finance"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Finance"
End Sub

' ============================================================
' TOOL 13 — Ratio Analysis Dashboard Builder        [TIER 2]
' Calculates key financial ratios from a sheet with labeled rows
' Looks for standard labels: Revenue, Net Income, Total Assets, etc.
' ============================================================
Sub RatioAnalysisDashboard()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    On Error GoTo ErrHandler
    UTL_TurboOn

    ' Search for key values by label
    Dim revenue As Double, netIncome As Double
    Dim grossProfit As Double, totalAssets As Double
    Dim currentAssets As Double, currentLiabilities As Double
    Dim totalEquity As Double, ebitda As Double

    Dim c As Range
    For Each c In ws.UsedRange.Columns(1).Cells
        Dim lbl As String
        lbl = LCase(Trim(CStr(c.Value)))
        Dim valCell As Range
        Set valCell = c.Offset(0, 1)
        If IsNumeric(valCell.Value) Then
            Select Case True
                Case InStr(lbl, "revenue") > 0 Or InStr(lbl, "net sales") > 0: revenue = CDbl(valCell.Value)
                Case InStr(lbl, "net income") > 0 Or InStr(lbl, "net profit") > 0: netIncome = CDbl(valCell.Value)
                Case InStr(lbl, "gross profit") > 0: grossProfit = CDbl(valCell.Value)
                Case InStr(lbl, "total assets") > 0: totalAssets = CDbl(valCell.Value)
                Case InStr(lbl, "current assets") > 0: currentAssets = CDbl(valCell.Value)
                Case InStr(lbl, "current liabilities") > 0: currentLiabilities = CDbl(valCell.Value)
                Case InStr(lbl, "equity") > 0 Or InStr(lbl, "stockholders") > 0: totalEquity = CDbl(valCell.Value)
                Case InStr(lbl, "ebitda") > 0: ebitda = CDbl(valCell.Value)
            End Select
        End If
    Next c

    ' Build ratio sheet
    Dim ratioWS As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Sheets("UTL Ratios").Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    Set ratioWS = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    ratioWS.Name = "UTL Ratios"

    With ratioWS
        .Range("A1").Value = "Financial Ratio Dashboard"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A2").Value = "Generated: " & Format(Now, "MM/DD/YYYY")
        .Range("A4").Value = "Ratio"
        .Range("B4").Value = "Value"
        .Range("C4").Value = "Benchmark"
        .Range("A4:C4").Font.Bold = True
        .Range("A4:C4").Interior.Color = RGB(31, 73, 125)
        .Range("A4:C4").Font.Color = RGB(255, 255, 255)

        Dim row As Long
        row = 5

        If revenue <> 0 Then
            .Cells(row, 1).Value = "Gross Margin"
            .Cells(row, 2).Value = IIf(grossProfit <> 0, grossProfit / revenue, "N/A")
            .Cells(row, 3).Value = "> 40% is strong"
            If grossProfit <> 0 Then .Cells(row, 2).NumberFormat = "0.0%"
            row = row + 1

            .Cells(row, 1).Value = "Net Profit Margin"
            .Cells(row, 2).Value = IIf(netIncome <> 0, netIncome / revenue, "N/A")
            .Cells(row, 3).Value = "> 10% is healthy"
            If netIncome <> 0 Then .Cells(row, 2).NumberFormat = "0.0%"
            row = row + 1

            If ebitda <> 0 Then
                .Cells(row, 1).Value = "EBITDA Margin"
                .Cells(row, 2).Value = ebitda / revenue
                .Cells(row, 2).NumberFormat = "0.0%"
                .Cells(row, 3).Value = "> 15% is healthy"
                row = row + 1
            End If
        End If

        If currentLiabilities <> 0 And currentAssets <> 0 Then
            .Cells(row, 1).Value = "Current Ratio"
            .Cells(row, 2).Value = currentAssets / currentLiabilities
            .Cells(row, 2).NumberFormat = "0.00x"
            .Cells(row, 3).Value = "> 1.5 is healthy"
            row = row + 1
        End If

        If totalEquity <> 0 And netIncome <> 0 Then
            .Cells(row, 1).Value = "Return on Equity (ROE)"
            .Cells(row, 2).Value = netIncome / totalEquity
            .Cells(row, 2).NumberFormat = "0.0%"
            .Cells(row, 3).Value = "> 15% is strong"
            row = row + 1
        End If

        If totalAssets <> 0 And netIncome <> 0 Then
            .Cells(row, 1).Value = "Return on Assets (ROA)"
            .Cells(row, 2).Value = netIncome / totalAssets
            .Cells(row, 2).NumberFormat = "0.0%"
            .Cells(row, 3).Value = "> 5% is healthy"
        End If

        .Columns("A:C").AutoFit
    End With

    UTL_TurboOff
    ratioWS.Activate
    MsgBox "Done! Ratio dashboard created." & Chr(10) & _
           "Note: Only ratios with matching labels found were calculated.", _
           vbInformation, "UTL Finance"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Finance"
End Sub

' ============================================================
' TOOL 14 — General Ledger Journal Mapper           [TIER 2]
' Transforms a raw trial balance into a journal entry upload template
' Output format: Date | Account | Description | Debit | Credit
' ============================================================
Sub GeneralLedgerJournalMapper()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim acctCol As String
    Dim descCol As String
    Dim balCol As String

    acctCol = InputBox("Which column has Account Number/Code? (e.g. A)", "UTL Finance")
    If acctCol = "" Then Exit Sub
    descCol = InputBox("Which column has Account Description? (e.g. B)", "UTL Finance")
    If descCol = "" Then Exit Sub
    balCol = InputBox("Which column has Balance? (positive=Debit, negative=Credit) (e.g. C)", "UTL Finance")
    If balCol = "" Then Exit Sub

    On Error GoTo ErrHandler
    UTL_TurboOn

    Dim jeDate As String
    jeDate = InputBox("Enter the journal entry date (MM/DD/YYYY):", "UTL Finance", Format(Date, "MM/DD/YYYY"))
    If jeDate = "" Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, acctCol).End(xlUp).Row

    ' Create JE template sheet
    Dim jeWS As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Sheets("JE Upload Template").Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    Set jeWS = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    jeWS.Name = "JE Upload Template"

    With jeWS
        .Range("A1").Value = "Date"
        .Range("B1").Value = "Account"
        .Range("C1").Value = "Description"
        .Range("D1").Value = "Debit"
        .Range("E1").Value = "Credit"
        .Range("A1:E1").Font.Bold = True
        .Range("A1:E1").Interior.Color = RGB(31, 73, 125)
        .Range("A1:E1").Font.Color = RGB(255, 255, 255)
    End With

    Dim jeRow As Long
    jeRow = 2
    Dim i As Long
    For i = 2 To lastRow
        Dim bal As Double
        If IsNumeric(ws.Cells(i, balCol).Value) And ws.Cells(i, acctCol).Value <> "" Then
            bal = CDbl(ws.Cells(i, balCol).Value)
            jeWS.Cells(jeRow, 1).Value = CDate(jeDate)
            jeWS.Cells(jeRow, 1).NumberFormat = "MM/DD/YYYY"
            jeWS.Cells(jeRow, 2).Value = ws.Cells(i, acctCol).Value
            jeWS.Cells(jeRow, 3).Value = ws.Cells(i, descCol).Value
            If bal >= 0 Then
                jeWS.Cells(jeRow, 4).Value = bal
                jeWS.Cells(jeRow, 4).NumberFormat = "$#,##0.00"
            Else
                jeWS.Cells(jeRow, 5).Value = Abs(bal)
                jeWS.Cells(jeRow, 5).NumberFormat = "$#,##0.00"
            End If
            jeRow = jeRow + 1
        End If
    Next i

    jeWS.Columns("A:E").AutoFit
    UTL_TurboOff
    jeWS.Activate
    MsgBox "Done! Journal entry upload template created with " & (jeRow - 2) & " entries.", _
           vbInformation, "UTL Finance"
    Exit Sub
ErrHandler:
    UTL_TurboOff
    MsgBox "Error: " & Err.Description, vbCritical, "UTL Finance"
End Sub
