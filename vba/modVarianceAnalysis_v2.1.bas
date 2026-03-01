Attribute VB_Name = "modVarianceAnalysis"
Option Explicit

'===============================================================================
' modVarianceAnalysis - Variance Detection, Reporting & Commentary
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Compare monthly summary sheets, flag items exceeding threshold,
'           and auto-generate English narratives for top variances.
'
' PUBLIC SUBS:
'   RunVarianceAnalysis   - MoM comparison of two functional P&L sheets
'   GenerateCommentary    - Auto-narrative top 5 FY-vs-Budget variances
'
' VERSION:  2.1.0
' CHANGES:  v2.0 -> v2.1:
'           + ISSUE-010: Added GenerateCommentary (Action #46)
'           + Added BuildNarrative private helper for contextual text
'           + FIX BUG-024: Alternating row color no longer overwrites flag
'
' PRIOR FIXES (v2.0):
'   BUG-001  usCol uses HDR_ROW_FUNC (row 4) not row 1
'   BUG-004  Loop starts at DATA_ROW_FUNC (row 5) not row 2
'===============================================================================

Private Const VAR_SHEET As String = "Variance Analysis"

'===============================================================================
' RunVarianceAnalysis - Compare two monthly summary sheets
'===============================================================================
Public Sub RunVarianceAnalysis()
    On Error GoTo ErrHandler
    
    ' Default: compare Jan vs Feb
    Dim sheet1 As String: sheet1 = SH_FUNC_JAN
    Dim sheet2 As String: sheet2 = SH_FUNC_FEB
    
    ' Validate both sheets exist
    If Not modConfig.SheetExists(sheet1) Or Not modConfig.SheetExists(sheet2) Then
        MsgBox "Required monthly summary sheets not found.", vbCritical, APP_NAME
        Exit Sub
    End If
    
    If MsgBox("Run variance analysis: " & sheet1 & " vs " & sheet2 & "?" & vbCrLf & _
              "Threshold: " & Format(VARIANCE_PCT, "0%") & " for flagging", _
              vbYesNo + vbQuestion, APP_NAME) = vbNo Then Exit Sub
    
    modPerformance.TurboOn
    modPerformance.UpdateStatus "Analyzing variances...", 0
    
    Dim ws1 As Worksheet: Set ws1 = ThisWorkbook.Worksheets(sheet1)
    Dim ws2 As Worksheet: Set ws2 = ThisWorkbook.Worksheets(sheet2)
    
    ' Build variance report sheet
    modConfig.SafeDeleteSheet VAR_SHEET
    
    Dim wsVar As Worksheet
    Set wsVar = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsVar.Name = VAR_SHEET
    wsVar.Tab.Color = RGB(0, 112, 192)
    
    ' Title
    wsVar.Range("A1").Value = "Month-over-Month Variance Analysis"
    wsVar.Range("A1").Font.Size = 14: wsVar.Range("A1").Font.Bold = True
    wsVar.Range("A2").Value = "Comparing: " & sheet1 & " -> " & sheet2
    wsVar.Range("A2").Font.Italic = True
    wsVar.Range("A3").Value = "Threshold: " & Format(VARIANCE_PCT, "0%") & _
        " | Generated: " & Format(Now, "yyyy-mm-dd hh:mm")
    
    ' Headers
    Dim headers As Variant
    headers = Array("Line Item", "Prior Month ($)", "Current Month ($)", _
                    "Variance ($)", "Variance (%)", "Status", "Flag")
    
    Dim c As Long
    For c = 0 To UBound(headers)
        wsVar.Cells(5, c + 1).Value = headers(c)
    Next c
    With wsVar.Range("A5:G5")
        .Font.Bold = True
        .Interior.Color = CLR_NAVY
        .Font.Color = CLR_WHITE
    End With
    
    ' Read the US/Consolidated column from both sheets
    Dim lastRow1 As Long: lastRow1 = modConfig.LastRow(ws1, 1)
    
    ' BUG-001 FIX: Use HDR_ROW_FUNC (row 4) to find last column
    Dim usCol As Long: usCol = modConfig.LastCol(ws1, HDR_ROW_FUNC)
    
    Dim row As Long: row = 6
    Dim flagCount As Long: flagCount = 0
    Dim r As Long
    
    ' BUG-004 FIX: Start at DATA_ROW_FUNC (row 5) instead of row 2
    For r = DATA_ROW_FUNC To lastRow1
        Dim label As String: label = Trim(CStr(ws1.Cells(r, 1).Value))
        If label = "" Then GoTo NextRow
        
        modPerformance.UpdateStatus "Analyzing " & label & "...", r / lastRow1
        
        ' Use SafeNum for robust numeric conversion
        Dim val1 As Double, val2 As Double
        val1 = modConfig.SafeNum(ws1.Cells(r, usCol).Value)
        val2 = modConfig.SafeNum(ws2.Cells(r, usCol).Value)
        
        Dim absDelta As Double: absDelta = val2 - val1
        Dim pctDelta As Double
        If val1 <> 0 Then
            pctDelta = absDelta / Abs(val1)
        Else
            pctDelta = 0
        End If
        
        ' Determine status
        Dim status As String, flagged As Boolean
        flagged = Abs(pctDelta) >= VARIANCE_PCT And val1 <> 0
        
        If absDelta > 0 Then
            status = "Favorable"
        ElseIf absDelta < 0 Then
            status = "Unfavorable"
        Else
            status = "Flat"
        End If
        
        ' Cost-line reversal: increase in cost = Unfavorable
        If InStr(1, label, "Cost", vbTextCompare) > 0 Or _
           InStr(1, label, "Expense", vbTextCompare) > 0 Or _
           InStr(1, label, "COGS", vbTextCompare) > 0 Or _
           InStr(1, label, "Depreciation", vbTextCompare) > 0 Or _
           InStr(1, label, "Amortization", vbTextCompare) > 0 Or _
           InStr(1, label, "Salary", vbTextCompare) > 0 Or _
           InStr(1, label, "Wages", vbTextCompare) > 0 Or _
           InStr(1, label, "Rent", vbTextCompare) > 0 Or _
           InStr(1, label, "AWS", vbTextCompare) > 0 Then
            If absDelta > 0 Then status = "Unfavorable"
            If absDelta < 0 Then status = "Favorable"
        End If
        
        ' Write row
        wsVar.Cells(row, 1).Value = label
        wsVar.Cells(row, 2).Value = val1
        wsVar.Cells(row, 3).Value = val2
        wsVar.Cells(row, 4).Value = absDelta
        wsVar.Cells(row, 5).Value = pctDelta
        wsVar.Cells(row, 6).Value = status
        wsVar.Cells(row, 7).Value = IIf(flagged, "FLAG", "")
        
        ' Format
        wsVar.Range(wsVar.Cells(row, 2), wsVar.Cells(row, 4)).NumberFormat = "$#,##0.00"
        wsVar.Cells(row, 5).NumberFormat = "0.0%"
        
        ' Color-code flagged rows
        If flagged Then
            wsVar.Range(wsVar.Cells(row, 1), wsVar.Cells(row, 7)).Interior.Color = RGB(255, 235, 156)
            wsVar.Cells(row, 7).Font.Color = RGB(200, 0, 0)
            wsVar.Cells(row, 7).Font.Bold = True
            flagCount = flagCount + 1
        ElseIf row Mod 2 = 0 Then
            ' BUG-024 FIX: Only apply alternating row color if NOT flagged
            wsVar.Range(wsVar.Cells(row, 1), wsVar.Cells(row, 7)).Interior.Color = CLR_ALT_ROW
        End If
        
        If status = "Favorable" Then
            wsVar.Cells(row, 6).Font.Color = RGB(0, 128, 0)
        ElseIf status = "Unfavorable" Then
            wsVar.Cells(row, 6).Font.Color = RGB(200, 0, 0)
        End If
        
        row = row + 1
NextRow:
    Next r
    
    wsVar.Columns("A:G").AutoFit
    wsVar.Activate
    
    modPerformance.TurboOff
    
    modLogger.LogAction "modVarianceAnalysis", "RunVarianceAnalysis", _
                        (row - 6) & " line items, " & flagCount & " flagged", _
                        modPerformance.ElapsedSeconds()
    
    MsgBox "Variance analysis complete!" & vbCrLf & _
           (row - 6) & " line items compared" & vbCrLf & _
           flagCount & " items flagged (>" & Format(VARIANCE_PCT, "0%") & " variance)", _
           vbInformation, APP_NAME
    Exit Sub
    
ErrHandler:
    modPerformance.TurboOff
    MsgBox "Variance analysis error: " & Err.Description, vbCritical, APP_NAME
End Sub


'===============================================================================
'
' ===  VARIANCE COMMENTARY (v2.1 — ISSUE-010)  ================================
'
'===============================================================================

'===============================================================================
' GenerateCommentary - Auto-generate English narratives for top variances
' Compares FY Total vs Budget on P&L Trend, ranks by absolute $ impact,
' generates contextual commentary for the top 5 variances.
' Ported from legacy T2 #11 GenerateVarianceCommentary.
'===============================================================================
Public Sub GenerateCommentary()
    On Error GoTo ErrHandler
    
    If Not modConfig.SheetExists(SH_PL_TREND) Then
        MsgBox "'" & SH_PL_TREND & "' not found.", vbCritical, APP_NAME
        Exit Sub
    End If
    
    modPerformance.TurboOn
    modPerformance.UpdateStatus "Scanning for variances...", 0.1
    
    Dim wsTrend As Worksheet: Set wsTrend = ThisWorkbook.Worksheets(SH_PL_TREND)
    Dim tLastRow As Long: tLastRow = wsTrend.Cells(wsTrend.Rows.Count, 1).End(xlUp).Row
    Dim tLastCol As Long: tLastCol = wsTrend.Cells(1, wsTrend.Columns.Count).End(xlToLeft).Column
    
    ' Find FY Total column
    Dim fyCol As Long: fyCol = 0
    Dim c As Long
    For c = 2 To tLastCol
        Dim hdr As String: hdr = LCase(Trim(CStr(wsTrend.Cells(1, c).Value)))
        If (InStr(hdr, "fy") > 0 And InStr(hdr, "total") > 0) Or _
           InStr(hdr, "fy" & FISCAL_YEAR_4) > 0 Or _
           InStr(hdr, FISCAL_YEAR_4 & " total") > 0 Then
            fyCol = c: Exit For
        End If
    Next c
    If fyCol = 0 Then fyCol = tLastCol
    
    ' Find Budget column
    Dim budCol As Long: budCol = 0
    For c = 2 To tLastCol
        hdr = LCase(Trim(CStr(wsTrend.Cells(1, c).Value)))
        If InStr(hdr, "budget") > 0 Then budCol = c: Exit For
    Next c
    If budCol = 0 Then budCol = tLastCol
    
    ' Collect all line items with non-zero variance
    Dim viMax As Long: viMax = tLastRow
    Dim viLineItem() As String: ReDim viLineItem(1 To viMax)
    Dim viActual() As Double:   ReDim viActual(1 To viMax)
    Dim viBudget() As Double:   ReDim viBudget(1 To viMax)
    Dim viVarDollar() As Double: ReDim viVarDollar(1 To viMax)
    Dim viVarPct() As Double:   ReDim viVarPct(1 To viMax)
    Dim viCount As Long: viCount = 0
    
    Dim r As Long
    For r = 2 To tLastRow
        Dim lineItem As String: lineItem = Trim(CStr(wsTrend.Cells(r, 1).Value))
        If lineItem = "" Then GoTo NextComRow
        
        Dim actVal As Double: actVal = modConfig.SafeNum(wsTrend.Cells(r, fyCol).Value)
        Dim budVal As Double: budVal = modConfig.SafeNum(wsTrend.Cells(r, budCol).Value)
        
        ' Only include items where both values exist and there's a variance
        If actVal = 0 And budVal = 0 Then GoTo NextComRow
        
        Dim varDollar As Double: varDollar = actVal - budVal
        If Abs(varDollar) < 1 Then GoTo NextComRow  ' Skip immaterial variances
        
        Dim varPct As Double
        If budVal <> 0 Then
            varPct = varDollar / Abs(budVal)
        Else
            varPct = 0
        End If
        
        viCount = viCount + 1
        viLineItem(viCount) = lineItem
        viActual(viCount) = actVal
        viBudget(viCount) = budVal
        viVarDollar(viCount) = varDollar
        viVarPct(viCount) = varPct
NextComRow:
    Next r
    
    If viCount = 0 Then
        modPerformance.TurboOff
        MsgBox "No material variances found between FY Actual and Budget.", _
               vbInformation, APP_NAME
        Exit Sub
    End If
    
    ' Sort by absolute $ variance (descending) — bubble sort
    Dim i As Long, j As Long
    For i = 1 To viCount - 1
        For j = i + 1 To viCount
            If Abs(viVarDollar(j)) > Abs(viVarDollar(i)) Then
                Dim tmpS As String, tmpD As Double
                tmpS = viLineItem(i): viLineItem(i) = viLineItem(j): viLineItem(j) = tmpS
                tmpD = viActual(i): viActual(i) = viActual(j): viActual(j) = tmpD
                tmpD = viBudget(i): viBudget(i) = viBudget(j): viBudget(j) = tmpD
                tmpD = viVarDollar(i): viVarDollar(i) = viVarDollar(j): viVarDollar(j) = tmpD
                tmpD = viVarPct(i): viVarPct(i) = viVarPct(j): viVarPct(j) = tmpD
            End If
        Next j
    Next i
    
    modPerformance.UpdateStatus "Generating commentary...", 0.5
    
    ' Create output sheet
    Dim comName As String: comName = "Variance Commentary"
    modConfig.SafeDeleteSheet comName
    
    Dim wsCom As Worksheet
    Set wsCom = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsCom.Name = comName
    
    wsCom.Range("A1").Value = "VARIANCE COMMENTARY - FY" & FISCAL_YEAR_4 & " Budget vs Actual"
    wsCom.Range("A1").Font.Bold = True
    wsCom.Range("A1").Font.Size = 14
    wsCom.Range("A1").Font.Color = CLR_NAVY

    modConfig.StyleHeader wsCom, 4, _
        Array("#", "Line Item", "FY Actual", "Budget", "Variance ($)", "Variance (%)", "Commentary")
    
    Dim topN As Long: topN = Application.Min(5, viCount)
    wsCom.Range("A3").Value = "Top " & topN & " variances by absolute dollar impact"
    wsCom.Range("A3").Font.Bold = True
    
    Dim outRow As Long: outRow = 5
    For i = 1 To topN
        wsCom.Cells(outRow, 1).Value = i
        wsCom.Cells(outRow, 1).HorizontalAlignment = xlCenter
        wsCom.Cells(outRow, 2).Value = viLineItem(i)
        wsCom.Cells(outRow, 2).Font.Bold = True
        wsCom.Cells(outRow, 3).Value = viActual(i)
        wsCom.Cells(outRow, 3).NumberFormat = "$#,##0"
        wsCom.Cells(outRow, 4).Value = viBudget(i)
        wsCom.Cells(outRow, 4).NumberFormat = "$#,##0"
        wsCom.Cells(outRow, 5).Value = viVarDollar(i)
        wsCom.Cells(outRow, 5).NumberFormat = "$#,##0;($#,##0)"
        wsCom.Cells(outRow, 6).Value = viVarPct(i)
        wsCom.Cells(outRow, 6).NumberFormat = "0.0%"
        
        ' Color variance
        If viVarDollar(i) > 0 Then
            wsCom.Cells(outRow, 5).Font.Color = RGB(0, 128, 0)
            wsCom.Cells(outRow, 6).Font.Color = RGB(0, 128, 0)
        Else
            wsCom.Cells(outRow, 5).Font.Color = RGB(192, 0, 0)
            wsCom.Cells(outRow, 6).Font.Color = RGB(192, 0, 0)
        End If
        
        ' Generate narrative
        wsCom.Cells(outRow, 7).Value = BuildNarrative(viLineItem(i), viVarDollar(i), viVarPct(i))
        wsCom.Cells(outRow, 7).WrapText = True
        
        outRow = outRow + 1
    Next i
    
    ' Format column widths and row heights
    wsCom.Columns("A").ColumnWidth = 4
    wsCom.Columns("B").ColumnWidth = 30
    wsCom.Columns("C").ColumnWidth = 14
    wsCom.Columns("D").ColumnWidth = 14
    wsCom.Columns("E").ColumnWidth = 14
    wsCom.Columns("F").ColumnWidth = 12
    wsCom.Columns("G").ColumnWidth = 70
    
    ' Add borders and row height for commentary readability
    If topN > 0 Then
        With wsCom.Range("A4:G" & 4 + topN).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(191, 191, 191)
        End With
        For i = 5 To 4 + topN
            wsCom.Rows(i).RowHeight = 60
        Next i
    End If
    
    wsCom.Tab.Color = RGB(0, 112, 192)
    wsCom.Activate
    
    Dim elapsed As Double: elapsed = modPerformance.ElapsedSeconds()
    modPerformance.TurboOff
    
    modLogger.LogAction "modVarianceAnalysis", "GenerateCommentary", _
        topN & " of " & viCount & " variances detailed", elapsed
    
    MsgBox "Variance Commentary Generated!" & vbCrLf & vbCrLf & _
           "  Variances analyzed: " & viCount & vbCrLf & _
           "  Top " & topN & " detailed on '" & comName & "'" & vbCrLf & _
           "  Source: " & SH_PL_TREND & " (FY Total vs Budget)", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modVarianceAnalysis", "ERROR-Commentary", Err.Description
    MsgBox "Commentary error: " & Err.Description, vbCritical, APP_NAME
End Sub


'===============================================================================
'
' ===  PRIVATE HELPERS  ========================================================
'
'===============================================================================

'===============================================================================
' BuildNarrative - Generate contextual English narrative for a variance
' Classifies the line item as Revenue, Expense, or Subtotal and generates
' directional commentary with suggested investigation areas.
'===============================================================================
Private Function BuildNarrative(ByVal lineItem As String, _
    ByVal varDollar As Double, ByVal varPct As Double) As String
    
    Dim direction As String
    If varDollar > 0 Then direction = "favorable" Else direction = "unfavorable"
    
    ' Classify line item for contextual commentary
    Dim lineLC As String: lineLC = LCase(lineItem)
    Dim category As String
    If InStr(lineLC, "revenue") > 0 Or InStr(lineLC, "sales") > 0 Then
        category = "Revenue"
    ElseIf InStr(lineLC, "total") > 0 Or InStr(lineLC, "net") > 0 Or _
           InStr(lineLC, "gross profit") > 0 Or InStr(lineLC, "margin") > 0 Then
        category = "Subtotal"
    Else
        category = "Expense"
    End If
    
    ' Build the base narrative
    Dim narrative As String
    narrative = lineItem & " shows a " & direction & " variance of " & _
        Format(Abs(varDollar), "$#,##0") & " (" & Format(Abs(varPct), "0.0%") & "). "
    
    ' Add context-specific commentary
    Select Case category
        Case "Revenue"
            If varDollar > 0 Then
                narrative = narrative & "Revenue exceeded plan, suggesting stronger-than-expected " & _
                    "demand or pricing uplift. Review by product line to identify drivers."
            Else
                narrative = narrative & "Revenue shortfall vs budget warrants investigation into " & _
                    "volume vs pricing drivers. Check product mix and customer retention."
            End If
        Case "Expense"
            If varDollar > 0 Then
                narrative = narrative & "Spending exceeded budget. Investigate whether timing-related " & _
                    "(pulled forward) or structural. Review vendor contracts and headcount."
            Else
                narrative = narrative & "Favorable to budget. Validate genuine efficiency vs " & _
                    "deferred spend that may catch up in later periods."
            End If
        Case "Subtotal"
            narrative = narrative & "This subtotal variance flows from line items above. " & _
                "See individual drivers for root cause."
    End Select
    
    BuildNarrative = narrative
End Function
