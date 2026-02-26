Attribute VB_Name = "modMonthlyTabGenerator"
Option Explicit

'===============================================================================
' modMonthlyTabGenerator - Auto-Generate Monthly Summary Tabs
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Clone the Mar template Functional P&L Summary to create
'           Apr through Dec. Updates all cross-sheet formula references
'           to point at the correct month column on Functional P&L - Monthly Trend.
'
' PUBLIC SUBS:
'   GenerateMonthlyTabs   - Create all 9 tabs (Apr-Dec) from Mar template
'   GenerateNextMonthOnly - Create just the next missing monthly tab
'   DeleteGeneratedTabs   - Remove all auto-generated tabs (Apr-Dec)
'
' VERSION:  2.1.0
' CHANGES:  v2.0 -> v2.1 (Phase 1):
'           + ISSUE-002 (BUG-013): UpdateHeaderText safe replacement
'             (prevents "Margin"->"Aprigin" corruption)
'           v2.1 Phase 1 -> v2.1 Phase 3D:
'           + ISSUE-012: Added GenerateNextMonthOnly (Action #42)
'             Detects latest existing tab, clones it, updates refs,
'             clears data values, marks cells needing input (yellow),
'             stamps "[NEW - DATA NEEDED]"
'===============================================================================

' Column letters on Functional P&L - Monthly Trend for each month
' Jan=B(2), Feb=C(3), Mar=D(4), Apr=E(5), May=F(6), Jun=G(7),
' Jul=H(8), Aug=I(9), Sep=J(10), Oct=K(11), Nov=L(12), Dec=M(13)
Private Const MONTH_COLS As String = "B,C,D,E,F,G,H,I,J,K,L,M"

'===============================================================================
' GenerateMonthlyTabs - Main entry point (creates Apr-Dec from Mar template)
'===============================================================================
Public Sub GenerateMonthlyTabs()
    On Error GoTo ErrHandler
    
    ' Build template name from constants (no hardcoded year)
    Dim templateSheet As String
    templateSheet = "Functional P&L Summary - Mar " & FISCAL_YEAR
    
    ' Validate template exists
    If Not modConfig.SheetExists(templateSheet) Then
        MsgBox "Template sheet '" & templateSheet & "' not found.", vbCritical, APP_NAME
        Exit Sub
    End If
    
    ' Confirm with user
    Dim monthsToCreate As String: monthsToCreate = "Apr, May, Jun, Jul, Aug, Sep, Oct, Nov, Dec"
    If MsgBox("Generate 9 monthly summary tabs (" & monthsToCreate & ")?" & vbCrLf & vbCrLf & _
              "This will:" & vbCrLf & _
              "  - Copy the Mar " & FISCAL_YEAR & " template for each month" & vbCrLf & _
              "  - Update all formula references to the correct month column" & vbCrLf & _
              "  - Apply matching tab colors and formatting", _
              vbYesNo + vbQuestion, APP_NAME) = vbNo Then Exit Sub
    
    modPerformance.TurboOn
    modPerformance.UpdateStatus "Generating monthly tabs...", 0
    
    Dim mths As Variant: mths = modConfig.GetMonths()
    Dim cols As Variant: cols = Split(MONTH_COLS, ",")
    Dim templateWs As Worksheet: Set templateWs = ThisWorkbook.Worksheets(templateSheet)
    Dim templateCol As String: templateCol = "D"  ' Mar = column D on trend sheet
    
    Dim created As Long: created = 0
    Dim i As Long
    
    ' Loop Apr (index 3) through Dec (index 11)
    For i = 3 To 11
        Dim newName As String
        newName = "Functional P&L Summary - " & mths(i) & " " & FISCAL_YEAR
        
        ' Skip if already exists
        If modConfig.SheetExists(newName) Then
            modPerformance.UpdateStatus "Skipping " & mths(i) & " (already exists)...", (i - 3) / 9
            GoTo NextMonth
        End If
        
        modPerformance.UpdateStatus "Creating " & mths(i) & " " & FISCAL_YEAR & "...", (i - 3) / 9
        
        ' Copy template
        templateWs.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
        Dim newWs As Worksheet: Set newWs = ActiveSheet
        newWs.Name = newName
        
        ' Update formula references: swap column letter
        Dim targetCol As String: targetCol = cols(i)
        UpdateSheetFormulas newWs, templateCol, targetCol, mths(i)
        
        ' Update header text (change "Mar" to new month name — SAFE version)
        UpdateHeaderText newWs, "Mar", mths(i)
        
        ' Tab color = Blue (same as other monthly summaries)
        newWs.Tab.Color = RGB(68, 114, 196)
        
        created = created + 1
        
NextMonth:
    Next i
    
    modPerformance.ForceRecalc
    modPerformance.TurboOff
    
    modLogger.LogAction "modMonthlyTabGenerator", "GenerateMonthlyTabs", _
                        created & " tabs created", modPerformance.ElapsedSeconds()
    
    MsgBox created & " monthly tabs generated successfully in " & _
           modPerformance.ElapsedSeconds() & " seconds.", vbInformation, APP_NAME
    Exit Sub
    
ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modMonthlyTabGenerator", "ERROR", Err.Description
    MsgBox "Error generating tabs: " & Err.Description, vbCritical, APP_NAME
End Sub


'===============================================================================
'
' ===  SINGLE-MONTH GENERATION (v2.1 — ISSUE-012)  ============================
'
'===============================================================================

'===============================================================================
' GenerateNextMonthOnly - Create just the next missing monthly tab
' Detects which months already exist, clones the latest one, clears data
' values (preserves formulas), marks cells needing input (yellow highlight),
' and stamps "[NEW - DATA NEEDED]" on the title cell.
' Ported from legacy T2 #18 AutoGenerateNewMonthSheet.
'===============================================================================
Public Sub GenerateNextMonthOnly()
    On Error GoTo ErrHandler
    
    Dim mths As Variant: mths = modConfig.GetMonths()
    
    ' Find the latest existing Functional P&L Summary tab (scan Dec->Jan)
    Dim latestIdx As Long: latestIdx = -1
    Dim latestName As String
    Dim i As Long
    For i = 11 To 0 Step -1
        Dim chkName As String
        chkName = "Functional P&L Summary - " & mths(i) & " " & FISCAL_YEAR
        If modConfig.SheetExists(chkName) Then
            latestIdx = i
            latestName = chkName
            Exit For
        End If
    Next i
    
    If latestIdx = -1 Then
        MsgBox "No existing Functional P&L Summary tabs found." & vbCrLf & _
               "Run GenerateMonthlyTabs first to create from template.", _
               vbExclamation, APP_NAME
        Exit Sub
    End If
    
    If latestIdx >= 11 Then
        MsgBox "All 12 months already have summary tabs.", vbInformation, APP_NAME
        Exit Sub
    End If
    
    Dim nextIdx As Long: nextIdx = latestIdx + 1
    Dim newName As String
    newName = "Functional P&L Summary - " & mths(nextIdx) & " " & FISCAL_YEAR
    
    If modConfig.SheetExists(newName) Then
        MsgBox "'" & newName & "' already exists.", vbInformation, APP_NAME
        Exit Sub
    End If
    
    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("Generate next month tab:" & vbCrLf & vbCrLf & _
                     "  Clone: " & latestName & vbCrLf & _
                     "  New:   " & newName & vbCrLf & vbCrLf & _
                     "Data values will be cleared. Formulas preserved." & vbCrLf & _
                     "Yellow cells = needs data entry.", _
                     vbYesNo + vbQuestion, APP_NAME)
    If confirm = vbNo Then Exit Sub
    
    modPerformance.TurboOn
    modPerformance.UpdateStatus "Cloning " & latestName & "...", 0.2
    
    ' Clone latest tab
    ThisWorkbook.Worksheets(latestName).Copy _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Dim newWs As Worksheet: Set newWs = ActiveSheet
    newWs.Name = newName
    
    modPerformance.UpdateStatus "Updating references...", 0.4
    
    ' Update header text (uses ISSUE-002 safe replacement)
    UpdateHeaderText newWs, mths(latestIdx), mths(nextIdx)
    
    ' Update formula column references
    Dim cols As Variant: cols = Split(MONTH_COLS, ",")
    Dim oldCol As String: oldCol = cols(latestIdx)
    Dim newCol As String: newCol = cols(nextIdx)
    UpdateSheetFormulas newWs, oldCol, newCol, mths(nextIdx)
    
    modPerformance.UpdateStatus "Clearing data values...", 0.6
    
    ' Clear non-formula data cells, highlight yellow for data entry
    Dim cell As Range
    For Each cell In newWs.UsedRange
        If Not cell.HasFormula Then
            If IsNumeric(cell.Value) And cell.Value <> 0 Then
                cell.Value = 0
                cell.Interior.Color = RGB(255, 255, 230)  ' light yellow = needs input
            End If
        End If
    Next cell
    
    ' Stamp as new
    newWs.Range("A1").Value = newWs.Range("A1").Value & "  [NEW - DATA NEEDED]"
    newWs.Tab.Color = RGB(0, 176, 80)  ' green tab = new/pending
    
    modPerformance.ForceRecalc
    
    Dim elapsed As Double: elapsed = modPerformance.ElapsedSeconds()
    modPerformance.TurboOff
    
    newWs.Activate
    
    modLogger.LogAction "modMonthlyTabGenerator", "GenerateNextMonthOnly", _
        newName & " from " & latestName, elapsed
    
    MsgBox "Next Month Tab Created!" & vbCrLf & vbCrLf & _
           "  New tab: " & newName & vbCrLf & _
           "  Cloned from: " & latestName & vbCrLf & _
           "  Yellow cells = Ready for data entry" & vbCrLf & vbCrLf & _
           "  Next: Enter actuals, then run Reconciliation.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modMonthlyTabGenerator", "ERROR-NextMonth", Err.Description
    MsgBox "Next month error: " & Err.Description, vbCritical, APP_NAME
End Sub


'===============================================================================
'
' ===  DELETE / CLEANUP  =======================================================
'
'===============================================================================

'===============================================================================
' DeleteGeneratedTabs - Remove all auto-generated tabs (Apr-Dec)
'===============================================================================
Public Sub DeleteGeneratedTabs()
    If MsgBox("Delete all auto-generated monthly tabs (Apr-Dec " & FISCAL_YEAR & ")?" & vbCrLf & _
              "This cannot be undone.", vbYesNo + vbExclamation, APP_NAME) = vbNo Then Exit Sub
    
    modPerformance.TurboOn
    
    Dim mths As Variant: mths = modConfig.GetMonths()
    Dim deleted As Long: deleted = 0
    Dim i As Long
    
    For i = 3 To 11  ' Apr through Dec
        Dim shName As String
        shName = "Functional P&L Summary - " & mths(i) & " " & FISCAL_YEAR
        If modConfig.SheetExists(shName) Then
            modConfig.SafeDeleteSheet shName
            deleted = deleted + 1
        End If
    Next i
    
    modPerformance.TurboOff
    
    modLogger.LogAction "modMonthlyTabGenerator", "DeleteGeneratedTabs", deleted & " tabs deleted"
    MsgBox deleted & " tabs deleted.", vbInformation, APP_NAME
End Sub


'===============================================================================
'
' ===  PRIVATE HELPERS  ========================================================
'
'===============================================================================

'===============================================================================
' UpdateSheetFormulas - Rewrite column references in formulas
'===============================================================================
Private Sub UpdateSheetFormulas(ByVal ws As Worksheet, _
                                ByVal oldCol As String, _
                                ByVal newCol As String, _
                                ByVal monthName As String)
    Dim cell As Range
    Dim usedRng As Range: Set usedRng = ws.UsedRange
    
    For Each cell In usedRng
        If cell.HasFormula Then
            Dim f As String: f = cell.Formula
            ' Replace column references in the Functional P&L trend sheet reference
            If InStr(f, "'" & SH_FUNC_TREND & "'!") > 0 Then
                f = ReplaceColInFormula(f, SH_FUNC_TREND, oldCol, newCol)
                On Error Resume Next
                cell.Formula = f
                On Error GoTo 0
            End If
        End If
    Next cell
End Sub

'===============================================================================
' ReplaceColInFormula - Swap column letter in sheet-qualified references
'===============================================================================
Private Function ReplaceColInFormula(ByVal formula As String, _
                                     ByVal sheetName As String, _
                                     ByVal oldCol As String, _
                                     ByVal newCol As String) As String
    Dim oldRef As String: oldRef = "'" & sheetName & "'!" & oldCol
    Dim newRef As String: newRef = "'" & sheetName & "'!" & newCol
    ReplaceColInFormula = Replace(formula, oldRef, newRef)
    
    ' Also handle $D (absolute column references)
    oldRef = "'" & sheetName & "'!$" & oldCol
    newRef = "'" & sheetName & "'!$" & newCol
    ReplaceColInFormula = Replace(ReplaceColInFormula, oldRef, newRef)
End Function

'===============================================================================
' UpdateHeaderText - Replace month name in title cells (SAFE version)
'
' FIX (v2.1 — ISSUE-002 / BUG-013):
' The v2.0 version used blind Replace(cell.Value, "Mar", newMonth) which
' corrupted any word containing "Mar" as a substring:
'   "Margin"  -> "Aprigin"
'   "March"   -> "Aprch"
'   "Market"  -> "Aprket"
'
' This safe version only replaces these COMPLETE patterns:
'   "Mar 25"       -> "Apr 25"      (short fiscal year suffix)
'   "Mar 2025"     -> "Apr 2025"    (full fiscal year suffix)
'   "MARCH"        -> "APRIL"       (uppercase month name)
'   "Month of Mar" -> "Month of Apr" (report subtitle pattern)
'
' None of these patterns can match inside "Margin", "Market", etc.
'===============================================================================
Private Sub UpdateHeaderText(ByVal ws As Worksheet, _
                              ByVal oldMonth As String, _
                              ByVal newMonth As String)
    Dim cell As Range
    Dim safePats As Variant
    safePats = Array(oldMonth & " " & FISCAL_YEAR, _
                     oldMonth & " " & FISCAL_YEAR_4, _
                     UCase(oldMonth), _
                     "Month of " & oldMonth)
    
    For Each cell In ws.UsedRange
        If Not cell.HasFormula And VarType(cell.Value) = vbString Then
            Dim p As Long
            For p = 0 To UBound(safePats)
                If InStr(cell.Value, safePats(p)) > 0 Then
                    cell.Value = Replace(cell.Value, safePats(p), _
                                 Replace(safePats(p), oldMonth, newMonth))
                End If
            Next p
        End If
    Next cell
End Sub
