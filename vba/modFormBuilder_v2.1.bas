Attribute VB_Name = "modFormBuilder"
Option Explicit

'===============================================================================
' modFormBuilder - UserForm Command Center Builder & Launcher
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Resolves BUG-038 (InputBox character limit) by replacing the
'           InputBox menu with a professional UserForm command center.
'
'           Two usage modes:
'           (A) AUTOMATIC: Run BuildCommandCenter once to create the form.
'               Requires: Trust access to VBA project object model
'               (File > Options > Trust Center > Macro Settings > check the box)
'           (B) MANUAL: Follow instructions in CreateFormManually to build
'               the form by hand in the VBE designer, then paste the code
'               from GetFormCode into the form's code module.
'
' PUBLIC SUBS:
'   BuildCommandCenter   - One-time form builder (needs Trust Access)
'   LaunchCommandCenter  - Show the form (or fallback to InputBox)
'   ExecuteAction        - Central router called by the form
'   CreateFormManually   - MsgBox with manual build steps
'   GetFormCodeForManual - Prints form code to Immediate Window
'   GetFormInstallGuide  - Prints complete install guide to Immediate Window
'
' VERSION:  2.1.0
' CHANGES:  v2.1.0:
'           + Added GetFormInstallGuide for step-by-step install to Immediate
'           + Verified all 50 items in ExecuteAction match training reference
'           + Version header updated
'===============================================================================

' --- Category & Action Data ---
Private Const CAT_COUNT As Long = 15
Private Const ACT_COUNT As Long = 62

'===============================================================================
' LaunchCommandCenter - Show UserForm or fall back to InputBox
'===============================================================================
Public Sub LaunchCommandCenter()
    On Error GoTo Fallback
    
    ' Try to show the UserForm
    Dim frm As Object
    Set frm = Nothing
    
    On Error Resume Next
    Set frm = VBA.UserForms.Add("frmCommandCenter")
    On Error GoTo Fallback
    
    If frm Is Nothing Then GoTo Fallback
    
    frm.Show vbModal
    Exit Sub

Fallback:
    ' Form doesn't exist yet - offer to build or use InputBox
    On Error GoTo FallbackErr
    
    Dim resp As VbMsgBoxResult
    resp = MsgBox("The Command Center form is not yet installed." & vbCrLf & vbCrLf & _
                  "YES = Try to build it automatically" & vbCrLf & _
                  "       (requires Trust Access to VBA project)" & vbCrLf & vbCrLf & _
                  "NO  = Use classic InputBox menu instead", _
                  vbYesNo + vbQuestion, APP_NAME)
    
    If resp = vbYes Then
        BuildCommandCenter
    Else
        modMasterMenu.ShowMasterMenu
    End If
    Exit Sub

FallbackErr:
    ' Ultimate fallback
    modMasterMenu.ShowMasterMenu
End Sub

'===============================================================================
' BuildCommandCenter - Programmatically create the UserForm
'===============================================================================
Public Sub BuildCommandCenter()
    On Error GoTo ErrHandler
    
    ' Check if form already exists
    Dim vbComp As Object
    Dim formExists As Boolean: formExists = False
    
    On Error Resume Next
    Set vbComp = ThisWorkbook.VBProject.VBComponents("frmCommandCenter")
    If Not vbComp Is Nothing Then formExists = True
    On Error GoTo ErrHandler
    
    If formExists Then
        If MsgBox("frmCommandCenter already exists. Rebuild it?", _
                  vbYesNo + vbQuestion, APP_NAME) = vbNo Then
            ' Just show it
            VBA.UserForms.Add("frmCommandCenter").Show vbModal
            Exit Sub
        End If
        ThisWorkbook.VBProject.VBComponents.Remove vbComp
        Set vbComp = Nothing
    End If
    
    Application.StatusBar = "Building Command Center form..."
    
    ' Create new UserForm
    Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(3)  ' vbext_ct_MSForm
    
    With vbComp
        .Name = "frmCommandCenter"
        
        ' --- Form Properties ---
        .Properties("Caption") = "Keystone BenefitTech - Command Center v" & APP_VERSION
        .Properties("Width") = 540
        .Properties("Height") = 440
        .Properties("StartUpPosition") = 1  ' CenterOwner
        .Properties("BackColor") = &HFFFFFF  ' White
        
        Dim frm As Object: Set frm = .Designer
        
        ' ===================================================================
        ' CONTROLS - Title Section
        ' ===================================================================
        Dim ctrl As Object
        
        ' Title Label
        Set ctrl = frm.Controls.Add("Forms.Label.1", "lblTitle")
        With ctrl
            .Caption = "AUTOMATION COMMAND CENTER"
            .Left = 12: .Top = 8: .Width = 390: .Height = 22
            .Font.Size = 14: .Font.Bold = True
            .ForeColor = &H794E1F  ' Navy
        End With
        
        ' Version Label
        Set ctrl = frm.Controls.Add("Forms.Label.1", "lblVersion")
        With ctrl
            .Caption = "v" & APP_VERSION & " | " & ThisWorkbook.Worksheets.Count & " sheets | 62 actions"
            .Left = 12: .Top = 30: .Width = 350: .Height = 14
            .Font.Size = 8: .Font.Italic = True
            .ForeColor = &H808080
        End With
        
        ' ===================================================================
        ' CONTROLS - Search Box
        ' ===================================================================
        Set ctrl = frm.Controls.Add("Forms.Label.1", "lblSearch")
        With ctrl
            .Caption = "Search:"
            .Left = 12: .Top = 52: .Width = 44: .Height = 16
            .Font.Size = 9
        End With
        
        Set ctrl = frm.Controls.Add("Forms.TextBox.1", "txtSearch")
        With ctrl
            .Left = 58: .Top = 50: .Width = 456: .Height = 20
            .Font.Size = 9
        End With
        
        ' ===================================================================
        ' CONTROLS - Category ListBox (left panel)
        ' ===================================================================
        Set ctrl = frm.Controls.Add("Forms.Label.1", "lblCats")
        With ctrl
            .Caption = "Categories"
            .Left = 12: .Top = 76: .Width = 140: .Height = 14
            .Font.Size = 9: .Font.Bold = True
        End With
        
        Set ctrl = frm.Controls.Add("Forms.ListBox.1", "lstCategories")
        With ctrl
            .Left = 12: .Top = 92: .Width = 146: .Height = 260
            .Font.Size = 9
        End With
        
        ' ===================================================================
        ' CONTROLS - Actions ListBox (right panel)
        ' ===================================================================
        Set ctrl = frm.Controls.Add("Forms.Label.1", "lblActions")
        With ctrl
            .Caption = "Available Actions"
            .Left = 168: .Top = 76: .Width = 200: .Height = 14
            .Font.Size = 9: .Font.Bold = True
        End With
        
        Set ctrl = frm.Controls.Add("Forms.ListBox.1", "lstActions")
        With ctrl
            .Left = 168: .Top = 92: .Width = 346: .Height = 260
            .Font.Size = 9
            .ColumnCount = 2
            .ColumnWidths = "30;310"
        End With
        
        ' ===================================================================
        ' CONTROLS - Buttons
        ' ===================================================================
        Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnRun")
        With ctrl
            .Caption = "Run Selected"
            .Left = 168: .Top = 360: .Width = 110: .Height = 28
            .Font.Size = 10: .Font.Bold = True
            .BackColor = &H80FF80  ' Light green
        End With
        
        Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnRunClose")
        With ctrl
            .Caption = "Run && Close"
            .Left = 286: .Top = 360: .Width = 100: .Height = 28
            .Font.Size = 9
        End With
        
        Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnClose")
        With ctrl
            .Caption = "Close"
            .Left = 440: .Top = 360: .Width = 74: .Height = 28
            .Font.Size = 9
            .Cancel = True
        End With
        
        ' ===================================================================
        ' CONTROLS - Status Bar
        ' ===================================================================
        Set ctrl = frm.Controls.Add("Forms.Label.1", "lblStatus")
        With ctrl
            .Caption = "Select a category, then double-click an action to run it."
            .Left = 12: .Top = 398: .Width = 502: .Height = 14
            .Font.Size = 8: .Font.Italic = True
            .ForeColor = &H808080
        End With
        
        ' ===================================================================
        ' INJECT CODE into form's code module
        ' ===================================================================
        Dim codeLines As String
        codeLines = GetFormCode()
        
        ' Clear any existing code
        If .CodeModule.CountOfLines > 0 Then
            .CodeModule.DeleteLines 1, .CodeModule.CountOfLines
        End If
        .CodeModule.InsertLines 1, codeLines
        
    End With
    
    Application.StatusBar = False
    
    modLogger.LogAction "modFormBuilder", "BuildCommandCenter", "Form created successfully"
    
    MsgBox "Command Center form built successfully!" & vbCrLf & vbCrLf & _
           "The form will now open. In future, use:" & vbCrLf & _
           "  Ctrl+Shift+M  or  modFormBuilder.LaunchCommandCenter", _
           vbInformation, APP_NAME
    
    ' Show it
    VBA.UserForms.Add("frmCommandCenter").Show vbModal
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    If Err.Number = 1004 Or InStr(LCase(Err.Description), "programmatic access") > 0 Or _
       InStr(LCase(Err.Description), "trust") > 0 Then
        MsgBox "Cannot create form automatically." & vbCrLf & vbCrLf & _
               "To enable: File > Options > Trust Center >" & vbCrLf & _
               "Trust Center Settings > Macro Settings >" & vbCrLf & _
               "Check 'Trust access to the VBA project object model'" & vbCrLf & vbCrLf & _
               "OR: Follow the manual instructions in" & vbCrLf & _
               "    modFormBuilder.CreateFormManually", _
               vbExclamation, APP_NAME
        ' Fall back to InputBox
        modMasterMenu.ShowMasterMenu
    Else
        MsgBox "Error building form:" & vbCrLf & Err.Description, vbExclamation, APP_NAME
    End If
End Sub

'===============================================================================
' GetFormCode - Returns the VBA code to inject into the form module
'===============================================================================
Private Function GetFormCode() As String
    Dim s As String
    
    s = "Option Explicit" & vbCrLf & vbCrLf
    
    ' --- Type for action entries ---
    s = s & "Private Type ActionEntry" & vbCrLf
    s = s & "    Num As Long" & vbCrLf
    s = s & "    Category As String" & vbCrLf
    s = s & "    Label As String" & vbCrLf
    s = s & "End Type" & vbCrLf & vbCrLf
    
    s = s & "Private m_Actions() As ActionEntry" & vbCrLf
    s = s & "Private m_FilteredNums() As Long" & vbCrLf
    s = s & "Private m_ActionCount As Long" & vbCrLf & vbCrLf
    
    ' --- Initialize ---
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    LoadActions" & vbCrLf
    s = s & "    LoadCategories" & vbCrLf
    s = s & "    lstCategories.ListIndex = 0" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    ' --- AddAction helper (avoids Array() line-continuation limit) ---
    s = s & "Private Sub AddAction(n As Long, cat As String, lbl As String)" & vbCrLf
    s = s & "    m_Actions(n).Num = n" & vbCrLf
    s = s & "    m_Actions(n).Category = cat" & vbCrLf
    s = s & "    m_Actions(n).Label = lbl" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf

    ' --- LoadActions - one line per action, no line continuations ---
    s = s & "Private Sub LoadActions()" & vbCrLf
    s = s & "    m_ActionCount = 62" & vbCrLf
    s = s & "    ReDim m_Actions(1 To 62)" & vbCrLf
    s = s & "    AddAction 1, ""Monthly Operations"", ""Generate Monthly Tabs""" & vbCrLf
    s = s & "    AddAction 2, ""Monthly Operations"", ""Delete Generated Tabs""" & vbCrLf
    s = s & "    AddAction 3, ""Monthly Operations"", ""Run Reconciliation Checks""" & vbCrLf
    s = s & "    AddAction 4, ""Monthly Operations"", ""Export Reconciliation Report""" & vbCrLf
    s = s & "    AddAction 5, ""Analysis"", ""Run Sensitivity Analysis""" & vbCrLf
    s = s & "    AddAction 6, ""Analysis"", ""Run Variance Analysis""" & vbCrLf
    s = s & "    AddAction 7, ""Data Quality"", ""Scan for Data Quality Issues""" & vbCrLf
    s = s & "    AddAction 8, ""Data Quality"", ""Fix Text-Stored Numbers""" & vbCrLf
    s = s & "    AddAction 9, ""Data Quality"", ""Fix Duplicate Rows""" & vbCrLf
    s = s & "    AddAction 10, ""Reporting"", ""Export Report Package (PDF)""" & vbCrLf
    s = s & "    AddAction 11, ""Reporting"", ""Export Active Sheet (PDF)""" & vbCrLf
    s = s & "    AddAction 12, ""Reporting"", ""Build Dashboard Charts""" & vbCrLf
    s = s & "    AddAction 13, ""Utilities"", ""Refresh Table of Contents""" & vbCrLf
    s = s & "    AddAction 14, ""Utilities"", ""Recalculate AWS Allocations""" & vbCrLf
    s = s & "    AddAction 15, ""Utilities"", ""Quick Jump to Sheet""" & vbCrLf
    s = s & "    AddAction 16, ""Utilities"", ""Go Home (Report-->)""" & vbCrLf
    s = s & "    AddAction 17, ""Data & Import"", ""Import GL Data Pipeline""" & vbCrLf
    s = s & "    AddAction 18, ""Forecasting"", ""Rolling Forecast""" & vbCrLf
    s = s & "    AddAction 19, ""Forecasting"", ""Append Month to Trend""" & vbCrLf
    s = s & "    AddAction 20, ""Scenarios"", ""Save Current Scenario""" & vbCrLf
    s = s & "    AddAction 21, ""Scenarios"", ""Load Scenario""" & vbCrLf
    s = s & "    AddAction 22, ""Scenarios"", ""Compare Scenarios""" & vbCrLf
    s = s & "    AddAction 23, ""Scenarios"", ""Delete Scenario""" & vbCrLf
    s = s & "    AddAction 24, ""Allocation"", ""Run Allocation Engine""" & vbCrLf
    s = s & "    AddAction 25, ""Allocation"", ""Allocation Scenario Preview""" & vbCrLf
    s = s & "    AddAction 26, ""Consolidation"", ""Consolidation Menu""" & vbCrLf
    s = s & "    AddAction 27, ""Consolidation"", ""Add Entity File""" & vbCrLf
    s = s & "    AddAction 28, ""Consolidation"", ""Generate Consolidated P&L""" & vbCrLf
    s = s & "    AddAction 29, ""Consolidation"", ""View Loaded Entities""" & vbCrLf
    s = s & "    AddAction 30, ""Consolidation"", ""Add Elimination Entry""" & vbCrLf
    s = s & "    AddAction 31, ""Version Control"", ""Version Control Menu""" & vbCrLf
    s = s & "    AddAction 32, ""Version Control"", ""Save Version""" & vbCrLf
    s = s & "    AddAction 33, ""Version Control"", ""Compare Versions""" & vbCrLf
    s = s & "    AddAction 34, ""Version Control"", ""Restore Version""" & vbCrLf
    s = s & "    AddAction 35, ""Version Control"", ""List Versions""" & vbCrLf
    s = s & "    AddAction 36, ""Governance"", ""Auto-Documentation""" & vbCrLf
    s = s & "    AddAction 37, ""Governance"", ""Change Management Menu""" & vbCrLf
    s = s & "    AddAction 38, ""Governance"", ""Add Change Request""" & vbCrLf
    s = s & "    AddAction 39, ""Governance"", ""Update CR Status""" & vbCrLf
    s = s & "    AddAction 40, ""Governance"", ""CR Summary Report""" & vbCrLf
    s = s & "    AddAction 41, ""Admin & Testing"", ""View Audit Log""" & vbCrLf
    s = s & "    AddAction 42, ""Admin & Testing"", ""Export Audit Log""" & vbCrLf
    s = s & "    AddAction 43, ""Admin & Testing"", ""Clear Audit Log""" & vbCrLf
    s = s & "    AddAction 44, ""Admin & Testing"", ""Full Integration Test""" & vbCrLf
    s = s & "    AddAction 45, ""Admin & Testing"", ""Quick Health Check""" & vbCrLf
    s = s & "    AddAction 46, ""Advanced"", ""Variance Commentary""" & vbCrLf
    s = s & "    AddAction 47, ""Advanced"", ""Cross-Sheet Validation""" & vbCrLf
    s = s & "    AddAction 48, ""Advanced"", ""Executive Mode Toggle""" & vbCrLf
    s = s & "    AddAction 49, ""Advanced"", ""Force Recalculate All""" & vbCrLf
    s = s & "    AddAction 50, ""Advanced"", ""About This Toolkit""" & vbCrLf
    s = s & "    AddAction 51, ""Sheet Tools"", ""Delete All Blank Rows""" & vbCrLf
    s = s & "    AddAction 52, ""Sheet Tools"", ""Unhide All Worksheets""" & vbCrLf
    s = s & "    AddAction 53, ""Sheet Tools"", ""Sort Sheets Alphabetically""" & vbCrLf
    s = s & "    AddAction 54, ""Sheet Tools"", ""Toggle Freeze Panes""" & vbCrLf
    s = s & "    AddAction 55, ""Sheet Tools"", ""Convert Formulas to Values""" & vbCrLf
    s = s & "    AddAction 56, ""Sheet Tools"", ""AutoFit All Columns""" & vbCrLf
    s = s & "    AddAction 57, ""Sheet Tools"", ""Protect All Sheets""" & vbCrLf
    s = s & "    AddAction 58, ""Sheet Tools"", ""Unprotect All Sheets""" & vbCrLf
    s = s & "    AddAction 59, ""Sheet Tools"", ""Find & Replace (All Sheets)""" & vbCrLf
    s = s & "    AddAction 60, ""Sheet Tools"", ""Highlight Hardcoded Numbers""" & vbCrLf
    s = s & "    AddAction 61, ""Sheet Tools"", ""Toggle Presentation Mode""" & vbCrLf
    s = s & "    AddAction 62, ""Sheet Tools"", ""Unmerge and Fill Down""" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    ' --- LoadCategories ---
    s = s & "Private Sub LoadCategories()" & vbCrLf
    s = s & "    lstCategories.Clear" & vbCrLf
    s = s & "    lstCategories.AddItem ""All Actions""" & vbCrLf
    s = s & "    Dim cats As Variant" & vbCrLf
    s = s & "    cats = Array(""Monthly Operations"", ""Analysis"", ""Data Quality"", _" & vbCrLf
    s = s & "        ""Reporting"", ""Utilities"", ""Data & Import"", ""Forecasting"", _" & vbCrLf
    s = s & "        ""Scenarios"", ""Allocation"", ""Consolidation"", _" & vbCrLf
    s = s & "        ""Version Control"", ""Governance"", ""Admin & Testing"", ""Advanced"", ""Sheet Tools"")" & vbCrLf
    s = s & "    Dim c As Variant" & vbCrLf
    s = s & "    For Each c In cats" & vbCrLf
    s = s & "        lstCategories.AddItem CStr(c)" & vbCrLf
    s = s & "    Next c" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    ' --- Category selection ---
    s = s & "Private Sub lstCategories_Click()" & vbCrLf
    s = s & "    FilterActions" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    ' --- Search ---
    s = s & "Private Sub txtSearch_Change()" & vbCrLf
    s = s & "    FilterActions" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    ' --- FilterActions - Apply category + search filters ---
    s = s & "Private Sub FilterActions()" & vbCrLf
    s = s & "    lstActions.Clear" & vbCrLf
    s = s & "    Dim selCat As String" & vbCrLf
    s = s & "    If lstCategories.ListIndex >= 0 Then" & vbCrLf
    s = s & "        selCat = lstCategories.List(lstCategories.ListIndex)" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        selCat = ""All Actions""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Dim srch As String: srch = LCase(Trim(txtSearch.Text))" & vbCrLf
    s = s & "    Dim cnt As Long: cnt = 0" & vbCrLf
    s = s & "    ReDim m_FilteredNums(1 To 62)" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    For i = 1 To m_ActionCount" & vbCrLf
    s = s & "        Dim catMatch As Boolean" & vbCrLf
    s = s & "        catMatch = (selCat = ""All Actions"") Or (m_Actions(i).Category = selCat)" & vbCrLf
    s = s & "        Dim srchMatch As Boolean" & vbCrLf
    s = s & "        srchMatch = (srch = """") Or _" & vbCrLf
    s = s & "            (InStr(1, LCase(m_Actions(i).Label), srch) > 0) Or _" & vbCrLf
    s = s & "            (InStr(1, LCase(m_Actions(i).Category), srch) > 0)" & vbCrLf
    s = s & "        If catMatch And srchMatch Then" & vbCrLf
    s = s & "            cnt = cnt + 1" & vbCrLf
    s = s & "            m_FilteredNums(cnt) = m_Actions(i).Num" & vbCrLf
    s = s & "            lstActions.AddItem m_Actions(i).Num" & vbCrLf
    s = s & "            lstActions.List(lstActions.ListCount - 1, 1) = m_Actions(i).Label" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    lblStatus.Caption = cnt & "" actions shown""" & vbCrLf
    s = s & "    If cnt > 0 Then lstActions.ListIndex = 0" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    ' --- Double-click to run ---
    s = s & "Private Sub lstActions_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
    s = s & "    RunSelected" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    ' --- Button handlers ---
    s = s & "Private Sub btnRun_Click()" & vbCrLf
    s = s & "    RunSelected" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    s = s & "Private Sub btnRunClose_Click()" & vbCrLf
    s = s & "    Dim actNum As Long: actNum = GetSelectedNum()" & vbCrLf
    s = s & "    If actNum = 0 Then Exit Sub" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "    modFormBuilder.ExecuteAction actNum" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    ' --- Run helper ---
    s = s & "Private Sub RunSelected()" & vbCrLf
    s = s & "    Dim actNum As Long: actNum = GetSelectedNum()" & vbCrLf
    s = s & "    If actNum = 0 Then Exit Sub" & vbCrLf
    s = s & "    lblStatus.Caption = ""Running #"" & actNum & ""...""" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "    modFormBuilder.ExecuteAction actNum" & vbCrLf
    s = s & "    Me.Show" & vbCrLf
    s = s & "    lblStatus.Caption = ""Last run: #"" & actNum & "" - Done""" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    ' --- Get selected action number ---
    s = s & "Private Function GetSelectedNum() As Long" & vbCrLf
    s = s & "    GetSelectedNum = 0" & vbCrLf
    s = s & "    If lstActions.ListIndex < 0 Then" & vbCrLf
    s = s & "        MsgBox ""Select an action first."", vbExclamation" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    GetSelectedNum = CLng(lstActions.List(lstActions.ListIndex, 0))" & vbCrLf
    s = s & "End Function" & vbCrLf & vbCrLf
    
    ' --- Escape key closes ---
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = 0 Then Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    
    GetFormCode = s
End Function

'===============================================================================
' ExecuteAction - Central router (called by form, same as modMasterMenu)
' ALL 50 ITEMS — single routing table. modMasterMenu delegates here.
'===============================================================================
Public Sub ExecuteAction(ByVal num As Long)
    On Error GoTo ErrHandler
    
    Select Case num
        ' --- Monthly Operations (1-4) ---
        Case 1: modMonthlyTabGenerator.GenerateMonthlyTabs
        Case 2: modMonthlyTabGenerator.DeleteGeneratedTabs
        Case 3: modReconciliation.RunAllChecks
        Case 4: modReconciliation.ExportCheckResults
        
        ' --- Analysis (5-6) ---
        Case 5: modSensitivity.RunSensitivityAnalysis
        Case 6: modVarianceAnalysis.RunVarianceAnalysis
        
        ' --- Data Quality (7-9) ---
        Case 7: modDataQuality.ScanAll
        Case 8: modDataQuality.FixTextNumbers
        Case 9: modDataQuality.FixDuplicates
        
        ' --- Reporting (10-12) ---
        Case 10: modPDFExport.ExportReportPackage
        Case 11: modPDFExport.ExportSingleSheet
        Case 12: modDashboard.BuildDashboard
        
        ' --- Utilities (13-16) ---
        Case 13: modNavigation.RefreshTableOfContents
        Case 14: modAWSRecompute.ValidateAndRecalcAWS
        Case 15: modNavigation.QuickJump
        Case 16: modNavigation.GoHome
        
        ' --- Data Import (17) ---
        Case 17: modImport.ImportDataPipeline
        
        ' --- Forecasting (18-19) ---
        Case 18: modForecast.RollingForecast
        Case 19: modForecast.AppendToTrend
        
        ' --- Scenarios (20-23) ---
        Case 20: modScenario.SaveScenario
        Case 21: modScenario.LoadScenario
        Case 22: modScenario.CompareScenarios
        Case 23: modScenario.DeleteScenario
        
        ' --- Allocation (24-25) ---
        Case 24: modAllocation.RunAllocationEngine
        Case 25: modAllocation.AllocationPreview
        
        ' --- Consolidation (26-30) ---
        Case 26: modConsolidation.ShowConsolidationMenu
        Case 27: modConsolidation.AddEntity
        Case 28: modConsolidation.GenerateConsolidated
        Case 29: modConsolidation.ListEntities
        Case 30: modConsolidation.AddElimination
        
        ' --- Version Control (31-35) ---
        Case 31: modVersionControl.ShowVersionMenu
        Case 32: modVersionControl.SaveVersion
        Case 33: modVersionControl.CompareVersions
        Case 34: modVersionControl.RestoreVersion
        Case 35: modVersionControl.ListVersions
        
        ' --- Governance (36-40) ---
        Case 36: modAdmin.GenerateDocumentation
        Case 37: modAdmin.ShowChangeMenu
        Case 38: modAdmin.AddChangeRequest
        Case 39: modAdmin.UpdateChangeStatus
        Case 40: modAdmin.ChangeManagementSummary
        
        ' --- Admin & Testing (41-45) ---
        Case 41: modLogger.ViewLog
        Case 42: modLogger.ExportLog
        Case 43: modLogger.ClearLog
        Case 44: modIntegrationTest.RunFullTest
        Case 45: modIntegrationTest.QuickHealthCheck
        
        ' --- Advanced (46-50) ---
        Case 46: modVarianceAnalysis.GenerateCommentary
        Case 47: modReconciliation.ValidateCrossSheet
        Case 48: modNavigation.ToggleExecutiveMode
        Case 49: modPerformance.ForceRecalc
        Case 50: ShowAbout

        ' --- Sheet Tools (51-62) ---
        Case 51: modUtilities.DeleteBlankRows
        Case 52: modUtilities.UnhideAllSheets
        Case 53: modUtilities.SortSheetsAlphabetically
        Case 54: modUtilities.ToggleFreezePanes
        Case 55: modUtilities.ConvertToValues
        Case 56: modUtilities.AutoFitAllColumns
        Case 57: modUtilities.ProtectAllSheets
        Case 58: modUtilities.UnprotectAllSheets
        Case 59: modUtilities.FindReplaceAllSheets
        Case 60: modUtilities.HighlightHardcodedNumbers
        Case 61: modUtilities.TogglePresentationMode
        Case 62: modUtilities.UnmergeAndFillDown

        Case Else
            MsgBox "Unknown action #" & num, vbExclamation, APP_NAME
    End Select
    Exit Sub

ErrHandler:
    MsgBox "Error running action #" & num & ":" & vbCrLf & Err.Description, _
           vbExclamation, APP_NAME
End Sub

'===============================================================================
' ShowAbout - Toolkit info
'===============================================================================
Private Sub ShowAbout()
    MsgBox "KEYSTONE BENEFITECH AUTOMATION TOOLKIT" & vbCrLf & _
           "Version " & APP_VERSION & " | Build " & APP_BUILD_DATE & vbCrLf & vbCrLf & _
           "30 VBA modules | 62 menu options" & vbCrLf & _
           "Built for: Keystone BenefitTech, Inc." & vbCrLf & _
           "Model: P&L Reporting & Allocation" & vbCrLf & vbCrLf & _
           "UserForm Command Center (BUG-038 resolved)" & vbCrLf & _
           "Keyboard: Ctrl+Shift+M to open", _
           vbInformation, APP_NAME
End Sub

'===============================================================================
' CreateFormManually - Instructions for manual form creation (Mode B)
'===============================================================================
Public Sub CreateFormManually()
    MsgBox "MANUAL FORM CREATION INSTRUCTIONS" & vbCrLf & _
           String(45, "=") & vbCrLf & vbCrLf & _
           "1. In VBE: Insert > UserForm" & vbCrLf & _
           "2. Rename to: frmCommandCenter" & vbCrLf & _
           "3. Set form size: 540 x 440" & vbCrLf & _
           "4. Add these controls:" & vbCrLf & _
           "   - Label 'lblTitle' (top, bold, 14pt)" & vbCrLf & _
           "   - Label 'lblVersion' (below title, 8pt)" & vbCrLf & _
           "   - TextBox 'txtSearch' (below, full width)" & vbCrLf & _
           "   - ListBox 'lstCategories' (left, 146w x 260h)" & vbCrLf & _
           "   - ListBox 'lstActions' (right, 346w x 260h, 2 cols)" & vbCrLf & _
           "   - Button 'btnRun' (Run Selected)" & vbCrLf & _
           "   - Button 'btnRunClose' (Run & Close)" & vbCrLf & _
           "   - Button 'btnClose' (Close, Cancel=True)" & vbCrLf & _
           "   - Label 'lblStatus' (bottom)" & vbCrLf & _
           "5. Paste code from modFormBuilder.GetFormCodeForManual" & vbCrLf & _
           "   into the form's code module" & vbCrLf & _
           "6. Done! Run modFormBuilder.LaunchCommandCenter", _
           vbInformation, APP_NAME
End Sub

'===============================================================================
' GetFormCodeForManual - Print form code to Immediate window for copy/paste
'===============================================================================
Public Sub GetFormCodeForManual()
    Debug.Print "============================================"
    Debug.Print "PASTE THIS CODE INTO frmCommandCenter MODULE"
    Debug.Print "============================================"
    Debug.Print GetFormCode()
    MsgBox "Form code has been printed to the Immediate Window." & vbCrLf & _
           "(View > Immediate Window, or Ctrl+G)" & vbCrLf & vbCrLf & _
           "Select all text and paste into the form's code module.", _
           vbInformation, APP_NAME
End Sub

'===============================================================================
' GetFormInstallGuide - Print complete step-by-step install guide
'                       to the Immediate Window (v2.1 addition)
'===============================================================================
Public Sub GetFormInstallGuide()
    Dim g As String
    
    g = String(70, "=") & vbCrLf
    g = g & "  frmCommandCenter INSTALLATION GUIDE" & vbCrLf
    g = g & "  Keystone BenefitTech Automation Toolkit v" & APP_VERSION & vbCrLf
    g = g & String(70, "=") & vbCrLf & vbCrLf
    
    g = g & "MODE A - AUTOMATIC (recommended)" & vbCrLf
    g = g & String(40, "-") & vbCrLf
    g = g & "Prerequisites:" & vbCrLf
    g = g & "  1. File > Options > Trust Center > Trust Center Settings" & vbCrLf
    g = g & "  2. Macro Settings > check 'Trust access to the VBA project" & vbCrLf
    g = g & "     object model'" & vbCrLf
    g = g & "  3. Click OK twice to close dialogs" & vbCrLf & vbCrLf
    g = g & "Steps:" & vbCrLf
    g = g & "  1. Press Alt+F8 (Macros dialog)" & vbCrLf
    g = g & "  2. Select 'BuildCommandCenter', click Run" & vbCrLf
    g = g & "  3. Form is built and opens automatically" & vbCrLf
    g = g & "  4. Press Ctrl+Shift+M anytime to reopen" & vbCrLf & vbCrLf
    
    g = g & "MODE B - MANUAL (if Trust Access is unavailable)" & vbCrLf
    g = g & String(40, "-") & vbCrLf
    g = g & "Step 1: Create the UserForm" & vbCrLf
    g = g & "  - Press Alt+F11 to open VBE" & vbCrLf
    g = g & "  - Insert > UserForm" & vbCrLf
    g = g & "  - In Properties window (F4), set:" & vbCrLf
    g = g & "      (Name)         = frmCommandCenter" & vbCrLf
    g = g & "      Caption        = Keystone BenefitTech - Command Center" & vbCrLf
    g = g & "      Width          = 540" & vbCrLf
    g = g & "      Height         = 440" & vbCrLf
    g = g & "      StartUpPosition = 1 (CenterOwner)" & vbCrLf & vbCrLf
    
    g = g & "Step 2: Add Controls (use Toolbox to drag onto form)" & vbCrLf
    g = g & "  Label   'lblTitle'       L=12  T=8   W=390 H=22  Bold 14pt" & vbCrLf
    g = g & "  Label   'lblVersion'     L=12  T=30  W=350 H=14  Italic 8pt" & vbCrLf
    g = g & "  Label   'lblSearch'      L=12  T=52  W=44  H=16  Caption='Search:'" & vbCrLf
    g = g & "  TextBox 'txtSearch'      L=58  T=50  W=456 H=20" & vbCrLf
    g = g & "  Label   'lblCats'        L=12  T=76  W=140 H=14  Bold Caption='Categories'" & vbCrLf
    g = g & "  ListBox 'lstCategories'  L=12  T=92  W=146 H=260" & vbCrLf
    g = g & "  Label   'lblActions'     L=168 T=76  W=200 H=14  Bold Caption='Available Actions'" & vbCrLf
    g = g & "  ListBox 'lstActions'     L=168 T=92  W=346 H=260 ColumnCount=2 ColumnWidths='30;310'" & vbCrLf
    g = g & "  Button  'btnRun'         L=168 T=360 W=110 H=28  Caption='Run Selected'" & vbCrLf
    g = g & "  Button  'btnRunClose'    L=286 T=360 W=100 H=28  Caption='Run & Close'" & vbCrLf
    g = g & "  Button  'btnClose'       L=440 T=360 W=74  H=28  Caption='Close' Cancel=True" & vbCrLf
    g = g & "  Label   'lblStatus'      L=12  T=398 W=502 H=14  Italic 8pt" & vbCrLf & vbCrLf
    
    g = g & "Step 3: Add Code" & vbCrLf
    g = g & "  - Double-click the form to open its code module" & vbCrLf
    g = g & "  - Delete any auto-generated code" & vbCrLf
    g = g & "  - Run modFormBuilder.GetFormCodeForManual (Alt+F8)" & vbCrLf
    g = g & "  - Press Ctrl+G to open Immediate Window" & vbCrLf
    g = g & "  - Select ALL text, Ctrl+C to copy" & vbCrLf
    g = g & "  - Paste into the form's code module" & vbCrLf & vbCrLf
    
    g = g & "Step 4: Verify" & vbCrLf
    g = g & "  - Save the workbook (Ctrl+S)" & vbCrLf
    g = g & "  - Press Ctrl+Shift+M from any sheet" & vbCrLf
    g = g & "  - Command Center should open with 62 actions" & vbCrLf
    g = g & "  - Try clicking categories, searching, running an action" & vbCrLf & vbCrLf
    
    g = g & String(70, "=") & vbCrLf
    g = g & "  END OF INSTALL GUIDE" & vbCrLf
    g = g & String(70, "=") & vbCrLf
    
    Debug.Print g
    
    MsgBox "Install guide printed to the Immediate Window." & vbCrLf & _
           "(View > Immediate Window, or Ctrl+G)" & vbCrLf & vbCrLf & _
           "You can copy the guide text for reference.", _
           vbInformation, APP_NAME
End Sub
