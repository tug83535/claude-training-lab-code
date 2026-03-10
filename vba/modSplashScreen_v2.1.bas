Attribute VB_Name = "modSplashScreen"
Option Explicit

'===============================================================================
' modSplashScreen - Branded Splash Screen on Workbook Open
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Displays a professional iPipeline-branded welcome screen when the
'           workbook opens. Shows version number, tool count, and a button to
'           launch the Command Center. Auto-dismisses after 5 seconds or on
'           click. Makes the file feel like a product, not a spreadsheet.
'
' USAGE:    Call ShowSplash from Workbook_Open event:
'             Private Sub Workbook_Open()
'                 modSplashScreen.ShowSplash
'             End Sub
'
' PUBLIC SUBS:
'   ShowSplash           - Display the splash screen (called on open)
'   ShowSplashManual     - Same as ShowSplash but always shows (for testing)
'
' DEPENDENCIES: modConfig, modFormBuilder (for LaunchCommandCenter)
' VERSION:  2.1.0
'===============================================================================

' --- Splash screen colors (iPipeline brand) ---
Private Const SPLASH_BG      As Long = 1132409   ' RGB(9,71,17) -> actually use CLR_NAVY
Private Const SPLASH_ACCENT  As Long = 13303487   ' RGB(191,241,140) Lime accent

'===============================================================================
' ShowSplash - Display branded splash screen
' Creates a temporary sheet, shows it for 5 seconds, then removes it.
' Uses a modeless approach via Application.OnTime for auto-dismiss.
'===============================================================================
Public Sub ShowSplash()
    On Error GoTo ErrHandler

    ' Don't show splash if the user has already dismissed it this session
    ' (use a module-level flag if needed, but for now always show on open)

    Dim frm As Object
    On Error Resume Next
    Set frm = VBA.UserForms.Add("frmSplash")
    On Error GoTo ErrHandler

    ' If the UserForm exists, show it
    If Not frm Is Nothing Then
        frm.Show vbModeless
        Application.OnTime Now + TimeSerial(0, 0, 5), "modSplashScreen.DismissSplash"
        Exit Sub
    End If

    ' Fallback: Show a styled MsgBox splash if UserForm not available
    ShowMsgBoxSplash
    Exit Sub

ErrHandler:
    ' Splash should never crash anything - fail silently
    On Error Resume Next
    ShowMsgBoxSplash
    On Error GoTo 0
End Sub

'===============================================================================
' ShowSplashManual - Always show splash (for testing from the Command Center)
'===============================================================================
Public Sub ShowSplashManual()
    ShowMsgBoxSplash
End Sub

'===============================================================================
' DismissSplash - Called by OnTime to auto-close the splash form
'===============================================================================
Public Sub DismissSplash()
    On Error Resume Next
    Unload VBA.UserForms("frmSplash")
    On Error GoTo 0
End Sub

'===============================================================================
' ShowMsgBoxSplash - Fallback splash using MsgBox (always works, no UserForm)
'===============================================================================
Private Sub ShowMsgBoxSplash()
    Dim msg As String
    Dim border As String: border = String(50, Chr(9472))

    msg = border & vbCrLf & vbCrLf
    msg = msg & "   KEYSTONE BENEFITECH" & vbCrLf
    msg = msg & "   P&L Reporting & Allocation Model" & vbCrLf & vbCrLf
    msg = msg & border & vbCrLf & vbCrLf
    msg = msg & "   Version: " & APP_VERSION & "  |  Build: " & APP_BUILD_DATE & vbCrLf & vbCrLf
    msg = msg & "   34 VBA Modules  |  62 Command Center Actions" & vbCrLf
    msg = msg & "   14 Python Scripts  |  100+ Automation Tools" & vbCrLf & vbCrLf
    msg = msg & border & vbCrLf & vbCrLf
    msg = msg & "   Press Ctrl+Shift+M to open the Command Center" & vbCrLf
    msg = msg & "   or click OK to get started." & vbCrLf & vbCrLf
    msg = msg & "   Built for iPipeline Finance & Accounting" & vbCrLf

    MsgBox msg, vbInformation, APP_NAME & " v" & APP_VERSION

    ' Offer to launch Command Center
    Dim resp As VbMsgBoxResult
    resp = MsgBox("Launch the Command Center now?", _
                  vbYesNo + vbQuestion, APP_NAME)
    If resp = vbYes Then
        modFormBuilder.LaunchCommandCenter
    End If
End Sub

'===============================================================================
' GetSplashFormCode - Returns VBA code for frmSplash UserForm
' Use this to manually create the splash form if needed.
' Run this sub, then copy from Immediate Window into frmSplash code module.
'===============================================================================
Public Sub GetSplashFormCode()
    Dim s As String

    s = "' === frmSplash Code Module ===" & vbCrLf & vbCrLf

    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    Me.Caption = """"" & vbCrLf
    s = s & "    Me.Width = 420" & vbCrLf
    s = s & "    Me.Height = 300" & vbCrLf
    s = s & "    Me.BackColor = RGB(11, 71, 121)" & vbCrLf
    s = s & "    Me.StartUpPosition = 1" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf

    s = s & "Private Sub UserForm_Click()" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf

    s = s & "Private Sub btnLaunch_Click()" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "    modFormBuilder.LaunchCommandCenter" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf

    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf

    Debug.Print s
    MsgBox "Splash form code printed to Immediate Window (Ctrl+G).", _
           vbInformation, APP_NAME
End Sub

'===============================================================================
' BuildSplashForm - Programmatically create frmSplash UserForm
' Requires Trust Access to VBA project object model.
'===============================================================================
Public Sub BuildSplashForm()
    On Error GoTo ErrHandler

    ' Check if form already exists
    Dim vbComp As Object
    Dim formExists As Boolean: formExists = False

    On Error Resume Next
    Set vbComp = ThisWorkbook.VBProject.VBComponents("frmSplash")
    If Not vbComp Is Nothing Then formExists = True
    On Error GoTo ErrHandler

    If formExists Then
        If MsgBox("frmSplash already exists. Rebuild it?", _
                  vbYesNo + vbQuestion, APP_NAME) = vbNo Then Exit Sub
        ThisWorkbook.VBProject.VBComponents.Remove vbComp
    End If

    ' Create new UserForm
    Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(3)  ' 3 = vbext_ct_MSForm
    vbComp.Name = "frmSplash"

    With vbComp.Properties
        .Item("Caption") = ""
        .Item("Width") = 420
        .Item("Height") = 300
        .Item("BackColor") = RGB(11, 71, 121)
        .Item("StartUpPosition") = 1
    End With

    ' Add title label
    Dim ctrl As Object
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1", "lblTitle")
    With ctrl
        .Left = 30: .Top = 40: .Width = 360: .Height = 30
        .Caption = "KEYSTONE BENEFITECH"
        .Font.Size = 20: .Font.Bold = True
        .ForeColor = RGB(255, 255, 255)
        .BackColor = RGB(11, 71, 121)
        .BackStyle = 0  ' Transparent
        .TextAlign = 2  ' Center
    End With

    ' Add subtitle
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1", "lblSubtitle")
    With ctrl
        .Left = 30: .Top = 72: .Width = 360: .Height = 20
        .Caption = "P&L Reporting & Allocation Model"
        .Font.Size = 11: .Font.Italic = True
        .ForeColor = RGB(191, 241, 140)
        .BackColor = RGB(11, 71, 121)
        .BackStyle = 0
        .TextAlign = 2
    End With

    ' Add version label
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1", "lblVersion")
    With ctrl
        .Left = 30: .Top = 110: .Width = 360: .Height = 16
        .Caption = "Version " & APP_VERSION & "  |  Build " & APP_BUILD_DATE
        .Font.Size = 9
        .ForeColor = RGB(200, 200, 200)
        .BackColor = RGB(11, 71, 121)
        .BackStyle = 0
        .TextAlign = 2
    End With

    ' Add stats label
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1", "lblStats")
    With ctrl
        .Left = 30: .Top = 135: .Width = 360: .Height = 32
        .Caption = "34 VBA Modules  |  62 Actions  |  14 Python Scripts" & vbCrLf & _
                   "100+ Automation Tools  |  Built for iPipeline"
        .Font.Size = 9
        .ForeColor = RGB(180, 180, 180)
        .BackColor = RGB(11, 71, 121)
        .BackStyle = 0
        .TextAlign = 2
    End With

    ' Add launch button
    Set ctrl = vbComp.Designer.Controls.Add("Forms.CommandButton.1", "btnLaunch")
    With ctrl
        .Left = 110: .Top = 190: .Width = 200: .Height = 34
        .Caption = "Launch Command Center"
        .Font.Size = 11: .Font.Bold = True
        .ForeColor = RGB(11, 71, 121)
        .BackColor = RGB(191, 241, 140)
    End With

    ' Add close label
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1", "lblClose")
    With ctrl
        .Left = 130: .Top = 240: .Width = 160: .Height = 16
        .Caption = "Click anywhere to dismiss"
        .Font.Size = 8: .Font.Italic = True
        .ForeColor = RGB(150, 150, 150)
        .BackColor = RGB(11, 71, 121)
        .BackStyle = 0
        .TextAlign = 2
    End With

    ' Add code to the form
    Dim codeText As String
    codeText = "Private Sub UserForm_Click()" & vbCrLf & _
               "    Unload Me" & vbCrLf & _
               "End Sub" & vbCrLf & vbCrLf & _
               "Private Sub btnLaunch_Click()" & vbCrLf & _
               "    Unload Me" & vbCrLf & _
               "    modFormBuilder.LaunchCommandCenter" & vbCrLf & _
               "End Sub" & vbCrLf & vbCrLf & _
               "Private Sub UserForm_Initialize()" & vbCrLf & _
               "    Me.Caption = """"" & vbCrLf & _
               "End Sub"
    vbComp.CodeModule.AddFromString codeText

    modLogger.LogAction "modSplashScreen", "BuildSplashForm", "frmSplash created successfully"
    MsgBox "Splash screen form created!" & vbCrLf & vbCrLf & _
           "To enable on workbook open, add this to ThisWorkbook:" & vbCrLf & _
           "  Private Sub Workbook_Open()" & vbCrLf & _
           "      modSplashScreen.ShowSplash" & vbCrLf & _
           "  End Sub", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    MsgBox "BuildSplashForm error: " & Err.Description & vbCrLf & vbCrLf & _
           "Make sure Trust Access is enabled:" & vbCrLf & _
           "File > Options > Trust Center > Macro Settings" & vbCrLf & _
           "Check 'Trust access to the VBA project object model'", _
           vbCritical, APP_NAME
End Sub
