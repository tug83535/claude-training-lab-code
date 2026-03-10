Attribute VB_Name = "modProgressBar"
Option Explicit

'===============================================================================
' modProgressBar - Animated Progress Bar for Long-Running Macros
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Provides a professional progress bar UserForm with percentage,
'           task description, elapsed time, and iPipeline branding. Replaces
'           simple status bar messages for any macro that takes a few seconds.
'
' USAGE:    Any module can use the progress bar in 3 steps:
'             1. Call modProgressBar.StartProgress "Task Name", totalSteps
'             2. Call modProgressBar.UpdateProgress currentStep, "Detail text"
'             3. Call modProgressBar.EndProgress
'
'           Example:
'             modProgressBar.StartProgress "Exporting PDF", 7
'             For i = 1 To 7
'                 ' ... export sheet i ...
'                 modProgressBar.UpdateProgress i, "Sheet " & i & " of 7"
'             Next i
'             modProgressBar.EndProgress
'
' PUBLIC SUBS:
'   StartProgress    - Initialize and show the progress bar
'   UpdateProgress   - Update percentage, detail text, and elapsed time
'   EndProgress      - Close and clean up the progress bar
'   BuildProgressForm - Programmatically create frmProgress UserForm
'
' DEPENDENCIES: modConfig (for APP_NAME, CLR_NAVY)
' VERSION:  2.1.0
'===============================================================================

Private m_TotalSteps   As Long
Private m_StartTime    As Double
Private m_TaskName     As String
Private m_UseForm      As Boolean
Private m_FormRef      As Object

'===============================================================================
' StartProgress - Initialize the progress bar
'===============================================================================
Public Sub StartProgress(ByVal taskName As String, ByVal totalSteps As Long)
    m_TaskName = taskName
    m_TotalSteps = totalSteps
    If m_TotalSteps < 1 Then m_TotalSteps = 1
    m_StartTime = Timer

    ' Try to show the UserForm progress bar
    On Error Resume Next
    Set m_FormRef = VBA.UserForms.Add("frmProgress")
    If Not m_FormRef Is Nothing Then
        m_UseForm = True
        m_FormRef.Caption = APP_NAME & " - " & taskName
        m_FormRef.lblTask.Caption = taskName
        m_FormRef.lblPercent.Caption = "0%"
        m_FormRef.lblDetail.Caption = "Starting..."
        m_FormRef.lblElapsed.Caption = "Elapsed: 0s"
        m_FormRef.lblBar.Width = 0
        m_FormRef.Show vbModeless
        DoEvents
    Else
        ' Fallback to status bar
        m_UseForm = False
        Application.StatusBar = taskName & " - 0% complete"
    End If
    On Error GoTo 0
End Sub

'===============================================================================
' UpdateProgress - Update the progress bar with current step
'===============================================================================
Public Sub UpdateProgress(ByVal currentStep As Long, Optional ByVal detail As String = "")
    If currentStep < 0 Then currentStep = 0
    If currentStep > m_TotalSteps Then currentStep = m_TotalSteps

    Dim pct As Double
    pct = currentStep / m_TotalSteps

    Dim elapsed As Double
    elapsed = Timer - m_StartTime
    If elapsed < 0 Then elapsed = elapsed + 86400  ' midnight rollover

    Dim elapsedText As String
    If elapsed < 60 Then
        elapsedText = "Elapsed: " & Round(elapsed, 0) & "s"
    Else
        elapsedText = "Elapsed: " & Int(elapsed / 60) & "m " & Round(elapsed Mod 60, 0) & "s"
    End If

    ' Estimate remaining time
    Dim etaText As String
    If currentStep > 0 And pct < 1 Then
        Dim remaining As Double
        remaining = (elapsed / pct) * (1 - pct)
        If remaining < 60 Then
            etaText = " | ETA: ~" & Round(remaining, 0) & "s"
        Else
            etaText = " | ETA: ~" & Int(remaining / 60) & "m " & Round(remaining Mod 60, 0) & "s"
        End If
    Else
        etaText = ""
    End If

    If m_UseForm And Not m_FormRef Is Nothing Then
        On Error Resume Next
        m_FormRef.lblPercent.Caption = Format(pct, "0%")
        m_FormRef.lblDetail.Caption = detail
        m_FormRef.lblElapsed.Caption = elapsedText & etaText

        ' Animate the bar: lblBar sits on top of lblBarBG
        ' lblBarBG is full width (380), lblBar grows proportionally
        Dim maxWidth As Long: maxWidth = 380
        m_FormRef.lblBar.Width = CLng(pct * maxWidth)
        DoEvents
        On Error GoTo 0
    Else
        ' Status bar fallback
        Dim statusMsg As String
        statusMsg = m_TaskName & " - " & Format(pct, "0%") & " complete"
        If detail <> "" Then statusMsg = statusMsg & " | " & detail
        statusMsg = statusMsg & " | " & elapsedText & etaText
        Application.StatusBar = statusMsg
        DoEvents
    End If
End Sub

'===============================================================================
' EndProgress - Close the progress bar and clean up
'===============================================================================
Public Sub EndProgress()
    If m_UseForm And Not m_FormRef Is Nothing Then
        On Error Resume Next
        m_FormRef.lblPercent.Caption = "100%"
        m_FormRef.lblDetail.Caption = "Complete!"
        m_FormRef.lblBar.Width = 380
        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 1)  ' Brief pause to show 100%
        Unload m_FormRef
        Set m_FormRef = Nothing
        On Error GoTo 0
    Else
        Application.StatusBar = m_TaskName & " - Complete!"
        DoEvents
        Application.StatusBar = False
    End If

    m_UseForm = False
    m_TotalSteps = 0
    m_TaskName = ""
End Sub

'===============================================================================
' BuildProgressForm - Programmatically create frmProgress UserForm
' Requires Trust Access to VBA project object model.
'===============================================================================
Public Sub BuildProgressForm()
    On Error GoTo ErrHandler

    ' Check if form already exists
    Dim vbComp As Object
    Dim formExists As Boolean: formExists = False

    On Error Resume Next
    Set vbComp = ThisWorkbook.VBProject.VBComponents("frmProgress")
    If Not vbComp Is Nothing Then formExists = True
    On Error GoTo ErrHandler

    If formExists Then
        If MsgBox("frmProgress already exists. Rebuild it?", _
                  vbYesNo + vbQuestion, APP_NAME) = vbNo Then Exit Sub
        ThisWorkbook.VBProject.VBComponents.Remove vbComp
    End If

    ' Create new UserForm
    Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(3)  ' vbext_ct_MSForm
    vbComp.Name = "frmProgress"

    With vbComp.Properties
        .Item("Caption") = APP_NAME
        .Item("Width") = 440
        .Item("Height") = 160
        .Item("BackColor") = RGB(255, 255, 255)
        .Item("StartUpPosition") = 1
    End With

    Dim ctrl As Object

    ' Task name label (top)
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1", "lblTask")
    With ctrl
        .Left = 20: .Top = 12: .Width = 400: .Height = 20
        .Caption = "Processing..."
        .Font.Size = 11: .Font.Bold = True
        .ForeColor = RGB(11, 71, 121)
    End With

    ' Detail label
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1", "lblDetail")
    With ctrl
        .Left = 20: .Top = 34: .Width = 300: .Height = 16
        .Caption = ""
        .Font.Size = 9
        .ForeColor = RGB(100, 100, 100)
    End With

    ' Percentage label (right-aligned)
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1", "lblPercent")
    With ctrl
        .Left = 340: .Top = 34: .Width = 70: .Height = 16
        .Caption = "0%"
        .Font.Size = 10: .Font.Bold = True
        .ForeColor = RGB(11, 71, 121)
        .TextAlign = 3  ' Right
    End With

    ' Bar background (gray track)
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1", "lblBarBG")
    With ctrl
        .Left = 20: .Top = 58: .Width = 380: .Height = 28
        .Caption = ""
        .BackColor = RGB(230, 230, 230)
        .BackStyle = 1  ' Opaque
    End With

    ' Bar fill (colored progress indicator)
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1", "lblBar")
    With ctrl
        .Left = 20: .Top = 58: .Width = 0: .Height = 28
        .Caption = ""
        .BackColor = RGB(11, 71, 121)
        .BackStyle = 1
    End With

    ' Elapsed time label
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1", "lblElapsed")
    With ctrl
        .Left = 20: .Top = 96: .Width = 380: .Height = 16
        .Caption = "Elapsed: 0s"
        .Font.Size = 9: .Font.Italic = True
        .ForeColor = RGB(120, 120, 120)
    End With

    ' Branding label (bottom)
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1", "lblBrand")
    With ctrl
        .Left = 20: .Top = 118: .Width = 380: .Height = 14
        .Caption = "Keystone BenefitTech Automation Toolkit v" & APP_VERSION
        .Font.Size = 8: .Font.Italic = True
        .ForeColor = RGB(180, 180, 180)
        .TextAlign = 2  ' Center
    End With

    MsgBox "Progress bar form (frmProgress) created!" & vbCrLf & vbCrLf & _
           "Usage in any module:" & vbCrLf & _
           "  modProgressBar.StartProgress ""Task"", 10" & vbCrLf & _
           "  modProgressBar.UpdateProgress i, ""Step "" & i" & vbCrLf & _
           "  modProgressBar.EndProgress", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    MsgBox "BuildProgressForm error: " & Err.Description & vbCrLf & vbCrLf & _
           "Make sure Trust Access is enabled:" & vbCrLf & _
           "File > Options > Trust Center > Macro Settings" & vbCrLf & _
           "Check 'Trust access to the VBA project object model'", _
           vbCritical, APP_NAME
End Sub
