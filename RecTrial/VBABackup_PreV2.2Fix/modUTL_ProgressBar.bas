Attribute VB_Name = "modUTL_ProgressBar"
Option Explicit

'===============================================================================
' modUTL_ProgressBar - Universal Animated Progress Bar
' Universal Toolkit - Works on ANY Excel file
'===============================================================================
' PURPOSE:  Provides a reusable progress bar for any long-running macro.
'           Falls back to the status bar if UserForm not available.
'           Any VBA module can use this in 3 simple calls.
'
' USAGE:
'   modUTL_ProgressBar.StartProgress "Task Name", totalSteps
'   modUTL_ProgressBar.UpdateProgress currentStep, "Detail text"
'   modUTL_ProgressBar.EndProgress
'
' DEPENDENCIES: None (fully standalone)
' VERSION:  1.0.0
'===============================================================================

Private m_TotalSteps   As Long
Private m_StartTime    As Double
Private m_TaskName     As String

'===============================================================================
' StartProgress - Initialize the progress bar (status bar mode)
'===============================================================================
Public Sub StartProgress(ByVal taskName As String, ByVal totalSteps As Long)
    m_TaskName = taskName
    m_TotalSteps = totalSteps
    If m_TotalSteps < 1 Then m_TotalSteps = 1
    m_StartTime = Timer
    Application.StatusBar = taskName & " - 0% complete"
    DoEvents
End Sub

'===============================================================================
' UpdateProgress - Update the progress bar
'===============================================================================
Public Sub UpdateProgress(ByVal currentStep As Long, Optional ByVal detail As String = "")
    If currentStep < 0 Then currentStep = 0
    If currentStep > m_TotalSteps Then currentStep = m_TotalSteps

    Dim pct As Double: pct = currentStep / m_TotalSteps

    Dim elapsed As Double: elapsed = Timer - m_StartTime
    If elapsed < 0 Then elapsed = elapsed + 86400

    Dim elapsedText As String
    If elapsed < 60 Then
        elapsedText = Round(elapsed, 0) & "s"
    Else
        elapsedText = Int(elapsed / 60) & "m " & Round(elapsed Mod 60, 0) & "s"
    End If

    Dim etaText As String: etaText = ""
    If currentStep > 0 And pct < 1 Then
        Dim remaining As Double: remaining = (elapsed / pct) * (1 - pct)
        If remaining < 60 Then
            etaText = " | ETA: ~" & Round(remaining, 0) & "s"
        Else
            etaText = " | ETA: ~" & Int(remaining / 60) & "m"
        End If
    End If

    ' Build visual bar: [=========>          ] 45%
    Dim barWidth As Long: barWidth = 20
    Dim filled As Long: filled = CLng(pct * barWidth)
    Dim bar As String: bar = "[" & String(filled, "=")
    If filled < barWidth Then bar = bar & ">"
    bar = bar & String(barWidth - filled, " ") & "]"

    Dim statusMsg As String
    statusMsg = m_TaskName & " " & bar & " " & Format(pct, "0%")
    If detail <> "" Then statusMsg = statusMsg & " | " & detail
    statusMsg = statusMsg & " | " & elapsedText & etaText

    Application.StatusBar = statusMsg
    DoEvents
End Sub

'===============================================================================
' EndProgress - Close the progress bar
'===============================================================================
Public Sub EndProgress()
    Application.StatusBar = m_TaskName & " - Complete!"
    DoEvents
    Application.Wait Now + TimeSerial(0, 0, 1)
    Application.StatusBar = False
    m_TotalSteps = 0
    m_TaskName = ""
End Sub
