Attribute VB_Name = "modPerformance"
Option Explicit

'===============================================================================
' modPerformance - TurboMode & Timer Utilities
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Wrap Application-level toggles that accelerate VBA execution.
'           TurboOn/TurboOff are the bookends for every long-running macro.
'           Includes precision timer for benchmarking.
'
' USAGE:    Call TurboOn at the start of any macro, TurboOff in the exit
'           and error handler. Use ElapsedSeconds() for performance logging.
'
' VERSION:  2.1.0
' CHANGES:  v2.0 -> v2.1:
'           + Fixed ISSUE-005 (BUG-004): ElapsedSeconds midnight rollover.
'             Timer resets to 0 at midnight; elapsed goes negative. Now adds
'             86400 (seconds per day) when elapsed < 0 to handle the wrap.
'===============================================================================

Private m_CalcMode      As XlCalculation
Private m_ScreenUpdate  As Boolean
Private m_EnableEvents  As Boolean
Private m_DisplayAlerts As Boolean
Private m_StartTime     As Double

'===============================================================================
' TurboOn - Suppress UI refreshes, disable events, manual calc
'===============================================================================
Public Sub TurboOn()
    With Application
        m_CalcMode = .Calculation
        m_ScreenUpdate = .ScreenUpdating
        m_EnableEvents = .EnableEvents
        m_DisplayAlerts = .DisplayAlerts
        
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .Cursor = xlWait
    End With
    m_StartTime = Timer
End Sub

'===============================================================================
' TurboOff - Restore all settings to pre-TurboOn state
'===============================================================================
Public Sub TurboOff()
    With Application
        .Calculation = m_CalcMode
        .ScreenUpdating = m_ScreenUpdate
        .EnableEvents = m_EnableEvents
        .DisplayAlerts = m_DisplayAlerts
        .Cursor = xlDefault
        .StatusBar = False
    End With
End Sub

'===============================================================================
' ForceRecalc - Full workbook recalculation
'===============================================================================
Public Sub ForceRecalc()
    Application.Calculate
    DoEvents
End Sub

'===============================================================================
' ElapsedSeconds - Seconds since TurboOn was called
' FIX (v2.1): VBA Timer resets to 0 at midnight. If a macro starts at 23:59:50
' and finishes at 00:00:05, raw elapsed = -86395. Adding 86400 corrects this.
'===============================================================================
Public Function ElapsedSeconds() As Double
    Dim elapsed As Double
    elapsed = Timer - m_StartTime
    If elapsed < 0 Then elapsed = elapsed + 86400  ' midnight rollover
    ElapsedSeconds = Round(elapsed, 2)
End Function

'===============================================================================
' UpdateStatus - Write to status bar with optional percentage
'===============================================================================
Public Sub UpdateStatus(ByVal msg As String, Optional ByVal pct As Double = -1)
    If pct >= 0 Then
        Application.StatusBar = msg & " - " & Format(pct, "0%") & " complete"
    Else
        Application.StatusBar = msg
    End If
    DoEvents
End Sub
