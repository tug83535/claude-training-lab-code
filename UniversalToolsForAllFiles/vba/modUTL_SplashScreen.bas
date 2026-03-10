Attribute VB_Name = "modUTL_SplashScreen"
Option Explicit

'===============================================================================
' modUTL_SplashScreen - Universal Branded Splash Screen
' Universal Toolkit - Works on ANY Excel file
'===============================================================================
' PURPOSE:  Displays a professional branded welcome screen when any workbook
'           opens. Shows workbook name, sheet count, and a friendly greeting.
'           Coworkers can add this to any file for a polished feel.
'
' USAGE:    Add to ThisWorkbook module:
'             Private Sub Workbook_Open()
'                 modUTL_SplashScreen.ShowSplash
'             End Sub
'
' CUSTOMIZATION:
'   Change SPLASH_TITLE and SPLASH_SUBTITLE constants below to match your
'   department or project. The splash auto-detects workbook name and stats.
'
' DEPENDENCIES: None (fully standalone)
' VERSION:  1.0.0
'===============================================================================

' --- Customize These ---
Private Const SPLASH_TITLE    As String = "iPipeline Finance & Accounting"
Private Const SPLASH_SUBTITLE As String = "Excel Automation Toolkit"
Private Const SPLASH_COLOR    As Long = 7364913  ' RGB(49,113,112) iPipeline teal

'===============================================================================
' ShowSplash - Display branded welcome message
'===============================================================================
Public Sub ShowSplash()
    On Error GoTo ErrHandler

    Dim wbName As String: wbName = ThisWorkbook.Name
    Dim sheetCount As Long: sheetCount = ThisWorkbook.Worksheets.Count

    Dim border As String: border = String(50, "=")
    Dim msg As String

    msg = border & vbCrLf & vbCrLf
    msg = msg & "   " & SPLASH_TITLE & vbCrLf
    msg = msg & "   " & SPLASH_SUBTITLE & vbCrLf & vbCrLf
    msg = msg & border & vbCrLf & vbCrLf
    msg = msg & "   Workbook: " & wbName & vbCrLf
    msg = msg & "   Sheets: " & sheetCount & vbCrLf
    msg = msg & "   Opened: " & Format(Now, "mmmm d, yyyy h:mm AM/PM") & vbCrLf & vbCrLf
    msg = msg & border & vbCrLf & vbCrLf
    msg = msg & "   Welcome! Click OK to get started." & vbCrLf

    MsgBox msg, vbInformation, SPLASH_TITLE

    Exit Sub

ErrHandler:
    ' Splash should never crash anything
    On Error Resume Next
End Sub

'===============================================================================
' ShowSplashCustom - Display splash with custom title and subtitle
' For coworkers who want to customize without editing constants.
'===============================================================================
Public Sub ShowSplashCustom(ByVal title As String, ByVal subtitle As String)
    On Error GoTo ErrHandler

    Dim wbName As String: wbName = ThisWorkbook.Name
    Dim sheetCount As Long: sheetCount = ThisWorkbook.Worksheets.Count

    Dim border As String: border = String(50, "=")
    Dim msg As String

    msg = border & vbCrLf & vbCrLf
    msg = msg & "   " & title & vbCrLf
    msg = msg & "   " & subtitle & vbCrLf & vbCrLf
    msg = msg & border & vbCrLf & vbCrLf
    msg = msg & "   Workbook: " & wbName & vbCrLf
    msg = msg & "   Sheets: " & sheetCount & vbCrLf
    msg = msg & "   Opened: " & Format(Now, "mmmm d, yyyy h:mm AM/PM") & vbCrLf & vbCrLf
    msg = msg & border & vbCrLf & vbCrLf
    msg = msg & "   Welcome! Click OK to get started." & vbCrLf

    MsgBox msg, vbInformation, title

    Exit Sub

ErrHandler:
    On Error Resume Next
End Sub
