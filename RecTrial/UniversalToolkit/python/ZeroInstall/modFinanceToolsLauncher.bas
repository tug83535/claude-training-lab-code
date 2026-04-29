Attribute VB_Name = "modFinanceToolsLauncher"
Option Explicit

' ---------------------------------------------------------------------------
' modFinanceToolsLauncher
' Launches finance_automation_launcher.py via bundled Python 3.11 embeddable.
'
' FOLDER STRUCTURE REQUIRED (relative to FinanceTools.xlsm):
'   FinanceTools.xlsm          <- this workbook
'   python\
'     python-embedded\
'       python.exe             <- bundled Python 3.11 (no system install needed)
'   scripts\
'     finance_automation_launcher.py   <- the CLI menu script
'   outputs\                   <- created automatically on first run
'
' HOW TO WIRE UP THE BUTTON:
'   1. Open FinanceTools.xlsm in Excel
'   2. Developer tab > Insert > Button (Form Control)
'   3. Draw the button on the sheet
'   4. In "Assign Macro" dialog, pick LaunchFinanceTools
'   5. Right-click the button > Edit Text > type "Finance Tools"
'   6. Done. Clicking the button opens the numbered menu.
' ---------------------------------------------------------------------------

Public Sub LaunchFinanceTools()

    Dim pyExe    As String
    Dim pyScript As String
    Dim cmd      As String

    ' Build absolute paths relative to this workbook — works from any machine
    pyExe    = ThisWorkbook.Path & "\python\python-embedded\python.exe"
    pyScript = ThisWorkbook.Path & "\scripts\finance_automation_launcher.py"

    ' --- Guard: Python executable missing ---
    If Dir(pyExe) = "" Then
        MsgBox "Finance Tools could not start." & vbNewLine & vbNewLine & _
               "Python not found at:" & vbNewLine & _
               "  " & pyExe & vbNewLine & vbNewLine & _
               "Make sure the zip was fully unzipped and FinanceTools.xlsm" & vbNewLine & _
               "is in the same folder as the python\ and scripts\ folders." & vbNewLine & vbNewLine & _
               "Contact Connor in Finance & Accounting if this persists.", _
               vbCritical, "Finance Tools"
        Exit Sub
    End If

    ' --- Guard: launcher script missing ---
    If Dir(pyScript) = "" Then
        MsgBox "Finance Tools could not start." & vbNewLine & vbNewLine & _
               "Launcher script not found at:" & vbNewLine & _
               "  " & pyScript & vbNewLine & vbNewLine & _
               "Contact Connor in Finance & Accounting.", _
               vbCritical, "Finance Tools"
        Exit Sub
    End If

    ' --- Launch ---
    ' /k keeps the CMD window open after Python exits so coworkers can read output.
    ' Paths quoted with Chr(34) to handle spaces in Windows user paths.
    cmd = "cmd.exe /k " & Chr(34) & pyExe & Chr(34) & " " & Chr(34) & pyScript & Chr(34)
    Shell cmd, vbNormalFocus

End Sub
