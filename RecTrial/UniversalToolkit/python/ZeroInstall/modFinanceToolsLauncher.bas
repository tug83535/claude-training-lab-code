Attribute VB_Name = "modFinanceToolsLauncher"
Option Explicit

' ---------------------------------------------------------------------------
' modFinanceToolsLauncher
' Launches finance_automation_launcher.py via bundled Python 3.11 embeddable.
'
' FOLDER STRUCTURE REQUIRED (relative to FinanceTools.xlsm):
'   FinanceTools.xlsm
'   python\
'     python-embedded\
'       python.exe
'   scripts\
'     finance_automation_launcher.py
'   outputs\
' ---------------------------------------------------------------------------

Public Sub LaunchFinanceTools()

    Dim pyExe    As String
    Dim pyScript As String
    Dim wsh      As Object
    Dim fso      As Object

    pyExe    = ThisWorkbook.Path & "\python\python-embedded\python.exe"
    pyScript = ThisWorkbook.Path & "\scripts\finance_automation_launcher.py"

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(pyExe) Then
        MsgBox "Finance Tools could not start." & vbNewLine & vbNewLine & _
               "Python not found at:" & vbNewLine & _
               "  " & pyExe & vbNewLine & vbNewLine & _
               "Make sure FinanceTools.xlsm is in the same folder as python\ and scripts\." & vbNewLine & vbNewLine & _
               "Contact Connor in Finance & Accounting if this persists.", _
               vbCritical, "Finance Tools"
        Set fso = Nothing
        Exit Sub
    End If

    If Not fso.FileExists(pyScript) Then
        MsgBox "Finance Tools could not start." & vbNewLine & vbNewLine & _
               "Launcher script not found at:" & vbNewLine & _
               "  " & pyScript & vbNewLine & vbNewLine & _
               "Contact Connor in Finance & Accounting.", _
               vbCritical, "Finance Tools"
        Set fso = Nothing
        Exit Sub
    End If

    Set fso = Nothing

    ' Set working directory to the workbook folder, then use relative paths.
    ' This avoids all CMD quoting issues — no spaces in relative paths.
    Set wsh = CreateObject("WScript.Shell")
    wsh.CurrentDirectory = ThisWorkbook.Path
    wsh.Run "cmd.exe /k python\python-embedded\python.exe scripts\finance_automation_launcher.py", 1, False
    Set wsh = Nothing

End Sub
