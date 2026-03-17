Attribute VB_Name = "modMasterMenu"
Option Explicit

'===============================================================================
' modMasterMenu - Central Command Panel (InputBox Fallback)
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Provides the 50-item InputBox menu as a fallback when the
'           frmCommandCenter UserForm is not installed.
'
'           Primary flow:
'             Ctrl+Shift+M -> modFormBuilder.LaunchCommandCenter
'               -> tries frmCommandCenter UserForm
'               -> falls back to modMasterMenu.ShowMasterMenu (this module)
'
'           All action routing delegates to modFormBuilder.ExecuteAction
'           so there is ONE routing table to maintain.
'
' VERSION:  2.1.0
' CHANGES:  v2.0 -> v2.1:
'           + ISSUE-006: Expanded from 36 items to 62 items
'           + Split into 4-page InputBox to stay within character limits
'           + All routing now delegates to modFormBuilder.ExecuteAction
'             (single-point routing — no duplicate Select Case)
'           + Ctrl+Shift+M now routes to LaunchCommandCenter (primary)
'           + ShowMasterMenu is the InputBox fallback only
'===============================================================================

'===============================================================================
' ShowMasterMenu - 4-page InputBox fallback (62 items)
' Called by modFormBuilder.LaunchCommandCenter when UserForm unavailable.
'===============================================================================
Public Sub ShowMasterMenu()
    On Error GoTo ErrHandler
    
    Dim choice As String
    Dim page As Long: page = 1
    
PageLoop:
    Select Case page
        Case 1: choice = ShowPage1()
        Case 2: choice = ShowPage2()
        Case 3: choice = ShowPage3()
        Case 4: choice = ShowPage4()
    End Select
    
    ' User cancelled
    If choice = "" Then Exit Sub
    
    ' Page navigation
    If UCase(choice) = "N" Or UCase(choice) = "NEXT" Then
        If page < 4 Then page = page + 1 Else page = 1
        GoTo PageLoop
    End If
    If UCase(choice) = "P" Or UCase(choice) = "PREV" Then
        If page > 1 Then page = page - 1 Else page = 4
        GoTo PageLoop
    End If

    ' Validate numeric input
    If Not IsNumeric(choice) Then
        MsgBox "Please enter an action number (1-62), N for next page, or P for previous.", _
               vbExclamation, APP_NAME
        GoTo PageLoop
    End If

    Dim num As Long: num = CLng(choice)
    If num < 1 Or num > 62 Then
        MsgBox "Please enter a number between 1 and 62.", vbExclamation, APP_NAME
        GoTo PageLoop
    End If
    
    ' Delegate to single routing table in modFormBuilder
    modFormBuilder.ExecuteAction num
    Exit Sub
    
ErrHandler:
    MsgBox "Menu error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' ShowPage1 - Items 1-20 (Operations, Analysis, Data Quality, Reporting, Utilities)
'===============================================================================
Private Function ShowPage1() As String
    Dim msg As String
    msg = "KEYSTONE BENEFITTECH TOOLKIT v" & APP_VERSION & "  [Page 1/4]" & vbCrLf
    msg = msg & String(50, "=") & vbCrLf & vbCrLf
    msg = msg & "MONTHLY OPERATIONS" & vbCrLf
    msg = msg & "  1.  Generate Monthly Tabs (Apr-Dec)" & vbCrLf
    msg = msg & "  2.  Delete Generated Tabs" & vbCrLf
    msg = msg & "  3.  Run Reconciliation Checks" & vbCrLf
    msg = msg & "  4.  Export Reconciliation Report" & vbCrLf & vbCrLf
    msg = msg & "ANALYSIS" & vbCrLf
    msg = msg & "  5.  Run Sensitivity Analysis" & vbCrLf
    msg = msg & "  6.  Run Variance Analysis" & vbCrLf & vbCrLf
    msg = msg & "DATA QUALITY" & vbCrLf
    msg = msg & "  7.  Scan for Data Quality Issues" & vbCrLf
    msg = msg & "  8.  Fix Text-Stored Numbers" & vbCrLf
    msg = msg & "  9.  Fix Duplicate Rows" & vbCrLf & vbCrLf
    msg = msg & "REPORTING" & vbCrLf
    msg = msg & "  10. Export Report Package (PDF)" & vbCrLf
    msg = msg & "  11. Export Active Sheet (PDF)" & vbCrLf
    msg = msg & "  12. Build Dashboard Charts" & vbCrLf & vbCrLf
    msg = msg & "UTILITIES" & vbCrLf
    msg = msg & "  13. Refresh Table of Contents" & vbCrLf
    msg = msg & "  14. Recalculate AWS Allocations" & vbCrLf
    msg = msg & "  15. Quick Jump to Sheet" & vbCrLf
    msg = msg & "  16. Go Home (Report-->)" & vbCrLf & vbCrLf
    msg = msg & "DATA & IMPORT" & vbCrLf
    msg = msg & "  17. Import GL Data (CSV/Excel)" & vbCrLf & vbCrLf
    msg = msg & "FORECASTING" & vbCrLf
    msg = msg & "  18. Update Rolling Forecast" & vbCrLf
    msg = msg & "  19. Append Month to Trend" & vbCrLf & vbCrLf
    msg = msg & "SCENARIOS" & vbCrLf
    msg = msg & "  20. Save Scenario" & vbCrLf & vbCrLf
    msg = msg & "Enter # (1-62) | N=Next Page | Cancel=Exit:"
    ShowPage1 = InputBox(msg, APP_NAME & " - Command Center")
End Function

'===============================================================================
' ShowPage2 - Items 21-40 (Scenarios, Allocation, Consolidation, VC, Governance)
'===============================================================================
Private Function ShowPage2() As String
    Dim msg As String
    msg = "KEYSTONE BENEFITTECH TOOLKIT v" & APP_VERSION & "  [Page 2/4]" & vbCrLf
    msg = msg & String(50, "=") & vbCrLf & vbCrLf
    msg = msg & "SCENARIOS (cont.)" & vbCrLf
    msg = msg & "  21. Load Scenario" & vbCrLf
    msg = msg & "  22. Compare Scenarios" & vbCrLf
    msg = msg & "  23. Delete Scenario" & vbCrLf & vbCrLf
    msg = msg & "ALLOCATION" & vbCrLf
    msg = msg & "  24. Run Allocation Engine" & vbCrLf
    msg = msg & "  25. Allocation Preview" & vbCrLf & vbCrLf
    msg = msg & "CONSOLIDATION" & vbCrLf
    msg = msg & "  26. Consolidation Menu" & vbCrLf
    msg = msg & "  27. Add Entity" & vbCrLf
    msg = msg & "  28. Generate Consolidated" & vbCrLf
    msg = msg & "  29. List Entities" & vbCrLf
    msg = msg & "  30. Add Elimination" & vbCrLf & vbCrLf
    msg = msg & "VERSION CONTROL" & vbCrLf
    msg = msg & "  31. Version Control Menu" & vbCrLf
    msg = msg & "  32. Save Version" & vbCrLf
    msg = msg & "  33. Compare Versions" & vbCrLf
    msg = msg & "  34. Restore Version" & vbCrLf
    msg = msg & "  35. List Versions" & vbCrLf & vbCrLf
    msg = msg & "GOVERNANCE" & vbCrLf
    msg = msg & "  36. Auto-Documentation" & vbCrLf
    msg = msg & "  37. Change Management Menu" & vbCrLf
    msg = msg & "  38. Add Change Request" & vbCrLf
    msg = msg & "  39. Update CR Status" & vbCrLf
    msg = msg & "  40. CR Summary Report" & vbCrLf & vbCrLf
    msg = msg & "Enter # (1-62) | N=Next | P=Prev | Cancel=Exit:"
    ShowPage2 = InputBox(msg, APP_NAME & " - Command Center")
End Function

'===============================================================================
' ShowPage3 - Items 41-50 (Admin & Testing, Advanced)
'===============================================================================
Private Function ShowPage3() As String
    Dim msg As String
    msg = "KEYSTONE BENEFITTECH TOOLKIT v" & APP_VERSION & "  [Page 3/4]" & vbCrLf
    msg = msg & String(50, "=") & vbCrLf & vbCrLf
    msg = msg & "ADMIN & TESTING" & vbCrLf
    msg = msg & "  41. View Audit Log" & vbCrLf
    msg = msg & "  42. Export Audit Log" & vbCrLf
    msg = msg & "  43. Clear Audit Log" & vbCrLf
    msg = msg & "  44. Full Integration Test" & vbCrLf
    msg = msg & "  45. Quick Health Check" & vbCrLf & vbCrLf
    msg = msg & "ADVANCED" & vbCrLf
    msg = msg & "  46. Variance Commentary" & vbCrLf
    msg = msg & "  47. Cross-Sheet Validation" & vbCrLf
    msg = msg & "  48. Executive Mode Toggle" & vbCrLf
    msg = msg & "  49. Force Recalculate All" & vbCrLf
    msg = msg & "  50. About This Toolkit" & vbCrLf & vbCrLf
    msg = msg & String(50, "-") & vbCrLf
    msg = msg & "TIP: Install the Command Center UserForm for" & vbCrLf
    msg = msg & "a better experience. Press Ctrl+Shift+M after" & vbCrLf
    msg = msg & "running modFormBuilder.BuildCommandCenter." & vbCrLf & vbCrLf
    msg = msg & "Enter # (1-62) | N=Next | P=Prev | Cancel=Exit:"
    ShowPage3 = InputBox(msg, APP_NAME & " - Command Center")
End Function

'===============================================================================
' ShowPage4 - Items 51-62 (Sheet Tools)
'===============================================================================
Private Function ShowPage4() As String
    Dim msg As String
    msg = "KEYSTONE BENEFITTECH TOOLKIT v" & APP_VERSION & "  [Page 4/4]" & vbCrLf
    msg = msg & String(50, "=") & vbCrLf & vbCrLf
    msg = msg & "SHEET TOOLS" & vbCrLf
    msg = msg & "  51. Delete All Blank Rows" & vbCrLf
    msg = msg & "  52. Unhide All Worksheets" & vbCrLf
    msg = msg & "  53. Sort Sheets Alphabetically" & vbCrLf
    msg = msg & "  54. Toggle Freeze Panes" & vbCrLf
    msg = msg & "  55. Convert Formulas to Values" & vbCrLf
    msg = msg & "  56. AutoFit All Columns" & vbCrLf
    msg = msg & "  57. Protect All Sheets" & vbCrLf
    msg = msg & "  58. Unprotect All Sheets" & vbCrLf
    msg = msg & "  59. Find & Replace (All Sheets)" & vbCrLf
    msg = msg & "  60. Highlight Hardcoded Numbers" & vbCrLf
    msg = msg & "  61. Toggle Presentation Mode" & vbCrLf
    msg = msg & "  62. Unmerge and Fill Down" & vbCrLf & vbCrLf
    msg = msg & String(50, "-") & vbCrLf
    msg = msg & "TIP: Install the Command Center UserForm for" & vbCrLf
    msg = msg & "a better experience. Press Ctrl+Shift+M after" & vbCrLf
    msg = msg & "running modFormBuilder.BuildCommandCenter." & vbCrLf & vbCrLf
    msg = msg & "Enter # (1-62) | P=Prev | Cancel=Exit:"
    ShowPage4 = InputBox(msg, APP_NAME & " - Command Center")
End Function
