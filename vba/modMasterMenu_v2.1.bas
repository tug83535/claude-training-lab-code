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
'           + ISSUE-006: Expanded from 36 items to 50 items
'           + Split into 3-page InputBox to stay within character limits
'           + All routing now delegates to modFormBuilder.ExecuteAction
'             (single-point routing — no duplicate Select Case)
'           + Ctrl+Shift+M now routes to LaunchCommandCenter (primary)
'           + ShowMasterMenu is the InputBox fallback only
'===============================================================================

'===============================================================================
' ShowMasterMenu - 3-page InputBox fallback (50 items)
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
    End Select
    
    ' User cancelled
    If choice = "" Then Exit Sub
    
    ' Page navigation
    If UCase(choice) = "N" Or UCase(choice) = "NEXT" Then
        If page < 3 Then page = page + 1 Else page = 1
        GoTo PageLoop
    End If
    If UCase(choice) = "P" Or UCase(choice) = "PREV" Then
        If page > 1 Then page = page - 1 Else page = 3
        GoTo PageLoop
    End If
    
    ' Validate numeric input
    If Not IsNumeric(choice) Then
        MsgBox "Please enter an action number (1-50), N for next page, or P for previous.", _
               vbExclamation, APP_NAME
        GoTo PageLoop
    End If
    
    Dim num As Long: num = CLng(choice)
    If num < 1 Or num > 50 Then
        MsgBox "Please enter a number between 1 and 50.", vbExclamation, APP_NAME
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
    ShowPage1 = InputBox( _
        "KEYSTONE BENEFITTECH TOOLKIT v" & APP_VERSION & "  [Page 1/3]" & vbCrLf & _
        String(50, "=") & vbCrLf & vbCrLf & _
        "MONTHLY OPERATIONS" & vbCrLf & _
        "  1.  Generate Monthly Tabs (Apr-Dec)" & vbCrLf & _
        "  2.  Delete Generated Tabs" & vbCrLf & _
        "  3.  Run Reconciliation Checks" & vbCrLf & _
        "  4.  Export Reconciliation Report" & vbCrLf & vbCrLf & _
        "ANALYSIS" & vbCrLf & _
        "  5.  Run Sensitivity Analysis" & vbCrLf & _
        "  6.  Run Variance Analysis" & vbCrLf & vbCrLf & _
        "DATA QUALITY" & vbCrLf & _
        "  7.  Scan for Data Quality Issues" & vbCrLf & _
        "  8.  Fix Text-Stored Numbers" & vbCrLf & _
        "  9.  Fix Duplicate Rows" & vbCrLf & vbCrLf & _
        "REPORTING" & vbCrLf & _
        "  10. Export Report Package (PDF)" & vbCrLf & _
        "  11. Export Active Sheet (PDF)" & vbCrLf & _
        "  12. Build Dashboard Charts" & vbCrLf & vbCrLf & _
        "UTILITIES" & vbCrLf & _
        "  13. Refresh Table of Contents" & vbCrLf & _
        "  14. Recalculate AWS Allocations" & vbCrLf & _
        "  15. Quick Jump to Sheet" & vbCrLf & _
        "  16. Go Home (Report-->)" & vbCrLf & vbCrLf & _
        "DATA & IMPORT" & vbCrLf & _
        "  17. Import GL Data (CSV/Excel)" & vbCrLf & vbCrLf & _
        "FORECASTING" & vbCrLf & _
        "  18. Update Rolling Forecast" & vbCrLf & _
        "  19. Append Month to Trend" & vbCrLf & vbCrLf & _
        "SCENARIOS" & vbCrLf & _
        "  20. Save Scenario" & vbCrLf & vbCrLf & _
        "Enter # (1-50) | N=Next Page | Cancel=Exit:", _
        APP_NAME & " - Command Center")
End Function

'===============================================================================
' ShowPage2 - Items 21-40 (Scenarios, Allocation, Consolidation, VC, Governance)
'===============================================================================
Private Function ShowPage2() As String
    ShowPage2 = InputBox( _
        "KEYSTONE BENEFITTECH TOOLKIT v" & APP_VERSION & "  [Page 2/3]" & vbCrLf & _
        String(50, "=") & vbCrLf & vbCrLf & _
        "SCENARIOS (cont.)" & vbCrLf & _
        "  21. Load Scenario" & vbCrLf & _
        "  22. Compare Scenarios" & vbCrLf & _
        "  23. Delete Scenario" & vbCrLf & vbCrLf & _
        "ALLOCATION" & vbCrLf & _
        "  24. Run Allocation Engine" & vbCrLf & _
        "  25. Allocation Preview" & vbCrLf & vbCrLf & _
        "CONSOLIDATION" & vbCrLf & _
        "  26. Consolidation Menu" & vbCrLf & _
        "  27. Add Entity" & vbCrLf & _
        "  28. Generate Consolidated" & vbCrLf & _
        "  29. List Entities" & vbCrLf & _
        "  30. Add Elimination" & vbCrLf & vbCrLf & _
        "VERSION CONTROL" & vbCrLf & _
        "  31. Version Control Menu" & vbCrLf & _
        "  32. Save Version" & vbCrLf & _
        "  33. Compare Versions" & vbCrLf & _
        "  34. Restore Version" & vbCrLf & _
        "  35. List Versions" & vbCrLf & vbCrLf & _
        "GOVERNANCE" & vbCrLf & _
        "  36. Auto-Documentation" & vbCrLf & _
        "  37. Change Management Menu" & vbCrLf & _
        "  38. Add Change Request" & vbCrLf & _
        "  39. Update CR Status" & vbCrLf & _
        "  40. CR Summary Report" & vbCrLf & vbCrLf & _
        "Enter # (1-50) | N=Next | P=Prev | Cancel=Exit:", _
        APP_NAME & " - Command Center")
End Function

'===============================================================================
' ShowPage3 - Items 41-50 (Admin & Testing, Advanced)
'===============================================================================
Private Function ShowPage3() As String
    ShowPage3 = InputBox( _
        "KEYSTONE BENEFITTECH TOOLKIT v" & APP_VERSION & "  [Page 3/3]" & vbCrLf & _
        String(50, "=") & vbCrLf & vbCrLf & _
        "ADMIN & TESTING" & vbCrLf & _
        "  41. View Audit Log" & vbCrLf & _
        "  42. Export Audit Log" & vbCrLf & _
        "  43. Clear Audit Log" & vbCrLf & _
        "  44. Full Integration Test" & vbCrLf & _
        "  45. Quick Health Check" & vbCrLf & vbCrLf & _
        "ADVANCED" & vbCrLf & _
        "  46. Variance Commentary" & vbCrLf & _
        "  47. Cross-Sheet Validation" & vbCrLf & _
        "  48. Executive Mode Toggle" & vbCrLf & _
        "  49. Force Recalculate All" & vbCrLf & _
        "  50. About This Toolkit" & vbCrLf & vbCrLf & _
        String(50, "-") & vbCrLf & _
        "TIP: Install the Command Center UserForm for" & vbCrLf & _
        "a better experience. Press Ctrl+Shift+M after" & vbCrLf & _
        "running modFormBuilder.BuildCommandCenter." & vbCrLf & vbCrLf & _
        "Enter # (1-50) | P=Prev | Cancel=Exit:", _
        APP_NAME & " - Command Center")
End Function
