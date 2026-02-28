Attribute VB_Name = "modAWSRecompute"
Option Explicit

'===============================================================================
' modAWSRecompute - AWS Allocation Recalculation & Validation
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Reads the AWS Allocation sheet, validates that compute share
'           percentages sum to 100%, recalculates each product's allocated
'           AWS cost, and writes results back to the sheet.
'
' PUBLIC SUBS:
'   ValidateAndRecalcAWS - Main entry (Action #14)
'
' DEPENDENCIES: modConfig, modPerformance, modLogger
' VERSION:  2.1.0
'===============================================================================

'===============================================================================
' ValidateAndRecalcAWS - Validate shares and recalculate allocations
'===============================================================================
Public Sub ValidateAndRecalcAWS()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_AWS) Then
        MsgBox "AWS Allocation sheet not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Validating AWS allocations...", 0.1

    Dim wsAWS As Worksheet: Set wsAWS = ThisWorkbook.Worksheets(SH_AWS)
    Dim lastRow As Long: lastRow = modConfig.LastRow(wsAWS, 1)

    If lastRow < DATA_ROW_AWS Then
        modPerformance.TurboOff
        MsgBox "No AWS allocation data found.", vbInformation, APP_NAME
        Exit Sub
    End If

    ' Find key columns by header
    Dim colProduct As Long: colProduct = 1
    Dim colShare As Long: colShare = modConfig.FindColByHeader(wsAWS, "Share", HDR_ROW_AWS)
    If colShare = 0 Then colShare = modConfig.FindColByHeader(wsAWS, "%", HDR_ROW_AWS)
    If colShare = 0 Then colShare = 2  ' Default

    Dim colPool As Long: colPool = modConfig.FindColByHeader(wsAWS, "Pool", HDR_ROW_AWS)
    If colPool = 0 Then colPool = modConfig.FindColByHeader(wsAWS, "AWS", HDR_ROW_AWS)
    If colPool = 0 Then colPool = 3  ' Default

    Dim colAllocated As Long: colAllocated = modConfig.FindColByHeader(wsAWS, "Allocated", HDR_ROW_AWS)
    If colAllocated = 0 Then colAllocated = modConfig.LastCol(wsAWS, HDR_ROW_AWS)

    ' Read products, shares, and pool
    Dim totalShare As Double: totalShare = 0
    Dim productCount As Long: productCount = 0
    Dim issues As String: issues = ""
    Dim r As Long

    For r = DATA_ROW_AWS To lastRow
        Dim pName As String: pName = Trim(CStr(wsAWS.Cells(r, colProduct).Value))
        If pName <> "" Then
            productCount = productCount + 1
            Dim share As Double: share = modConfig.SafeNum(wsAWS.Cells(r, colShare).Value)

            ' Convert to decimal if stored as whole number (e.g. 55 instead of 0.55)
            If share > 1 Then share = share / 100

            totalShare = totalShare + share

            ' Validate individual share
            If share <= 0 Then
                issues = issues & "  - " & pName & " has zero or negative share" & vbCrLf
            End If
        End If
    Next r

    modPerformance.UpdateStatus "Recalculating allocations...", 0.5

    ' Validate total shares
    Dim shareOK As Boolean: shareOK = (Abs(totalShare - 1) < 0.005)
    If Not shareOK Then
        issues = issues & "  - Share percentages sum to " & Format(totalShare * 100, "0.0") & _
                 "% (should be 100%)" & vbCrLf
    End If

    ' Recalculate allocated amounts
    ' Find the monthly pool amount (look for a total/pool row or use a known cell)
    Dim monthlyPool As Double: monthlyPool = 0
    For r = DATA_ROW_AWS To lastRow
        Dim poolVal As Double: poolVal = modConfig.SafeNum(wsAWS.Cells(r, colPool).Value)
        If poolVal > monthlyPool Then monthlyPool = poolVal
    Next r

    ' Write recalculated allocations
    Dim recalcCount As Long: recalcCount = 0
    For r = DATA_ROW_AWS To lastRow
        Dim prod As String: prod = Trim(CStr(wsAWS.Cells(r, colProduct).Value))
        If prod <> "" Then
            Dim pShare As Double: pShare = modConfig.SafeNum(wsAWS.Cells(r, colShare).Value)
            If pShare > 1 Then pShare = pShare / 100

            Dim allocated As Double: allocated = monthlyPool * pShare
            Dim oldVal As Double: oldVal = modConfig.SafeNum(wsAWS.Cells(r, colAllocated).Value)

            ' Write the recalculated value
            wsAWS.Cells(r, colAllocated).Value = allocated
            wsAWS.Cells(r, colAllocated).NumberFormat = "$#,##0.00"

            ' Highlight if changed
            If Abs(allocated - oldVal) > 0.01 Then
                wsAWS.Cells(r, colAllocated).Interior.Color = RGB(255, 255, 200)
                recalcCount = recalcCount + 1
            End If
        End If
    Next r

    wsAWS.Activate

    Dim elapsed As Double: elapsed = modPerformance.ElapsedSeconds()
    modPerformance.TurboOff

    ' Build result message
    Dim msg As String
    msg = "AWS ALLOCATION RECALCULATION" & vbCrLf & _
          String(35, "=") & vbCrLf & vbCrLf & _
          "Products:       " & productCount & vbCrLf & _
          "Share Total:    " & Format(totalShare * 100, "0.0") & "%" & vbCrLf & _
          "Monthly Pool:   " & Format(monthlyPool, "$#,##0") & vbCrLf & _
          "Values Updated: " & recalcCount & vbCrLf

    If issues <> "" Then
        msg = msg & vbCrLf & "ISSUES FOUND:" & vbCrLf & issues
    Else
        msg = msg & vbCrLf & "All validations passed."
    End If

    Dim logStatus As String
    logStatus = IIf(issues = "", "OK", "WARN")
    modLogger.LogAction "modAWSRecompute", "ValidateAndRecalcAWS", _
        productCount & " products, " & recalcCount & " updated. Share sum=" & Format(totalShare * 100, "0.0") & "%", _
        logStatus

    Dim icon As VbMsgBoxStyle
    icon = IIf(issues = "", vbInformation, vbExclamation)
    MsgBox msg, icon, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modAWSRecompute", "ERROR", Err.Description
    MsgBox "AWS recalculation error: " & Err.Description, vbCritical, APP_NAME
End Sub
