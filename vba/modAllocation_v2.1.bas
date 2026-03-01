Attribute VB_Name = "modAllocation"
Option Explicit

'===============================================================================
' modAllocation - Cost Allocation Engine
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Reads GL transactions and allocates shared costs across products
'           using the shares defined in the Assumptions sheet. Outputs an
'           Allocation Output sheet with the full breakdown.
'
' PUBLIC SUBS:
'   RunAllocationEngine  - Full allocation run (Action #24)
'   AllocationPreview    - What-if preview with modified shares (Action #25)
'
' DEPENDENCIES: modConfig, modPerformance, modLogger
' VERSION:  2.1.0
'===============================================================================

'===============================================================================
' RunAllocationEngine - Allocate shared costs to products
'===============================================================================
Public Sub RunAllocationEngine()
    On Error GoTo ErrHandler

    If Not modConfig.SheetExists(SH_GL) Then
        MsgBox "GL data sheet (" & SH_GL & ") not found.", vbCritical, APP_NAME
        Exit Sub
    End If

    modPerformance.TurboOn
    modPerformance.UpdateStatus "Running allocation engine...", 0.05

    Dim wsGL As Worksheet: Set wsGL = ThisWorkbook.Worksheets(SH_GL)
    Dim glLastRow As Long: glLastRow = modConfig.LastRow(wsGL, COL_GL_ID)

    If glLastRow < DATA_ROW_GL Then
        modPerformance.TurboOff
        MsgBox "No GL data found.", vbInformation, APP_NAME
        Exit Sub
    End If

    ' Get products and build allocation shares from Assumptions
    Dim products As Variant: products = modConfig.GetProducts()
    Dim prodCount As Long: prodCount = UBound(products) + 1
    Dim shares() As Double
    ReDim shares(0 To prodCount - 1)

    ' Default equal shares
    Dim p As Long
    For p = 0 To prodCount - 1
        shares(p) = 1 / prodCount
    Next p

    ' Try to read shares from Assumptions
    If modConfig.SheetExists(SH_ASSUMPTIONS) Then
        Dim wsA As Worksheet: Set wsA = ThisWorkbook.Worksheets(SH_ASSUMPTIONS)
        Dim aLastRow As Long: aLastRow = modConfig.LastRow(wsA, 1)
        Dim r As Long
        For r = DATA_ROW_ASSUME To aLastRow
            Dim drvName As String: drvName = LCase(Trim(CStr(wsA.Cells(r, 1).Value)))
            For p = 0 To prodCount - 1
                If InStr(drvName, LCase(CStr(products(p)))) > 0 And _
                   InStr(drvName, "share") > 0 Then
                    Dim sVal As Double: sVal = modConfig.SafeNum(wsA.Cells(r, 2).Value)
                    If sVal > 1 Then sVal = sVal / 100
                    If sVal > 0 Then shares(p) = sVal
                End If
            Next p
        Next r
    End If

    modPerformance.UpdateStatus "Calculating allocations...", 0.3

    ' Sum GL by category (shared vs direct)
    Dim totalAmount As Double: totalAmount = 0
    Dim directByProd() As Double
    ReDim directByProd(0 To prodCount - 1)
    Dim sharedAmount As Double: sharedAmount = 0

    For r = DATA_ROW_GL To glLastRow
        Dim amt As Double: amt = modConfig.SafeNum(wsGL.Cells(r, COL_GL_AMOUNT).Value)
        Dim prod As String: prod = Trim(CStr(wsGL.Cells(r, COL_GL_PRODUCT).Value))
        totalAmount = totalAmount + amt

        ' Check if directly assigned to a product
        Dim isDirect As Boolean: isDirect = False
        For p = 0 To prodCount - 1
            If InStr(1, prod, CStr(products(p)), vbTextCompare) > 0 Then
                directByProd(p) = directByProd(p) + amt
                isDirect = True
                Exit For
            End If
        Next p
        If Not isDirect Then sharedAmount = sharedAmount + amt
    Next r

    ' Create output sheet
    modConfig.SafeDeleteSheet SH_ALLOC_OUT
    Dim wsOut As Worksheet
    Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsOut.Name = SH_ALLOC_OUT

    ' Title
    wsOut.Range("A1").Value = "COST ALLOCATION OUTPUT"
    wsOut.Range("A1").Font.Bold = True
    wsOut.Range("A1").Font.Size = 14
    wsOut.Range("A1").Font.Color = CLR_NAVY
    wsOut.Range("A2").Value = "Run Date: " & Format(Now, "yyyy-mm-dd hh:mm") & _
        "  |  Total GL: " & Format(totalAmount, "$#,##0")

    modConfig.StyleHeader wsOut, 4, _
        Array("Product", "Direct Costs", "Share %", "Allocated Shared", "Total Allocated")

    Dim outRow As Long: outRow = 5
    For p = 0 To prodCount - 1
        wsOut.Cells(outRow, 1).Value = CStr(products(p))
        wsOut.Cells(outRow, 1).Font.Bold = True
        wsOut.Cells(outRow, 2).Value = directByProd(p)
        wsOut.Cells(outRow, 2).NumberFormat = "$#,##0"
        wsOut.Cells(outRow, 3).Value = shares(p)
        wsOut.Cells(outRow, 3).NumberFormat = "0.0%"
        wsOut.Cells(outRow, 4).Value = sharedAmount * shares(p)
        wsOut.Cells(outRow, 4).NumberFormat = "$#,##0"
        wsOut.Cells(outRow, 5).Value = directByProd(p) + (sharedAmount * shares(p))
        wsOut.Cells(outRow, 5).NumberFormat = "$#,##0"
        wsOut.Cells(outRow, 5).Font.Bold = True

        If outRow Mod 2 = 1 Then
            wsOut.Range("A" & outRow & ":E" & outRow).Interior.Color = CLR_ALT_ROW
        End If
        outRow = outRow + 1
    Next p

    ' Totals row
    wsOut.Cells(outRow, 1).Value = "TOTAL"
    wsOut.Cells(outRow, 1).Font.Bold = True
    wsOut.Cells(outRow, 2).Value = totalAmount - sharedAmount
    wsOut.Cells(outRow, 3).Value = 1
    wsOut.Cells(outRow, 4).Value = sharedAmount
    wsOut.Cells(outRow, 5).Value = totalAmount
    Dim totRng As Range: Set totRng = wsOut.Range("A" & outRow & ":E" & outRow)
    totRng.Font.Bold = True
    totRng.NumberFormat = "$#,##0"
    totRng.Cells(1, 3).NumberFormat = "0.0%"
    totRng.Borders(xlEdgeTop).LineStyle = xlDouble

    wsOut.Columns("A:E").AutoFit
    wsOut.Tab.Color = RGB(0, 176, 80)
    wsOut.Activate

    Dim elapsed As Double: elapsed = modPerformance.ElapsedSeconds()
    modPerformance.TurboOff

    modLogger.LogAction "modAllocation", "RunAllocationEngine", _
        prodCount & " products, " & Format(totalAmount, "$#,##0") & " allocated"

    MsgBox "ALLOCATION COMPLETE" & vbCrLf & String(25, "=") & vbCrLf & vbCrLf & _
           "Total GL Amount:  " & Format(totalAmount, "$#,##0") & vbCrLf & _
           "Direct Costs:     " & Format(totalAmount - sharedAmount, "$#,##0") & vbCrLf & _
           "Shared Costs:     " & Format(sharedAmount, "$#,##0") & vbCrLf & _
           "Products:         " & prodCount & vbCrLf & vbCrLf & _
           "Results on '" & SH_ALLOC_OUT & "' sheet.", _
           vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    modPerformance.TurboOff
    modLogger.LogAction "modAllocation", "ERROR", Err.Description
    MsgBox "Allocation error: " & Err.Description, vbCritical, APP_NAME
End Sub

'===============================================================================
' AllocationPreview - What-if with user-modified shares
'===============================================================================
Public Sub AllocationPreview()
    On Error GoTo ErrHandler

    Dim products As Variant: products = modConfig.GetProducts()

    ' Ask user for modified shares
    Dim msg As String
    msg = "Enter allocation shares (must total 100%):" & vbCrLf & vbCrLf & _
          "Format: product1%,product2%,product3%,product4%" & vbCrLf & _
          "Example: 55,28,12,5" & vbCrLf & vbCrLf & _
          "Products: "
    Dim p As Long
    For p = 0 To UBound(products)
        msg = msg & CStr(products(p))
        If p < UBound(products) Then msg = msg & ", "
    Next p

    Dim userInput As String
    userInput = InputBox(msg, APP_NAME & " - Allocation Preview")
    If userInput = "" Then Exit Sub

    ' Parse shares
    Dim parts As Variant: parts = Split(userInput, ",")
    If UBound(parts) <> UBound(products) Then
        MsgBox "Expected " & (UBound(products) + 1) & " values, got " & (UBound(parts) + 1) & ".", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim shares() As Double
    ReDim shares(0 To UBound(parts))
    Dim total As Double: total = 0
    For p = 0 To UBound(parts)
        shares(p) = Val(Trim(CStr(parts(p)))) / 100
        total = total + shares(p)
    Next p

    If Abs(total - 1) > 0.01 Then
        MsgBox "Shares total " & Format(total * 100, "0.0") & "% — must equal 100%.", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    ' Show preview
    msg = "ALLOCATION PREVIEW" & vbCrLf & String(25, "=") & vbCrLf & vbCrLf
    For p = 0 To UBound(products)
        msg = msg & CStr(products(p)) & ": " & Format(shares(p) * 100, "0.0") & "%" & vbCrLf
    Next p

    msg = msg & vbCrLf & "Run the full allocation engine (Action #24)" & vbCrLf & _
          "with these shares by updating the Assumptions sheet first."

    modLogger.LogAction "modAllocation", "AllocationPreview", "Preview with custom shares"
    MsgBox msg, vbInformation, APP_NAME
    Exit Sub

ErrHandler:
    MsgBox "Preview error: " & Err.Description, vbCritical, APP_NAME
End Sub
