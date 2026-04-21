Attribute VB_Name = "modInvoicePDFGenerator"
'===============================================================================
' modInvoicePDFGenerator
' PURPOSE: Generate one fully-branded PDF invoice per customer row in Excel.
'          Each PDF is saved to a per-customer folder and optionally emailed.
'
' WHY THIS IS NOT NATIVE: OneDrive/Excel will "Save as PDF" a single sheet.
'          It will NOT iterate a customer list, render a template, swap in
'          that customer's line items, save a named PDF per customer, and
'          email it. That requires VBA automation.
'
' USE CASE (software business):
'   Billing team generates 300 monthly SaaS invoices — each with the customer's
'   logo, their per-seat usage line items, their account manager signature block,
'   and their contract-specific discount — in 90 seconds.
'
' SHEETS REQUIRED:
'   "InvoiceTemplate"  - the blank styled invoice. Placeholders in [brackets]:
'                        [InvoiceNo], [CustomerName], [InvoiceDate], [DueDate],
'                        [Subtotal], [Tax], [Total], and a "LINES:" marker.
'   "Customers"        - one row per customer: A=CustomerID, B=Name, C=Email,
'                        D=Address, E=TaxRate, F=AMName, G=OutputFolder
'   "LineItems"        - A=CustomerID, B=Description, C=Qty, D=UnitPrice, E=Amount
'===============================================================================
Option Explicit

Private Const START_INVOICE_NO As Long = 24001

'-------------------------------------------------------------------------------
' GenerateAllInvoices - Loops customers, builds invoice, saves PDF, optionally emails.
'-------------------------------------------------------------------------------
Public Sub GenerateAllInvoices(Optional ByVal emailThem As Boolean = False)
    Dim wsC As Worksheet, wsL As Worksheet, wsT As Worksheet
    Dim custRow As Long, lastCust As Long, invNo As Long
    Dim custID As String, outFolder As String, pdfPath As String
    Dim subtotal As Double, tax As Double, total As Double

    Set wsC = ThisWorkbook.Worksheets("Customers")
    Set wsL = ThisWorkbook.Worksheets("LineItems")
    Set wsT = ThisWorkbook.Worksheets("InvoiceTemplate")

    lastCust = wsC.Cells(wsC.Rows.Count, "A").End(xlUp).Row
    invNo = START_INVOICE_NO + GetMaxIssuedInvoiceNo()

    Application.ScreenUpdating = False

    For custRow = 2 To lastCust
        custID = CStr(wsC.Cells(custRow, "A").Value)
        outFolder = CStr(wsC.Cells(custRow, "G").Value)
        If Len(outFolder) = 0 Then outFolder = Environ("USERPROFILE") & "\Documents\Invoices\"
        EnsureFolder outFolder

        ' Render template for this customer
        subtotal = WriteInvoiceFromTemplate(wsT, wsC, wsL, custRow, invNo, custID)
        tax = subtotal * CDbl(wsC.Cells(custRow, "E").Value)
        total = subtotal + tax
        wsT.Range("Subtotal").Value = subtotal
        wsT.Range("Tax").Value = tax
        wsT.Range("Total").Value = total

        ' Export PDF
        pdfPath = outFolder & "Invoice_" & invNo & "_" & SafeFileName(CStr(wsC.Cells(custRow, "B").Value)) & ".pdf"
        wsT.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, _
            Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False

        ' Record in a tracker sheet
        RecordIssuedInvoice invNo, custID, pdfPath, total

        If emailThem Then
            EmailInvoice CStr(wsC.Cells(custRow, "C").Value), _
                         CStr(wsC.Cells(custRow, "B").Value), _
                         invNo, total, pdfPath, _
                         CStr(wsC.Cells(custRow, "F").Value)
        End If

        invNo = invNo + 1
        Application.StatusBar = "Invoice " & invNo & " done..."
    Next custRow

    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox (lastCust - 1) & " invoices generated.", vbInformation
End Sub

'-------------------------------------------------------------------------------
' WriteInvoiceFromTemplate - populates invoice template placeholders and lines.
' Returns the subtotal.
'-------------------------------------------------------------------------------
Private Function WriteInvoiceFromTemplate(wsT As Worksheet, wsC As Worksheet, _
                                          wsL As Worksheet, custRow As Long, _
                                          invNo As Long, custID As String) As Double
    Dim linesStart As Range, r As Long, lastLine As Long, destRow As Long
    Dim subtotal As Double
    subtotal = 0

    ' Named ranges in the template get filled in directly.
    On Error Resume Next
    wsT.Range("InvoiceNo").Value = invNo
    wsT.Range("CustomerName").Value = wsC.Cells(custRow, "B").Value
    wsT.Range("CustomerAddress").Value = wsC.Cells(custRow, "D").Value
    wsT.Range("InvoiceDate").Value = Date
    wsT.Range("DueDate").Value = Date + 30
    wsT.Range("AccountManager").Value = wsC.Cells(custRow, "F").Value
    On Error GoTo 0

    ' Clear any prior line items between markers ROW_LINES_START / ROW_LINES_END
    Set linesStart = wsT.Range("LinesStart")  ' named range at the first blank line row
    Dim endRow As Long
    endRow = wsT.Range("LinesEnd").Row - 1
    wsT.Range(linesStart.Offset(0, 0), wsT.Cells(endRow, 5)).ClearContents

    destRow = linesStart.Row
    lastLine = wsL.Cells(wsL.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastLine
        If CStr(wsL.Cells(r, "A").Value) = custID Then
            wsT.Cells(destRow, 1).Value = wsL.Cells(r, "B").Value
            wsT.Cells(destRow, 2).Value = wsL.Cells(r, "C").Value
            wsT.Cells(destRow, 3).Value = wsL.Cells(r, "D").Value
            wsT.Cells(destRow, 4).Value = wsL.Cells(r, "C").Value * wsL.Cells(r, "D").Value
            subtotal = subtotal + wsT.Cells(destRow, 4).Value
            destRow = destRow + 1
            If destRow > endRow Then Exit For
        End If
    Next r
    WriteInvoiceFromTemplate = subtotal
End Function

Private Sub EmailInvoice(toAddr As String, custName As String, invNo As Long, _
                         total As Double, pdfPath As String, amName As String)
    Dim ol As Object, mail As Object
    On Error Resume Next
    Set ol = GetObject(, "Outlook.Application")
    If ol Is Nothing Then Set ol = CreateObject("Outlook.Application")
    Set mail = ol.CreateItem(0)
    mail.To = toAddr
    mail.Subject = "Invoice " & invNo & " - " & custName
    mail.HTMLBody = "<div style='font-family:Arial;font-size:11pt;'>" & _
                    "Hi " & Split(custName, " ")(0) & ",<br><br>" & _
                    "Attached is invoice <b>#" & invNo & "</b> for $" & _
                    Format(total, "#,##0.00") & ", due in 30 days.<br><br>" & _
                    "Please let me know if you have any questions.<br><br>" & _
                    "Thanks,<br>" & amName & "</div>"
    mail.Attachments.Add pdfPath
    mail.Send
    On Error GoTo 0
End Sub

Private Sub EnsureFolder(ByVal path As String)
    If Right(path, 1) <> "\" Then path = path & "\"
    If Dir(path, vbDirectory) = "" Then MkDir path
End Sub

Private Function SafeFileName(ByVal s As String) As String
    Dim bad As Variant, b As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", ",")
    For Each b In bad
        s = Replace(s, CStr(b), "_")
    Next b
    SafeFileName = Left(s, 60)
End Function

Private Function GetMaxIssuedInvoiceNo() As Long
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("InvoiceRegister")
    If ws Is Nothing Then Exit Function
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Function
    GetMaxIssuedInvoiceNo = Application.WorksheetFunction.Max(ws.Range("A2:A" & lastRow)) - START_INVOICE_NO + 1
End Function

Private Sub RecordIssuedInvoice(invNo As Long, custID As String, pdf As String, total As Double)
    Dim ws As Worksheet, newRow As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("InvoiceRegister")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "InvoiceRegister"
        ws.Range("A1:E1").Value = Array("InvoiceNo", "CustomerID", "PDFPath", "Total", "IssuedAt")
        ws.Range("A1:E1").Font.Bold = True
    End If
    On Error GoTo 0
    newRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    ws.Cells(newRow, 1).Value = invNo
    ws.Cells(newRow, 2).Value = custID
    ws.Cells(newRow, 3).Value = pdf
    ws.Cells(newRow, 4).Value = total
    ws.Cells(newRow, 5).Value = Now
End Sub
