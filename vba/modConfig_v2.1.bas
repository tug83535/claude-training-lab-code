Attribute VB_Name = "modConfig"
Option Explicit

'===============================================================================
' modConfig - Workbook Constants & Configuration
' Keystone BenefitTech, Inc. - P&L Reporting & Allocation Model
'===============================================================================
' PURPOSE:  Single source of truth for all workbook-specific constants.
'           Change a sheet name, row offset, or threshold here and every
'           module picks it up automatically.
'
' VERSION:  2.1.0
' CHANGES:  v2.0 -> v2.1:
'           + Added SH_GL alias for SH_HIDDEN (used by modImport, modAdmin)
'           + Added SH_TECH_DOC, SH_CHANGE_LOG, SH_TEST_REPORT, SH_ALLOC_OUT
'             constants for v2.1 module generated sheets (modAdmin, modAllocation,
'             modIntegrationTest)
'           + Centralized SH_SENSITIVITY, SH_VARIANCE, SH_DQ_REPORT, SH_SEARCH,
'             SH_VAL_REPORT (were Private Const in individual modules)
'           + Added APP_BUILD_DATE constant
'           + Added SafeDeleteSheet helper (used by modAdmin, modAllocation,
'             modIntegrationTest)
'           + Added StyleHeader helper (used by modAdmin, modAllocation,
'             modIntegrationTest, modDataQuality)
'           + Updated APP_VERSION to 2.1.0
'===============================================================================

' === FISCAL YEAR ====================== CHANGE THIS each January ============
Public Const FISCAL_YEAR     As String = "25"          ' <-- CHANGE THIS: "25"=2025, "26"=2026
Public Const FISCAL_YEAR_4   As String = "2025"        ' <-- CHANGE THIS: full 4-digit year
' ==========================================================================

' --- Sheet Names (must match tab names exactly) --- CHANGE THIS if tabs renamed
Public Const SH_HIDDEN       As String = "CrossfireHiddenWorksheet"
Public Const SH_ASSUMPTIONS  As String = "Assumptions"
Public Const SH_DATADICT     As String = "Data Dictionary"
Public Const SH_AWS          As String = "AWS Allocation"
Public Const SH_REPORT       As String = "Report-->"
Public Const SH_PL_TREND     As String = "P&L - Monthly Trend"
Public Const SH_PROD_SUMMARY As String = "Product Line Summary"
Public Const SH_FUNC_TREND   As String = "Functional P&L - Monthly Trend"
Public Const SH_FUNC_JAN     As String = "Functional P&L Summary - Jan 25"
Public Const SH_FUNC_FEB     As String = "Functional P&L Summary - Feb 25"
Public Const SH_FUNC_MAR     As String = "Functional P&L Summary - Mar 25"
Public Const SH_NATURAL      As String = "US January 2025 Natural P&L"
Public Const SH_CHECKS       As String = "Checks"
Public Const SH_LOG          As String = "VBA_AuditLog"

'===============================================================================
' v2.1 SHEET NAME CONSTANTS — Generated/Utility Sheets
'===============================================================================
' Alias for GL detail sheet (used by modImport, modAdmin, modAllocation)
Public Const SH_GL           As String = "CrossfireHiddenWorksheet"

' Sheets created by v2.1 advanced modules
Public Const SH_TECH_DOC     As String = "Tech Documentation"
Public Const SH_CHANGE_LOG   As String = "Change Management Log"
Public Const SH_TEST_REPORT  As String = "Integration Test Report"
Public Const SH_ALLOC_OUT    As String = "Allocation Output"

' Centralized generated sheet names (were Private Const in individual modules)
Public Const SH_SENSITIVITY  As String = "Sensitivity Analysis"
Public Const SH_VARIANCE     As String = "Variance Analysis"
Public Const SH_DQ_REPORT    As String = "Data Quality Report"
Public Const SH_SEARCH       As String = "Search Results"
Public Const SH_VAL_REPORT   As String = "Validation Report"

' --- Products --- CHANGE THIS if product lines added/removed
Public Const PRODUCTS_CSV    As String = "iGO,Affirm,InsureSight,DocFast"

' --- Departments --- CHANGE THIS if org structure changes
Public Const DEPTS_CSV       As String = "NetOps,Security,Support,Partners,Content,R&D,Product Management"

' --- Months ---
Public Const MONTHS_CSV      As String = "Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec"

'===============================================================================
' ROW & COLUMN LAYOUT CONSTANTS
'===============================================================================
' These define where headers and data start on each sheet type.
' Every module should use these instead of hardcoding row/column numbers.
'
' STANDARD REPORT LAYOUT (P&L Trend, Functional P&L, Product Summary):
'   Row 1 = Company title ("Keystone BenefitTech, Inc.") - COLUMN A ONLY
'   Row 2 = Sheet subtitle
'   Row 3 = Blank
'   Row 4 = Column headers (spans all data columns)
'   Row 5+ = Data
'===============================================================================

' --- Report Sheets (P&L Trend, Functional P&L Summary, Product Line Summary) ---
Public Const HDR_ROW_REPORT  As Long = 4     ' Column headers on these sheets
Public Const DATA_ROW_REPORT As Long = 5     ' First data row

' --- Functional P&L Summary (Jan/Feb/Mar) ---
Public Const HDR_ROW_FUNC    As Long = 4     ' Line Item | iGO | Affirm | InsureSight | US
Public Const DATA_ROW_FUNC   As Long = 5     ' First data row
Public Const COL_US_TOTAL    As Long = 5     ' Column E = US consolidated total

' --- Assumptions ---
Public Const HDR_ROW_ASSUME  As Long = 5     ' Driver Name | Value | Type | Notes
Public Const DATA_ROW_ASSUME As Long = 6     ' First driver row

' --- Checks ---
Public Const HDR_ROW_CHECKS  As Long = 4     ' Check Name | Sheet A | Sheet B | Diff | Status
Public Const DATA_ROW_CHECKS As Long = 5     ' First check row
Public Const COL_CHECK_STATUS As Long = 5    ' Column E = PASS/FAIL

' --- AWS Allocation ---
Public Const HDR_ROW_AWS     As Long = 5     ' Product | Compute Share % | Monthly AWS Pool ($)
Public Const DATA_ROW_AWS    As Long = 6     ' First product row

' --- CrossfireHiddenWorksheet (GL Detail) ---
Public Const HDR_ROW_GL      As Long = 1     ' GL has headers in row 1 (no title row)
Public Const DATA_ROW_GL     As Long = 2     ' First transaction row
Public Const COL_GL_ID       As Long = 1     ' Column A
Public Const COL_GL_DATE     As Long = 2     ' Column B
Public Const COL_GL_DEPT     As Long = 3     ' Column C
Public Const COL_GL_PRODUCT  As Long = 4     ' Column D
Public Const COL_GL_CATEGORY As Long = 5     ' Column E
Public Const COL_GL_VENDOR   As Long = 6     ' Column F
Public Const COL_GL_AMOUNT   As Long = 7     ' Column G

' --- Formatting Constants ---
Public Const CLR_NAVY        As Long = 2050943   ' RGB(31,78,121) = #1F4E79
Public Const CLR_LIGHT_GRAY  As Long = 15921906  ' RGB(242,242,242) = #F2F2F2
Public Const CLR_ALT_ROW     As Long = 15651567  ' RGB(237,242,249) = #EDF2F9
Public Const CLR_GREEN_PASS  As Long = 5287936   ' RGB(0,176,80)
Public Const CLR_RED_FAIL    As Long = 255        ' RGB(255,0,0)
Public Const CLR_WHITE       As Long = 16777215

' --- PDF Export Path ---
Public Const PDF_SUBFOLDER   As String = "\PDF_Exports\"

' --- Variance Threshold --- CHANGE THIS to adjust sensitivity
Public Const VARIANCE_PCT    As Double = 0.15  ' 15% MoM threshold

' --- Reconciliation Tolerance --- CHANGE THIS to adjust pass/fail sensitivity
Public Const RECON_TOLERANCE As Double = 1     ' $1 tolerance for cross-sheet validation

' --- Application Info ---
Public Const APP_NAME        As String = "Keystone BenefitTech Automation Toolkit"
Public Const APP_VERSION     As String = "2.1.0"
Public Const APP_BUILD_DATE  As String = "2026-02-18"

'===============================================================================
' CSV SPLITTERS
'===============================================================================
Public Function GetProducts() As Variant
    GetProducts = Split(PRODUCTS_CSV, ",")
End Function

Public Function GetDepartments() As Variant
    GetDepartments = Split(DEPTS_CSV, ",")
End Function

Public Function GetMonths() As Variant
    GetMonths = Split(MONTHS_CSV, ",")
End Function

Public Function GetMonthSheetNames() As Variant
    ' Returns array of functional summary sheet names for each month
    ' Uses FISCAL_YEAR constant instead of hardcoded " 25"
    Dim arr(1 To 12) As String
    Dim mths As Variant: mths = GetMonths()
    Dim i As Long
    For i = 0 To 11
        arr(i + 1) = "Functional P&L Summary - " & mths(i) & " " & FISCAL_YEAR
    Next i
    GetMonthSheetNames = arr
End Function

'===============================================================================
' SAFE SHEET REFERENCE
'===============================================================================
Public Function GetSheet(ByVal shName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(shName)
    On Error GoTo 0
End Function

Public Function SheetExists(ByVal shName As String) As Boolean
    SheetExists = Not GetSheet(shName) Is Nothing
End Function

'===============================================================================
' SAFE SHEET DELETION (v2.1)
' Deletes a sheet by name if it exists. Suppresses alerts. No error if missing.
' Used by: modAdmin, modAllocation, modIntegrationTest
'===============================================================================
Public Sub SafeDeleteSheet(ByVal shName As String)
    If Not SheetExists(shName) Then Exit Sub
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(shName).Delete
    Application.DisplayAlerts = True
End Sub

'===============================================================================
' STYLE HEADER ROW (v2.1)
' Writes a headers array to a row with navy background and white bold text.
' Used by: modAdmin, modAllocation, modIntegrationTest, modDataQuality
'
' Parameters:
'   ws        - Target worksheet
'   headerRow - Row number to write headers into
'   headers   - Variant array of header strings (0-based)
'===============================================================================
Public Sub StyleHeader(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headers As Variant)
    Dim i As Long
    For i = 0 To UBound(headers)
        ws.Cells(headerRow, i + 1).Value = headers(i)
    Next i
    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, UBound(headers) + 1))
        .Font.Bold = True
        .Interior.Color = CLR_NAVY
        .Font.Color = CLR_WHITE
    End With
End Sub

'===============================================================================
' LAYOUT HELPER: LastRow
' Returns the last non-empty row in a given column.
' Use this instead of ws.Cells(ws.Rows.Count, col).End(xlUp).Row
'===============================================================================
Public Function LastRow(ByRef ws As Worksheet, Optional ByVal col As Long = 1) As Long
    LastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

'===============================================================================
' LAYOUT HELPER: LastCol
' Returns the last non-empty column in a given row.
'
' CRITICAL: Always pass the HEADER ROW, not row 1.
'           Row 1 on report sheets contains only the title in column A.
'           Row 4 (HDR_ROW_REPORT) contains headers spanning all data columns.
'
' Example:  lastC = LastCol(ws, HDR_ROW_REPORT)  ' Returns 18 on P&L Trend
'           lastC = LastCol(ws, 1)                ' Returns 1 (WRONG for reports!)
'===============================================================================
Public Function LastCol(ByRef ws As Worksheet, Optional ByVal headerRow As Long = 4) As Long
    LastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
End Function

'===============================================================================
' LAYOUT HELPER: FindColByHeader
' Scans a header row for a column containing the given keyword (case-insensitive).
' Returns the column number, or 0 if not found.
'
' Example:  usCol = FindColByHeader(ws, "US", HDR_ROW_FUNC)
'           totalCol = FindColByHeader(ws, "2025 Total", HDR_ROW_REPORT)
'===============================================================================
Public Function FindColByHeader(ByRef ws As Worksheet, _
                                ByVal keyword As String, _
                                Optional ByVal headerRow As Long = 4) As Long
    Dim lc As Long
    lc = LastCol(ws, headerRow)
    
    Dim c As Long
    Dim kw As String: kw = LCase(Trim(keyword))
    
    For c = 1 To lc
        If InStr(1, LCase(Trim(CStr(ws.Cells(headerRow, c).Value))), kw) > 0 Then
            FindColByHeader = c
            Exit Function
        End If
    Next c
    
    FindColByHeader = 0
End Function

'===============================================================================
' LAYOUT HELPER: FindRowByLabel
' Scans column A for a row containing the given keyword (case-insensitive).
' startRow defaults to 1; set it to DATA_ROW_REPORT (5) to skip title rows.
' Returns the row number, or 0 if not found.
'
' Example:  revRow = FindRowByLabel(ws, "Total Revenue", DATA_ROW_REPORT)
'===============================================================================
Public Function FindRowByLabel(ByRef ws As Worksheet, _
                               ByVal keyword As String, _
                               Optional ByVal startRow As Long = 1, _
                               Optional ByVal col As Long = 1) As Long
    Dim lr As Long
    lr = LastRow(ws, col)
    
    Dim r As Long
    Dim kw As String: kw = LCase(Trim(keyword))
    
    For r = startRow To lr
        If InStr(1, LCase(Trim(CStr(ws.Cells(r, col).Value))), kw) > 0 Then
            FindRowByLabel = r
            Exit Function
        End If
    Next r
    
    FindRowByLabel = 0
End Function

'===============================================================================
' SAFE CONVERSION: SafeNum
' Converts any cell value to Double safely. Returns 0 on error.
' Handles text-stored numbers, empty cells, and error values.
'===============================================================================
Public Function SafeNum(ByVal v As Variant) As Double
    On Error Resume Next
    If IsEmpty(v) Or v = "" Then
        SafeNum = 0
    ElseIf IsError(v) Then
        SafeNum = 0
    Else
        SafeNum = CDbl(v)
    End If
    If Err.Number <> 0 Then SafeNum = 0
    On Error GoTo 0
End Function

'===============================================================================
' SAFE CONVERSION: SafeStr
' Converts any cell value to trimmed String safely. Returns "" on error.
'===============================================================================
Public Function SafeStr(ByVal v As Variant) As String
    On Error Resume Next
    If IsEmpty(v) Or IsError(v) Then
        SafeStr = ""
    Else
        SafeStr = Trim(CStr(v))
    End If
    If Err.Number <> 0 Then SafeStr = ""
    On Error GoTo 0
End Function
