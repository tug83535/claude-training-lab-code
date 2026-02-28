#!/usr/bin/env python3
"""
Keystone BenefitTech P&L Model — Charts & Visuals Builder
Fortune 100 Executive Dashboard Layout

Architecture:
- Two-column grid layout with generous whitespace
- Row heights and column widths set explicitly for clean alignment
- Data Validation dropdown in B6 drives Section 1 charts via IF() lookup
- Data tables pushed to row 200+ (off-screen, not visible to user)
- 8 charts across 3 clearly separated sections

Run: python3 python/build_charts.py
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import (
    Paragraph, ParagraphProperties, CharacterProperties,
    Font as DrawingFont, RichTextProperties,
)
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

# ============================================================
# 1. LOAD WORKBOOK
# ============================================================
INPUT_FILE = 'excel/KeystoneBenefitTech_PL_Model.xlsx'
OUTPUT_FILE = 'excel/KeystoneBenefitTech_PL_Model.xlsx'

wb = openpyxl.load_workbook(INPUT_FILE)
ws = wb['Charts & Visuals']

# Unmerge first, then clear
for merge in list(ws.merged_cells.ranges):
    ws.unmerge_cells(str(merge))

for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=20):
    for cell in row:
        cell.value = None
        cell.font = Font()
        cell.fill = PatternFill()
        cell.alignment = Alignment()
        cell.border = Border()

ws._charts = []

# Remove any existing data validations (from previous runs)
ws.data_validations.dataValidation = []

# ============================================================
# 2. BRAND CONSTANTS
# ============================================================
IPIPELINE_BLUE = '0B4779'
ARCTIC_WHITE = 'F9F9F9'
SOFT_NEUTRAL = 'F0F0EE'
NAVY_BLUE = '112E51'
INNOVATION_BLUE = '4B9BCB'
LIME_GREEN = 'BFF18C'
AQUA = '2BCCD3'
CHARCOAL = '161616'

CLR_BLUE = '0B4779'
CLR_NAVY = '112E51'
CLR_AQUA = '2BCCD3'
CLR_LIME = 'BFF18C'
CLR_INNOVATION = '4B9BCB'
CLR_ORANGE = 'E8833A'
CLR_TEAL = '1A8A8A'

PRODUCT_COLORS = [CLR_BLUE, CLR_AQUA, CLR_LIME, CLR_INNOVATION]
DEPT_COLORS = [CLR_BLUE, CLR_AQUA, CLR_LIME, CLR_INNOVATION, CLR_ORANGE, CLR_NAVY, CLR_TEAL]

# Fonts
font_title = Font(name='Arial', bold=True, color=ARCTIC_WHITE, size=18)
font_subtitle = Font(name='Arial', bold=True, color=ARCTIC_WHITE, size=11)
font_date = Font(name='Arial', italic=True, color=INNOVATION_BLUE, size=9)
font_section_bar = Font(name='Arial', bold=True, color=ARCTIC_WHITE, size=12)
font_dropdown_label = Font(name='Arial', bold=True, color=IPIPELINE_BLUE, size=11)
font_dropdown_value = Font(name='Arial', bold=True, color=NAVY_BLUE, size=13)
font_instruction = Font(name='Arial', italic=True, color=INNOVATION_BLUE, size=9)
font_data_label = Font(name='Arial', color=CHARCOAL, size=9)
font_chart_desc = Font(name='Arial', italic=True, color='666666', size=9)

# Fills
fill_header = PatternFill(start_color=IPIPELINE_BLUE, end_color=IPIPELINE_BLUE, fill_type='solid')
fill_navy = PatternFill(start_color=NAVY_BLUE, end_color=NAVY_BLUE, fill_type='solid')
fill_white = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
fill_dropdown = PatternFill(start_color='E8F4FD', end_color='E8F4FD', fill_type='solid')
fill_spacer = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
fill_light_gray = PatternFill(start_color='F5F5F5', end_color='F5F5F5', fill_type='solid')

# Borders
medium_blue = Side(style='medium', color=IPIPELINE_BLUE)
thick_blue = Side(style='thick', color=IPIPELINE_BLUE)
border_dropdown = Border(
    left=medium_blue, right=medium_blue,
    top=medium_blue, bottom=medium_blue,
)

FMT_CURRENCY = '$#,##0.0"K"'
FMT_PERCENT = '0.0%'

# ============================================================
# 3. COLUMN WIDTHS — Dashboard Grid
# ============================================================
# A = narrow left gutter
# B-H = left chart zone (7 cols)
# I = center gutter
# J-P = right chart zone (7 cols)
# Q = narrow right gutter
col_widths = {
    'A': 3, 'B': 12, 'C': 12, 'D': 12, 'E': 12, 'F': 12, 'G': 12, 'H': 12,
    'I': 3,
    'J': 12, 'K': 12, 'L': 12, 'M': 12, 'N': 12, 'O': 12, 'P': 12,
    'Q': 3,
}
for col_letter, width in col_widths.items():
    ws.column_dimensions[col_letter].width = width


# ============================================================
# 4. HELPER: Fill a row range white
# ============================================================
MAX_COL = 17  # A through Q

def fill_row_white(row_num, height=None):
    """Fill an entire row with white background."""
    for c in range(1, MAX_COL + 1):
        ws.cell(row=row_num, column=c).fill = fill_white
    if height:
        ws.row_dimensions[row_num].height = height


def fill_rows_white(start, end, height=None):
    """Fill a range of rows white."""
    for r in range(start, end + 1):
        fill_row_white(r, height)


def make_section_bar(row_num, text, fill=None):
    """Create a full-width colored section header bar."""
    if fill is None:
        fill = fill_header
    ws.merge_cells(start_row=row_num, start_column=2, end_row=row_num, end_column=16)
    ws.cell(row=row_num, column=2).value = text
    ws.cell(row=row_num, column=2).font = font_section_bar
    ws.cell(row=row_num, column=2).fill = fill
    ws.cell(row=row_num, column=2).alignment = Alignment(horizontal='left', vertical='center')
    ws.cell(row=row_num, column=2).border = Border(
        left=Side(style='medium', color=fill.start_color.rgb[2:] if len(fill.start_color.rgb) > 6 else fill.start_color.rgb),
    )
    # Fill all cells in the merged range
    for c in range(2, 17):
        ws.cell(row=row_num, column=c).fill = fill
    # Gutters
    ws.cell(row=row_num, column=1).fill = fill_white
    ws.cell(row=row_num, column=17).fill = fill_white
    ws.row_dimensions[row_num].height = 32


# ============================================================
# 5. CHART STYLING HELPERS
# ============================================================
def make_chart_font(size=10, bold=False, color='161616'):
    return CharacterProperties(
        latin=DrawingFont(typeface='Arial'),
        sz=size * 100,
        b=bold,
        solidFill=color,
    )


def style_chart(chart, title_text, width=13.5, height=9.5):
    """Apply Fortune 100 styling to a chart."""
    chart.width = width
    chart.height = height
    chart.style = 2

    # Title
    chart.title = title_text
    chart.title.txPr = RichText(
        p=[Paragraph(
            pPr=ParagraphProperties(
                defRPr=make_chart_font(size=11, bold=True, color=IPIPELINE_BLUE)
            ),
            endParaRPr=make_chart_font(size=11, bold=True, color=IPIPELINE_BLUE),
        )]
    )

    # Legend
    if chart.legend:
        chart.legend.txPr = RichText(
            p=[Paragraph(
                pPr=ParagraphProperties(
                    defRPr=make_chart_font(size=8, color=CHARCOAL)
                ),
                endParaRPr=make_chart_font(size=8, color=CHARCOAL),
            )]
        )

    # No chart border
    chart.graphical_properties = GraphicalProperties()
    chart.graphical_properties.line = LineProperties(noFill=True)

    return chart


def style_axis(axis, fmt=None, title_text=None):
    if axis is None:
        return
    axis.txPr = RichText(
        p=[Paragraph(
            pPr=ParagraphProperties(
                defRPr=make_chart_font(size=8, color=CHARCOAL)
            ),
            endParaRPr=make_chart_font(size=8, color=CHARCOAL),
        )]
    )
    if fmt:
        axis.numFmt = fmt
    if title_text:
        axis.title = title_text


def color_series(series, hex_color):
    series.graphicalProperties.solidFill = hex_color
    series.graphicalProperties.line.solidFill = hex_color


# ============================================================
# 6. DASHBOARD HEADER (Rows 1-4)
# ============================================================
# Row 1: Company title bar
ws.merge_cells('B1:P1')
ws['B1'] = 'Keystone BenefitTech, Inc.'
ws['B1'].font = font_title
ws['B1'].fill = fill_header
ws['B1'].alignment = Alignment(horizontal='left', vertical='center')
for c in range(2, 17):
    ws.cell(row=1, column=c).fill = fill_header
ws.cell(row=1, column=1).fill = fill_white
ws.cell(row=1, column=17).fill = fill_white
ws.row_dimensions[1].height = 42

# Row 2: Subtitle bar
ws.merge_cells('B2:N2')
ws['B2'] = 'Charts & Visuals Dashboard — FY2025'
ws['B2'].font = font_subtitle
ws['B2'].fill = fill_header
ws['B2'].alignment = Alignment(horizontal='left', vertical='center')
ws['O2'] = 'Last Updated: 2026-02-28'
ws['O2'].font = Font(name='Arial', italic=True, color=INNOVATION_BLUE, size=8)
ws['O2'].fill = fill_header
ws['O2'].alignment = Alignment(horizontal='right', vertical='center')
for c in range(2, 17):
    ws.cell(row=2, column=c).fill = fill_header
ws.cell(row=2, column=1).fill = fill_white
ws.cell(row=2, column=17).fill = fill_white
ws.row_dimensions[2].height = 26

# Row 3: Thin accent line
for c in range(2, 17):
    ws.cell(row=3, column=c).fill = PatternFill(start_color=AQUA, end_color=AQUA, fill_type='solid')
ws.cell(row=3, column=1).fill = fill_white
ws.cell(row=3, column=17).fill = fill_white
ws.row_dimensions[3].height = 4

# Row 4: Spacer
fill_row_white(4, height=12)

# ============================================================
# 7. PRODUCT SELECTOR (Rows 5-7)
# ============================================================
make_section_bar(5, 'PRODUCT SELECTOR — Choose a product to filter Section 1 charts')

# Row 6: Dropdown row
fill_row_white(6, height=36)
ws['B6'] = 'Product:'
ws['B6'].font = font_dropdown_label
ws['B6'].alignment = Alignment(horizontal='right', vertical='center')
ws['B6'].fill = fill_white

ws['C6'] = 'iGO'
ws['C6'].font = font_dropdown_value
ws['C6'].fill = fill_dropdown
ws['C6'].border = border_dropdown
ws['C6'].alignment = Alignment(horizontal='center', vertical='center')

# Data Validation
dv = DataValidation(
    type='list',
    formula1='"iGO,Affirm,InsureSight,DocFast"',
    allow_blank=False,
    showDropDown=False,
)
dv.error = 'Please select a valid product'
dv.errorTitle = 'Invalid Product'
dv.prompt = 'Select a product to view its dashboard'
dv.promptTitle = 'Product Selector'
ws.add_data_validation(dv)
dv.add(ws['C6'])

ws['D6'] = 'Click the dropdown arrow to select iGO, Affirm, InsureSight, or DocFast'
ws['D6'].font = font_instruction
ws['D6'].alignment = Alignment(vertical='center')
ws['D6'].fill = fill_white

# Row 7: Spacer
fill_row_white(7, height=14)

# ============================================================
# 8. SECTION 1: SELECTED PRODUCT DASHBOARD (Rows 8-46)
# ============================================================
make_section_bar(8, 'SECTION 1  |  SELECTED PRODUCT DASHBOARD', fill_navy)

# Row 9: Spacer
fill_row_white(9, height=8)

# Chart descriptions
fill_row_white(10, height=16)
ws['B10'] = 'Revenue vs Cost of Revenue — Monthly ($K)'
ws['B10'].font = font_chart_desc
ws['B10'].fill = fill_white
ws['J10'] = 'Contribution Margin % — Monthly Trend'
ws['J10'].font = font_chart_desc
ws['J10'].fill = fill_white

# Rows 11-27: Chart 1 (left) + Chart 2 (right) — side by side
fill_rows_white(11, 28)

# Row 29: Spacer
fill_row_white(29, height=10)

# Chart description
fill_row_white(30, height=16)
ws['B30'] = 'Expense Breakdown by Department — Full Year'
ws['B30'].font = font_chart_desc
ws['B30'].fill = fill_white

# Rows 31-47: Chart 3 (centered left half)
fill_rows_white(31, 48)

# ============================================================
# 9. SECTION 2: ALL-PRODUCTS COMPARISON (Rows 49-87)
# ============================================================
fill_row_white(49, height=14)
make_section_bar(50, 'SECTION 2  |  ALL-PRODUCTS COMPARISON', fill_navy)
fill_row_white(51, height=8)

# Descriptions
fill_row_white(52, height=16)
ws['B52'] = 'Monthly Revenue by Product ($K)'
ws['B52'].font = font_chart_desc
ws['B52'].fill = fill_white
ws['J52'] = 'Contribution Margin % by Product'
ws['J52'].font = font_chart_desc
ws['J52'].fill = fill_white

# Rows 53-69: Chart 4 (left) + Chart 5 (right)
fill_rows_white(53, 70)

fill_row_white(71, height=10)

# Description
fill_row_white(72, height=16)
ws['B72'] = 'Full Year Revenue Mix — FY2025'
ws['B72'].font = font_chart_desc
ws['B72'].fill = fill_white

# Rows 73-89: Chart 6 (centered left half)
fill_rows_white(73, 90)

# ============================================================
# 10. SECTION 3: ADVANCED ANALYTICS (Rows 91-130)
# ============================================================
fill_row_white(91, height=14)
make_section_bar(92, 'SECTION 3  |  ADVANCED ANALYTICS', fill_navy)
fill_row_white(93, height=8)

# Descriptions
fill_row_white(94, height=16)
ws['B94'] = 'Revenue vs COGS vs Contribution Margin — By Product'
ws['B94'].font = font_chart_desc
ws['B94'].fill = fill_white
ws['J94'] = 'Department Cost Distribution Across Products'
ws['J94'].font = font_chart_desc
ws['J94'].fill = fill_white

# Rows 95-111: Chart 7 (left) + Chart 8 (right)
fill_rows_white(95, 112)

# Footer spacer
fill_rows_white(113, 118)

# ============================================================
# 11. DATA TABLES (Row 200+, far off-screen)
# ============================================================
PLS = "'Product Line Summary'"
LT_START = 200

months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

ws.cell(row=LT_START - 1, column=1).value = 'LOOKUP TABLES — DO NOT EDIT (used by charts above)'
ws.cell(row=LT_START - 1, column=1).font = Font(name='Arial', bold=True, color='999999', size=8)

# Month headers
ws.cell(row=LT_START, column=1).value = 'Metric'
ws.cell(row=LT_START, column=1).font = font_data_label
for i, m in enumerate(months):
    ws.cell(row=LT_START, column=i + 2).value = m
    ws.cell(row=LT_START, column=i + 2).font = font_data_label


def build_if_formula(col_letter, igo_row, affirm_row, insure_row, docfast_row, sheet=PLS):
    return (
        f'=IF($C$6="iGO",{sheet}!{col_letter}{igo_row},'
        f'IF($C$6="Affirm",{sheet}!{col_letter}{affirm_row},'
        f'IF($C$6="InsureSight",{sheet}!{col_letter}{insure_row},'
        f'{sheet}!{col_letter}{docfast_row})))'
    )


# Metric lookup rows
metric_rows = {
    'Revenue':               (8, 16, 24, 32),
    'Cost of Revenue':       (9, 17, 25, 33),
    'Contribution Margin $': (10, 18, 26, 34),
    'Contribution Margin %': (11, 19, 27, 35),
    'CM + R&D $':            (12, 20, 28, 36),
    'CM + R&D %':            (13, 21, 29, 37),
}

row_offset = 1
for metric_name, (igo_r, aff_r, ins_r, doc_r) in metric_rows.items():
    r = LT_START + row_offset
    ws.cell(row=r, column=1).value = metric_name
    ws.cell(row=r, column=1).font = font_data_label
    for col_idx in range(2, 14):
        cl = get_column_letter(col_idx)
        ws.cell(row=r, column=col_idx).value = build_if_formula(cl, igo_r, aff_r, ins_r, doc_r)
        ws.cell(row=r, column=col_idx).number_format = FMT_PERCENT if '%' in metric_name else FMT_CURRENCY
    row_offset += 1

# Department expense lookup
row_offset += 1
dept_label_row = LT_START + row_offset
ws.cell(row=dept_label_row, column=1).value = 'Department'
ws.cell(row=dept_label_row, column=1).font = font_data_label
for i, m in enumerate(months):
    ws.cell(row=dept_label_row, column=i + 2).value = m
    ws.cell(row=dept_label_row, column=i + 2).font = font_data_label

dept_rows = {
    'NetOps':               (43, 53, 63, 73),
    'Security':             (44, 54, 64, 74),
    'Support':              (45, 55, 65, 75),
    'Partners':             (46, 56, 66, 76),
    'Content':              (47, 57, 67, 77),
    'Research & Development': (48, 58, 68, 78),
    'Product Management':   (49, 59, 69, 79),
}

row_offset += 1
DEPT_START_ROW = LT_START + row_offset
for dept_name, (igo_r, aff_r, ins_r, doc_r) in dept_rows.items():
    r = LT_START + row_offset
    ws.cell(row=r, column=1).value = dept_name
    ws.cell(row=r, column=1).font = font_data_label
    for col_idx in range(2, 14):
        cl = get_column_letter(col_idx)
        ws.cell(row=r, column=col_idx).value = build_if_formula(cl, igo_r, aff_r, ins_r, doc_r)
        ws.cell(row=r, column=col_idx).number_format = FMT_CURRENCY
    row_offset += 1

DEPT_END_ROW = LT_START + row_offset - 1

# FY totals for dept pie (column N)
for r in range(DEPT_START_ROW, DEPT_END_ROW + 1):
    ws.cell(row=r, column=14).value = f'=SUM(B{r}:M{r})'
    ws.cell(row=r, column=14).number_format = FMT_CURRENCY

# Full Year Revenue per product for pie chart
row_offset += 2
PIE_DATA_ROW = LT_START + row_offset
for label, col_idx in [('Product', 1), ('FY Revenue', 2), ('FY COGS', 3), ('FY CM$', 4)]:
    ws.cell(row=PIE_DATA_ROW, column=col_idx).value = label
    ws.cell(row=PIE_DATA_ROW, column=col_idx).font = font_data_label

products_fy = [('iGO', 8, 9, 10), ('Affirm', 16, 17, 18), ('InsureSight', 24, 25, 26), ('DocFast', 32, 33, 34)]
for i, (prod, rev_r, cogs_r, cm_r) in enumerate(products_fy):
    r = PIE_DATA_ROW + 1 + i
    ws.cell(row=r, column=1).value = prod
    ws.cell(row=r, column=1).font = font_data_label
    ws.cell(row=r, column=2).value = f'={PLS}!R{rev_r}'
    ws.cell(row=r, column=2).number_format = FMT_CURRENCY
    ws.cell(row=r, column=3).value = f'={PLS}!R{cogs_r}'
    ws.cell(row=r, column=3).number_format = FMT_CURRENCY
    ws.cell(row=r, column=4).value = f'={PLS}!R{cm_r}'
    ws.cell(row=r, column=4).number_format = FMT_CURRENCY

PIE_PROD_START = PIE_DATA_ROW + 1
PIE_PROD_END = PIE_DATA_ROW + 4

# Department totals across ALL products
DEPT_ALL_ROW = PIE_PROD_END + 2
for label, col_idx in [('Department', 1), ('iGO', 2), ('Affirm', 3), ('InsureSight', 4), ('DocFast', 5)]:
    ws.cell(row=DEPT_ALL_ROW, column=col_idx).value = label
    ws.cell(row=DEPT_ALL_ROW, column=col_idx).font = font_data_label

dept_names_list = ['NetOps', 'Security', 'Support', 'Partners', 'Content', 'R&D', 'Product Mgmt']
igo_dept = [43, 44, 45, 46, 47, 48, 49]
aff_dept = [53, 54, 55, 56, 57, 58, 59]
ins_dept = [63, 64, 65, 66, 67, 68, 69]
doc_dept = [73, 74, 75, 76, 77, 78, 79]

DEPT_ALL_START = DEPT_ALL_ROW + 1
for i, dept_name in enumerate(dept_names_list):
    r = DEPT_ALL_START + i
    ws.cell(row=r, column=1).value = dept_name
    ws.cell(row=r, column=1).font = font_data_label
    ws.cell(row=r, column=2).value = f'={PLS}!R{igo_dept[i]}'
    ws.cell(row=r, column=2).number_format = FMT_CURRENCY
    ws.cell(row=r, column=3).value = f'={PLS}!R{aff_dept[i]}'
    ws.cell(row=r, column=3).number_format = FMT_CURRENCY
    ws.cell(row=r, column=4).value = f'={PLS}!R{ins_dept[i]}'
    ws.cell(row=r, column=4).number_format = FMT_CURRENCY
    ws.cell(row=r, column=5).value = f'={PLS}!R{doc_dept[i]}'
    ws.cell(row=r, column=5).number_format = FMT_CURRENCY
DEPT_ALL_END = DEPT_ALL_START + 6

# Revenue by product by month
MULTI_REV_ROW = DEPT_ALL_END + 2
ws.cell(row=MULTI_REV_ROW, column=1).value = 'Month'
ws.cell(row=MULTI_REV_ROW, column=1).font = font_data_label
for i, m in enumerate(months):
    ws.cell(row=MULTI_REV_ROW + 1 + i, column=1).value = m
    ws.cell(row=MULTI_REV_ROW + 1 + i, column=1).font = font_data_label

rev_rows_by_product = {'iGO': 8, 'Affirm': 16, 'InsureSight': 24, 'DocFast': 32}
for p_idx, (prod_name, pls_row) in enumerate(rev_rows_by_product.items()):
    col = p_idx + 2
    ws.cell(row=MULTI_REV_ROW, column=col).value = prod_name
    ws.cell(row=MULTI_REV_ROW, column=col).font = font_data_label
    for m_idx in range(12):
        pls_col = get_column_letter(m_idx + 2)
        r = MULTI_REV_ROW + 1 + m_idx
        ws.cell(row=r, column=col).value = f'={PLS}!{pls_col}{pls_row}'
        ws.cell(row=r, column=col).number_format = FMT_CURRENCY

MULTI_REV_START = MULTI_REV_ROW + 1
MULTI_REV_END = MULTI_REV_ROW + 12

# CM% by product by month
MULTI_CM_ROW = MULTI_REV_END + 2
ws.cell(row=MULTI_CM_ROW, column=1).value = 'Month'
ws.cell(row=MULTI_CM_ROW, column=1).font = font_data_label
for i, m in enumerate(months):
    ws.cell(row=MULTI_CM_ROW + 1 + i, column=1).value = m
    ws.cell(row=MULTI_CM_ROW + 1 + i, column=1).font = font_data_label

cm_pct_rows = {'iGO': 11, 'Affirm': 19, 'InsureSight': 27, 'DocFast': 35}
for p_idx, (prod_name, pls_row) in enumerate(cm_pct_rows.items()):
    col = p_idx + 2
    ws.cell(row=MULTI_CM_ROW, column=col).value = prod_name
    ws.cell(row=MULTI_CM_ROW, column=col).font = font_data_label
    for m_idx in range(12):
        pls_col = get_column_letter(m_idx + 2)
        r = MULTI_CM_ROW + 1 + m_idx
        ws.cell(row=r, column=col).value = f'={PLS}!{pls_col}{pls_row}'
        ws.cell(row=r, column=col).number_format = FMT_PERCENT

MULTI_CM_START = MULTI_CM_ROW + 1
MULTI_CM_END = MULTI_CM_ROW + 12

# Revenue vs COGS vs CM per product
WATERFALL_ROW = MULTI_CM_END + 2
for label, col_idx in [('Product', 1), ('Revenue', 2), ('Cost of Revenue', 3), ('Contribution Margin', 4)]:
    ws.cell(row=WATERFALL_ROW, column=col_idx).value = label
    ws.cell(row=WATERFALL_ROW, column=col_idx).font = font_data_label

waterfall_data = [('iGO', 8, 9, 10), ('Affirm', 16, 17, 18), ('InsureSight', 24, 25, 26), ('DocFast', 32, 33, 34)]
WATERFALL_START = WATERFALL_ROW + 1
for i, (prod, rev_r, cogs_r, cm_r) in enumerate(waterfall_data):
    r = WATERFALL_START + i
    ws.cell(row=r, column=1).value = prod
    ws.cell(row=r, column=1).font = font_data_label
    ws.cell(row=r, column=2).value = f'={PLS}!R{rev_r}'
    ws.cell(row=r, column=2).number_format = FMT_CURRENCY
    ws.cell(row=r, column=3).value = f'={PLS}!R{cogs_r}'
    ws.cell(row=r, column=3).number_format = FMT_CURRENCY
    ws.cell(row=r, column=4).value = f'={PLS}!R{cm_r}'
    ws.cell(row=r, column=4).number_format = FMT_CURRENCY
WATERFALL_END = WATERFALL_START + 3


# ============================================================
# 12. BUILD CHARTS
# ============================================================
# Chart sizes for the two-column grid
HALF_W = 13.0   # width for side-by-side charts (fits B-H or J-P)
HALF_H = 10.0   # height for all charts
FULL_W = 13.0   # width for single-column charts

# --- SECTION 1 ---

# Chart 1: Revenue vs Cost of Revenue (left)
chart1 = BarChart()
chart1.type = 'col'
chart1.grouping = 'clustered'

rev_ref = Reference(ws, min_col=2, max_col=13, min_row=LT_START + 1)
rev_cats = Reference(ws, min_col=2, max_col=13, min_row=LT_START)
chart1.add_data(rev_ref, from_rows=True, titles_from_data=False)
chart1.set_categories(rev_cats)
chart1.series[0].title = openpyxl.chart.series.SeriesLabel(v='Revenue')
color_series(chart1.series[0], CLR_BLUE)

cogs_ref = Reference(ws, min_col=2, max_col=13, min_row=LT_START + 2)
chart1.add_data(cogs_ref, from_rows=True, titles_from_data=False)
chart1.series[1].title = openpyxl.chart.series.SeriesLabel(v='Cost of Revenue')
color_series(chart1.series[1], CLR_AQUA)

style_chart(chart1, 'Revenue vs Cost of Revenue', width=HALF_W, height=HALF_H)
style_axis(chart1.x_axis)
style_axis(chart1.y_axis, fmt=FMT_CURRENCY)
chart1.y_axis.delete = False
chart1.legend.position = 'b'
ws.add_chart(chart1, 'B11')

# Chart 2: CM% Trend (right)
chart2 = LineChart()

cm_pct_ref = Reference(ws, min_col=2, max_col=13, min_row=LT_START + 4)
cm_cats = Reference(ws, min_col=2, max_col=13, min_row=LT_START)
chart2.add_data(cm_pct_ref, from_rows=True, titles_from_data=False)
chart2.set_categories(cm_cats)
chart2.series[0].title = openpyxl.chart.series.SeriesLabel(v='Contribution Margin %')
color_series(chart2.series[0], CLR_BLUE)
chart2.series[0].graphicalProperties.line.width = 28000

style_chart(chart2, 'Contribution Margin %', width=HALF_W, height=HALF_H)
style_axis(chart2.x_axis)
style_axis(chart2.y_axis, fmt=FMT_PERCENT)
chart2.y_axis.delete = False
chart2.legend.position = 'b'
ws.add_chart(chart2, 'J11')

# Chart 3: Expense Breakdown Pie (left)
chart3 = PieChart()

dept_labels = Reference(ws, min_col=1, min_row=DEPT_START_ROW, max_row=DEPT_END_ROW)
dept_data = Reference(ws, min_col=14, min_row=DEPT_START_ROW, max_row=DEPT_END_ROW)
chart3.add_data(dept_data, titles_from_data=False)
chart3.set_categories(dept_labels)
chart3.series[0].title = openpyxl.chart.series.SeriesLabel(v='Expense by Department')

for i, clr in enumerate(DEPT_COLORS):
    pt = DataPoint(idx=i)
    pt.graphicalProperties.solidFill = clr
    chart3.series[0].data_points.append(pt)

chart3.series[0].dLbls = DataLabelList()
chart3.series[0].dLbls.showPercent = True
chart3.series[0].dLbls.showCatName = True
chart3.series[0].dLbls.showVal = False
chart3.series[0].dLbls.numFmt = '0.0%'
chart3.series[0].dLbls.txPr = RichText(
    p=[Paragraph(
        pPr=ParagraphProperties(defRPr=make_chart_font(size=8, color=CHARCOAL)),
        endParaRPr=make_chart_font(size=8, color=CHARCOAL),
    )]
)

style_chart(chart3, 'Expense Breakdown by Department', width=FULL_W, height=HALF_H)
ws.add_chart(chart3, 'B31')


# --- SECTION 2 ---

# Chart 4: Monthly Revenue by Product (left)
chart4 = BarChart()
chart4.type = 'col'
chart4.grouping = 'clustered'

multi_cats = Reference(ws, min_col=1, min_row=MULTI_REV_START, max_row=MULTI_REV_END)
for p_idx in range(4):
    col = p_idx + 2
    data_ref = Reference(ws, min_col=col, min_row=MULTI_REV_ROW, max_row=MULTI_REV_END)
    chart4.add_data(data_ref, titles_from_data=True)
chart4.set_categories(multi_cats)

for i, clr in enumerate(PRODUCT_COLORS):
    color_series(chart4.series[i], clr)

style_chart(chart4, 'Monthly Revenue by Product', width=HALF_W, height=HALF_H)
style_axis(chart4.x_axis)
style_axis(chart4.y_axis, fmt=FMT_CURRENCY)
chart4.y_axis.delete = False
chart4.legend.position = 'b'
ws.add_chart(chart4, 'B53')

# Chart 5: CM% by Product (right)
chart5 = LineChart()

multi_cm_cats = Reference(ws, min_col=1, min_row=MULTI_CM_START, max_row=MULTI_CM_END)
for p_idx in range(4):
    col = p_idx + 2
    data_ref = Reference(ws, min_col=col, min_row=MULTI_CM_ROW, max_row=MULTI_CM_END)
    chart5.add_data(data_ref, titles_from_data=True)
chart5.set_categories(multi_cm_cats)

for i, clr in enumerate(PRODUCT_COLORS):
    color_series(chart5.series[i], clr)
    chart5.series[i].graphicalProperties.line.width = 28000

style_chart(chart5, 'Contribution Margin % by Product', width=HALF_W, height=HALF_H)
style_axis(chart5.x_axis)
style_axis(chart5.y_axis, fmt=FMT_PERCENT)
chart5.y_axis.delete = False
chart5.legend.position = 'b'
ws.add_chart(chart5, 'J53')

# Chart 6: Revenue Mix Pie (left)
chart6 = PieChart()

pie_labels = Reference(ws, min_col=1, min_row=PIE_PROD_START, max_row=PIE_PROD_END)
pie_data = Reference(ws, min_col=2, min_row=PIE_PROD_START, max_row=PIE_PROD_END)
chart6.add_data(pie_data, titles_from_data=False)
chart6.set_categories(pie_labels)
chart6.series[0].title = openpyxl.chart.series.SeriesLabel(v='Revenue Share')

for i, clr in enumerate(PRODUCT_COLORS):
    pt = DataPoint(idx=i)
    pt.graphicalProperties.solidFill = clr
    chart6.series[0].data_points.append(pt)

chart6.series[0].dLbls = DataLabelList()
chart6.series[0].dLbls.showPercent = True
chart6.series[0].dLbls.showCatName = True
chart6.series[0].dLbls.showVal = False
chart6.series[0].dLbls.numFmt = '0.0%'
chart6.series[0].dLbls.txPr = RichText(
    p=[Paragraph(
        pPr=ParagraphProperties(defRPr=make_chart_font(size=9, color=CHARCOAL)),
        endParaRPr=make_chart_font(size=9, color=CHARCOAL),
    )]
)

style_chart(chart6, 'Full Year Revenue Mix', width=FULL_W, height=HALF_H)
ws.add_chart(chart6, 'B73')


# --- SECTION 3 ---

# Chart 7: Revenue vs COGS vs CM (left)
chart7 = BarChart()
chart7.type = 'col'
chart7.grouping = 'clustered'

wf_cats = Reference(ws, min_col=1, min_row=WATERFALL_START, max_row=WATERFALL_END)
wf_rev = Reference(ws, min_col=2, min_row=WATERFALL_ROW, max_row=WATERFALL_END)
chart7.add_data(wf_rev, titles_from_data=True)
color_series(chart7.series[0], CLR_BLUE)

wf_cogs = Reference(ws, min_col=3, min_row=WATERFALL_ROW, max_row=WATERFALL_END)
chart7.add_data(wf_cogs, titles_from_data=True)
color_series(chart7.series[1], CLR_AQUA)

wf_cm = Reference(ws, min_col=4, min_row=WATERFALL_ROW, max_row=WATERFALL_END)
chart7.add_data(wf_cm, titles_from_data=True)
color_series(chart7.series[2], CLR_LIME)

chart7.set_categories(wf_cats)

style_chart(chart7, 'Revenue / COGS / Margin by Product', width=HALF_W, height=HALF_H)
style_axis(chart7.x_axis)
style_axis(chart7.y_axis, fmt=FMT_CURRENCY)
chart7.y_axis.delete = False
chart7.legend.position = 'b'
ws.add_chart(chart7, 'B95')

# Chart 8: Dept Cost Stacked Bar (right)
chart8 = BarChart()
chart8.type = 'col'
chart8.grouping = 'stacked'

dept_all_cats = Reference(ws, min_col=1, min_row=DEPT_ALL_START, max_row=DEPT_ALL_END)
for p_idx in range(4):
    col = p_idx + 2
    data_ref = Reference(ws, min_col=col, min_row=DEPT_ALL_ROW, max_row=DEPT_ALL_END)
    chart8.add_data(data_ref, titles_from_data=True)
chart8.set_categories(dept_all_cats)

for i, clr in enumerate(PRODUCT_COLORS):
    color_series(chart8.series[i], clr)

style_chart(chart8, 'Dept Cost by Product', width=HALF_W, height=HALF_H)
style_axis(chart8.x_axis)
style_axis(chart8.y_axis, fmt=FMT_CURRENCY)
chart8.y_axis.delete = False
chart8.legend.position = 'b'
ws.add_chart(chart8, 'J95')


# ============================================================
# 13. FINAL CLEANUP
# ============================================================
ws.sheet_view.showGridLines = False
ws.freeze_panes = 'A4'
ws.sheet_properties.tabColor = IPIPELINE_BLUE

# Fill all visible rows white where not already set
for row_num in range(1, 120):
    for col_idx in range(1, MAX_COL + 1):
        c = ws.cell(row=row_num, column=col_idx)
        if c.fill == PatternFill() or (hasattr(c.fill.start_color, 'rgb') and c.fill.start_color.rgb == '00000000'):
            c.fill = fill_white

# ============================================================
# 14. SAVE
# ============================================================
wb.save(OUTPUT_FILE)
print(f'Charts & Visuals dashboard built → {OUTPUT_FILE}')
print(f'  8 charts in 2-column grid layout')
print(f'  3 sections with navy header bars')
print(f'  Dropdown in C6, data tables at row {LT_START}+')
