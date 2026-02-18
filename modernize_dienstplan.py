#!/usr/bin/env python3
"""
Modernize Dienstplanübersicht 2026.xlsm
- Modern font (Calibri instead of Arial)
- Clean, modern color palette for shift codes
- Better contrast, boomer-friendly large text
- Consistent thin borders
- Professional header styling
- Preserves all formulas, VBA macros, and data
"""

import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Border, Side, Alignment, NamedStyle
)
from openpyxl.utils import get_column_letter
from copy import copy
import warnings
warnings.filterwarnings("ignore")

# ============================================================
# MODERN COLOR PALETTE
# ============================================================
# Shift code colors - modern flat design, high contrast, distinguishable
SHIFT_COLORS = {
    # Ambulanz I
    'FI':   {'bg': '4472C4', 'fg': 'FFFFFF'},  # Strong blue
    'SI':   {'bg': 'ED7D31', 'fg': 'FFFFFF'},  # Warm orange
    'NI':   {'bg': '2F5496', 'fg': 'FFFFFF'},  # Dark navy

    # Ambulanz II
    'FII':  {'bg': '5B9BD5', 'fg': '000000'},  # Medium blue
    'SII':  {'bg': 'F4B183', 'fg': '000000'},  # Light orange

    # Diagnostik
    'DI':   {'bg': '00B0F0', 'fg': '000000'},  # Bright cyan
    'DII':  {'bg': '9DC3E6', 'fg': '000000'},  # Light blue

    # Sonderdienst
    'SD':   {'bg': 'BFBFBF', 'fg': '000000'},  # Medium gray
    'SD 4h': {'bg': 'D9D9D9', 'fg': '000000'}, # Light gray

    # Koordination
    'KT':   {'bg': '7030A0', 'fg': 'FFFFFF'},  # Purple
    'KU':   {'bg': 'BF8F00', 'fg': 'FFFFFF'},  # Dark gold/brown

    # Gleittag / Sonstiges
    'GT':   {'bg': 'FFD966', 'fg': '000000'},  # Warm yellow
    'NST':  {'bg': 'E2EFDA', 'fg': '000000'},  # Very light green
    'EH':   {'bg': '00B0F0', 'fg': 'FFFFFF'},  # Cyan (Erste Hilfe)

    # Urlaub
    'U':    {'bg': '70AD47', 'fg': 'FFFFFF'},  # Fresh green

    # Fortbildung
    'Fobi': {'bg': 'C00000', 'fg': 'FFFFFF'},  # Dark red
    'Fort': {'bg': 'C00000', 'fg': 'FFFFFF'},

    # Abwesend / Krank / Kur
    'A':    {'bg': 'FF6F61', 'fg': 'FFFFFF'},  # Coral red
    'Kur':  {'bg': 'FF6F61', 'fg': 'FFFFFF'},
    'FZC':  {'bg': 'D9D9D9', 'fg': '000000'},

    # RTW
    'RTW- Des.': {'bg': '305496', 'fg': 'FFFFFF'},

    # T-ZUG
    'T-ZUG': {'bg': '70AD47', 'fg': 'FFFFFF'},
}

# Part-time shifts (in parentheses)
PARTTIME_KEYWORDS = ['(DI)', '(DII)', '(FI)', '(FII)', '(SI)', '(SII)', '(NI)', '(NII)']

# Weekend/Holiday colors
WEEKEND_BG = 'F2F2F2'       # Light gray for Sa/So columns
HOLIDAY_BG = 'FCE4EC'       # Soft pink for holidays
HEADER_BG = '1F4E79'        # Dark blue header
HEADER_FG = 'FFFFFF'        # White text on headers
SUBHEADER_BG = 'D6E4F0'     # Light blue for KW/day rows
NAME_BG = 'E8EEF4'          # Very light blue for name cells
KW_BG = 'BDD7EE'            # Medium light blue for KW rows
MONTH_BG = '1F4E79'         # Same dark blue as header
YEAR_BG = '1F4E79'

# Modern thin border style
THIN_BORDER = Border(
    left=Side(style='thin', color='B4C6E7'),
    right=Side(style='thin', color='B4C6E7'),
    top=Side(style='thin', color='B4C6E7'),
    bottom=Side(style='thin', color='B4C6E7'),
)

THICK_BOTTOM = Border(
    left=Side(style='thin', color='B4C6E7'),
    right=Side(style='thin', color='B4C6E7'),
    top=Side(style='thin', color='B4C6E7'),
    bottom=Side(style='medium', color='1F4E79'),
)

NO_BORDER = Border()

# Font definitions
FONT_HEADER = Font(name='Calibri', size=22, bold=True, color=HEADER_FG)
FONT_MONTH = Font(name='Calibri', size=18, bold=True, color=HEADER_FG)
FONT_YEAR = Font(name='Calibri', size=18, bold=False, color='B4C6E7')
FONT_KW = Font(name='Calibri', size=10, bold=True, color='1F4E79')
FONT_DAY_NUM = Font(name='Calibri', size=13, bold=True, color='333333')
FONT_DAY_NAME = Font(name='Calibri', size=12, bold=False, color='555555')
FONT_EMPLOYEE = Font(name='Calibri', size=13, bold=True, color='1F4E79')
FONT_SHIFT = Font(name='Calibri', size=11, bold=True, color='000000')
FONT_SHIFT_WHITE = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
FONT_HOLIDAY = Font(name='Calibri', size=9, bold=True, color='C00000')
FONT_FERIEN = Font(name='Calibri', size=9, italic=True, color='C00000')

# Alignment
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=False)
CENTER_WRAP = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT_CENTER = Alignment(horizontal='left', vertical='center')


def get_shift_style(value):
    """Determine fill and font color based on shift code value."""
    if not value or not isinstance(value, str):
        return None, None

    val = value.strip()

    # Check for "U    alt" (Urlaub alt)
    if 'U' in val and 'alt' in val:
        return (
            PatternFill(start_color='70AD47', end_color='70AD47', fill_type='solid'),
            Font(name='Calibri', size=10, bold=True, color='FFFFFF')
        )

    # Check for part-time shifts like "(DII) 2h"
    for pt in PARTTIME_KEYWORDS:
        if val.startswith(pt):
            # Use parent shift color but lighter
            parent = pt.strip('()')
            if parent in SHIFT_COLORS:
                sc = SHIFT_COLORS[parent]
                return (
                    PatternFill(start_color=sc['bg'], end_color=sc['bg'], fill_type='solid'),
                    Font(name='Calibri', size=9, bold=False, color=sc['fg'])
                )

    # Check for "1.AT" (erster Arbeitstag)
    if val == '1.AT':
        return (
            PatternFill(start_color='FFD966', end_color='FFD966', fill_type='solid'),
            Font(name='Calibri', size=9, bold=True, color='000000')
        )

    # Direct shift code match
    if val in SHIFT_COLORS:
        sc = SHIFT_COLORS[val]
        return (
            PatternFill(start_color=sc['bg'], end_color=sc['bg'], fill_type='solid'),
            Font(name='Calibri', size=11, bold=True, color=sc['fg'])
        )

    return None, None


def is_weekend_day(ws, row, col, day_name_rows):
    """Check if a column corresponds to a weekend day (Sa or So)."""
    for dnr in day_name_rows:
        day_cell = ws.cell(row=dnr, column=col)
        if day_cell.value in ('Sa', 'So'):
            return True
    return False


def style_dienstplan(ws):
    """Apply modern styling to the Dienstplan sheet."""
    max_col = ws.max_column
    max_row = ws.max_row

    print(f"  Styling Dienstplan: {max_row} rows x {max_col} cols")

    # Identify the 3 horizontal strips:
    # Strip 1: rows 2-19 (Jan-Jun)
    # Strip 2: rows 20-36 (Jul-Dec)
    # Strip 3: rows 37-53 (Jan-Jun next year, if present)
    strips = [
        {'header_row': 2, 'month_row': 3, 'holiday_row': 4, 'kw_row': 5,
         'daynum_row': 6, 'dayname_row': 7, 'emp_start': 8, 'emp_end': 19},
        {'header_row': None, 'month_row': 20, 'holiday_row': 21, 'kw_row': 22,
         'daynum_row': 23, 'dayname_row': 24, 'emp_start': 25, 'emp_end': 36},
    ]
    # Check if strip 3 exists
    if max_row >= 37:
        strips.append(
            {'header_row': None, 'month_row': 37, 'holiday_row': 38, 'kw_row': 39,
             'daynum_row': 40, 'dayname_row': 41, 'emp_start': 42, 'emp_end': 53}
        )

    # Collect all day_name_rows for weekend detection
    day_name_rows = [s['dayname_row'] for s in strips]

    for strip_idx, strip in enumerate(strips):
        print(f"  Processing strip {strip_idx + 1}...")

        # --- YEAR/HEADER ROW (only strip 1) ---
        if strip.get('header_row'):
            hr = strip['header_row']
            # Style B cell (Jahr label)
            cell_b = ws.cell(row=hr, column=2)
            cell_b.font = FONT_HEADER
            cell_b.fill = PatternFill(start_color=HEADER_BG, end_color=HEADER_BG, fill_type='solid')
            cell_b.alignment = CENTER
            cell_b.border = THIN_BORDER

            # Style the year value cell and surrounding
            for c in range(3, max_col + 1):
                cell = ws.cell(row=hr, column=c)
                if cell.value is not None:
                    cell.font = FONT_HEADER
                cell.fill = PatternFill(start_color=HEADER_BG, end_color=HEADER_BG, fill_type='solid')
                cell.alignment = CENTER
                cell.border = THIN_BORDER

        # --- MONTH NAME ROW ---
        mr = strip['month_row']
        for c in range(2, max_col + 1):
            cell = ws.cell(row=mr, column=c)
            cell.fill = PatternFill(start_color=MONTH_BG, end_color=MONTH_BG, fill_type='solid')
            cell.border = THIN_BORDER
            cell.alignment = CENTER
            if cell.value:
                val = str(cell.value)
                if val.startswith('='):
                    # Year reference formula
                    cell.font = FONT_YEAR
                elif len(val) <= 4:
                    # Month abbreviation
                    cell.font = FONT_MONTH
                else:
                    cell.font = Font(name='Calibri', size=14, bold=True, color=HEADER_FG)

        # --- HOLIDAY ROW ---
        hlr = strip['holiday_row']
        for c in range(2, max_col + 1):
            cell = ws.cell(row=hlr, column=c)
            if cell.value and isinstance(cell.value, str) and 'ferien' in cell.value.lower():
                cell.font = FONT_FERIEN
                cell.fill = PatternFill(start_color=HOLIDAY_BG, end_color=HOLIDAY_BG, fill_type='solid')
            else:
                cell.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
            cell.border = THIN_BORDER
            cell.alignment = CENTER

        # --- KW ROW ---
        kwr = strip['kw_row']
        for c in range(2, max_col + 1):
            cell = ws.cell(row=kwr, column=c)
            cell.fill = PatternFill(start_color=KW_BG, end_color=KW_BG, fill_type='solid')
            cell.font = FONT_KW
            cell.border = THIN_BORDER
            cell.alignment = CENTER

        # --- DAY NUMBER ROW ---
        dnr = strip['daynum_row']
        for c in range(2, max_col + 1):
            cell = ws.cell(row=dnr, column=c)
            cell.fill = PatternFill(start_color=SUBHEADER_BG, end_color=SUBHEADER_BG, fill_type='solid')
            cell.font = FONT_DAY_NUM
            cell.border = THIN_BORDER
            cell.alignment = CENTER

        # --- DAY NAME ROW ---
        dnamer = strip['dayname_row']
        for c in range(2, max_col + 1):
            cell = ws.cell(row=dnamer, column=c)
            val = cell.value
            cell.border = THICK_BOTTOM
            cell.alignment = CENTER

            if val in ('Sa', 'So'):
                cell.fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
                cell.font = Font(name='Calibri', size=12, bold=True, color='C00000')
            else:
                cell.fill = PatternFill(start_color=SUBHEADER_BG, end_color=SUBHEADER_BG, fill_type='solid')
                cell.font = FONT_DAY_NAME

        # --- EMPLOYEE DATA ROWS ---
        for r in range(strip['emp_start'], strip['emp_end'] + 1):
            # Name cell (column B)
            name_cell = ws.cell(row=r, column=2)
            name_cell.font = FONT_EMPLOYEE
            name_cell.fill = PatternFill(start_color=NAME_BG, end_color=NAME_BG, fill_type='solid')
            name_cell.alignment = LEFT_CENTER
            name_cell.border = Border(
                left=Side(style='thin', color='B4C6E7'),
                right=Side(style='medium', color='1F4E79'),
                top=Side(style='thin', color='B4C6E7'),
                bottom=Side(style='thin', color='B4C6E7'),
            )

            # Data cells
            for c in range(3, max_col + 1):
                cell = ws.cell(row=r, column=c)
                cell.border = THIN_BORDER
                cell.alignment = CENTER

                val = cell.value
                if val is not None and isinstance(val, str) and val.strip():
                    shift_fill, shift_font = get_shift_style(val)
                    if shift_fill:
                        cell.fill = shift_fill
                        cell.font = shift_font
                    else:
                        # Unknown code - keep neutral
                        cell.fill = PatternFill(start_color='F5F5F5', end_color='F5F5F5', fill_type='solid')
                        cell.font = Font(name='Calibri', size=11, bold=False, color='333333')
                else:
                    # Empty cell - check if weekend
                    is_we = is_weekend_day(ws, r, c, day_name_rows)
                    if is_we:
                        cell.fill = PatternFill(start_color=WEEKEND_BG, end_color=WEEKEND_BG, fill_type='solid')
                    else:
                        cell.fill = PatternFill(fill_type=None)  # No fill for empty weekday cells

        # --- Column A (spacer) ---
        start_r = strip.get('header_row') or strip['month_row']
        for r in range(start_r, strip['emp_end'] + 1):
            ws.cell(row=r, column=1).fill = PatternFill(fill_type=None)

    # Adjust column widths
    ws.column_dimensions['A'].width = 1.5
    ws.column_dimensions['B'].width = 16
    for c in range(3, max_col + 1):
        cl = get_column_letter(c)
        ws.column_dimensions[cl].width = 5.8

    # Adjust row heights
    for strip in strips:
        if strip.get('header_row'):
            ws.row_dimensions[strip['header_row']].height = 32
        ws.row_dimensions[strip['month_row']].height = 28
        ws.row_dimensions[strip['holiday_row']].height = 16
        ws.row_dimensions[strip['kw_row']].height = 18
        ws.row_dimensions[strip['daynum_row']].height = 24
        ws.row_dimensions[strip['dayname_row']].height = 24
        for r in range(strip['emp_start'], strip['emp_end'] + 1):
            ws.row_dimensions[r].height = 26

    print("  Dienstplan styling complete!")


def style_monat(ws):
    """Apply modern styling to the Monat calculation sheet."""
    print("  Styling Monat sheet...")

    # Month blocks repeat every 16 rows starting at row 2
    month_starts = list(range(2, 194, 16))  # 2, 18, 34, ..., 178

    for ms in month_starts:
        if ms > ws.max_row:
            break

        # Month header row
        for c in range(1, 37):
            cell = ws.cell(row=ms, column=c)
            cell.font = Font(name='Calibri', size=11, bold=True, color='1F4E79')
            if cell.value and isinstance(cell.value, str) and cell.value in [
                'Januar', 'Februar', 'März', 'April', 'Mai', 'Juni',
                'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'
            ]:
                cell.font = Font(name='Calibri', size=13, bold=True, color='FFFFFF')
                cell.fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
            cell.border = THIN_BORDER

        # Column headers row (ms+1)
        header_r = ms + 1
        for c in range(1, 37):
            cell = ws.cell(row=header_r, column=c)
            cell.font = Font(name='Calibri', size=10, bold=True, color='1F4E79')
            cell.fill = PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid')
            cell.border = THIN_BORDER
            cell.alignment = CENTER

        # Data rows (ms+2 to ms+12)
        for r in range(ms + 2, min(ms + 14, ws.max_row + 1)):
            for c in range(1, 37):
                cell = ws.cell(row=r, column=c)
                cell.font = Font(name='Calibri', size=10, bold=False, color='333333')
                cell.border = THIN_BORDER
                cell.alignment = CENTER
                # Employee name column
                if c == 2:
                    cell.font = Font(name='Calibri', size=10, bold=True, color='1F4E79')

        # Summe row (ms+14)
        sum_r = ms + 14
        if sum_r <= ws.max_row:
            for c in range(1, 37):
                cell = ws.cell(row=sum_r, column=c)
                cell.font = Font(name='Calibri', size=10, bold=True, color='1F4E79')
                cell.fill = PatternFill(start_color='E8EEF4', end_color='E8EEF4', fill_type='solid')
                cell.border = THICK_BOTTOM

    # Reference table on the right (Y-AD columns)
    for r in range(2, 20):
        for c in range(25, 31):  # Y=25 to AD=30
            cell = ws.cell(row=r, column=c)
            cell.font = Font(name='Calibri', size=10, bold=(r <= 3), color='333333')
            cell.border = THIN_BORDER
            cell.alignment = CENTER

    print("  Monat styling complete!")


def style_jahr(ws):
    """Apply modern styling to the Jahr summary sheet."""
    print("  Styling Jahr sheet...")

    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = THIN_BORDER
            cell.alignment = CENTER

            if r <= 3:
                cell.font = Font(name='Calibri', size=11, bold=True, color='1F4E79')
                cell.fill = PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid')
            elif r == ws.max_row or (cell.value and str(cell.value) == 'Summe'):
                cell.font = Font(name='Calibri', size=10, bold=True, color='1F4E79')
                cell.fill = PatternFill(start_color='E8EEF4', end_color='E8EEF4', fill_type='solid')
            else:
                cell.font = Font(name='Calibri', size=10, bold=(c <= 2), color='333333')
                if c == 2:
                    cell.font = Font(name='Calibri', size=10, bold=True, color='1F4E79')

    print("  Jahr styling complete!")


def main():
    input_file = "Dienstplanübersicht 2026.xlsm"
    output_file = "Dienstplanübersicht 2026.xlsm"

    print(f"Loading {input_file}...")
    wb = openpyxl.load_workbook(input_file, keep_vba=True)

    print("\nModernizing styles...")
    style_dienstplan(wb['Dienstplan'])
    style_monat(wb['Monat'])
    style_jahr(wb['Jahr'])

    print(f"\nSaving to {output_file}...")
    wb.save(output_file)
    print("Done! File saved successfully.")
    print("\nColor Legend:")
    print("  FI  (Früh Amb.I)    = Strong Blue")
    print("  SI  (Spät Amb.I)    = Warm Orange")
    print("  NI  (Nacht Amb.I)   = Dark Navy")
    print("  FII (Früh Amb.II)   = Medium Blue")
    print("  SII (Spät Amb.II)   = Light Orange")
    print("  DI  (Diagnostik I)  = Bright Cyan")
    print("  DII (Diagnostik II) = Light Blue")
    print("  SD  (Sonderdienst)  = Gray")
    print("  KT  (Koordination)  = Purple")
    print("  KU  (Kuppenheim)    = Dark Gold")
    print("  GT  (Gleittag)      = Warm Yellow")
    print("  U   (Urlaub)        = Fresh Green")
    print("  EH  (Erste Hilfe)   = Cyan")
    print("  Fobi (Fortbildung)  = Dark Red")
    print("  NST (Nicht-Schicht) = Light Green")
    print("  A   (Abwesend)      = Coral")
    print("  Weekend (Sa/So)     = Light Gray bg")


if __name__ == '__main__':
    main()
