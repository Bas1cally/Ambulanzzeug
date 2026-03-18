#!/usr/bin/env python3
"""
Creates EH_Kurs_2026.xlsx – Consolidated First Aid Course Management.
New Bescheinigungen sheet: modern DGUV 2023 layout, 2 certs per A4 landscape
page, cut in the middle (each half = A5 portrait). 20 participants = 10 pages.

Sheets: Teilnehmer | Bescheinigungen | BG-Liste | Lehrgangsdoku | Hilfslisten
"""
import re
import os
import shutil
import zipfile
import openpyxl
from copy import copy
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.pagebreak import Break
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.drawing.image import Image as XLImage
import openpyxl.worksheet.properties

# ══════════════════════════════════════════════════════════════════════════════
# Constants
# ══════════════════════════════════════════════════════════════════════════════
MAX_TN = 20

INSTRUCTORS = [
    "Putschler, Walter",
    "Kaya, Michelle",
    "Krempl, Elke",
    "Breig, Bernd",
    "Zuber, Harry",
    "Jochim, Benjamin",
    "Siebert, Emanuel",
]

KURSARTEN = ["EH-Ausbildung", "EH-Fortbildung", "EH-Schulung"]

COMPANY         = "Mercedes-Benz AG"
COMPANY_STREET  = "Mercedesstr. 1"
COMPANY_PLZ     = "76437 Rastatt"
COMPANY_WERK    = "PKW-Werk Rastatt"
QSEH_KENNZIFFER = "7.0103"
LOCATION_DETAIL = "Gebäude 39, Gesundheitszentrum 1.OG"
ARZT            = "Dr. Schmidt, Sabine"
ARZT_TEL        = "07222/91-22111"

BASE_DIR  = "/home/user/Ambulanzzeug"
ORIG_BG   = f"{BASE_DIR}/Neu EH/210923_EH_BG_Liste.xlsx"
IMG_CROSS = f"{BASE_DIR}/Neu EH/eh_cross.png"
IMG_QR1   = f"{BASE_DIR}/Neu EH/eh_qr1.png"
IMG_QR2   = f"{BASE_DIR}/Neu EH/eh_qr2.png"

# ══════════════════════════════════════════════════════════════════════════════
# Styles – general
# ══════════════════════════════════════════════════════════════════════════════
DARK_BLUE    = "0F3460"
HEADER_BG    = PatternFill(start_color=DARK_BLUE, end_color=DARK_BLUE, fill_type="solid")
LIGHT_BLUE   = PatternFill(start_color="E8EAF6", end_color="E8EAF6", fill_type="solid")
LIGHT_GRAY   = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
WHITE_BG     = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
INPUT_BG     = PatternFill(start_color="FFF9C4", end_color="FFF9C4", fill_type="solid")
GREEN_BG     = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")

TITLE_FONT  = Font(name="Calibri", size=16, bold=True,  color="FFFFFF")
HEADER_FONT = Font(name="Calibri", size=11, bold=True,  color="FFFFFF")
LABEL_FONT  = Font(name="Calibri", size=10, bold=True,  color="333333")
VALUE_FONT  = Font(name="Calibri", size=11,             color="333333")
INPUT_FONT  = Font(name="Calibri", size=11,             color="0F3460")
SMALL_FONT  = Font(name="Calibri", size=9,              color="666666")

# Lehrgangsdoku fonts
LD_TITLE    = Font(name="Arial", size=18, bold=True, color="000000")
LD_SUBTITLE = Font(name="Arial", size=10,            color="000000")
LD_SECTION  = Font(name="Arial", size=12, bold=True, color="000000")
LD_LABEL    = Font(name="Arial", size=10,            color="000000")
LD_LABEL_SM = Font(name="Arial", size=8,             color="000000")
LD_VALUE    = Font(name="Arial", size=12,            color="000000")
LD_VALUE_UL = Font(name="Arial", size=12, underline="single", color="000000")
LD_FOOT     = Font(name="Arial", size=8,             color="666666")

thin_border = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)
bottom_line   = Border(bottom=Side(style="thin", color="000000"))
center        = Alignment(horizontal="center", vertical="center")
left_center   = Alignment(horizontal="left",   vertical="center")
right_center  = Alignment(horizontal="right",  vertical="center")
left_wrap     = Alignment(horizontal="left",   vertical="center", wrap_text=True)
center_wrap   = Alignment(horizontal="center", vertical="center", wrap_text=True)

# ── Certificate-specific styles ───────────────────────────────────────────────
CERT_GRAY  = PatternFill(start_color="EBEBEB", end_color="EBEBEB", fill_type="solid")
CERT_WHITE = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

CT_TITLE   = Font(name="Calibri", size=14, bold=True, color="404040")
CT_SUB     = Font(name="Calibri", size=9,             color="404040")
CT_CHECK   = Font(name="Calibri", size=9,             color="404040")
CT_BODY    = Font(name="Calibri", size=9,             color="000000")
CT_BODY_B  = Font(name="Calibri", size=9,  bold=True, color="000000")
CT_SMALL   = Font(name="Calibri", size=7.5,           color="555555")
CT_FOOT    = Font(name="Calibri", size=7,             color="555555")

# Dashed cut-line border (between left and right cert)
CUT_LEFT  = Border(right=Side(style="dashed", color="AAAAAA"))
CUT_RIGHT = Border(left=Side(style="dashed",  color="AAAAAA"))

# Underline border for fill-in lines
UL_BORDER = Border(bottom=Side(style="thin", color="555555"))

# ══════════════════════════════════════════════════════════════════════════════
# Helpers
# ══════════════════════════════════════════════════════════════════════════════
def set_cell(ws, row, col, value, font=None, fill=None, alignment=None,
             border=None, number_format=None):
    cell = ws.cell(row=row, column=col, value=value)
    if font:          cell.font          = font
    if fill:          cell.fill          = fill
    if alignment:     cell.alignment     = alignment
    if border:        cell.border        = border
    if number_format: cell.number_format = number_format
    return cell


def fill_row(ws, row, col_start, col_end, fill):
    """Apply a fill to a range of cells in one row."""
    for c in range(col_start, col_end + 1):
        ws.cell(row=row, column=c).fill = fill


def merge_set(ws, r1, c1, r2, c2, value=None, font=None, fill=None,
              alignment=None, border=None, number_format=None):
    """Merge cells and set value/style on the top-left cell."""
    ws.merge_cells(start_row=r1, start_column=c1,
                   end_row=r2, end_column=c2)
    cell = ws.cell(row=r1, column=c1)
    if value is not None:   cell.value         = value
    if font:                cell.font          = font
    if fill:                cell.fill          = fill
    if alignment:           cell.alignment     = alignment
    if border:              cell.border        = border
    if number_format:       cell.number_format = number_format
    # Apply fill to all cells in the merged range
    if fill:
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(row=r, column=c).fill = fill
    return cell


def apply_cut_line(ws, rows_start, rows_end, cut_col):
    """Draw dashed cut border on col cut_col (right side) and cut_col+1 (left side)."""
    for r in range(rows_start, rows_end + 1):
        ws.cell(row=r, column=cut_col).border     = CUT_LEFT
        ws.cell(row=r, column=cut_col + 1).border = CUT_RIGHT


# ══════════════════════════════════════════════════════════════════════════════
# Formula translation for BG-Liste
# ══════════════════════════════════════════════════════════════════════════════
def translate_formula(formula):
    if not isinstance(formula, str) or "[1]Sheet1!" not in formula:
        return formula
    formula = formula.replace("[1]Sheet1!$D$2", "Teilnehmer!C5")
    formula = formula.replace("[1]Sheet1!$E$2", "Teilnehmer!C3")
    formula = formula.replace("[1]Sheet1!$F$2", "Teilnehmer!C4")
    formula = formula.replace("[1]Sheet1!$G$2", "Teilnehmer!F3")
    formula = formula.replace("[1]Sheet1!$H$2", "Teilnehmer!H3")

    def replace_participant(match):
        col = match.group(1)
        row = int(match.group(2))
        col_map = {"A": "B", "B": "C", "C": "D"}
        return f"Teilnehmer!{col_map[col]}{row + 9}"

    formula = re.sub(r'\[1\]Sheet1!\$?([ABC])\$?(\d+)', replace_participant, formula)
    return formula


def fix_formulas(ws):
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.value = translate_formula(cell.value)


def copy_sheet(ws_src, ws_dst, row_offset=0):
    from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
    if row_offset == 0:
        for col_letter, dim in ws_src.column_dimensions.items():
            ws_dst.column_dimensions[col_letter].width  = dim.width
            ws_dst.column_dimensions[col_letter].hidden = dim.hidden
    for row_idx, dim in ws_src.row_dimensions.items():
        if dim.height is not None:
            ws_dst.row_dimensions[row_idx + row_offset].height = dim.height
    for row in ws_src.iter_rows():
        for cell in row:
            new_cell = ws_dst.cell(row=cell.row + row_offset, column=cell.column)
            new_cell.value = cell.value
            if cell.has_style:
                new_cell.font         = copy(cell.font)
                new_cell.fill         = copy(cell.fill)
                new_cell.alignment    = copy(cell.alignment)
                new_cell.border       = copy(cell.border)
                new_cell.number_format = cell.number_format
    for mcr in ws_src.merged_cells.ranges:
        if row_offset == 0:
            ws_dst.merge_cells(str(mcr))
        else:
            ws_dst.merge_cells(
                start_row=mcr.min_row + row_offset, start_column=mcr.min_col,
                end_row=mcr.max_row   + row_offset, end_column=mcr.max_col,
            )
    for img in ws_src._images:
        if row_offset == 0:
            ws_dst._images.append(img)
        else:
            new_img = copy(img)
            anchor  = img.anchor
            if isinstance(anchor, TwoCellAnchor):
                new_anchor       = TwoCellAnchor()
                new_anchor._from = AnchorMarker(
                    col=anchor._from.col, colOff=anchor._from.colOff,
                    row=anchor._from.row + row_offset, rowOff=anchor._from.rowOff)
                new_anchor.to    = AnchorMarker(
                    col=anchor.to.col, colOff=anchor.to.colOff,
                    row=anchor.to.row + row_offset, rowOff=anchor.to.rowOff)
                new_img.anchor   = new_anchor
            ws_dst._images.append(new_img)
    if row_offset == 0:
        ws_dst.page_setup.paperSize    = ws_src.page_setup.paperSize
        ws_dst.page_setup.orientation  = ws_src.page_setup.orientation
        ws_dst.page_setup.fitToWidth   = ws_src.page_setup.fitToWidth
        ws_dst.page_setup.fitToHeight  = ws_src.page_setup.fitToHeight
        src = ws_src.page_margins
        ws_dst.page_margins = PageMargins(
            left=src.left, right=src.right, top=src.top, bottom=src.bottom,
            header=src.header, footer=src.footer)
    if ws_src.row_breaks:
        for brk in ws_src.row_breaks.brk:
            ws_dst.row_breaks.append(Break(id=brk.id + row_offset, man=brk.man))


# ══════════════════════════════════════════════════════════════════════════════
# CERTIFICATE SHEET BUILDER
# ══════════════════════════════════════════════════════════════════════════════
# Layout: 2 certificates side by side on A4 landscape (each = A5 portrait).
# Cols 1-9  = left cert   (A-I)
# Col 10    = cut line    (J)
# Cols 11-19 = right cert (K-S)
#
# Column widths (symmetric, each cert = ~45 char units = ~142mm):
#   C1/C11: 1.5  outer margin
#   C2/C12: 4.5  logo area
#   C3/C13: 4.5  logo area
#   C4/C14: 9.5  main content
#   C5/C15: 9.5  main content
#   C6/C16: 7.0  content
#   C7/C17: 5.0  content
#   C8/C18: 3.0  content
#   C9/C19: 1.5  inner margin (near cut)
#   C10:    0.5  cut-line column
# ──────────────────────────────────────────────────────────────────────────────
#
# Rows per cert page block (CERT_ROWS = 35):
CERT_ROWS = 35

# Row heights in points (must sum to ≈ 554pt for A4 landscape 0.5cm margins)
CERT_ROW_H = [
    12,   # 1  top margin (gray header)
    28,   # 2  "Bescheinigung" title
    16,   # 3  "über die Teilnahme an einer Erste-Hilfe-"
    15,   # 4  □ Ausbildung*
    15,   # 5  □ Fortbildung
    15,   # 6  □ Schulung in Bildungs- …
    13,   # 7  … für Kinder
    10,   # 8  bottom header margin
    9,    # 9  gap (white)
    18,   # 10 Name / Vorname / geb. am
    9,    # 11 underline / labels row
    16,   # 12 "hat an dem 9 Unterrichtseinheiten…"
    9,    # 13 spacer
    16,   # 14 "am … in der Zeit von … Uhr bis … Uhr"
    9,    # 15 spacer
    16,   # 16 "unter der Leitung von … erfolgreich teilgenommen."
    9,    # 17 spacer
    14,   # 18 "Teilnehmerunterlagen ausgehändigt:  □ Ja  □ Nein"
    8,    # 19 spacer
    16,   # 20 QR code area
    16,   # 21
    16,   # 22
    16,   # 23
    16,   # 24
    13,   # 25 QR label line 1
    12,   # 26 QR label line 2
    8,    # 27 spacer
    20,   # 28 signature line
    13,   # 29 "Ort        Datum        Unterschrift der Lehrkraft"
    8,    # 30 spacer before footer
    3,    # 31 thin separator
    18,   # 32 Name der ermächtigten Stelle: ___________
    16,   # 33 Kennnummer ___
    16,   # 34 Registriernummer ___
    11,   # 35 footnote + "Stand Oktober 2023"
]
assert len(CERT_ROW_H) == CERT_ROWS, f"Expected {CERT_ROWS} row heights"

# Column definitions: (width, ) for cols 1-19
# Cols 1-9 = left cert, col 10 = cut, cols 11-19 = right cert (mirror of 1-9)
CERT_COL_W = [1.5, 4.5, 4.5, 9.5, 9.5, 7.0, 5.0, 3.0, 1.5,  # left cert
              0.5,                                               # cut line
              1.5, 4.5, 4.5, 9.5, 9.5, 7.0, 5.0, 3.0, 1.5]   # right cert


def tn_ref(tn_row, col):
    """Return a Teilnehmer sheet cell reference formula string."""
    return f"=Teilnehmer!{col}{tn_row}"


def _build_one_cert(ws, page_row, col_offset, tn_row):
    """
    Build one certificate.
    page_row: absolute row of the start of this cert page block (1-based)
    col_offset: 0 for left cert, 10 for right cert
    tn_row: Teilnehmer data row (11 for TN1 … 30 for TN20)
    """
    r = page_row  # convenience alias
    o = col_offset  # column offset

    # Columns used inside one cert (1-based, offset applied):
    cA = 1 + o   # outer margin / logo
    cB = 2 + o   # logo
    cC = 3 + o   # logo end
    cD = 4 + o   # content start
    cE = 5 + o
    cF = 6 + o
    cG = 7 + o
    cH = 8 + o
    cI = 9 + o   # inner margin

    # ── GRAY HEADER (rows 1-8) ────────────────────────────────────────────────
    gray_rows = list(range(r, r + 8))
    for gr in gray_rows:
        for c in range(cA, cI + 1):
            ws.cell(row=gr, column=c).fill = CERT_GRAY

    # Row 2: Title "Bescheinigung"
    merge_set(ws, r + 1, cD, r + 1, cI,
              value="Bescheinigung",
              font=CT_TITLE,
              fill=CERT_GRAY,
              alignment=Alignment(horizontal="left", vertical="center"))

    # Row 3: subtitle
    merge_set(ws, r + 2, cD, r + 2, cI,
              value="über die Teilnahme an einer Erste-Hilfe-",
              font=CT_SUB, fill=CERT_GRAY,
              alignment=left_center)

    # Rows 4-7: Kursart checkboxes with auto-tick from Teilnehmer!F4
    kursart_ref = "Teilnehmer!F4"
    for idx, (label, kursart) in enumerate([
        ("Ausbildung\u00b9 f\u00fcr betriebliche Ersthelfende", "EH-Ausbildung"),
        ("Fortbildung f\u00fcr betriebliche Ersthelfende",      "EH-Fortbildung"),
        ("Schulung in Bildungs- und Betreuungs-",               "EH-Schulung"),
        ("einrichtungen f\u00fcr Kinder",                       None),
    ]):
        row_r = r + 3 + idx
        for c in range(cA, cI + 1):
            ws.cell(row=row_r, column=c).fill = CERT_GRAY
        if kursart:
            # checkbox tick symbol via formula
            tick = ws.cell(row=row_r, column=cD)
            tick.value = f'=IF({kursart_ref}="{kursart}","\u2611","\u2610")'
            tick.font  = Font(name="Segoe UI Symbol", size=9, color="404040")
            tick.fill  = CERT_GRAY
            tick.alignment = Alignment(horizontal="center", vertical="center")
            merge_set(ws, row_r, cD + 1, row_r, cI,
                      value=label, font=CT_CHECK, fill=CERT_GRAY, alignment=left_center)
        else:
            merge_set(ws, row_r, cD, row_r, cI,
                      value=label, font=CT_CHECK, fill=CERT_GRAY,
                      alignment=Alignment(horizontal="left", vertical="center",
                                          indent=1))

    # ── WHITE BODY (rows 9-30) ────────────────────────────────────────────────
    body_start = r + 8
    body_end   = r + 29
    for br in range(body_start, body_end + 1):
        for c in range(cA, cI + 1):
            ws.cell(row=br, column=c).fill = CERT_WHITE

    # Row 10: Name / Vorname / geb. am: underline cells
    rn = r + 9  # name row
    # Name (cols cB..cE)
    merge_set(ws, rn, cB, rn, cE,
              value=f"=Teilnehmer!B{tn_row}",
              font=CT_BODY_B, fill=CERT_WHITE,
              alignment=left_center, border=UL_BORDER)
    # Vorname (cols cF..cG)
    merge_set(ws, rn, cF, rn, cG,
              value=f"=Teilnehmer!C{tn_row}",
              font=CT_BODY_B, fill=CERT_WHITE,
              alignment=left_center, border=UL_BORDER)
    # geb. am: label
    ws.cell(row=rn, column=cH, value="geb. am:").font = CT_BODY
    ws.cell(row=rn, column=cH).fill = CERT_WHITE
    ws.cell(row=rn, column=cH).alignment = left_center
    # geb. am value with border
    set_cell(ws, rn, cI, f"=TEXT(Teilnehmer!D{tn_row},\"DD.MM.YYYY\")",
             font=CT_BODY_B, fill=CERT_WHITE, alignment=left_center,
             border=UL_BORDER)

    # Row 11: labels under name line
    rl = r + 10
    for c in range(cA, cI + 1):
        ws.cell(row=rl, column=c).fill = CERT_WHITE
    set_cell(ws, rl, cB, "Name", font=CT_SMALL, fill=CERT_WHITE,
             alignment=left_center)
    set_cell(ws, rl, cF, "Vorname", font=CT_SMALL, fill=CERT_WHITE,
             alignment=left_center)

    # Row 12: "hat an dem 9 Unterrichtseinheiten…"
    r12 = r + 11
    merge_set(ws, r12, cB, r12, cI,
              value=("hat an dem 9 Unterrichtseinheiten "
                     "(Nettounterrichtszeit 9 \u00d7 45 Minuten) umfassenden Lehrgang"),
              font=CT_BODY, fill=CERT_WHITE, alignment=left_wrap)

    # Row 14: "am … in der Zeit von … Uhr bis … Uhr"
    r14 = r + 13
    # "am"
    set_cell(ws, r14, cB, "am", font=CT_BODY, fill=CERT_WHITE, alignment=left_center)
    # date value
    date_cell = ws.cell(row=r14, column=cC)
    date_cell.value       = f"=TEXT(Teilnehmer!C3,\"DD.MM.YYYY\")"
    date_cell.font        = CT_BODY_B
    date_cell.fill        = CERT_WHITE
    date_cell.alignment   = center
    date_cell.border      = UL_BORDER
    # "in der Zeit von"
    set_cell(ws, r14, cD, "in der Zeit von", font=CT_BODY, fill=CERT_WHITE,
             alignment=left_center)
    # start time
    t1 = ws.cell(row=r14, column=cE)
    t1.value     = f"=TEXT(Teilnehmer!F3,\"HH:MM\")"
    t1.font      = CT_BODY_B; t1.fill = CERT_WHITE
    t1.alignment = center; t1.border = UL_BORDER
    # "Uhr bis"
    set_cell(ws, r14, cF, "Uhr bis", font=CT_BODY, fill=CERT_WHITE, alignment=center)
    # end time
    t2 = ws.cell(row=r14, column=cG)
    t2.value     = f"=TEXT(Teilnehmer!H3,\"HH:MM\")"
    t2.font      = CT_BODY_B; t2.fill = CERT_WHITE
    t2.alignment = center; t2.border = UL_BORDER
    # "Uhr"
    set_cell(ws, r14, cH, "Uhr", font=CT_BODY, fill=CERT_WHITE, alignment=left_center)

    # Row 16: "unter der Leitung von … erfolgreich teilgenommen."
    r16 = r + 15
    set_cell(ws, r16, cB, "unter der Leitung von", font=CT_BODY,
             fill=CERT_WHITE, alignment=left_center)
    lk = ws.cell(row=r16, column=cD)
    lk.value = f"=Teilnehmer!C4"; lk.font = CT_BODY_B
    lk.fill = CERT_WHITE; lk.alignment = center; lk.border = UL_BORDER
    ws.merge_cells(start_row=r16, start_column=cD,
                   end_row=r16,   end_column=cG)
    set_cell(ws, r16, cH, "erfolgreich", font=CT_BODY, fill=CERT_WHITE,
             alignment=left_center)

    r16b = r + 15
    # "teilgenommen." on same row at cI if it fits, else wrap
    set_cell(ws, r16b, cI, "teilgen.", font=CT_BODY, fill=CERT_WHITE,
             alignment=left_center)

    # Row 18: "Teilnehmerunterlagen ausgehändigt: □ Ja □ Nein"
    r18 = r + 17
    merge_set(ws, r18, cB, r18, cI,
              value='Teilnehmerunterlagen ausgeh\u00e4ndigt:    \u2610 Ja    \u2610 Nein',
              font=CT_BODY, fill=CERT_WHITE, alignment=left_center)

    # Rows 20-24: QR codes (images placed later) – just fill white
    for qr_r in range(r + 19, r + 25):
        for c in range(cA, cI + 1):
            ws.cell(row=qr_r, column=c).fill = CERT_WHITE

    # Row 25-26: QR labels
    r25 = r + 24
    merge_set(ws, r25, cB, r25, cD,
              value="Erste Hilfe Handbuch\n(DGUV Information 204-007)",
              font=CT_SMALL, fill=CERT_WHITE,
              alignment=Alignment(horizontal="center", vertical="top",
                                  wrap_text=True))
    merge_set(ws, r25, cE, r25, cG,
              value="Handbuch zur Ersten Hilfe in Bildungs-\nund Betreuungseinrichtungen für Kinder\n(DGUV Information 204-008)",
              font=CT_SMALL, fill=CERT_WHITE,
              alignment=Alignment(horizontal="center", vertical="top",
                                  wrap_text=True))

    # Row 28: Signature line
    r28 = r + 27
    for c in range(cA, cI + 1):
        ws.cell(row=r28, column=c).fill = CERT_WHITE
    # Ort + Datum
    merge_set(ws, r28, cB, r28, cD,
              value=f"=Teilnehmer!C7&\", den \"&TEXT(Teilnehmer!C3,\"DD.MM.YYYY\")",
              font=CT_BODY_B, fill=CERT_WHITE,
              alignment=left_center, border=UL_BORDER)
    # Unterschrift
    merge_set(ws, r28, cE, r28, cI,
              font=CT_BODY, fill=CERT_WHITE,
              alignment=left_center, border=UL_BORDER)

    # Row 29: labels
    r29 = r + 28
    for c in range(cA, cI + 1):
        ws.cell(row=r29, column=c).fill = CERT_WHITE
    set_cell(ws, r29, cB, "Ort", font=CT_SMALL, fill=CERT_WHITE,
             alignment=left_center)
    set_cell(ws, r29, cD, "Datum", font=CT_SMALL, fill=CERT_WHITE,
             alignment=left_center)
    set_cell(ws, r29, cE, "Unterschrift der Lehrkraft",
             font=CT_SMALL, fill=CERT_WHITE, alignment=left_center)

    # ── GRAY FOOTER (rows 31-35) ──────────────────────────────────────────────
    for fr in range(r + 30, r + CERT_ROWS):
        for c in range(cA, cI + 1):
            ws.cell(row=fr, column=c).fill = CERT_GRAY

    # Row 31: thin separator line (already gray, just a visual row)
    # Row 32: Name der ermächtigten Stelle
    r32 = r + 31
    set_cell(ws, r32, cB, "Name der ermächtigten Stelle:",
             font=CT_BODY, fill=CERT_GRAY, alignment=left_center)
    merge_set(ws, r32, cD, r32, cI,
              value=COMPANY,
              font=CT_BODY_B, fill=CERT_GRAY,
              alignment=left_center, border=UL_BORDER)

    # Row 33: Kennnummer
    r33 = r + 32
    merge_set(ws, r33, cB, r33, cC,
              value="Kennnummer gemäß\n§ 26 DGUV Vorschrift 1:",
              font=CT_SMALL, fill=CERT_GRAY,
              alignment=Alignment(horizontal="left", vertical="center",
                                  wrap_text=True))
    merge_set(ws, r33, cD, r33, cI,
              value=QSEH_KENNZIFFER,
              font=CT_BODY_B, fill=CERT_GRAY,
              alignment=left_center, border=UL_BORDER)

    # Row 34: Registriernummer
    r34 = r + 33
    set_cell(ws, r34, cB, "Registriernummer der Schulung:",
             font=CT_BODY, fill=CERT_GRAY, alignment=left_center)
    merge_set(ws, r34, cD, r34, cI,
              value=f"=Teilnehmer!C5",
              font=CT_BODY_B, fill=CERT_GRAY,
              alignment=left_center, border=UL_BORDER)

    # Row 35: footnote + Stand
    r35 = r + 34
    merge_set(ws, r35, cA, r35, cI,
              value=("\u00b9 Die Teilnahme an der Ausbildung in betrieblicher Erster Hilfe "
                     "gilt als Schulung in Erster Hilfe gem. § 19 FeV.   |   Stand Oktober 2023"),
              font=CT_FOOT, fill=CERT_GRAY,
              alignment=Alignment(horizontal="left", vertical="center"))


def build_cert_sheet(wb):
    """Build Bescheinigungen sheet: 2 certs per A4 landscape page, 10 pages."""
    ws = wb.create_sheet(title="Bescheinigungen")
    ws.sheet_properties.tabColor = "1B5E20"

    # ── Column widths ─────────────────────────────────────────────────────────
    for col_idx, width in enumerate(CERT_COL_W, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # ── Build 10 pages (20 certs, 2 per page) ────────────────────────────────
    total_rows = CERT_ROWS * 10
    for row in range(1, total_rows + 1):
        page_local_row = ((row - 1) % CERT_ROWS)
        ws.row_dimensions[row].height = CERT_ROW_H[page_local_row]

    for page in range(10):              # 0-9
        page_row   = page * CERT_ROWS + 1
        tn_left    = 11 + page * 2      # e.g. page 0 → TN1 (row 11), TN2 (row 12)
        tn_right   = tn_left + 1

        _build_one_cert(ws, page_row, col_offset=0,  tn_row=tn_left)
        _build_one_cert(ws, page_row, col_offset=10, tn_row=tn_right)

        # Dashed cut line (col 10 = J)
        apply_cut_line(ws, page_row, page_row + CERT_ROWS - 1, cut_col=10)

        # Page break after each page (except the last)
        if page < 9:
            ws.row_breaks.append(Break(id=page_row + CERT_ROWS - 1, man=True))

    # ── Place images (cross + QR codes) ──────────────────────────────────────
    # EMU per point: 1pt = 12700 EMU; 1 char width ≈ 48000 EMU (Calibri 11 default)
    # We place each image using a named-cell anchor approach.
    # Green cross: sits in header rows 2-7 of each page block, cols cB-cC (2-3)
    # QR codes: rows 20-24, left cert: cols cB-cC and cD-cE; right: same offset

    for page in range(10):
        page_row = page * CERT_ROWS + 1

        # Helper: pixel size for image (approx)
        # cross ≈ 22mm each side. At 96dpi: 22mm × 96/25.4 ≈ 83px
        # QR    ≈ 14mm each side → 53px

        for col_offset, qr1_col_start, qr2_col_start in [
            (0,  1, 3),   # left cert:  QR1 at col B(2), QR2 at col D(4)
            (10, 11, 13), # right cert: QR1 at col L(12), QR2 at col N(14)
        ]:
            # Row/col indices for image anchors (0-based for openpyxl Image)
            cross_row = page_row + 1 - 1   # row index 0-based = page_row+1 -1
            cross_col = 1 + col_offset     # col B or L (0-based: col 1 or 11)
            qr_row    = page_row + 19 - 1  # row 20 of page block

            # Green cross
            img_cross = XLImage(IMG_CROSS)
            img_cross.width  = 52   # pixels ≈ 14mm
            img_cross.height = 52
            img_cross.anchor = f"{get_column_letter(cross_col + 1)}{cross_row + 1}"
            ws.add_image(img_cross)

            # QR code 1
            img_q1 = XLImage(IMG_QR1)
            img_q1.width  = 44
            img_q1.height = 44
            img_q1.anchor = f"{get_column_letter(qr1_col_start + 1)}{qr_row + 1}"
            ws.add_image(img_q1)

            # QR code 2
            img_q2 = XLImage(IMG_QR2)
            img_q2.width  = 44
            img_q2.height = 44
            img_q2.anchor = f"{get_column_letter(qr2_col_start + 1)}{qr_row + 1}"
            ws.add_image(img_q2)

    # ── Page setup: A4 landscape, fit 1 page wide × 10 pages tall ────────────
    ws.page_setup.paperSize   = 9           # A4
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 10
    ws.sheet_properties.pageSetUpPr = \
        openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
    ws.page_margins = PageMargins(
        left=0.2, right=0.2, top=0.2, bottom=0.2, header=0.1, footer=0.1)
    ws.print_area = f"A1:S{CERT_ROWS * 10}"

    print(f"  Bescheinigungen: {10 * 3 * 2} images placed "
          f"(10 pages × 2 certs × 3 images)")
    return ws


# ══════════════════════════════════════════════════════════════════════════════
# Start fresh workbook
# ══════════════════════════════════════════════════════════════════════════════
print("Creating workbook...")
wb = openpyxl.Workbook()
# Remove default sheet
if "Sheet" in wb.sheetnames:
    del wb["Sheet"]


# ══════════════════════════════════════════════════════════════════════════════
# SHEET: Hilfslisten (created early for named range)
# ══════════════════════════════════════════════════════════════════════════════
ws_help = wb.create_sheet(title="Hilfslisten")
ws_help.sheet_properties.tabColor = "9E9E9E"

set_cell(ws_help, 1, 1, "Lehrkräfte",
    font=HEADER_FONT, fill=HEADER_BG, alignment=center, border=thin_border)
ws_help.column_dimensions["A"].width = 25

for i, name in enumerate(INSTRUCTORS):
    set_cell(ws_help, 2 + i, 1, name,
             font=VALUE_FONT, alignment=left_center, border=thin_border)

dn = DefinedName("Lehrkraefte", attr_text="Hilfslisten!$A$2:$A$20")
wb.defined_names.add(dn)


# ══════════════════════════════════════════════════════════════════════════════
# SHEET: Teilnehmer
# ══════════════════════════════════════════════════════════════════════════════
ws_tn = wb.create_sheet(title="Teilnehmer")
ws_tn.sheet_properties.tabColor = "0F3460"

ws_tn.column_dimensions["A"].width = 5
ws_tn.column_dimensions["B"].width = 18
ws_tn.column_dimensions["C"].width = 22
ws_tn.column_dimensions["D"].width = 16
ws_tn.column_dimensions["E"].width = 18
ws_tn.column_dimensions["F"].width = 22
ws_tn.column_dimensions["G"].width = 8
ws_tn.column_dimensions["H"].width = 12

# Row 1: Title
ws_tn.row_dimensions[1].height = 38
ws_tn.merge_cells("A1:H1")
set_cell(ws_tn, 1, 1, "EH-Kurs \u2013 Teilnehmerliste 2026",
    font=TITLE_FONT, fill=HEADER_BG,
    alignment=Alignment(horizontal="left", vertical="center"))

# Rows 3-8: Course Metadata
for r in range(3, 9):
    ws_tn.row_dimensions[r].height = 22

set_cell(ws_tn, 3, 2, "Datum:",       font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 3, 3, None,           font=INPUT_FONT, fill=INPUT_BG,
         alignment=left_center, border=thin_border, number_format="DD.MM.YYYY")
set_cell(ws_tn, 3, 5, "Beginn:",      font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 3, 6, None,           font=INPUT_FONT, fill=INPUT_BG,
         alignment=left_center, border=thin_border, number_format="HH:MM")
set_cell(ws_tn, 3, 7, "Ende:",        font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 3, 8, None,           font=INPUT_FONT, fill=INPUT_BG,
         alignment=left_center, border=thin_border, number_format="HH:MM")

set_cell(ws_tn, 4, 2, "Lehrkraft:",   font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 4, 3, None,           font=INPUT_FONT, fill=INPUT_BG,
         alignment=left_center, border=thin_border)
set_cell(ws_tn, 4, 5, "Kursart:",     font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 4, 6, None,           font=INPUT_FONT, fill=INPUT_BG,
         alignment=left_center, border=thin_border)

set_cell(ws_tn, 5, 2, "Registriernr.:",    font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 5, 3, None,                font=INPUT_FONT, fill=INPUT_BG,
         alignment=left_center, border=thin_border)
set_cell(ws_tn, 5, 5, "QSEH-Kennziffer:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 5, 6, QSEH_KENNZIFFER,    font=INPUT_FONT, fill=LIGHT_BLUE,
         alignment=left_center, border=thin_border)

set_cell(ws_tn, 6, 2, "Lehrgangsort:", font=LABEL_FONT, alignment=right_center)
ws_tn.merge_cells("C6:E6")
set_cell(ws_tn, 6, 3, f"{COMPANY}, {COMPANY_WERK}",
         font=INPUT_FONT, fill=INPUT_BG, alignment=left_center, border=thin_border)

set_cell(ws_tn, 7, 2, "Verantw. Arzt:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 7, 3, ARZT,             font=INPUT_FONT, fill=LIGHT_BLUE,
         alignment=left_center, border=thin_border)
set_cell(ws_tn, 7, 5, "Tel.:",          font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 7, 6, ARZT_TEL,         font=INPUT_FONT, fill=LIGHT_BLUE,
         alignment=left_center, border=thin_border)

set_cell(ws_tn, 8, 2, "Masken-Charge:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 8, 3, None,             font=INPUT_FONT, fill=INPUT_BG,
         alignment=left_center, border=thin_border)

# Data validations
dv_instructor = DataValidation(type="list", formula1="=Lehrkraefte",
    allow_blank=True, showDropDown=False)
dv_instructor.add("C4")
ws_tn.add_data_validation(dv_instructor)

dv_kursart = DataValidation(type="list",
    formula1='"' + ",".join(KURSARTEN) + '"',
    allow_blank=True, showDropDown=False)
dv_kursart.add("F4")
ws_tn.add_data_validation(dv_kursart)

# Row 10: Table Headers
ws_tn.row_dimensions[10].height = 25
for c, h in [(1, "Nr"), (2, "Nachname"), (3, "Vorname"),
             (4, "Geburtsdatum"), (5, "Ersthelfer (x)")]:
    set_cell(ws_tn, 10, c, h,
        font=HEADER_FONT, fill=HEADER_BG, alignment=center, border=thin_border)

# Rows 11-30: Participant Rows
for i in range(MAX_TN):
    r = 11 + i
    ws_tn.row_dimensions[r].height = 22
    alt = LIGHT_GRAY if i % 2 == 0 else WHITE_BG
    set_cell(ws_tn, r, 1, i + 1, font=VALUE_FONT, fill=alt,
             alignment=center, border=thin_border)
    for c in range(2, 6):
        nf = "DD.MM.YYYY" if c == 4 else None
        set_cell(ws_tn, r, c, None, font=INPUT_FONT, fill=alt,
                 alignment=left_center if c != 5 else center,
                 border=thin_border, number_format=nf)

dv_eh = DataValidation(type="list", formula1='"x"', allow_blank=True,
                       showDropDown=False)
dv_eh.add(f"E11:E{10 + MAX_TN}")
ws_tn.add_data_validation(dv_eh)

# Row 32: Counters
ws_tn.row_dimensions[32].height = 25
set_cell(ws_tn, 32, 2, "Gesamtanzahl:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 32, 3, f"=COUNTA(B11:B{10+MAX_TN})",
    font=Font(name="Calibri", size=14, bold=True, color="0F3460"),
    fill=GREEN_BG, alignment=center, border=thin_border)
set_cell(ws_tn, 32, 4, "davon Ersthelfer:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 32, 5, f'=COUNTIF(E11:E{10+MAX_TN},"x")',
    font=Font(name="Calibri", size=14, bold=True, color="0F3460"),
    fill=GREEN_BG, alignment=center, border=thin_border)

# Pflichtfeld warnings
ws_tn.conditional_formatting.add(f"B11:B{10+MAX_TN}", FormulaRule(
    formula=['AND($B11="",$C11<>"")'],
    fill=PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid")))
ws_tn.conditional_formatting.add(f"C11:C{10+MAX_TN}", FormulaRule(
    formula=['AND($C11="",$B11<>"")'],
    fill=PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid")))

# Print Setup
ws_tn.page_setup.orientation = "portrait"
ws_tn.page_setup.paperSize   = ws_tn.PAPERSIZE_A4
ws_tn.page_setup.fitToWidth  = 1
ws_tn.page_setup.fitToHeight = 1
ws_tn.sheet_properties.pageSetUpPr = \
    openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
ws_tn.page_margins = PageMargins(left=0.5, right=0.5, top=0.5, bottom=0.5)
ws_tn.freeze_panes = "A11"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET: Bescheinigungen (new modern DGUV layout)
# ══════════════════════════════════════════════════════════════════════════════
print("Building Bescheinigungen sheet...")
ws_cert = build_cert_sheet(wb)


# ══════════════════════════════════════════════════════════════════════════════
# SHEET: BG-Liste
# ══════════════════════════════════════════════════════════════════════════════
print("Loading original BG-Liste...")
wb_bg_orig  = openpyxl.load_workbook(ORIG_BG)
BG_PAGE_ROWS = 43

ws_bg = wb.create_sheet(title="BG-Liste")
ws_bg.sheet_properties.tabColor = "FF6F00"

copy_sheet(wb_bg_orig["Blatt1"], ws_bg, row_offset=0)
fix_formulas(ws_bg)
ws_bg["D7"].number_format = "@"

ws_bg.row_breaks.append(Break(id=BG_PAGE_ROWS, man=True))

copy_sheet(wb_bg_orig["Blatt2"], ws_bg, row_offset=BG_PAGE_ROWS)
for row in ws_bg.iter_rows(min_row=BG_PAGE_ROWS + 1):
    for cell in row:
        if isinstance(cell.value, str) and cell.value.startswith("="):
            cell.value = translate_formula(cell.value)
ws_bg.cell(row=BG_PAGE_ROWS + 7, column=4).number_format = "@"

ws_bg.print_area = f"A1:E{BG_PAGE_ROWS * 2}"
ws_bg.page_setup.paperSize   = ws_bg.PAPERSIZE_A4
ws_bg.page_setup.orientation = "portrait"
ws_bg.page_setup.fitToWidth  = 1
ws_bg.page_setup.fitToHeight = 0
ws_bg.sheet_properties.pageSetUpPr = \
    openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)

print(f"  BG-Liste: {len(ws_bg._images)} images (2 pages combined)")
wb_bg_orig.close()


# ══════════════════════════════════════════════════════════════════════════════
# SHEET: Lehrgangsdoku
# ══════════════════════════════════════════════════════════════════════════════
ws_ld = wb.create_sheet(title="Lehrgangsdoku")
ws_ld.sheet_properties.tabColor = "6A1B9A"

ws_ld.column_dimensions["A"].width = 3
ws_ld.column_dimensions["B"].width = 22
ws_ld.column_dimensions["C"].width = 22
ws_ld.column_dimensions["D"].width = 4
ws_ld.column_dimensions["E"].width = 18
ws_ld.column_dimensions["F"].width = 22
ws_ld.column_dimensions["G"].width = 4

ws_ld.merge_cells("A1:G1")
ws_ld.row_dimensions[1].height = 30
set_cell(ws_ld, 1, 1, "Lehrgangsdokumentation", font=LD_TITLE, alignment=center)

ws_ld.merge_cells("A2:G2")
ws_ld.row_dimensions[2].height = 18
set_cell(ws_ld, 2, 1, "gem. Abschnitt 2.4.6 DGUV Grundsatz 304-001",
    font=LD_SUBTITLE, alignment=center)

# Section 1: Ausbildungsstelle
ws_ld.row_dimensions[4].height = 22
ws_ld.merge_cells("B4:F4")
set_cell(ws_ld, 4, 2, "Angaben zur Ausbildungsstelle",
         font=LD_SECTION, alignment=left_center)

ws_ld.row_dimensions[5].height = 14
set_cell(ws_ld, 5, 2, "Bezeichnung", font=LD_LABEL_SM, alignment=left_center)
ws_ld.row_dimensions[6].height = 20
set_cell(ws_ld, 6, 2, COMPANY,      font=LD_VALUE,    alignment=left_center)
set_cell(ws_ld, 6, 5, "QSEH-Kennziffer:", font=LD_LABEL, alignment=left_center)
set_cell(ws_ld, 6, 6, "=Teilnehmer!F5",   font=LD_VALUE, alignment=left_center)

ws_ld.row_dimensions[8].height = 14
set_cell(ws_ld, 8, 2, "Straße / Hausnr.", font=LD_LABEL_SM, alignment=left_center)
ws_ld.row_dimensions[9].height = 20
set_cell(ws_ld, 9, 2, COMPANY_STREET,     font=LD_VALUE,    alignment=left_center)

ws_ld.row_dimensions[11].height = 14
set_cell(ws_ld, 11, 2, "PLZ / Ort",   font=LD_LABEL_SM, alignment=left_center)
ws_ld.row_dimensions[12].height = 20
set_cell(ws_ld, 12, 2, COMPANY_PLZ,   font=LD_VALUE,    alignment=left_center)

# Section 2: Seminar
ws_ld.row_dimensions[14].height = 22
ws_ld.merge_cells("B14:F14")
set_cell(ws_ld, 14, 2, "Angaben zum Seminar",
         font=LD_SECTION, alignment=left_center)

ws_ld.row_dimensions[15].height = 20
ws_ld.merge_cells("B15:F15")
ws_ld.cell(row=15, column=2).value = (
    '=IF(Teilnehmer!F4="EH-Ausbildung","X","  ")&" EH-Ausbildung       "'
    '&IF(Teilnehmer!F4="EH-Fortbildung","X","  ")&" EH-Fortbildung       "'
    '&IF(Teilnehmer!F4="EH-Schulung","X","  ")&" EH-Schulung"'
)
ws_ld.cell(row=15, column=2).font      = LD_LABEL
ws_ld.cell(row=15, column=2).alignment = left_wrap

ws_ld.row_dimensions[17].height = 20
set_cell(ws_ld, 17, 2, "Registriernummer:",  font=LD_LABEL, alignment=left_center)
set_cell(ws_ld, 17, 3, "=Teilnehmer!C5",     font=LD_VALUE, alignment=left_center)
set_cell(ws_ld, 17, 5, "Lehrgangsort:",       font=LD_LABEL, alignment=left_center)
set_cell(ws_ld, 17, 6, f"{COMPANY}, {COMPANY_WERK}", font=LD_VALUE, alignment=left_center)

ws_ld.row_dimensions[18].height = 14
set_cell(ws_ld, 18, 2, "(aus dem QSEH-Portal)", font=LD_LABEL_SM, alignment=left_center)
set_cell(ws_ld, 18, 6, LOCATION_DETAIL,         font=LD_LABEL_SM, alignment=left_center)

ws_ld.row_dimensions[20].height = 20
set_cell(ws_ld, 20, 2, "Lehrgangsdatum:", font=LD_LABEL, alignment=left_center)
date_ld = set_cell(ws_ld, 20, 3, "=Teilnehmer!C3", font=LD_VALUE, alignment=left_center)
date_ld.number_format = "DD.MM.YYYY"
set_cell(ws_ld, 20, 5, "Uhrzeit:", font=LD_LABEL, alignment=left_center)
ws_ld.cell(row=20, column=6).value = (
    '="von "&TEXT(Teilnehmer!F3,"HH:MM")&" Uhr bis "&TEXT(Teilnehmer!H3,"HH:MM")&" Uhr"'
)
ws_ld.cell(row=20, column=6).font      = LD_VALUE
ws_ld.cell(row=20, column=6).alignment = left_center

ws_ld.row_dimensions[22].height = 20
set_cell(ws_ld, 22, 2, "Name der Lehrkraft:",   font=LD_LABEL, alignment=left_center)
set_cell(ws_ld, 22, 3, "=Teilnehmer!C4",        font=LD_VALUE, alignment=left_center)

ws_ld.row_dimensions[24].height = 20
set_cell(ws_ld, 24, 2, "Verantwortlicher Arzt:", font=LD_LABEL, alignment=left_center)
set_cell(ws_ld, 24, 3, "=Teilnehmer!C7",         font=LD_VALUE, alignment=left_center)

ws_ld.row_dimensions[26].height = 20
set_cell(ws_ld, 26, 2, "Masken-Charge:", font=LD_LABEL, alignment=left_center)
set_cell(ws_ld, 26, 3, "=Teilnehmer!C8", font=LD_VALUE, alignment=left_center,
         border=bottom_line)

# Section 3: Teilnehmerzahl
ws_ld.row_dimensions[28].height = 22
ws_ld.merge_cells("B28:F28")
set_cell(ws_ld, 28, 2, "Anzahl der Teilnehmenden",
         font=LD_SECTION, alignment=left_center)

ws_ld.row_dimensions[30].height = 20
ws_ld.merge_cells("B30:D30")
set_cell(ws_ld, 30, 2, "Gesamtanzahl der Teilnehmenden:", font=LD_LABEL,
         alignment=left_center)
set_cell(ws_ld, 30, 5, "=Teilnehmer!C32", font=LD_VALUE, alignment=center,
         border=bottom_line)

ws_ld.row_dimensions[31].height = 14
set_cell(ws_ld, 31, 2,
    "(entspricht der Anzahl der Teilnehmerdatenblätter)",
    font=LD_LABEL_SM, alignment=left_center)

ws_ld.row_dimensions[33].height = 20
ws_ld.merge_cells("B33:D33")
set_cell(ws_ld, 33, 2, "davon betriebliche Ersthelfende:", font=LD_LABEL,
         alignment=left_center)
set_cell(ws_ld, 33, 5, "=Teilnehmer!E32", font=LD_VALUE, alignment=center,
         border=bottom_line)

# Section 4: Anlagen
ws_ld.row_dimensions[35].height = 22
ws_ld.merge_cells("B35:F35")
set_cell(ws_ld, 35, 2, "Anlagen", font=LD_SECTION, alignment=left_center)

ws_ld.merge_cells("B36:F38")
for rr in (36, 37, 38):
    ws_ld.row_dimensions[rr].height = 16
set_cell(ws_ld, 36, 2,
    "Alle Teilnehmenden sind mit Namen, Vornamen, Geburtsdatum und Unterschrift "
    "zu erfassen. Für UVT-Teilnehmende sind zusätzlich der Name und die Anschrift "
    "des Arbeitgebers sowie der kostentragende UVT zu ergänzen.",
    font=LD_LABEL, alignment=left_wrap)

ws_ld.merge_cells("B40:F40")
ws_ld.row_dimensions[40].height = 16
set_cell(ws_ld, 40, 2,
    "Die Dokumentation ist fünf Jahre aufzubewahren und auf Anforderung dem "
    "Unfallversicherungsträger vorzulegen.",
    font=LD_LABEL, alignment=left_wrap)

# Signature block
ws_ld.row_dimensions[42].height = 18
set_cell(ws_ld, 42, 2, "Für die Richtigkeit der Angaben:",
         font=LD_LABEL, alignment=left_center)

ws_ld.row_dimensions[47].height = 18
ws_ld.merge_cells("B47:C47")
ws_ld.cell(row=47, column=2).value     = '="Rastatt, "&TEXT(Teilnehmer!C3,"DD.MM.YYYY")'
ws_ld.cell(row=47, column=2).font      = LD_VALUE_UL
ws_ld.cell(row=47, column=2).alignment = left_center
ws_ld.merge_cells("E47:F47")
ws_ld.cell(row=47, column=5).value     = '="Rastatt, "&TEXT(Teilnehmer!C3,"DD.MM.YYYY")'
ws_ld.cell(row=47, column=5).font      = LD_VALUE_UL
ws_ld.cell(row=47, column=5).alignment = left_center

ws_ld.row_dimensions[48].height = 14
ws_ld.merge_cells("B48:C48")
set_cell(ws_ld, 48, 2, "Ort, Datum     Unterschrift Lehrgangsleitung",
    font=LD_LABEL_SM, alignment=left_center)
ws_ld.merge_cells("E48:F48")
set_cell(ws_ld, 48, 5, "Ort, Datum     Unterschrift Ausbildungsstelle",
    font=LD_LABEL_SM, alignment=left_center)

ws_ld.row_dimensions[50].height = 14
set_cell(ws_ld, 50, 2, "Stand: 2026", font=LD_FOOT, alignment=left_center)

ws_ld.page_setup.orientation = "portrait"
ws_ld.page_setup.paperSize   = ws_ld.PAPERSIZE_A4
ws_ld.page_setup.fitToWidth  = 1
ws_ld.page_setup.fitToHeight = 1
ws_ld.sheet_properties.pageSetUpPr = \
    openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
ws_ld.page_margins = PageMargins(left=0.5, right=0.5, top=0.5, bottom=0.5)


# ══════════════════════════════════════════════════════════════════════════════
# Reorder sheets
# ══════════════════════════════════════════════════════════════════════════════
desired_order = ["Teilnehmer", "Bescheinigungen", "BG-Liste",
                 "Lehrgangsdoku", "Hilfslisten"]
for i, name in enumerate(desired_order):
    current_idx = wb.sheetnames.index(name)
    wb.move_sheet(name, offset=i - current_idx)


# ══════════════════════════════════════════════════════════════════════════════
# Save
# ══════════════════════════════════════════════════════════════════════════════
output_path = f"{BASE_DIR}/EH_Kurs_2026.xlsx"
wb.save(output_path)
print(f"\nSaved: {output_path}")
print(f"Sheets: {wb.sheetnames}")


# ══════════════════════════════════════════════════════════════════════════════
# Post-process: fix BG-Liste drawings (no more cert post-processing needed)
# ══════════════════════════════════════════════════════════════════════════════
def post_process_xlsx(xlsx_path, orig_bg_path):
    """Fix generated xlsx: rebuild BG drawing, remove external links."""
    temp_path = xlsx_path + ".tmp"

    with zipfile.ZipFile(orig_bg_path, 'r') as z_bg:
        bg_drawing1      = z_bg.read('xl/drawings/drawing1.xml').decode('utf-8')
        bg_drawing2      = z_bg.read('xl/drawings/drawing2.xml').decode('utf-8')
        bg_drawing1_rels = z_bg.read('xl/drawings/_rels/drawing1.xml.rels').decode('utf-8')
        bg_drawing2_rels = z_bg.read('xl/drawings/_rels/drawing2.xml.rels').decode('utf-8')
        orig_bg_images   = {}
        bg_img_rename    = {}
        for f in z_bg.namelist():
            if f.startswith('xl/media/'):
                old_name = f.split('/')[-1]
                new_name = 'bg_' + old_name
                bg_img_rename[old_name] = new_name
                orig_bg_images[new_name] = z_bg.read(f)
        for old, new in bg_img_rename.items():
            bg_drawing1_rels = bg_drawing1_rels.replace(old, new)
            bg_drawing2_rels = bg_drawing2_rels.replace(old, new)

    BG_ROW_OFFSET = 43

    bg1_anchors     = re.findall(r'<xdr:twoCellAnchor[^>]*>.*?</xdr:twoCellAnchor>',
                                 bg_drawing1, re.DOTALL)
    bg1_one_anchors = re.findall(r'<xdr:oneCellAnchor[^>]*>.*?</xdr:oneCellAnchor>',
                                 bg_drawing1, re.DOTALL)
    bg2_anchors     = re.findall(r'<xdr:twoCellAnchor[^>]*>.*?</xdr:twoCellAnchor>',
                                 bg_drawing2, re.DOTALL)
    bg2_one_anchors = re.findall(r'<xdr:oneCellAnchor[^>]*>.*?</xdr:oneCellAnchor>',
                                 bg_drawing2, re.DOTALL)

    def offset_rows(anchor_xml, offset):
        def add_offset(m):
            return f'<xdr:row>{int(m.group(1)) + offset}</xdr:row>'
        return re.sub(r'<xdr:row>(\d+)</xdr:row>', add_offset, anchor_xml)

    bg1_rid_map = {}
    for m in re.finditer(r'Id="([^"]+)"[^>]*Target="([^"]+)"', bg_drawing1_rels):
        bg1_rid_map[m.group(1)] = m.group(2).split('/')[-1]
    bg2_rid_map = {}
    for m in re.finditer(r'Id="([^"]+)"[^>]*Target="([^"]+)"', bg_drawing2_rels):
        bg2_rid_map[m.group(1)] = m.group(2).split('/')[-1]

    combined_rels = dict(bg1_rid_map)
    bg2_remap     = {}
    next_rid      = 100
    for old_rid, img_name in bg2_rid_map.items():
        new_rid               = f'rId{next_rid}'
        bg2_remap[old_rid]    = new_rid
        combined_rels[new_rid] = img_name
        next_rid              += 1

    def remap_rids(anchor_xml, remap):
        for old, new in remap.items():
            anchor_xml = anchor_xml.replace(f'r:embed="{old}"', f'r:embed="{new}"')
        return anchor_xml

    ns_xdr = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
    ns_a   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    ns_r   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    ns_rel = 'http://schemas.openxmlformats.org/package/2006/relationships'
    ns_img = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
    ns_drw = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'

    combined_bg_drawing = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<xdr:wsDr xmlns:xdr="{ns_xdr}" xmlns:a="{ns_a}" xmlns:r="{ns_r}">'
    )
    for a in bg1_anchors + bg1_one_anchors:
        combined_bg_drawing += a
    for a in bg2_anchors + bg2_one_anchors:
        combined_bg_drawing += remap_rids(offset_rows(a, BG_ROW_OFFSET), bg2_remap)
    combined_bg_drawing += '</xdr:wsDr>'

    combined_bg_rels = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<Relationships xmlns="{ns_rel}">'
    )
    for rid, img_name in combined_rels.items():
        combined_bg_rels += (
            f'<Relationship Id="{rid}" Type="{ns_img}" '
            f'Target="../media/{img_name}"/>')
    combined_bg_rels += '</Relationships>'

    with zipfile.ZipFile(xlsx_path, 'r') as zin:
        wb_xml   = zin.read('xl/workbook.xml').decode('utf-8')
        rels_xml = zin.read('xl/_rels/workbook.xml.rels').decode('utf-8')

        sheets = re.findall(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb_xml)
        rels   = {}
        for rel_tag in re.findall(r'<Relationship[^>]+/>', rels_xml):
            id_m  = re.search(r'Id="([^"]+)"',     rel_tag)
            tgt_m = re.search(r'Target="([^"]+)"', rel_tag)
            if id_m and tgt_m:
                rels[id_m.group(1)] = tgt_m.group(1)

        bg_file = None
        for name, rid in sheets:
            target = rels.get(rid, '')
            if not target:
                continue
            path = target.lstrip('/') if target.startswith('/') else 'xl/' + target
            if name == "BG-Liste":
                bg_file = path

        # Fix workbook: remove external links
        fixed_wb_rels = re.sub(
            r'<Relationship[^>]*externalLink[^>]*/>', '', rels_xml)
        fixed_wb_xml  = re.sub(
            r'<externalReferences>.*?</externalReferences>', '',
            wb_xml, flags=re.DOTALL)

        ct_xml    = zin.read('[Content_Types].xml').decode('utf-8')
        fixed_ct  = re.sub(r'<Override[^>]*externalLink[^>]*/>', '', ct_xml)

        # Ensure drawing content types present
        additions = ''
        if 'Extension="png"' not in fixed_ct:
            additions += '<Default Extension="png" ContentType="image/png"/>'
        if 'drawing1.xml' not in fixed_ct:
            additions += ('<Override PartName="/xl/drawings/drawing1.xml" '
                          'ContentType="application/vnd.openxmlformats-officedocument'
                          '.drawing+xml"/>')
        if additions:
            fixed_ct = fixed_ct.replace('</Types>', additions + '</Types>')

        skip_files = set()
        for f in zin.namelist():
            if 'externalLink' in f:
                skip_files.add(f)
        # Only skip the BG drawing – we rebuild it. Cert drawing (drawing1) stays.

        # We keep the cert drawings generated by openpyxl (for cross + QR)
        # but rebuild BG drawings

        def make_rels(drawing_target):
            return (
                f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<Relationships xmlns="{ns_rel}">'
                f'<Relationship Id="rId1" Type="{ns_drw}" Target="{drawing_target}"/>'
                f'</Relationships>'
            )

        bg_rels_path = (
            bg_file.replace('xl/worksheets/', 'xl/worksheets/_rels/') + '.rels'
            if bg_file else None
        )

        # Find what drawing number openpyxl assigned to BG sheet
        # Openpyxl typically puts the cert drawings first (many images),
        # then BG images. We need to find drawing1 for BG and leave cert drawings.
        # Since we now start from a fresh workbook (no old cert), openpyxl will
        # number drawings sequentially. Let's check existing drawing files.
        existing_drawings = [f for f in zin.namelist()
                             if re.match(r'xl/drawings/drawing\d+\.xml$', f)]
        # Find which drawing belongs to BG-Liste by checking sheet rels
        bg_drawing_file = None
        if bg_rels_path and bg_rels_path in zin.namelist():
            bg_rels_content = zin.read(bg_rels_path).decode('utf-8')
            m = re.search(r'Target="\.\./drawings/(drawing\d+\.xml)"', bg_rels_content)
            if m:
                bg_drawing_file     = 'xl/drawings/' + m.group(1)
                bg_drawing_rels_file = f'xl/drawings/_rels/{m.group(1)}.rels'

        with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                if item in skip_files:
                    continue
                elif item == 'xl/_rels/workbook.xml.rels':
                    zout.writestr(item, fixed_wb_rels.encode('utf-8'))
                elif item == 'xl/workbook.xml':
                    zout.writestr(item, fixed_wb_xml.encode('utf-8'))
                elif item == '[Content_Types].xml':
                    zout.writestr(item, fixed_ct.encode('utf-8'))
                elif bg_drawing_file and item == bg_drawing_file:
                    zout.writestr(item, combined_bg_drawing.encode('utf-8'))
                elif bg_drawing_file and item == bg_drawing_rels_file:
                    zout.writestr(item, combined_bg_rels.encode('utf-8'))
                elif item.startswith('xl/media/') and item.split('/')[-1].startswith('image'):
                    # Keep cert images (from openpyxl), skip old BG images
                    # BG images will be re-added with bg_ prefix
                    zout.writestr(item, zin.read(item))
                else:
                    zout.writestr(item, zin.read(item))

            # Write BG images with bg_ prefix
            for img_name, img_data in orig_bg_images.items():
                zout.writestr(f'xl/media/{img_name}', img_data)

    shutil.move(temp_path, xlsx_path)
    print("  External links removed")
    bg_count = len(bg1_anchors)+len(bg1_one_anchors)+len(bg2_anchors)+len(bg2_one_anchors)
    print(f"  BG drawing rebuilt ({bg_count} anchors)")


print("\nPost-processing xlsx...")
post_process_xlsx(output_path, ORIG_BG)
print("Done!")
