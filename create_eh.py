#!/usr/bin/env python3
"""
Creates EH_Kurs_2026.xlsx – Consolidated First Aid Course Management.
Replaces 4 separate files:
  1. Namensliste_EH_Kurs.xlsx  (participant input)
  2. EH_Bescheinigung_Teilnehmer.xlsx  (certificates)
  3. EH_BG_Liste.xlsx  (BG participant list)
  4. Lehrgangsdoku.docx  (course protocol)

APPROACH: Load original cert file as base workbook to preserve 100% of cert
formatting (images, columns, page setup, printer settings). Add other sheets.
Post-process the saved xlsx to fix any attributes openpyxl changes.
"""
import re
import os
import shutil
import zipfile
import openpyxl
from copy import copy
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.pagebreak import Break
from openpyxl.workbook.defined_name import DefinedName
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

COMPANY = "Mercedes-Benz AG"
COMPANY_STREET = "Mercedesstr. 1"
COMPANY_PLZ = "76437 Rastatt"
COMPANY_WERK = "PKW-Werk Rastatt"
QSEH_KENNZIFFER = "7.0103"
LOCATION_DETAIL = "Gebäude 39, Gesundheitszentrum 1.OG"
ARZT = "Dr. Schmidt, Sabine"
ARZT_TEL = "07222/91-22111"
ARZT_EMAIL = "sabine.m.schmidt@mercedes-benz.com"

ORIG_BG = "/home/user/Ambulanzzeug/Neu EH/210923_EH_BG_Liste.xlsx"
ORIG_CERT = "/home/user/Ambulanzzeug/Neu EH/190731_EH_Bescheinigung_Teilnehmer.xlsx"

# ══════════════════════════════════════════════════════════════════════════════
# Styles
# ══════════════════════════════════════════════════════════════════════════════
DARK_BLUE = "0F3460"
HEADER_BG = PatternFill(start_color=DARK_BLUE, end_color=DARK_BLUE, fill_type="solid")
LIGHT_BLUE_BG = PatternFill(start_color="E8EAF6", end_color="E8EAF6", fill_type="solid")
LIGHT_GRAY = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
WHITE_BG = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
INPUT_BG = PatternFill(start_color="FFF9C4", end_color="FFF9C4", fill_type="solid")
GREEN_BG = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")

TITLE_FONT = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
LABEL_FONT = Font(name="Calibri", size=10, bold=True, color="333333")
VALUE_FONT = Font(name="Calibri", size=11, color="333333")
INPUT_FONT = Font(name="Calibri", size=11, color="0F3460")
SMALL_FONT = Font(name="Calibri", size=9, color="666666")

# Lehrgangsdoku fonts
LD_TITLE = Font(name="Arial", size=18, bold=True, color="000000")
LD_SUBTITLE = Font(name="Arial", size=10, color="000000")
LD_SECTION = Font(name="Arial", size=12, bold=True, color="000000")
LD_LABEL = Font(name="Arial", size=10, color="000000")
LD_LABEL_SM = Font(name="Arial", size=8, color="000000")
LD_VALUE = Font(name="Arial", size=12, color="000000")
LD_VALUE_UL = Font(name="Arial", size=12, color="000000", underline="single")
LD_FOOT = Font(name="Arial", size=8, color="666666")

thin_border = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)
bottom_line = Border(bottom=Side(style="thin", color="000000"))

center = Alignment(horizontal="center", vertical="center")
left_center = Alignment(horizontal="left", vertical="center")
right_center = Alignment(horizontal="right", vertical="center")
left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)


def set_cell(ws, row, col, value, font=None, fill=None, alignment=None,
             border=None, number_format=None):
    cell = ws.cell(row=row, column=col, value=value)
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    if alignment:
        cell.alignment = alignment
    if border:
        cell.border = border
    if number_format:
        cell.number_format = number_format
    return cell


# ══════════════════════════════════════════════════════════════════════════════
# Formula translation: [1]Sheet1!ref → Teilnehmer!ref
# ══════════════════════════════════════════════════════════════════════════════
def translate_formula(formula):
    """Replace [1]Sheet1! references with Teilnehmer! references."""
    if not isinstance(formula, str) or "[1]Sheet1!" not in formula:
        return formula

    # Metadata: absolute references to row 2 columns D-H
    formula = formula.replace("[1]Sheet1!$D$2", "Teilnehmer!C5")
    formula = formula.replace("[1]Sheet1!$E$2", "Teilnehmer!C3")
    formula = formula.replace("[1]Sheet1!$F$2", "Teilnehmer!C4")
    formula = formula.replace("[1]Sheet1!$G$2", "Teilnehmer!F3")
    formula = formula.replace("[1]Sheet1!$H$2", "Teilnehmer!H3")

    # Participant references: A/B/C columns, rows 2-21 → B/C/D, rows 11-30
    def replace_participant(match):
        col = match.group(1)
        row = int(match.group(2))
        new_row = row + 9
        col_map = {"A": "B", "B": "C", "C": "D"}
        return f"Teilnehmer!{col_map[col]}{new_row}"

    formula = re.sub(
        r'\[1\]Sheet1!\$?([ABC])\$?(\d+)', replace_participant, formula
    )
    return formula


def fix_formulas(ws):
    """Fix all formulas in a worksheet."""
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.value = translate_formula(cell.value)


def copy_sheet(ws_src, ws_dst, row_offset=0):
    """Copy all cell values, styles, merged cells, column widths, row heights,
    images, page setup, and other properties from ws_src to ws_dst.
    row_offset: shift all rows by this amount (for appending page 2 below page 1)."""
    from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker

    # Column dimensions (only if first page / no offset)
    if row_offset == 0:
        for col_letter, dim in ws_src.column_dimensions.items():
            ws_dst.column_dimensions[col_letter].width = dim.width
            ws_dst.column_dimensions[col_letter].hidden = dim.hidden

    # Row dimensions
    for row_idx, dim in ws_src.row_dimensions.items():
        if dim.height is not None:
            ws_dst.row_dimensions[row_idx + row_offset].height = dim.height

    # Cell values and styles
    for row in ws_src.iter_rows():
        for cell in row:
            new_cell = ws_dst.cell(row=cell.row + row_offset, column=cell.column)
            new_cell.value = cell.value
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.fill = copy(cell.fill)
                new_cell.alignment = copy(cell.alignment)
                new_cell.border = copy(cell.border)
                new_cell.number_format = cell.number_format

    # Merged cells
    for mcr in ws_src.merged_cells.ranges:
        if row_offset == 0:
            ws_dst.merge_cells(str(mcr))
        else:
            ws_dst.merge_cells(
                start_row=mcr.min_row + row_offset,
                start_column=mcr.min_col,
                end_row=mcr.max_row + row_offset,
                end_column=mcr.max_col,
            )

    # Images - offset TwoCellAnchor rows if needed
    for img in ws_src._images:
        if row_offset == 0:
            ws_dst._images.append(img)
        else:
            new_img = copy(img)
            anchor = img.anchor
            if isinstance(anchor, TwoCellAnchor):
                new_anchor = TwoCellAnchor()
                new_anchor._from = AnchorMarker(
                    col=anchor._from.col, colOff=anchor._from.colOff,
                    row=anchor._from.row + row_offset,
                    rowOff=anchor._from.rowOff,
                )
                new_anchor.to = AnchorMarker(
                    col=anchor.to.col, colOff=anchor.to.colOff,
                    row=anchor.to.row + row_offset,
                    rowOff=anchor.to.rowOff,
                )
                new_img.anchor = new_anchor
            ws_dst._images.append(new_img)

    # Page setup (only on first page)
    if row_offset == 0:
        ws_dst.page_setup.paperSize = ws_src.page_setup.paperSize
        ws_dst.page_setup.orientation = ws_src.page_setup.orientation
        ws_dst.page_setup.fitToWidth = ws_src.page_setup.fitToWidth
        ws_dst.page_setup.fitToHeight = ws_src.page_setup.fitToHeight

        src_margins = ws_src.page_margins
        ws_dst.page_margins = PageMargins(
            left=src_margins.left, right=src_margins.right,
            top=src_margins.top, bottom=src_margins.bottom,
            header=src_margins.header, footer=src_margins.footer
        )

    # Row breaks
    if ws_src.row_breaks:
        for brk in ws_src.row_breaks.brk:
            ws_dst.row_breaks.append(Break(id=brk.id + row_offset, man=brk.man))


# ══════════════════════════════════════════════════════════════════════════════
# Load original cert file as BASE workbook (preserves all formatting)
# ══════════════════════════════════════════════════════════════════════════════
print("Loading original Bescheinigungen as base workbook...")
wb = openpyxl.load_workbook(ORIG_CERT)

# Rename Sheet1 → Bescheinigungen
ws_cert = wb["Sheet1"]
ws_cert.title = "Bescheinigungen"
ws_cert.sheet_properties.tabColor = "1B5E20"

# Fix formulas to reference internal Teilnehmer sheet
fix_formulas(ws_cert)

# Page setup: A4 landscape, fit to page width, print all 10 certificate pages
ws_cert.page_setup.paperSize = 9  # A4
ws_cert.page_setup.orientation = "landscape"
ws_cert.page_setup.fitToWidth = 1
ws_cert.page_setup.fitToHeight = 10  # 10 certificate pages (9 row breaks)
ws_cert.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
ws_cert.row_breaks.append(Break(id=350, man=True))  # missing break between pages 9/10

print(f"  Bescheinigungen: {len(ws_cert._images)} images preserved")


# ══════════════════════════════════════════════════════════════════════════════
# SHEET: Hilfslisten (created early for named range)
# ══════════════════════════════════════════════════════════════════════════════
ws_help = wb.create_sheet(title="Hilfslisten")
ws_help.sheet_properties.tabColor = "9E9E9E"

set_cell(ws_help, 1, 1, "Lehrkräfte",
    font=HEADER_FONT, fill=HEADER_BG, alignment=center, border=thin_border)
ws_help.column_dimensions["A"].width = 25

for i, name in enumerate(INSTRUCTORS):
    set_cell(ws_help, 2 + i, 1, name, font=VALUE_FONT, alignment=left_center,
             border=thin_border)

dn = DefinedName("Lehrkraefte", attr_text="Hilfslisten!$A$2:$A$20")
wb.defined_names.add(dn)


# ══════════════════════════════════════════════════════════════════════════════
# SHEET: Teilnehmer (Input Sheet)
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

set_cell(ws_tn, 3, 2, "Datum:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 3, 3, None, font=INPUT_FONT, fill=INPUT_BG,
         alignment=left_center, border=thin_border, number_format="DD.MM.YYYY")
set_cell(ws_tn, 3, 5, "Beginn:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 3, 6, None, font=INPUT_FONT, fill=INPUT_BG,
         alignment=left_center, border=thin_border, number_format="HH:MM")
set_cell(ws_tn, 3, 7, "Ende:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 3, 8, None, font=INPUT_FONT, fill=INPUT_BG,
         alignment=left_center, border=thin_border, number_format="HH:MM")

set_cell(ws_tn, 4, 2, "Lehrkraft:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 4, 3, None, font=INPUT_FONT, fill=INPUT_BG,
         alignment=left_center, border=thin_border)
set_cell(ws_tn, 4, 5, "Kursart:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 4, 6, None, font=INPUT_FONT, fill=INPUT_BG,
         alignment=left_center, border=thin_border)

set_cell(ws_tn, 5, 2, "Registriernr.:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 5, 3, None, font=INPUT_FONT, fill=INPUT_BG,
         alignment=left_center, border=thin_border)
set_cell(ws_tn, 5, 5, "QSEH-Kennziffer:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 5, 6, QSEH_KENNZIFFER, font=INPUT_FONT, fill=LIGHT_BLUE_BG,
         alignment=left_center, border=thin_border)

set_cell(ws_tn, 6, 2, "Lehrgangsort:", font=LABEL_FONT, alignment=right_center)
ws_tn.merge_cells("C6:E6")
set_cell(ws_tn, 6, 3, f"{COMPANY}, {COMPANY_WERK}",
         font=INPUT_FONT, fill=INPUT_BG, alignment=left_center, border=thin_border)

set_cell(ws_tn, 7, 2, "Verantw. Arzt:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 7, 3, ARZT, font=INPUT_FONT, fill=LIGHT_BLUE_BG,
         alignment=left_center, border=thin_border)
set_cell(ws_tn, 7, 5, "Tel.:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 7, 6, ARZT_TEL, font=INPUT_FONT, fill=LIGHT_BLUE_BG,
         alignment=left_center, border=thin_border)

set_cell(ws_tn, 8, 2, "Masken-Charge:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 8, 3, None, font=INPUT_FONT, fill=INPUT_BG,
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
ws_tn.page_setup.paperSize = ws_tn.PAPERSIZE_A4
ws_tn.page_setup.fitToWidth = 1
ws_tn.page_setup.fitToHeight = 1
ws_tn.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
ws_tn.page_margins = PageMargins(left=0.5, right=0.5, top=0.5, bottom=0.5)
ws_tn.freeze_panes = "A11"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET: BG-Liste (both pages on one sheet, preserves all 27 images per page)
# ══════════════════════════════════════════════════════════════════════════════
print("Loading original BG-Liste...")
wb_bg_orig = openpyxl.load_workbook(ORIG_BG)

BG_PAGE_ROWS = 43  # rows per page in original

ws_bg = wb.create_sheet(title="BG-Liste")
ws_bg.sheet_properties.tabColor = "FF6F00"

# Page 1 (participants 1-10) from Blatt1
copy_sheet(wb_bg_orig["Blatt1"], ws_bg, row_offset=0)
fix_formulas(ws_bg)
ws_bg["D7"].number_format = "@"

# Page break between pages
ws_bg.row_breaks.append(Break(id=BG_PAGE_ROWS, man=True))

# Page 2 (participants 11-20) from Blatt2, offset by BG_PAGE_ROWS
copy_sheet(wb_bg_orig["Blatt2"], ws_bg, row_offset=BG_PAGE_ROWS)
# Fix formulas for the newly added page 2 cells
for row in ws_bg.iter_rows(min_row=BG_PAGE_ROWS + 1):
    for cell in row:
        if isinstance(cell.value, str) and cell.value.startswith("="):
            cell.value = translate_formula(cell.value)
ws_bg.cell(row=BG_PAGE_ROWS + 7, column=4).number_format = "@"

# Fix print area to cover both pages
ws_bg.print_area = f"A1:E{BG_PAGE_ROWS * 2}"

# Page setup: A4 portrait, fit to page width (original used scale=56%)
ws_bg.page_setup.paperSize = ws_bg.PAPERSIZE_A4
ws_bg.page_setup.orientation = "portrait"
ws_bg.page_setup.fitToWidth = 1
ws_bg.page_setup.fitToHeight = 0
ws_bg.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)

print(f"  BG-Liste: {len(ws_bg._images)} images (2 pages combined)")

wb_bg_orig.close()


# ══════════════════════════════════════════════════════════════════════════════
# SHEET: Lehrgangsdoku (Course Protocol)
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

# Title
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
set_cell(ws_ld, 4, 2, "Angaben zur Ausbildungsstelle", font=LD_SECTION,
         alignment=left_center)

ws_ld.row_dimensions[5].height = 14
set_cell(ws_ld, 5, 2, "Bezeichnung", font=LD_LABEL_SM, alignment=left_center)
ws_ld.row_dimensions[6].height = 20
set_cell(ws_ld, 6, 2, COMPANY, font=LD_VALUE, alignment=left_center)
set_cell(ws_ld, 6, 5, "QSEH-Kennziffer:", font=LD_LABEL, alignment=left_center)
set_cell(ws_ld, 6, 6, "=Teilnehmer!F5", font=LD_VALUE, alignment=left_center)

ws_ld.row_dimensions[8].height = 14
set_cell(ws_ld, 8, 2, "Straße / Hausnr.", font=LD_LABEL_SM, alignment=left_center)
ws_ld.row_dimensions[9].height = 20
set_cell(ws_ld, 9, 2, COMPANY_STREET, font=LD_VALUE, alignment=left_center)

ws_ld.row_dimensions[11].height = 14
set_cell(ws_ld, 11, 2, "PLZ / Ort", font=LD_LABEL_SM, alignment=left_center)
ws_ld.row_dimensions[12].height = 20
set_cell(ws_ld, 12, 2, COMPANY_PLZ, font=LD_VALUE, alignment=left_center)

# Section 2: Seminar
ws_ld.row_dimensions[14].height = 22
ws_ld.merge_cells("B14:F14")
set_cell(ws_ld, 14, 2, "Angaben zum Seminar", font=LD_SECTION,
         alignment=left_center)

ws_ld.row_dimensions[15].height = 20
ws_ld.merge_cells("B15:F15")
set_cell(ws_ld, 15, 2, None, font=LD_LABEL, alignment=left_wrap)
ws_ld.cell(row=15, column=2).value = (
    '=IF(Teilnehmer!F4="EH-Ausbildung","X","  ")&" EH-Ausbildung       "'
    '&IF(Teilnehmer!F4="EH-Fortbildung","X","  ")&" EH-Fortbildung       "'
    '&IF(Teilnehmer!F4="EH-Schulung","X","  ")&" EH-Schulung"'
)

ws_ld.row_dimensions[17].height = 20
set_cell(ws_ld, 17, 2, "Registriernummer:", font=LD_LABEL, alignment=left_center)
set_cell(ws_ld, 17, 3, "=Teilnehmer!C5", font=LD_VALUE, alignment=left_center)
set_cell(ws_ld, 17, 5, "Lehrgangsort:", font=LD_LABEL, alignment=left_center)
set_cell(ws_ld, 17, 6, f"{COMPANY}, {COMPANY_WERK}", font=LD_VALUE,
         alignment=left_center)

ws_ld.row_dimensions[18].height = 14
set_cell(ws_ld, 18, 2, "(aus dem QSEH-Portal)", font=LD_LABEL_SM,
         alignment=left_center)
set_cell(ws_ld, 18, 6, LOCATION_DETAIL, font=LD_LABEL_SM, alignment=left_center)

ws_ld.row_dimensions[20].height = 20
set_cell(ws_ld, 20, 2, "Lehrgangsdatum:", font=LD_LABEL, alignment=left_center)
date_ld = set_cell(ws_ld, 20, 3, "=Teilnehmer!C3", font=LD_VALUE,
                   alignment=left_center)
date_ld.number_format = "DD.MM.YYYY"
set_cell(ws_ld, 20, 5, "Uhrzeit:", font=LD_LABEL, alignment=left_center)
ws_ld.cell(row=20, column=6).value = (
    '="von "&TEXT(Teilnehmer!F3,"HH:MM")&" Uhr bis "&TEXT(Teilnehmer!H3,"HH:MM")&" Uhr"'
)
ws_ld.cell(row=20, column=6).font = LD_VALUE
ws_ld.cell(row=20, column=6).alignment = left_center

ws_ld.row_dimensions[22].height = 20
set_cell(ws_ld, 22, 2, "Name der Lehrkraft:", font=LD_LABEL, alignment=left_center)
set_cell(ws_ld, 22, 3, "=Teilnehmer!C4", font=LD_VALUE, alignment=left_center)

ws_ld.row_dimensions[24].height = 20
set_cell(ws_ld, 24, 2, "Verantwortlicher Arzt:", font=LD_LABEL, alignment=left_center)
set_cell(ws_ld, 24, 3, "=Teilnehmer!C7", font=LD_VALUE, alignment=left_center)

ws_ld.row_dimensions[26].height = 20
set_cell(ws_ld, 26, 2, "Masken-Charge:", font=LD_LABEL, alignment=left_center)
set_cell(ws_ld, 26, 3, "=Teilnehmer!C8", font=LD_VALUE, alignment=left_center,
         border=bottom_line)

# Section 3: Teilnehmerzahl
ws_ld.row_dimensions[28].height = 22
ws_ld.merge_cells("B28:F28")
set_cell(ws_ld, 28, 2, "Anzahl der Teilnehmenden", font=LD_SECTION,
         alignment=left_center)

ws_ld.row_dimensions[30].height = 20
ws_ld.merge_cells("B30:D30")
set_cell(ws_ld, 30, 2, "Gesamtanzahl der Teilnehmenden:", font=LD_LABEL,
         alignment=left_center)
set_cell(ws_ld, 30, 5, "=Teilnehmer!C32", font=LD_VALUE, alignment=center,
         border=bottom_line)

ws_ld.row_dimensions[31].height = 14
set_cell(ws_ld, 31, 2, "(entspricht der Anzahl der Teilnehmerdatenblätter)",
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
    "des Arbeitgebers sowie der kostentragende UVT zu ergänzen. Dies kann durch "
    "einzelne Teilnehmerdatenblätter erfolgen.",
    font=LD_LABEL, alignment=left_wrap)

ws_ld.merge_cells("B40:F40")
ws_ld.row_dimensions[40].height = 16
set_cell(ws_ld, 40, 2,
    "Die Dokumentation ist fünf Jahre aufzubewahren und auf Anforderung dem "
    "Unfallversicherungsträger vorzulegen.",
    font=LD_LABEL, alignment=left_wrap)

# Signature block
ws_ld.row_dimensions[42].height = 18
set_cell(ws_ld, 42, 2, "Für die Richtigkeit der Angaben:", font=LD_LABEL,
         alignment=left_center)

ws_ld.row_dimensions[47].height = 18
ws_ld.merge_cells("B47:C47")
ws_ld.cell(row=47, column=2).value = '="Rastatt, "&TEXT(Teilnehmer!C3,"DD.MM.YYYY")'
ws_ld.cell(row=47, column=2).font = LD_VALUE_UL
ws_ld.cell(row=47, column=2).alignment = left_center
ws_ld.merge_cells("E47:F47")
ws_ld.cell(row=47, column=5).value = '="Rastatt, "&TEXT(Teilnehmer!C3,"DD.MM.YYYY")'
ws_ld.cell(row=47, column=5).font = LD_VALUE_UL
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

# Print setup
ws_ld.page_setup.orientation = "portrait"
ws_ld.page_setup.paperSize = ws_ld.PAPERSIZE_A4
ws_ld.page_setup.fitToWidth = 1
ws_ld.page_setup.fitToHeight = 1
ws_ld.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
ws_ld.page_margins = PageMargins(left=0.5, right=0.5, top=0.5, bottom=0.5)


# ══════════════════════════════════════════════════════════════════════════════
# Reorder sheets: Teilnehmer, Bescheinigungen, BG-Liste, Lehrgangsdoku, Hilfslisten
# ══════════════════════════════════════════════════════════════════════════════
desired_order = ["Teilnehmer", "Bescheinigungen", "BG-Liste",
                 "Lehrgangsdoku", "Hilfslisten"]
for i, name in enumerate(desired_order):
    current_idx = wb.sheetnames.index(name)
    wb.move_sheet(name, offset=i - current_idx)


# ══════════════════════════════════════════════════════════════════════════════
# Save
# ══════════════════════════════════════════════════════════════════════════════
output_path = "/home/user/Ambulanzzeug/EH_Kurs_2026.xlsx"
wb.save(output_path)
print(f"\nSaved: {output_path}")
print(f"Sheets: {wb.sheetnames}")


# ══════════════════════════════════════════════════════════════════════════════
# Post-process: restore cert sheet XML attributes that openpyxl may change
# ══════════════════════════════════════════════════════════════════════════════
def post_process_xlsx(xlsx_path, orig_cert_path, orig_bg_path):
    """Fix generated xlsx: cert sheet XML, drawings, and remove external links."""
    temp_path = xlsx_path + ".tmp"

    # ── Load original cert data ──────────────────────────────────────────
    with zipfile.ZipFile(orig_cert_path, 'r') as z_orig:
        orig_sheet_xml = z_orig.read('xl/worksheets/sheet1.xml').decode('utf-8')
        orig_cert_drawing = z_orig.read('xl/drawings/drawing1.xml')
        orig_cert_drawing_rels = z_orig.read('xl/drawings/_rels/drawing1.xml.rels')
        orig_cert_images = {}
        for f in z_orig.namelist():
            if f.startswith('xl/media/'):
                orig_cert_images[f] = z_orig.read(f)

    # Extract key sections from original cert sheet XML
    orig_cols = re.search(r'<cols>.*?</cols>', orig_sheet_xml, re.DOTALL).group(0)
    orig_fmt = re.search(r'<sheetFormatPr[^/]*/>', orig_sheet_xml).group(0)
    orig_setup = re.search(r'<pageSetup[^/]*/>', orig_sheet_xml).group(0)
    orig_margins = re.search(r'<pageMargins[^/]*/>', orig_sheet_xml).group(0)
    orig_setup_clean = re.sub(r'\s*r:id="[^"]*"', '', orig_setup)
    orig_fmt = re.sub(r'\s*x14ac:dyDescent="[^"]*"', '', orig_fmt)

    # ── Load original BG drawing data ────────────────────────────────────
    # Rename BG images to bg_imageN.png to avoid conflicts with cert images
    with zipfile.ZipFile(orig_bg_path, 'r') as z_bg:
        bg_drawing1 = z_bg.read('xl/drawings/drawing1.xml').decode('utf-8')
        bg_drawing2 = z_bg.read('xl/drawings/drawing2.xml').decode('utf-8')
        bg_drawing1_rels = z_bg.read('xl/drawings/_rels/drawing1.xml.rels').decode('utf-8')
        bg_drawing2_rels = z_bg.read('xl/drawings/_rels/drawing2.xml.rels').decode('utf-8')
        orig_bg_images = {}  # new_name -> data
        bg_img_rename = {}   # old_name -> new_name
        for f in z_bg.namelist():
            if f.startswith('xl/media/'):
                old_name = f.split('/')[-1]
                new_name = 'bg_' + old_name
                bg_img_rename[old_name] = new_name
                orig_bg_images[new_name] = z_bg.read(f)
        # Apply renames to rels
        for old, new in bg_img_rename.items():
            bg_drawing1_rels = bg_drawing1_rels.replace(old, new)
            bg_drawing2_rels = bg_drawing2_rels.replace(old, new)

    # ── Build combined BG drawing ────────────────────────────────────────
    BG_ROW_OFFSET = 43

    # Extract anchors from BG Blatt1 drawing (rows as-is)
    bg1_anchors = re.findall(
        r'<xdr:twoCellAnchor[^>]*>.*?</xdr:twoCellAnchor>', bg_drawing1, re.DOTALL)
    bg1_one_anchors = re.findall(
        r'<xdr:oneCellAnchor[^>]*>.*?</xdr:oneCellAnchor>', bg_drawing1, re.DOTALL)

    # Extract anchors from BG Blatt2 drawing (need row offset)
    bg2_anchors = re.findall(
        r'<xdr:twoCellAnchor[^>]*>.*?</xdr:twoCellAnchor>', bg_drawing2, re.DOTALL)
    bg2_one_anchors = re.findall(
        r'<xdr:oneCellAnchor[^>]*>.*?</xdr:oneCellAnchor>', bg_drawing2, re.DOTALL)

    def offset_rows(anchor_xml, offset):
        """Add offset to <xdr:row> values in anchor XML."""
        def add_offset(m):
            return f'<xdr:row>{int(m.group(1)) + offset}</xdr:row>'
        return re.sub(r'<xdr:row>(\d+)</xdr:row>', add_offset, anchor_xml)

    # Remap Blatt2 image rIds to avoid conflicts with Blatt1
    # Parse both rels to find image mappings
    bg1_rid_map = {}
    for m in re.finditer(r'Id="([^"]+)"[^>]*Target="([^"]+)"', bg_drawing1_rels):
        bg1_rid_map[m.group(1)] = m.group(2).split('/')[-1]
    bg2_rid_map = {}
    for m in re.finditer(r'Id="([^"]+)"[^>]*Target="([^"]+)"', bg_drawing2_rels):
        bg2_rid_map[m.group(1)] = m.group(2).split('/')[-1]

    # Combined rels: keep Blatt1 rIds, remap Blatt2 to rId100+
    combined_rels = dict(bg1_rid_map)  # rId -> image filename
    bg2_remap = {}  # old rId -> new rId
    next_rid = 100
    for old_rid, img_name in bg2_rid_map.items():
        new_rid = f'rId{next_rid}'
        bg2_remap[old_rid] = new_rid
        combined_rels[new_rid] = img_name
        next_rid += 1

    # Apply rId remapping to Blatt2 anchors
    def remap_rids(anchor_xml, remap):
        for old, new in remap.items():
            anchor_xml = anchor_xml.replace(f'r:embed="{old}"', f'r:embed="{new}"')
        return anchor_xml

    # Build combined drawing XML
    ns_xdr = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
    ns_a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    ns_r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    combined_bg_drawing = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<xdr:wsDr xmlns:xdr="{ns_xdr}" xmlns:a="{ns_a}" xmlns:r="{ns_r}">'
    )
    for a in bg1_anchors + bg1_one_anchors:
        combined_bg_drawing += a
    for a in bg2_anchors + bg2_one_anchors:
        combined_bg_drawing += remap_rids(offset_rows(a, BG_ROW_OFFSET), bg2_remap)
    combined_bg_drawing += '</xdr:wsDr>'

    # Build combined BG drawing rels
    ns_rel = 'http://schemas.openxmlformats.org/package/2006/relationships'
    ns_img_type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
    combined_bg_rels = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    combined_bg_rels += f'<Relationships xmlns="{ns_rel}">'
    for rid, img_name in combined_rels.items():
        combined_bg_rels += f'<Relationship Id="{rid}" Type="{ns_img_type}" Target="../media/{img_name}"/>'
    combined_bg_rels += '</Relationships>'

    # ── Process the generated xlsx ───────────────────────────────────────
    with zipfile.ZipFile(xlsx_path, 'r') as zin:
        wb_xml = zin.read('xl/workbook.xml').decode('utf-8')
        sheets = re.findall(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb_xml)
        rels_xml = zin.read('xl/_rels/workbook.xml.rels').decode('utf-8')
        rels = {}
        for rel_tag in re.findall(r'<Relationship[^>]+/>', rels_xml):
            id_m = re.search(r'Id="([^"]+)"', rel_tag)
            tgt_m = re.search(r'Target="([^"]+)"', rel_tag)
            if id_m and tgt_m:
                rels[id_m.group(1)] = tgt_m.group(1)

        # Find cert and BG sheet files
        cert_file = None
        bg_file = None
        for name, rid in sheets:
            target = rels.get(rid, '')
            if not target:
                continue
            path = target.lstrip('/') if target.startswith('/') else 'xl/' + target
            if name == "Bescheinigungen":
                cert_file = path
            elif name == "BG-Liste":
                bg_file = path

        if not cert_file:
            print("  WARNING: Could not find Bescheinigungen for post-processing")
            return

        # Fix cert sheet XML
        ns_r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        cert_xml = zin.read(cert_file).decode('utf-8')
        # Add r: namespace to root element (needed for drawing reference)
        cert_xml = cert_xml.replace(
            'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
            'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
            ' xmlns:r="%s"' % ns_r)
        cert_xml = re.sub(r'<cols>.*?</cols>', orig_cols, cert_xml, flags=re.DOTALL)
        cert_xml = re.sub(r'<sheetFormatPr[^/]*/>', orig_fmt, cert_xml)
        # A3 landscape, fit to 1 page wide, unlimited pages tall
        cert_xml = re.sub(
            r'<pageSetup[^/]*/>',
            '<pageSetup paperSize="9" fitToWidth="1" fitToHeight="10" orientation="landscape"/>',
            cert_xml)
        cert_xml = re.sub(r'<pageMargins[^/]*/>', orig_margins, cert_xml)
        # Enable fitToPage in sheet properties
        cert_xml = cert_xml.replace('<pageSetUpPr/>', '<pageSetUpPr fitToPage="1"/>')
        # Add drawing reference so Excel can find the cert images
        cert_xml = cert_xml.replace('</worksheet>',
            '<drawing r:id="rId1"/></worksheet>')

        # Add drawing reference to BG-Liste sheet
        if bg_file:
            bg_xml = zin.read(bg_file).decode('utf-8')
            bg_xml = bg_xml.replace(
                'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
                'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
                ' xmlns:r="%s"' % ns_r)
            bg_xml = bg_xml.replace('</worksheet>',
                '<drawing r:id="rId1"/></worksheet>')

        # 1. Remove external link references from workbook.xml.rels
        fixed_wb_rels = re.sub(
            r'<Relationship[^>]*externalLink[^>]*/>', '', rels_xml)

        # Remove externalReferences from workbook.xml
        fixed_wb_xml = re.sub(
            r'<externalReferences>.*?</externalReferences>', '', wb_xml, flags=re.DOTALL)

        # Expand print area to cover all 10 certificate pages (was only row 1-37)
        fixed_wb_xml = fixed_wb_xml.replace(
            "'Bescheinigungen'!$A$1:$AZ$37",
            "'Bescheinigungen'!$A$1:$AZ$388")

        # Remove external link from content types and add drawing/image types
        ct_xml = zin.read('[Content_Types].xml').decode('utf-8')
        fixed_ct = re.sub(r'<Override[^>]*externalLink[^>]*/>', '', ct_xml)
        ct_additions = ''
        if 'Extension="png"' not in fixed_ct:
            ct_additions += '<Default Extension="png" ContentType="image/png"/>'
        if 'drawing1.xml' not in fixed_ct:
            ct_additions += '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
        if 'drawing2.xml' not in fixed_ct:
            ct_additions += '<Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
        if ct_additions:
            fixed_ct = fixed_ct.replace('</Types>', ct_additions + '</Types>')

        # Collect files to skip (external links + openpyxl-generated drawings + images)
        skip_files = set()
        for f in zin.namelist():
            if 'externalLink' in f:
                skip_files.add(f)
        # We'll replace drawings with originals
        skip_files.add('xl/drawings/drawing1.xml')
        skip_files.add('xl/drawings/_rels/drawing1.xml.rels')
        skip_files.add('xl/drawings/drawing2.xml')
        skip_files.add('xl/drawings/_rels/drawing2.xml.rels')
        # Skip ALL openpyxl-generated images (1-94) - we provide originals
        for f in zin.namelist():
            if f.startswith('xl/media/image'):
                skip_files.add(f)

        # Sheet rels: link sheets → drawings
        ns_rel = 'http://schemas.openxmlformats.org/package/2006/relationships'
        ns_drawing_type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
        def make_sheet_rels(drawing_target):
            return (
                f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<Relationships xmlns="{ns_rel}">'
                f'<Relationship Id="rId1" Type="{ns_drawing_type}" Target="{drawing_target}"/>'
                f'</Relationships>'
            )

        # Determine rels paths from sheet file paths
        cert_rels_path = cert_file.replace('xl/worksheets/', 'xl/worksheets/_rels/') + '.rels'
        bg_rels_path = bg_file.replace('xl/worksheets/', 'xl/worksheets/_rels/') + '.rels' if bg_file else None

        with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                if item in skip_files:
                    continue
                elif item == cert_file:
                    zout.writestr(item, cert_xml.encode('utf-8'))
                elif bg_file and item == bg_file:
                    zout.writestr(item, bg_xml.encode('utf-8'))
                elif item == 'xl/_rels/workbook.xml.rels':
                    zout.writestr(item, fixed_wb_rels.encode('utf-8'))
                elif item == 'xl/workbook.xml':
                    zout.writestr(item, fixed_wb_xml.encode('utf-8'))
                elif item == '[Content_Types].xml':
                    zout.writestr(item, fixed_ct.encode('utf-8'))
                else:
                    zout.writestr(item, zin.read(item))

            # Write sheet rels (link sheets to their drawings)
            zout.writestr(cert_rels_path, make_sheet_rels('../drawings/drawing1.xml'))
            if bg_rels_path:
                zout.writestr(bg_rels_path, make_sheet_rels('../drawings/drawing2.xml'))

            # Write original cert drawing + rels + images
            zout.writestr('xl/drawings/drawing1.xml', orig_cert_drawing)
            zout.writestr('xl/drawings/_rels/drawing1.xml.rels', orig_cert_drawing_rels)
            for img_path, img_data in orig_cert_images.items():
                zout.writestr(img_path, img_data)

            # Write combined BG drawing + rels + original BG images (bg_imageN.png)
            zout.writestr('xl/drawings/drawing2.xml', combined_bg_drawing.encode('utf-8'))
            zout.writestr('xl/drawings/_rels/drawing2.xml.rels', combined_bg_rels.encode('utf-8'))
            for img_name, img_data in orig_bg_images.items():
                zout.writestr(f'xl/media/{img_name}', img_data)

    shutil.move(temp_path, xlsx_path)
    print("  Removed external link references")
    print("  Replaced cert drawing with original (2 images, proper XML)")
    print(f"  Built combined BG drawing ({len(bg1_anchors)+len(bg1_one_anchors)}+"
          f"{len(bg2_anchors)+len(bg2_one_anchors)} anchors)")


print("\nPost-processing xlsx...")
post_process_xlsx(output_path, ORIG_CERT, ORIG_BG)
print("Done!")
