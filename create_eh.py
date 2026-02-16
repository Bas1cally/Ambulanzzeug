#!/usr/bin/env python3
"""
Creates EH_Kurs_2026.xlsx – Consolidated First Aid Course Management.
Replaces 4 separate files:
  1. Namensliste_EH_Kurs.xlsx  (participant input)
  2. EH_Bescheinigung_Teilnehmer.xlsx  (certificates)
  3. EH_BG_Liste.xlsx  (BG participant list)
  4. Lehrgangsdoku.docx  (course protocol)
All in one file, one input point, three print-ready output sheets.
"""
import openpyxl
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
MAX_TN = 20  # Max participants

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
BG_NAME = "BG-Metall Nord-Süd"

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

# Certificate fonts
CERT_NAME = Font(name="Calibri", size=12, bold=True, color="000000")
CERT_BODY = Font(name="Calibri", size=9, color="000000")
CERT_VALUE = Font(name="Calibri", size=12, color="000000")
CERT_SMALL = Font(name="Calibri", size=7, color="000000")
CERT_LABEL = Font(name="Calibri", size=9, color="000000")
CERT_BOLD_11 = Font(name="Calibri", size=11, bold=True, color="000000")

# BG fonts
BG_TITLE = Font(name="Calibri", size=14, bold=True, color=DARK_BLUE)
BG_LABEL = Font(name="Calibri", size=10, color="333333")
BG_VALUE = Font(name="Calibri", size=11, color="000000")
BG_NAME_FONT = Font(name="Calibri", size=11, color="000000")
BG_SECTION = Font(name="Calibri", size=10, bold=True, color=DARK_BLUE)

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
top_line = Border(top=Side(style="thin", color="000000"))
box_border = Border(
    left=Side(style="thin", color="000000"),
    right=Side(style="thin", color="000000"),
    top=Side(style="thin", color="000000"),
    bottom=Side(style="thin", color="000000"),
)
green_border = Border(
    left=Side(style="medium", color="70D2A6"),
    right=Side(style="medium", color="70D2A6"),
    top=Side(style="medium", color="70D2A6"),
    bottom=Side(style="medium", color="70D2A6"),
)

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


wb = openpyxl.Workbook()

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 5: Hilfslisten (created first so named range is available)
# ══════════════════════════════════════════════════════════════════════════════
ws_help = wb.create_sheet(title="Hilfslisten")
ws_help.sheet_properties.tabColor = "9E9E9E"

set_cell(ws_help, 1, 1, "Lehrkräfte",
    font=HEADER_FONT, fill=HEADER_BG, alignment=center, border=thin_border)
ws_help.column_dimensions["A"].width = 25

for i, name in enumerate(INSTRUCTORS):
    set_cell(ws_help, 2 + i, 1, name, font=VALUE_FONT, alignment=left_center,
             border=thin_border)

# Named range for instructors (generous range to allow additions)
dn = DefinedName("Lehrkraefte", attr_text="Hilfslisten!$A$2:$A$20")
wb.defined_names.add(dn)


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 1: Teilnehmer (Input Sheet)
# ══════════════════════════════════════════════════════════════════════════════
ws_tn = wb.active
ws_tn.title = "Teilnehmer"
ws_tn.sheet_properties.tabColor = "0F3460"

# Column widths
ws_tn.column_dimensions["A"].width = 5
ws_tn.column_dimensions["B"].width = 18
ws_tn.column_dimensions["C"].width = 22
ws_tn.column_dimensions["D"].width = 16
ws_tn.column_dimensions["E"].width = 18
ws_tn.column_dimensions["F"].width = 22
ws_tn.column_dimensions["G"].width = 8
ws_tn.column_dimensions["H"].width = 12

# ── Row 1: Title
ws_tn.row_dimensions[1].height = 38
ws_tn.merge_cells("A1:H1")
set_cell(ws_tn, 1, 1, "EH-Kurs \u2013 Teilnehmerliste 2026",
    font=TITLE_FONT, fill=HEADER_BG,
    alignment=Alignment(horizontal="left", vertical="center"))

# ── Rows 3-8: Course Metadata
for r in range(3, 9):
    ws_tn.row_dimensions[r].height = 22

# Row 3: Datum / Beginn / Ende
set_cell(ws_tn, 3, 2, "Datum:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 3, 3, None, font=INPUT_FONT, fill=INPUT_BG,
         alignment=left_center, border=thin_border, number_format="DD.MM.YYYY")
set_cell(ws_tn, 3, 5, "Beginn:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 3, 6, None, font=INPUT_FONT, fill=INPUT_BG,
         alignment=left_center, border=thin_border, number_format="HH:MM")
set_cell(ws_tn, 3, 7, "Ende:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 3, 8, None, font=INPUT_FONT, fill=INPUT_BG,
         alignment=left_center, border=thin_border, number_format="HH:MM")

# Row 4: Lehrkraft / Kursart
set_cell(ws_tn, 4, 2, "Lehrkraft:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 4, 3, None, font=INPUT_FONT, fill=INPUT_BG,
         alignment=left_center, border=thin_border)
set_cell(ws_tn, 4, 5, "Kursart:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 4, 6, None, font=INPUT_FONT, fill=INPUT_BG,
         alignment=left_center, border=thin_border)

# Row 5: Registriernr / QSEH
set_cell(ws_tn, 5, 2, "Registriernr.:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 5, 3, None, font=INPUT_FONT, fill=INPUT_BG,
         alignment=left_center, border=thin_border)
set_cell(ws_tn, 5, 5, "QSEH-Kennziffer:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 5, 6, QSEH_KENNZIFFER, font=INPUT_FONT, fill=LIGHT_BLUE_BG,
         alignment=left_center, border=thin_border)

# Row 6: Lehrgangsort
set_cell(ws_tn, 6, 2, "Lehrgangsort:", font=LABEL_FONT, alignment=right_center)
ws_tn.merge_cells("C6:E6")
set_cell(ws_tn, 6, 3, f"{COMPANY}, {COMPANY_WERK}",
         font=INPUT_FONT, fill=INPUT_BG, alignment=left_center, border=thin_border)

# Row 7: Verantw. Arzt / Tel
set_cell(ws_tn, 7, 2, "Verantw. Arzt:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 7, 3, ARZT, font=INPUT_FONT, fill=LIGHT_BLUE_BG,
         alignment=left_center, border=thin_border)
set_cell(ws_tn, 7, 5, "Tel.:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 7, 6, ARZT_TEL, font=INPUT_FONT, fill=LIGHT_BLUE_BG,
         alignment=left_center, border=thin_border)

# Row 8: Masken-Charge
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

# ── Row 10: Table Headers
ws_tn.row_dimensions[10].height = 25
for c, h in [(1, "Nr"), (2, "Nachname"), (3, "Vorname"),
             (4, "Geburtsdatum"), (5, "Ersthelfer (x)")]:
    set_cell(ws_tn, 10, c, h,
        font=HEADER_FONT, fill=HEADER_BG, alignment=center, border=thin_border)

# ── Rows 11-30: Participant Rows
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

# ── Row 32: Counters
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
# SHEET 2: Bescheinigungen (Certificate Print Sheet)
# Layout: 7 wide columns per certificate, 1 spacer, 7 wide columns
# ══════════════════════════════════════════════════════════════════════════════
ws_cert = wb.create_sheet(title="Bescheinigungen")
ws_cert.sheet_properties.tabColor = "1B5E20"

CC = 7       # columns per certificate
SPACER = 1   # spacer column between left & right
ROWS_PER_PAGE = 30

# Column widths: generous so nothing gets cut off or shows ####
CERT_WIDTHS = [18, 16, 16, 10, 14, 16, 16]  # per cert

for side in [0, CC + SPACER]:          # left cert offset=0, right cert offset=8
    for i, w in enumerate(CERT_WIDTHS):
        ws_cert.column_dimensions[get_column_letter(side + i + 1)].width = w

# Spacer column
ws_cert.column_dimensions[get_column_letter(CC + 1)].width = 3


def write_cert(ws, start_row, col_offset, tn_num):
    """Write one certificate at (start_row, col_offset).
    col_offset: 0 for left, CC+SPACER for right
    tn_num: 1-20 → Teilnehmer row = tn_num + 10
    """
    c = col_offset + 1   # 1-based first column
    r = start_row
    tn = tn_num + 10     # row in Teilnehmer sheet

    def mc(row, cs, ce):
        ws.merge_cells(start_row=row, start_column=c + cs,
                       end_row=row, end_column=c + ce)

    def sc(row, co, value, **kw):
        return set_cell(ws, row, c + co, value, **kw)

    def tn_if(col_letter):
        return f'=IF(Teilnehmer!B{tn}<>"",Teilnehmer!{col_letter}{tn},"")'

    def tn_ref(ref):
        return f"=Teilnehmer!{ref}"

    # ── Row 0: Name / Vorname / geb. am / DOB ──────────────────────────
    mc(r, 0, 1)
    sc(r, 0, tn_if("B"), font=CERT_NAME, border=bottom_line, alignment=left_center)
    mc(r, 2, 3)
    sc(r, 2, tn_if("C"), font=CERT_NAME, border=bottom_line, alignment=left_center)
    sc(r, 4, "geb. am:", font=CERT_BODY, alignment=right_center)
    mc(r, 5, 6)
    dob = sc(r, 5, tn_if("D"), font=CERT_VALUE, border=bottom_line, alignment=center)
    dob.number_format = "DD.MM.YYYY"

    # ── Row 1: Labels ──────────────────────────────────────────────────
    mc(r+1, 0, 1)
    sc(r+1, 0, "Name", font=CERT_LABEL, border=top_line, alignment=center)
    mc(r+1, 2, 3)
    sc(r+1, 2, "Vorname", font=CERT_LABEL, border=top_line, alignment=center)

    # ── Row 3: Body text ───────────────────────────────────────────────
    mc(r+3, 0, 6)
    sc(r+3, 0,
       "hat an dem 9 Unterrichtseinheiten (Nettounterrichtszeit 9 x 45 Minuten) "
       "umfassenden Lehrgang",
       font=CERT_BODY, alignment=left_wrap)
    ws.row_dimensions[r+3].height = 24

    # ── Row 5: Date / Time ─────────────────────────────────────────────
    sc(r+5, 0, "am", font=CERT_BODY, alignment=left_center)
    d1 = sc(r+5, 1, tn_ref("C3"), font=CERT_VALUE, border=bottom_line,
            alignment=center)
    d1.number_format = "DD.MM.YYYY"
    sc(r+5, 2, "in der Zeit von", font=CERT_BODY, alignment=center)
    tf = sc(r+5, 3, tn_ref("F3"), font=CERT_VALUE, border=bottom_line,
            alignment=center)
    tf.number_format = "HH:MM"
    sc(r+5, 4, "Uhr bis", font=CERT_BODY, alignment=center)
    tt = sc(r+5, 5, tn_ref("H3"), font=CERT_VALUE, border=bottom_line,
            alignment=center)
    tt.number_format = "HH:MM"
    sc(r+5, 6, "Uhr", font=CERT_BODY, alignment=left_center)

    # ── Row 7: Instructor ──────────────────────────────────────────────
    mc(r+7, 0, 1)
    sc(r+7, 0, "unter der Leitung von", font=CERT_BODY, alignment=left_center)
    mc(r+7, 2, 4)
    sc(r+7, 2, tn_ref("C4"), font=CERT_NAME, alignment=left_center)
    mc(r+7, 5, 6)
    sc(r+7, 5, "erfolgreich teilgenommen.", font=CERT_BODY, alignment=left_center)

    # ── Row 9: Checkbox ────────────────────────────────────────────────
    mc(r+9, 0, 3)
    sc(r+9, 0, "Teilnehmerunterlagen ausgehändigt:", font=CERT_BODY,
       alignment=left_center)
    sc(r+9, 4, "X", font=CERT_BOLD_11, border=box_border, alignment=center)
    sc(r+9, 5, "ja", font=CERT_BODY, alignment=left_center)
    sc(r+9, 6, "nein", font=CERT_BODY, alignment=left_center)

    # ── Row 12: Location / Date ────────────────────────────────────────
    sc(r+12, 0, "Rastatt", font=CERT_VALUE, alignment=left_center)
    sc(r+12, 1, ", den", font=CERT_BODY, alignment=left_center)
    d2 = sc(r+12, 2, tn_ref("C3"), font=CERT_VALUE, border=bottom_line,
            alignment=center)
    d2.number_format = "DD.MM.YYYY"

    # ── Row 13: Signature labels ───────────────────────────────────────
    sc(r+13, 0, "Ort", font=CERT_LABEL, border=top_line, alignment=center)
    sc(r+13, 2, "Datum", font=CERT_LABEL, border=top_line, alignment=center)
    mc(r+13, 4, 6)
    sc(r+13, 4, "Unterschrift der Lehrkraft", font=CERT_LABEL,
       border=top_line, alignment=center)

    # ── Row 16: Ausbildungsverantwortlicher ────────────────────────────
    mc(r+16, 4, 6)
    sc(r+16, 4, "Ausbildungsverantwortlicher", font=CERT_LABEL,
       border=top_line, alignment=center)

    # ── Row 18: Ermächtigte Stelle ─────────────────────────────────────
    mc(r+18, 0, 2)
    sc(r+18, 0, "Name der ermächtigten Stelle:", font=CERT_BODY,
       alignment=left_center)
    mc(r+18, 3, 6)
    sc(r+18, 3, f"{COMPANY}, Werk Rastatt", font=CERT_VALUE,
       alignment=left_center)

    # ── Row 20-21: Kennziffer ──────────────────────────────────────────
    mc(r+20, 0, 3)
    sc(r+20, 0, "Kennziffer der ermächtigten Stelle", font=CERT_BODY,
       alignment=left_center)
    mc(r+21, 0, 3)
    sc(r+21, 0, "gemäß § 26 DGUV Vorschrift:", font=CERT_BODY,
       alignment=left_center)
    mc(r+21, 4, 5)
    sc(r+21, 4, QSEH_KENNZIFFER, font=CERT_VALUE, alignment=left_center)

    # ── Row 23: Registriernummer ───────────────────────────────────────
    mc(r+23, 0, 3)
    sc(r+23, 0, "Registriernummer der Schulung:", font=CERT_BODY,
       alignment=left_center)
    mc(r+23, 4, 6)
    sc(r+23, 4, tn_ref("C5"), font=CERT_VALUE, alignment=left_center)

    # ── Row 25-26: Footnote ────────────────────────────────────────────
    mc(r+25, 0, 6)
    sc(r+25, 0,
       "* Die Teilnahme an der Ausbildung in betrieblicher Erster Hilfe "
       "gilt als Schulung in Erster Hilfe gem. § 19 Fahrerlaubnis-",
       font=CERT_SMALL, alignment=left_center)
    mc(r+26, 0, 6)
    sc(r+26, 0, "   Verordnung (FeV).", font=CERT_SMALL, alignment=left_center)


# Generate 10 pages × 2 certificates per page
for page in range(10):
    page_start = page * ROWS_PER_PAGE + 1
    write_cert(ws_cert, page_start + 1, 0, page * 2 + 1)           # left
    write_cert(ws_cert, page_start + 1, CC + SPACER, page * 2 + 2)  # right
    for ro in range(ROWS_PER_PAGE):
        ws_cert.row_dimensions[page_start + ro].height = 14.25

# Page breaks
for page in range(1, 10):
    ws_cert.row_breaks.append(Break(id=page * ROWS_PER_PAGE))

# Print setup
ws_cert.page_setup.orientation = "landscape"
ws_cert.page_setup.paperSize = 8  # A3
ws_cert.page_setup.fitToWidth = 1
ws_cert.page_setup.fitToHeight = 0
ws_cert.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
ws_cert.page_margins = PageMargins(left=0.4, right=0.4, top=0.4, bottom=0.4,
                                    header=0.3, footer=0.3)
ws_cert.oddHeader.left.text = "&1Internal"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 3: BG-Liste (BG Participant List – original design with images)
# ══════════════════════════════════════════════════════════════════════════════
from openpyxl.drawing.image import Image as XlImage
import os

ws_bg = wb.create_sheet(title="BG-Liste")
ws_bg.sheet_properties.tabColor = "FF6F00"

# Original column widths
ws_bg.column_dimensions["A"].width = 10.66
ws_bg.column_dimensions["B"].width = 62.44
ws_bg.column_dimensions["C"].width = 38.55
ws_bg.column_dimensions["D"].width = 40.66
ws_bg.column_dimensions["E"].width = 21.66
ws_bg.column_dimensions["F"].width = 4.33

# Original row heights (per page)
BG_ROW_HEIGHTS = {
    1: 42, 2: 41.25, 3: 38.25, 4: 38.25, 5: 42, 6: 41.25, 7: 32.25,
    8: 41.25, 9: 34.5, 10: 39.75, 11: 39.75, 12: 39.75, 13: 39.75,
    14: 40.5, 15: 39.75, 16: 40.5, 17: 39, 18: 39.75, 19: 39.75,
    20: 42, 21: 27.75, 22: 32.25, 23: 32.25, 24: 32.25, 25: 30,
    26: 25.5, 27: 31.5, 28: 8.25, 29: 25.5, 30: 16.5, 31: 25.5,
    32: 21, 33: 26.25, 34: 15, 35: 15, 36: 29.25, 37: 29.25,
    38: 29.25, 39: 29.25, 40: 29.25, 41: 21, 42: 9, 43: 9,
}

# Green medium border (original style)
bg_green_r = Border(
    right=Side(style="medium", color="70D2A6"),
    top=Side(style="medium", color="70D2A6"),
    bottom=Side(style="medium", color="70D2A6"))
bg_green_l = Border(
    left=Side(style="medium", color="70D2A6"),
    top=Side(style="medium", color="70D2A6"),
    bottom=Side(style="medium", color="70D2A6"))

# Image definitions: (filename, from_col_0, from_row_0, cx_emu, cy_emu)
IMG_DIR = "/home/user/Ambulanzzeug/Neu EH/extracted_images/xl/media"
BG_IMAGES = [
    ("image3.png", 0, 0, 11374437, 2048161),
    ("image4.png", 0, 0, 514422, 4277322),
    ("image6.png", 0, 4, 5601482, 314369),
    ("image7.png", 1, 2, 1810003, 2181529),
    ("image8.png", 0, 5, 5401429, 238158),
    ("image9.png", 4, 2, 438211, 2486372),
    ("image10.png", 2, 5, 5544324, 304843),
    ("image11.png", 2, 3, 5410955, 476316),
    ("image5.png", 0, 6, 11336332, 1086002),
    ("image1.png", 0, 7, 533474, 6049219),
    ("image2.png", 3, 8, 1409897, 5296639),
    ("image12.png", 0, 19, 11355385, 771633),
    ("image13.png", 0, 20, 1381318, 2734057),
    ("image14.png", 0, 23, 6763694, 666843),
    ("image15.png", 2, 20, 200053, 1943371),
    ("image16.png", 2, 24, 790685, 314369),
    ("image17.png", 3, 24, 3877216, 876422),
    ("image18.png", 4, 20, 504895, 2162477),
    ("image19.png", 0, 25, 11279174, 760400),
    ("image20.png", 3, 25, 1458493, 4727593),
    ("image21.png", 0, 31, 10470280, 255764),
    ("image22.png", 2, 27, 152421, 4242233),
    ("image23.png", 0, 29, 10918018, 258657),
    ("image24.png", 0, 33, 10813228, 275661),
    ("image25.png", 0, 39, 10565543, 550839),
    ("image26.png", 0, 26, 114316, 4196409),
    ("image27.png", 1, 39, 724001, 266737),
]

BG_PAGE_ROWS = 43  # rows per page (matches original)

# Fonts matching original
BG_F24 = Font(name="Calibri", size=24, color="000000")
BG_F20 = Font(name="Calibri", size=20, color="000000")
BG_F18 = Font(name="Calibri", size=18, color="000000")
BG_F14B = Font(name="Calibri", size=14, bold=True, color="000000")
BG_F11 = Font(name="Calibri", size=11, color="000000")


def write_bg_page(ws, row_off, tn_start):
    """Write one BG-Liste page. row_off = 0 for page1, BG_PAGE_ROWS for page2."""
    o = row_off  # row offset (0-based addition to 1-based rows)

    # Set row heights
    for rel_row, h in BG_ROW_HEIGHTS.items():
        ws.row_dimensions[rel_row + o].height = h

    # ── Company / BG info (rows 5-7) ──────────────────────────────────
    set_cell(ws, 5 + o, 2, "Mercedes Benz AG",
        font=BG_F24, alignment=Alignment(horizontal="left", vertical="top"))
    set_cell(ws, 5 + o, 4, BG_NAME,
        font=BG_F20, alignment=left_center)
    set_cell(ws, 6 + o, 2, "Mercedesstr. 1", font=BG_F20)
    set_cell(ws, 7 + o, 2, COMPANY_PLZ, font=BG_F20)
    set_cell(ws, 7 + o, 4, None, font=BG_F18,
        alignment=Alignment(horizontal="left", vertical="top"),
        number_format="@")

    # ── Participant list (rows 10-19) ──────────────────────────────────
    for i in range(10):
        r = 10 + i + o
        tn_row = tn_start + i + 10  # Teilnehmer sheet row

        set_cell(ws, r, 1, i + 1, font=BG_F11, alignment=center)

        # Name: "Nachname, Vorname"
        name_cell = set_cell(ws, r, 2, None, font=BG_F20,
            alignment=Alignment(vertical="center"), border=bg_green_r)
        name_cell.value = (
            f'=IF(Teilnehmer!B{tn_row}<>"",'
            f'Teilnehmer!B{tn_row}&", "&Teilnehmer!C{tn_row},"")'
        )

        # DOB
        dob_cell = set_cell(ws, r, 3, None, font=BG_F20,
            alignment=center, border=bg_green_r, number_format="DD.MM.YYYY")
        dob_cell.value = f'=IF(Teilnehmer!D{tn_row}<>"",Teilnehmer!D{tn_row},"")'

    # Conditional formatting: hide empty rows (white text on ", " or 0)
    name_rng = f"B{10+o}:B{19+o}"
    dob_rng = f"C{10+o}:C{19+o}"
    ws.conditional_formatting.add(name_rng, CellIsRule(
        operator="equal", formula=['", "'],
        font=Font(color="FFFFFF")))
    ws.conditional_formatting.add(dob_rng, CellIsRule(
        operator="equal", formula=["0"],
        font=Font(color="FFFFFF")))

    # ── Contact (rows 22-24) ──────────────────────────────────────────
    set_cell(ws, 22 + o, 2, ARZT, font=BG_F14B, alignment=center)
    set_cell(ws, 23 + o, 2, ARZT_TEL, font=BG_F14B, alignment=center)
    set_cell(ws, 24 + o, 2, ARZT_EMAIL, font=BG_F14B,
        alignment=Alignment(horizontal="right", vertical="center"))

    # ── Dates / course info ────────────────────────────────────────────
    d25 = set_cell(ws, 25 + o, 4, "=Teilnehmer!C3",
        font=BG_F18, alignment=center, number_format="DD.MM.YYYY")

    ws.merge_cells(start_row=29+o, start_column=3, end_row=29+o, end_column=4)
    set_cell(ws, 29 + o, 3, "=Teilnehmer!C3",
        font=BG_F18, alignment=center, number_format="DD.MM.YYYY")

    set_cell(ws, 31 + o, 2, f"  {QSEH_KENNZIFFER}",
        font=BG_F18, alignment=left_center, number_format="@")
    ws.merge_cells(start_row=31+o, start_column=3, end_row=31+o, end_column=4)
    set_cell(ws, 31 + o, 3, "=Teilnehmer!C4", font=BG_F18, alignment=center)

    set_cell(ws, 33 + o, 2, None, font=BG_F18)
    ws.cell(row=33+o, column=2).value = '="  "&Teilnehmer!C5'
    ws.merge_cells(start_row=33+o, start_column=3, end_row=33+o, end_column=4)
    set_cell(ws, 33 + o, 3, "Mercedes Benz AG, PKW Werk",
        font=BG_F18, alignment=center)

    # ── Stamp/address block (rows 36-40) ──────────────────────────────
    stamp_data = [
        (36, "Mercedes Benz AG"),
        (37, "Werksärztlicher Dienst"),
        (38, "PKW-Werk Rastatt"),
        (39, "Mercedesstrasse 1"),
    ]
    for sr, text in stamp_data:
        ws.merge_cells(start_row=sr+o, start_column=3, end_row=sr+o, end_column=4)
        set_cell(ws, sr + o, 3, text, font=BG_F18, alignment=center)

    set_cell(ws, 40 + o, 2, "=Teilnehmer!C3",
        font=BG_F18, alignment=center, number_format="DD.MM.YYYY")
    ws.merge_cells(start_row=40+o, start_column=3, end_row=40+o, end_column=4)
    set_cell(ws, 40 + o, 3, COMPANY_PLZ, font=BG_F18, alignment=center)

    # ── Place all 27 images ───────────────────────────────────────────
    for fname, fcol, frow, cx, cy in BG_IMAGES:
        img_path = os.path.join(IMG_DIR, fname)
        if os.path.exists(img_path):
            img = XlImage(img_path)
            img.width = cx / 9525
            img.height = cy / 9525
            anchor = f"{get_column_letter(fcol + 1)}{frow + 1 + o}"
            ws.add_image(img, anchor)


# Page 1: Participants 1-10
write_bg_page(ws_bg, 0, 1)
# Page 2: Participants 11-20
write_bg_page(ws_bg, BG_PAGE_ROWS, 11)

ws_bg.row_breaks.append(Break(id=BG_PAGE_ROWS))

# Print setup (matches original)
ws_bg.page_setup.orientation = "portrait"
ws_bg.page_setup.paperSize = ws_bg.PAPERSIZE_A4
ws_bg.page_setup.fitToWidth = 1
ws_bg.page_setup.fitToHeight = 0
ws_bg.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
ws_bg.page_margins = PageMargins(left=0.197, right=0.079, top=0.394, bottom=0.394)
ws_bg.print_area = f"A1:E{BG_PAGE_ROWS * 2}"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 4: Lehrgangsdoku (Course Protocol – replaces Word doc)
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

# ── Title
ws_ld.merge_cells("A1:G1")
ws_ld.row_dimensions[1].height = 30
set_cell(ws_ld, 1, 1, "Lehrgangsdokumentation", font=LD_TITLE, alignment=center)

ws_ld.merge_cells("A2:G2")
ws_ld.row_dimensions[2].height = 18
set_cell(ws_ld, 2, 1, "gem. Abschnitt 2.4.6 DGUV Grundsatz 304-001",
    font=LD_SUBTITLE, alignment=center)

# ── Section 1: Ausbildungsstelle
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

# ── Section 2: Seminar
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

# ── Section 3: Teilnehmerzahl
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

# ── Section 4: Anlagen
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

# ── Signature block
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
wb.move_sheet("Hilfslisten", offset=4)

# ══════════════════════════════════════════════════════════════════════════════
# Save
# ══════════════════════════════════════════════════════════════════════════════
output_path = "/home/user/Ambulanzzeug/EH_Kurs_2026.xlsx"
wb.save(output_path)
print(f"Saved: {output_path}")
print(f"Sheets: {wb.sheetnames}")
print("Done!")
