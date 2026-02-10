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
INPUT_BG = PatternFill(start_color="FFF9C4", end_color="FFF9C4", fill_type="solid")  # light yellow
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
BG_COMPANY = Font(name="Calibri", size=22, color="000000")
BG_ADDR = Font(name="Calibri", size=18, color="000000")
BG_NAME_FONT = Font(name="Calibri", size=18, color="000000")
BG_CONTACT = Font(name="Calibri", size=14, bold=True, color="000000")
BG_FIELD = Font(name="Calibri", size=16, color="000000")

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
bg_border = Border(
    right=Side(style="medium", color="70D2A6"),
    top=Side(style="medium", color="70D2A6"),
    bottom=Side(style="medium", color="70D2A6"),
)
bg_border_left = Border(
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
# SHEET 1: Teilnehmer (Input Sheet)
# ══════════════════════════════════════════════════════════════════════════════
ws_tn = wb.active
ws_tn.title = "Teilnehmer"
ws_tn.sheet_properties.tabColor = "0F3460"

# Column widths
ws_tn.column_dimensions["A"].width = 5    # Nr
ws_tn.column_dimensions["B"].width = 18   # Labels / Nachname
ws_tn.column_dimensions["C"].width = 22   # Values / Vorname
ws_tn.column_dimensions["D"].width = 16   # Geburtsdatum
ws_tn.column_dimensions["E"].width = 18   # Labels2 / Ersthelfer
ws_tn.column_dimensions["F"].width = 22   # Values2
ws_tn.column_dimensions["G"].width = 8    # Label3
ws_tn.column_dimensions["H"].width = 12   # Value3

# ── Row 1: Title ────────────────────────────────────────────────────────────
ws_tn.row_dimensions[1].height = 38
ws_tn.merge_cells("A1:H1")
set_cell(ws_tn, 1, 1, "EH-Kurs – Teilnehmerliste 2026",
    font=TITLE_FONT, fill=HEADER_BG,
    alignment=Alignment(horizontal="left", vertical="center"))

# ── Rows 3-8: Course Metadata ──────────────────────────────────────────────
meta_rows = [
    # (row, label_col, label, value_col, merge_end, label2_col, label2, value2_col, merge2_end)
    (3, 2, "Datum:", 3, None, 5, "Beginn:", 6, None),
    (4, 2, "Lehrkraft:", 3, None, 5, "Kursart:", 6, None),
    (5, 2, "Registriernr.:", 3, None, 5, "QSEH-Kennziffer:", 6, None),
    (6, 2, "Lehrgangsort:", 3, 5, None, None, None, None),
    (7, 2, "Verantw. Arzt:", 3, None, 5, "Tel.:", 6, None),
    (8, 2, "Masken-Charge:", 3, None, None, None, None, None),
]

# Row 3 also has Ende
for r, lc, label, vc, me, l2c, l2, v2c, m2e in meta_rows:
    ws_tn.row_dimensions[r].height = 22
    set_cell(ws_tn, r, lc, label, font=LABEL_FONT, alignment=right_center)
    if me:
        ws_tn.merge_cells(start_row=r, start_column=vc, end_row=r, end_column=me)
    set_cell(ws_tn, r, vc, None, font=INPUT_FONT, fill=INPUT_BG,
             alignment=left_center, border=thin_border)
    if l2c and l2:
        set_cell(ws_tn, r, l2c, l2, font=LABEL_FONT, alignment=right_center)
    if v2c:
        if m2e:
            ws_tn.merge_cells(start_row=r, start_column=v2c, end_row=r, end_column=m2e)
        set_cell(ws_tn, r, v2c, None, font=INPUT_FONT, fill=INPUT_BG,
                 alignment=left_center, border=thin_border)

# Row 3: add Ende label + cell
set_cell(ws_tn, 3, 7, "Ende:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 3, 8, None, font=INPUT_FONT, fill=INPUT_BG,
         alignment=left_center, border=thin_border)

# Pre-fill readonly values
ws_tn.cell(row=5, column=6).value = QSEH_KENNZIFFER
ws_tn.cell(row=5, column=6).fill = LIGHT_BLUE_BG
ws_tn.cell(row=6, column=3).value = f"{COMPANY}, {COMPANY_WERK}"
ws_tn.cell(row=7, column=3).value = ARZT
ws_tn.cell(row=7, column=3).fill = LIGHT_BLUE_BG
ws_tn.cell(row=7, column=6).value = ARZT_TEL
ws_tn.cell(row=7, column=6).fill = LIGHT_BLUE_BG

# Number formats for date/time
ws_tn["C3"].number_format = "DD.MM.YYYY"
ws_tn["F3"].number_format = "HH:MM"
ws_tn["H3"].number_format = "HH:MM"

# Data validations
dv_date = DataValidation(type="date", operator="greaterThanOrEqual",
                         formula1="1", allow_blank=True,
                         showErrorMessage=True,
                         errorTitle="Ungültiges Datum",
                         error="Bitte ein gültiges Datum eingeben.")
dv_date.add("C3")
ws_tn.add_data_validation(dv_date)

dv_instructor = DataValidation(type="list",
    formula1='"' + ",".join(INSTRUCTORS) + '"',
    allow_blank=True, showDropDown=False)
dv_instructor.add("C4")
ws_tn.add_data_validation(dv_instructor)

dv_kursart = DataValidation(type="list",
    formula1='"' + ",".join(KURSARTEN) + '"',
    allow_blank=True, showDropDown=False)
dv_kursart.add("F4")
ws_tn.add_data_validation(dv_kursart)

# ── Row 10: Table Headers ──────────────────────────────────────────────────
ws_tn.row_dimensions[10].height = 25
headers = [
    (1, "Nr"),
    (2, "Nachname"),
    (3, "Vorname"),
    (4, "Geburtsdatum"),
    (5, "Ersthelfer (x)"),
]
for c, h in headers:
    set_cell(ws_tn, 10, c, h,
        font=HEADER_FONT, fill=HEADER_BG, alignment=center, border=thin_border)

# ── Rows 11-30: Participant Rows ───────────────────────────────────────────
for i in range(MAX_TN):
    r = 11 + i
    ws_tn.row_dimensions[r].height = 22
    alt = LIGHT_GRAY if i % 2 == 0 else WHITE_BG
    set_cell(ws_tn, r, 1, i + 1, font=VALUE_FONT, fill=alt, alignment=center,
             border=thin_border)
    for c in range(2, 6):
        nf = "DD.MM.YYYY" if c == 4 else None
        set_cell(ws_tn, r, c, None, font=INPUT_FONT, fill=alt,
                 alignment=left_center if c != 5 else center,
                 border=thin_border, number_format=nf)

# Ersthelfer validation (x or empty)
dv_eh = DataValidation(type="list", formula1='"x"', allow_blank=True,
                       showDropDown=False)
dv_eh.add(f"E11:E{10 + MAX_TN}")
ws_tn.add_data_validation(dv_eh)

# ── Row 32-33: Counters ───────────────────────────────────────────────────
ws_tn.row_dimensions[32].height = 25
set_cell(ws_tn, 32, 2, "Gesamtanzahl:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 32, 3, f"=COUNTA(B11:B{10+MAX_TN})",
    font=Font(name="Calibri", size=14, bold=True, color="0F3460"),
    fill=GREEN_BG, alignment=center, border=thin_border)
set_cell(ws_tn, 32, 4, "davon Ersthelfer:", font=LABEL_FONT, alignment=right_center)
set_cell(ws_tn, 32, 5, f'=COUNTIF(E11:E{10+MAX_TN},"x")',
    font=Font(name="Calibri", size=14, bold=True, color="0F3460"),
    fill=GREEN_BG, alignment=center, border=thin_border)

# ── Conditional formatting: highlight empty required fields ────────────────
# Pflichtfeld warning for Nachname if row has data in other cols
ws_tn.conditional_formatting.add(f"B11:C{10+MAX_TN}", FormulaRule(
    formula=[f'AND($B11="",$C11<>"")'],
    fill=PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid")))
ws_tn.conditional_formatting.add(f"C11:C{10+MAX_TN}", FormulaRule(
    formula=[f'AND($C11="",$B11<>"")'],
    fill=PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid")))

# ── Print Setup ───────────────────────────────────────────────────────────
ws_tn.page_setup.orientation = "portrait"
ws_tn.page_setup.paperSize = ws_tn.PAPERSIZE_A4
ws_tn.page_setup.fitToWidth = 1
ws_tn.page_setup.fitToHeight = 1
ws_tn.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
ws_tn.page_margins = PageMargins(left=0.5, right=0.5, top=0.5, bottom=0.5)
ws_tn.freeze_panes = "A11"
ws_tn.print_title_rows = "10:10"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 2: Bescheinigungen (Certificate Print Sheet)
# ══════════════════════════════════════════════════════════════════════════════
ws_cert = wb.create_sheet(title="Bescheinigungen")
ws_cert.sheet_properties.tabColor = "1B5E20"

CERT_COLS = 12   # columns per certificate
SPACER = 1       # spacer column
ROWS_PER_PAGE = 30

# Column widths
for cert_side in [0, CERT_COLS + SPACER]:
    base = cert_side + 1
    ws_cert.column_dimensions[get_column_letter(base)].width = 1.0       # left margin
    ws_cert.column_dimensions[get_column_letter(base + 1)].width = 5.5
    ws_cert.column_dimensions[get_column_letter(base + 2)].width = 5.5
    ws_cert.column_dimensions[get_column_letter(base + 3)].width = 5.5
    ws_cert.column_dimensions[get_column_letter(base + 4)].width = 4.5
    ws_cert.column_dimensions[get_column_letter(base + 5)].width = 4.5
    ws_cert.column_dimensions[get_column_letter(base + 6)].width = 4.5
    ws_cert.column_dimensions[get_column_letter(base + 7)].width = 5.5
    ws_cert.column_dimensions[get_column_letter(base + 8)].width = 4.0
    ws_cert.column_dimensions[get_column_letter(base + 9)].width = 4.0
    ws_cert.column_dimensions[get_column_letter(base + 10)].width = 4.0
    ws_cert.column_dimensions[get_column_letter(base + 11)].width = 1.0  # right margin

# Spacer column
spacer_col = CERT_COLS + 1
ws_cert.column_dimensions[get_column_letter(spacer_col)].width = 1.5


def write_certificate(ws, start_row, col_offset, tn_num):
    """Write one certificate at position (start_row, col_offset).
    col_offset: 0-based column offset (0=left cert, 13=right cert)
    tn_num: participant number 1-20 → Teilnehmer row = tn_num + 10
    """
    c = col_offset + 1  # 1-based column number (first col of this cert)
    r = start_row
    tn_r = tn_num + 10  # Row in Teilnehmer sheet

    def mc(row, c_start, c_end):
        ws.merge_cells(start_row=row, start_column=c + c_start,
                       end_row=row, end_column=c + c_end)

    def sc(row, col_off, value, **kwargs):
        return set_cell(ws, row, c + col_off, value, **kwargs)

    # Conditional wrapper: show field only if participant name exists
    def tn_if(field):
        return f'=IF(Teilnehmer!B{tn_r}<>"",Teilnehmer!{field}{tn_r},"")'

    def tn_shared(cell_ref):
        return f"=Teilnehmer!{cell_ref}"

    # ── Row 0: Name, Vorname, geb. am, DOB ─────────────────────────────
    mc(r, 0, 3)
    sc(r, 0, tn_if("B"), font=CERT_NAME, border=bottom_line, alignment=left_center)
    mc(r, 4, 6)
    sc(r, 4, tn_if("C"), font=CERT_NAME, border=bottom_line, alignment=left_center)
    sc(r, 7, "geb. am:", font=Font(name="Calibri", size=9, color="000000"),
       alignment=right_center)
    mc(r, 8, 10)
    cell_dob = sc(r, 8, tn_if("D"), font=CERT_VALUE, border=bottom_line, alignment=center)
    cell_dob.number_format = "DD.MM.YYYY"

    # ── Row 1: Labels ──────────────────────────────────────────────────
    mc(r+1, 1, 3)
    sc(r+1, 1, "Name", font=CERT_LABEL, border=top_line, alignment=center)
    mc(r+1, 5, 6)
    sc(r+1, 5, "Vorname", font=CERT_LABEL, border=top_line, alignment=center)

    # ── Row 3: Body text ───────────────────────────────────────────────
    mc(r+3, 0, 11)
    sc(r+3, 0,
       "hat an dem 9 Unterrichtseinheiten (Nettounterrichtszeit 9 x 45 Minuten) "
       "umfassenden Lehrgang",
       font=CERT_BODY, alignment=left_wrap)
    ws.row_dimensions[r+3].height = 24

    # ── Row 5: Date / Time ─────────────────────────────────────────────
    sc(r+5, 0, "am", font=CERT_BODY, alignment=left_center)
    mc(r+5, 1, 2)
    date_cell = sc(r+5, 1, tn_shared("C3"), font=CERT_VALUE, alignment=center,
                   border=bottom_line)
    date_cell.number_format = "DD.MM.YYYY"
    mc(r+5, 3, 5)
    sc(r+5, 3, "in der Zeit von", font=CERT_BODY, alignment=center)
    time_from = sc(r+5, 6, tn_shared("F3"), font=CERT_VALUE, alignment=center,
                   border=bottom_line)
    time_from.number_format = "HH:MM"
    sc(r+5, 7, "Uhr bis", font=CERT_BODY, alignment=center)
    time_to = sc(r+5, 8, tn_shared("H3"), font=CERT_VALUE, alignment=center,
                 border=bottom_line)
    time_to.number_format = "HH:MM"
    sc(r+5, 9, "Uhr", font=CERT_BODY, alignment=left_center)

    # ── Row 7: Instructor ──────────────────────────────────────────────
    mc(r+7, 0, 3)
    sc(r+7, 0, "unter der Leitung von", font=CERT_BODY, alignment=left_center)
    mc(r+7, 4, 7)
    sc(r+7, 4, tn_shared("C4"), font=CERT_NAME, alignment=left_center)
    mc(r+7, 8, 11)
    sc(r+7, 8, "erfolgreich teilgenommen.", font=CERT_BODY, alignment=left_center)

    # ── Row 9: Checkbox ────────────────────────────────────────────────
    mc(r+9, 0, 5)
    sc(r+9, 0, "Teilnehmerunterlagen ausgehändigt:", font=CERT_BODY,
       alignment=left_center)
    sc(r+9, 6, "X", font=CERT_BOLD_11, border=box_border, alignment=center)
    sc(r+9, 7, "ja", font=CERT_BODY, alignment=left_center)
    sc(r+9, 9, "nein", font=CERT_BODY, alignment=left_center)

    # ── Row 12: Location / Date ────────────────────────────────────────
    mc(r+12, 0, 2)
    sc(r+12, 0, "Rastatt", font=CERT_VALUE, alignment=left_center)
    sc(r+12, 3, ", den", font=CERT_BODY, alignment=left_center)
    mc(r+12, 4, 5)
    date2 = sc(r+12, 4, tn_shared("C3"), font=CERT_VALUE, alignment=center,
               border=bottom_line)
    date2.number_format = "DD.MM.YYYY"

    # ── Row 13: Labels under signature lines ───────────────────────────
    mc(r+13, 0, 2)
    sc(r+13, 0, "Ort", font=CERT_LABEL, alignment=center, border=top_line)
    mc(r+13, 4, 5)
    sc(r+13, 4, "Datum", font=CERT_LABEL, alignment=center, border=top_line)
    mc(r+13, 7, 11)
    sc(r+13, 7, "Unterschrift der Lehrkraft", font=CERT_LABEL,
       alignment=center, border=top_line)

    # ── Row 16: Ausbildungsverantwortlicher ────────────────────────────
    mc(r+16, 7, 11)
    sc(r+16, 7, "Ausbildungsverantwortlicher", font=CERT_LABEL,
       alignment=center, border=top_line)

    # ── Row 18: Ermächtigte Stelle ─────────────────────────────────────
    mc(r+18, 0, 4)
    sc(r+18, 0, "Name der ermächtigten Stelle:", font=CERT_BODY,
       alignment=left_center)
    mc(r+18, 5, 11)
    sc(r+18, 5, f"{COMPANY}, Werk Rastatt", font=CERT_VALUE,
       alignment=left_center)

    # ── Row 20-21: Kennziffer ──────────────────────────────────────────
    mc(r+20, 0, 4)
    sc(r+20, 0, "Kennziffer der ermächtigten Stelle", font=CERT_BODY,
       alignment=left_center)
    mc(r+21, 0, 4)
    sc(r+21, 0, "gemäß § 26 DGUV Vorschrift:", font=CERT_BODY,
       alignment=left_center)
    mc(r+21, 5, 7)
    sc(r+21, 5, QSEH_KENNZIFFER, font=CERT_VALUE, alignment=left_center)

    # ── Row 23: Registriernummer ───────────────────────────────────────
    mc(r+23, 0, 4)
    sc(r+23, 0, "Registriernummer der Schulung:", font=CERT_BODY,
       alignment=left_center)
    mc(r+23, 5, 8)
    sc(r+23, 5, tn_shared("C5"), font=CERT_VALUE, alignment=left_center)

    # ── Row 25-26: Footnote ────────────────────────────────────────────
    mc(r+25, 0, 11)
    sc(r+25, 0,
       "* Die Teilnahme an der Ausbildung in betrieblicher Erster Hilfe "
       "gilt als Schulung in Erster Hilfe gem. § 19 Fahrerlaubnis-",
       font=CERT_SMALL, alignment=left_center)
    mc(r+26, 0, 11)
    sc(r+26, 0, "   Verordnung (FeV).", font=CERT_SMALL, alignment=left_center)


# Generate 10 pages × 2 certificates per page
for page in range(10):
    page_start = page * ROWS_PER_PAGE + 1
    # Left certificate
    left_tn = page * 2 + 1
    write_certificate(ws_cert, page_start + 1, 0, left_tn)
    # Right certificate
    right_tn = page * 2 + 2
    write_certificate(ws_cert, page_start + 1, CERT_COLS + SPACER, right_tn)
    # Set row heights for this page
    for row_offset in range(ROWS_PER_PAGE):
        ws_cert.row_dimensions[page_start + row_offset].height = 14.25

# Page breaks between pages
for page in range(1, 10):
    brk = Break(id=page * ROWS_PER_PAGE)
    ws_cert.row_breaks.append(brk)

# Print setup
ws_cert.page_setup.orientation = "landscape"
ws_cert.page_setup.paperSize = ws_cert.PAPERSIZE_A4
ws_cert.page_setup.fitToWidth = 1
ws_cert.page_setup.fitToHeight = 0
ws_cert.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
ws_cert.page_margins = PageMargins(left=0.3, right=0.3, top=0.4, bottom=0.4,
                                    header=0.3, footer=0.3)
# Header: Internal classification (like original)
ws_cert.oddHeader.left.text = "&1Internal"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 3: BG-Liste (BG Participant List – 2 pages)
# ══════════════════════════════════════════════════════════════════════════════
ws_bg = wb.create_sheet(title="BG-Liste")
ws_bg.sheet_properties.tabColor = "FF6F00"

ws_bg.column_dimensions["A"].width = 6
ws_bg.column_dimensions["B"].width = 45
ws_bg.column_dimensions["C"].width = 22
ws_bg.column_dimensions["D"].width = 30
ws_bg.column_dimensions["E"].width = 15

BG_PAGE_ROWS = 42


def write_bg_page(ws, page_start, tn_start):
    """Write one BG list page (10 participants).
    page_start: first row of this page
    tn_start: first participant number (1 or 11)
    """
    r = page_start

    # Row 2: time reference (hidden, from original)
    ws.row_dimensions[r + 1].height = 6

    # Row 5: Company name and BG name
    ws.row_dimensions[r + 4].height = 38
    set_cell(ws, r + 4, 2, COMPANY, font=BG_COMPANY, alignment=left_center)
    set_cell(ws, r + 4, 4, BG_NAME, font=Font(name="Calibri", size=18, color="000000"),
             alignment=left_center)

    # Row 6: Address
    ws.row_dimensions[r + 5].height = 30
    set_cell(ws, r + 5, 2, COMPANY_STREET, font=BG_ADDR, alignment=left_center)

    # Row 7: PLZ + BG membership
    ws.row_dimensions[r + 6].height = 30
    set_cell(ws, r + 6, 2, COMPANY_PLZ, font=BG_ADDR, alignment=left_center)

    # Row 9: spacing
    ws.row_dimensions[r + 8].height = 30

    # Rows 10-19: Participant list
    for i in range(10):
        pr = r + 9 + i
        tn_num = tn_start + i
        tn_row = tn_num + 10  # Row in Teilnehmer sheet
        ws.row_dimensions[pr].height = 34

        # Number
        set_cell(ws, pr, 1, i + 1, font=VALUE_FONT, alignment=center)

        # Name: "Nachname, Vorname" via CONCATENATE
        name_cell = set_cell(ws, pr, 2, None, font=BG_NAME_FONT,
            alignment=Alignment(horizontal="left", vertical="center"),
            border=bg_border)
        name_cell.value = (
            f'=IF(Teilnehmer!B{tn_row}<>"",'
            f'Teilnehmer!B{tn_row}&", "&Teilnehmer!C{tn_row},"")'
        )

        # DOB
        dob_cell = set_cell(ws, pr, 3, None, font=BG_NAME_FONT,
            alignment=center, border=bg_border)
        dob_cell.value = (
            f'=IF(Teilnehmer!D{tn_row}<>"",Teilnehmer!D{tn_row},"")'
        )
        dob_cell.number_format = "DD.MM.YYYY"

    # Conditional formatting: hide empty rows (white text)
    name_range = f"B{r+9}:B{r+18}"
    dob_range = f"C{r+9}:C{r+18}"
    ws.conditional_formatting.add(name_range, CellIsRule(
        operator="equal", formula=['", "'],
        font=Font(color="FFFFFF")))
    ws.conditional_formatting.add(dob_range, CellIsRule(
        operator="equal", formula=["0"],
        font=Font(color="FFFFFF")))

    # Row 21: spacing
    ws.row_dimensions[r + 20].height = 24

    # Row 22-24: Contact person
    set_cell(ws, r + 21, 2, ARZT, font=BG_CONTACT, alignment=center)
    set_cell(ws, r + 22, 2, ARZT_TEL, font=BG_CONTACT, alignment=center)
    set_cell(ws, r + 23, 2, ARZT_EMAIL, font=BG_CONTACT,
             alignment=Alignment(horizontal="right", vertical="center"))

    # Row 25: Date
    ws.row_dimensions[r + 24].height = 26
    date_cell = set_cell(ws, r + 24, 4, "=Teilnehmer!C3", font=BG_FIELD,
                         alignment=center)
    date_cell.number_format = "DD.MM.YYYY"

    # Row 28: spacing
    ws.row_dimensions[r + 27].height = 8

    # Row 29: Date repeat
    date2 = set_cell(ws, r + 28, 3, "=Teilnehmer!C3", font=BG_FIELD,
                     alignment=center)
    date2.number_format = "DD.MM.YYYY"
    ws.merge_cells(start_row=r + 28, start_column=3, end_row=r + 28, end_column=4)

    # Row 31: Kennziffer + Lehrkraft
    set_cell(ws, r + 30, 2, f"  {QSEH_KENNZIFFER}", font=BG_FIELD,
             alignment=left_center)
    ws.merge_cells(start_row=r + 30, start_column=3, end_row=r + 30, end_column=4)
    set_cell(ws, r + 30, 3, "=Teilnehmer!C4", font=BG_FIELD, alignment=center)

    # Row 33: Registriernummer + Ausbildungsstelle
    set_cell(ws, r + 32, 2, None, font=BG_FIELD, alignment=left_center)
    ws.cell(row=r + 32, column=2).value = '="  "&Teilnehmer!C5'
    ws.merge_cells(start_row=r + 32, start_column=3, end_row=r + 32, end_column=4)
    set_cell(ws, r + 32, 3, f"{COMPANY}, {COMPANY_WERK}",
             font=BG_FIELD, alignment=center)

    # Rows 36-40: Stamp/address section
    ws.merge_cells(start_row=r + 35, start_column=3, end_row=r + 35, end_column=4)
    set_cell(ws, r + 35, 3, COMPANY, font=BG_FIELD, alignment=center)
    ws.merge_cells(start_row=r + 36, start_column=3, end_row=r + 36, end_column=4)
    set_cell(ws, r + 36, 3, "Werksärztlicher Dienst", font=BG_FIELD,
             alignment=center)
    ws.merge_cells(start_row=r + 37, start_column=3, end_row=r + 37, end_column=4)
    set_cell(ws, r + 37, 3, COMPANY_WERK, font=BG_FIELD, alignment=center)
    ws.merge_cells(start_row=r + 38, start_column=3, end_row=r + 38, end_column=4)
    set_cell(ws, r + 38, 3, COMPANY_STREET, font=BG_FIELD, alignment=center)

    # Date + PLZ
    date3 = set_cell(ws, r + 39, 2, "=Teilnehmer!C3", font=BG_FIELD,
                     alignment=center)
    date3.number_format = "DD.MM.YYYY"
    ws.merge_cells(start_row=r + 39, start_column=3, end_row=r + 39, end_column=4)
    set_cell(ws, r + 39, 3, COMPANY_PLZ, font=BG_FIELD, alignment=center)


# Page 1: Participants 1-10
write_bg_page(ws_bg, 1, 1)
# Page 2: Participants 11-20
write_bg_page(ws_bg, 1 + BG_PAGE_ROWS, 11)

# Page break between pages
ws_bg.row_breaks.append(Break(id=BG_PAGE_ROWS))

# Print setup
ws_bg.page_setup.orientation = "portrait"
ws_bg.page_setup.paperSize = ws_bg.PAPERSIZE_A4
ws_bg.page_setup.fitToWidth = 1
ws_bg.page_setup.fitToHeight = 0
ws_bg.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
ws_bg.page_margins = PageMargins(left=0.2, right=0.1, top=0.4, bottom=0.4)


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 4: Lehrgangsdoku (Course Protocol – replaces Word doc)
# ══════════════════════════════════════════════════════════════════════════════
ws_ld = wb.create_sheet(title="Lehrgangsdoku")
ws_ld.sheet_properties.tabColor = "6A1B9A"

ws_ld.column_dimensions["A"].width = 3    # margin
ws_ld.column_dimensions["B"].width = 22   # labels
ws_ld.column_dimensions["C"].width = 22   # values
ws_ld.column_dimensions["D"].width = 4    # spacing
ws_ld.column_dimensions["E"].width = 18   # labels2
ws_ld.column_dimensions["F"].width = 22   # values2
ws_ld.column_dimensions["G"].width = 4    # margin

# ── Title ────────────────────────────────────────────────────────────────────
ws_ld.merge_cells("A1:G1")
ws_ld.row_dimensions[1].height = 30
set_cell(ws_ld, 1, 1, "Lehrgangsdokumentation", font=LD_TITLE, alignment=center)

ws_ld.merge_cells("A2:G2")
ws_ld.row_dimensions[2].height = 18
set_cell(ws_ld, 2, 1, "gem. Abschnitt 2.4.6 DGUV Grundsatz 304-001",
    font=LD_SUBTITLE, alignment=center)

# ── Section 1: Angaben zur Ausbildungsstelle ─────────────────────────────────
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

# ── Section 2: Angaben zum Seminar ───────────────────────────────────────────
ws_ld.row_dimensions[14].height = 22
ws_ld.merge_cells("B14:F14")
set_cell(ws_ld, 14, 2, "Angaben zum Seminar", font=LD_SECTION,
         alignment=left_center)

# Kursart checkboxes (auto-filled from Teilnehmer dropdown)
ws_ld.row_dimensions[15].height = 20
ws_ld.merge_cells("B15:F15")
# Build checkbox formula: show X next to selected Kursart
set_cell(ws_ld, 15, 2, None, font=LD_LABEL, alignment=left_wrap)
ws_ld.cell(row=15, column=2).value = (
    '=IF(Teilnehmer!F4="EH-Ausbildung","X","  ")&" EH-Ausbildung       "'
    '&IF(Teilnehmer!F4="EH-Fortbildung","X","  ")&" EH-Fortbildung       "'
    '&IF(Teilnehmer!F4="EH-Schulung","X","  ")&" EH-Schulung"'
)

# Registriernummer + Lehrgangsort
ws_ld.row_dimensions[17].height = 20
set_cell(ws_ld, 17, 2, "Registriernummer:", font=LD_LABEL, alignment=left_center)
set_cell(ws_ld, 17, 3, "=Teilnehmer!C5", font=LD_VALUE, alignment=left_center)
set_cell(ws_ld, 17, 5, "Lehrgangsort:", font=LD_LABEL, alignment=left_center)
set_cell(ws_ld, 17, 6, f"{COMPANY}, {COMPANY_WERK}", font=LD_VALUE,
         alignment=left_center)

ws_ld.row_dimensions[18].height = 14
set_cell(ws_ld, 18, 2, "(aus dem QSEH-Portal)", font=LD_LABEL_SM,
         alignment=left_center)
set_cell(ws_ld, 18, 6, LOCATION_DETAIL, font=LD_LABEL_SM,
         alignment=left_center)

# Lehrgangsdatum + Uhrzeit
ws_ld.row_dimensions[20].height = 20
set_cell(ws_ld, 20, 2, "Lehrgangsdatum:", font=LD_LABEL, alignment=left_center)
date_ld = set_cell(ws_ld, 20, 3, "=Teilnehmer!C3", font=LD_VALUE,
                   alignment=left_center)
date_ld.number_format = "DD.MM.YYYY"
set_cell(ws_ld, 20, 5, "Uhrzeit:", font=LD_LABEL, alignment=left_center)
# Time range formula
ws_ld.cell(row=20, column=6).value = (
    '="von "&TEXT(Teilnehmer!F3,"HH:MM")&" Uhr bis "&TEXT(Teilnehmer!H3,"HH:MM")&" Uhr"'
)
ws_ld.cell(row=20, column=6).font = LD_VALUE
ws_ld.cell(row=20, column=6).alignment = left_center

# Lehrkraft
ws_ld.row_dimensions[22].height = 20
set_cell(ws_ld, 22, 2, "Name der Lehrkraft:", font=LD_LABEL, alignment=left_center)
set_cell(ws_ld, 22, 3, "=Teilnehmer!C4", font=LD_VALUE, alignment=left_center)

# Verantwortlicher Arzt
ws_ld.row_dimensions[24].height = 20
set_cell(ws_ld, 24, 2, "Verantwortlicher Arzt:", font=LD_LABEL,
         alignment=left_center)
set_cell(ws_ld, 24, 3, "=Teilnehmer!C7", font=LD_VALUE, alignment=left_center)

# Masken-Charge
ws_ld.row_dimensions[26].height = 20
set_cell(ws_ld, 26, 2, "Masken-Charge:", font=LD_LABEL, alignment=left_center)
set_cell(ws_ld, 26, 3, "=Teilnehmer!C8", font=LD_VALUE, alignment=left_center,
         border=bottom_line)

# ── Section 3: Anzahl der Teilnehmenden ──────────────────────────────────────
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

# ── Section 4: Anlagen ───────────────────────────────────────────────────────
ws_ld.row_dimensions[35].height = 22
ws_ld.merge_cells("B35:F35")
set_cell(ws_ld, 35, 2, "Anlagen", font=LD_SECTION, alignment=left_center)

ws_ld.merge_cells("B36:F38")
ws_ld.row_dimensions[36].height = 16
ws_ld.row_dimensions[37].height = 16
ws_ld.row_dimensions[38].height = 16
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

# ── Signature block ──────────────────────────────────────────────────────────
ws_ld.row_dimensions[42].height = 18
set_cell(ws_ld, 42, 2, "Für die Richtigkeit der Angaben:", font=LD_LABEL,
         alignment=left_center)

# Signature lines
ws_ld.row_dimensions[47].height = 18
# Left signature
ws_ld.merge_cells("B47:C47")
ws_ld.cell(row=47, column=2).value = '="Rastatt, "&TEXT(Teilnehmer!C3,"DD.MM.YYYY")'
ws_ld.cell(row=47, column=2).font = LD_VALUE_UL
ws_ld.cell(row=47, column=2).alignment = left_center
# Right signature
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

# ── Footer ───────────────────────────────────────────────────────────────────
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
# Save
# ══════════════════════════════════════════════════════════════════════════════
output_path = "/home/user/Ambulanzzeug/EH_Kurs_2026.xlsx"
wb.save(output_path)
print(f"Saved: {output_path}")
print(f"Sheets: {wb.sheetnames}")
print("Done!")
