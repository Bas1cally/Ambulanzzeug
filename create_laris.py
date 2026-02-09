#!/usr/bin/env python3
"""
Creates a fresh LARIS_2026.xlsx without macros, ActiveX, or sheet protection.
All dropdowns use native Excel data validation.
Includes: Dashboard, conditional formatting, date validation, print setup, overdue markers.
"""
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.page import PageMargins
from copy import copy
import datetime
import urllib.parse

wb = openpyxl.Workbook()

# ── Colors & Styles ──────────────────────────────────────────────────────────
HEADER_BG = PatternFill(start_color="0F3460", end_color="0F3460", fill_type="solid")
LIGHT_GRAY_BG = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
WHITE_BG = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

TITLE_FONT = Font(name="Calibri", size=18, bold=True, color="0F3460")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
NORMAL_FONT = Font(name="Calibri", size=10, color="333333")
BOLD_FONT = Font(name="Calibri", size=10, bold=True, color="333333")
BULLET_FONT = Font(name="Calibri", size=10, bold=True, color="0F3460")
COUNTER_FONT = Font(name="Calibri", size=14, bold=True, color="0F3460")
INFO_TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="333333")

thin_border = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)

wrap_align = Alignment(wrap_text=True, vertical="top")
center_align = Alignment(horizontal="center", vertical="center")
center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)

# ── Names list ───────────────────────────────────────────────────────────────
NAMES = [
    "Dr.Frei", "Dr.Schmidt", "Dr. Vitez", "Breig", "Krempl",
    "Siebert", "Putschler", "Jochim", "Zuber", "Göpfrich",
    "Zeller", "Schmidt", "Zieger-Buchta", "Müller-Horn",
    "Radimersky", "Wunsch",
]

REKO_OPTIONS = [
    "REKO",
    "Nicht erforderlich",
    "Besprochen mit  Dr.Frei",
    "Besprochen mit  Dr.Schmidt",
    "Besprochen mit  Dr. Vitez",
    "Besprochen mit  Breig",
    "Besprochen mit  Krempl",
    "Besprochen mit  Siebert",
    "Besprochen mit  Putschler",
    "Besprochen mit  Jochim",
    "Besprochen mit  Zuber",
    "Besprochen mit  Göpfrich",
    "Besprochen mit  Wunsch",
    "Besprochen mit  Zeller",
    "Besprochen mit  Schmidt",
    "Besprochen mit  Zieger-Buchta",
    "Besprochen mit  Müller-Horn",
    "Besprochen mit  Radimersky",
]

# ── Bereiche with leaders and existing data ──────────────────────────────────
TAB_COLORS = ["1976D2", "43A047", "FB8C00", "8E24AA", "E53935", "00ACC1", "6D4C41", "C0CA33", "F06292"]

BEREICHE = [
    {"name": "Sekretariat", "leiter": "Emily Schmidt", "leiter_email": "emily_kim.schmidt@mercedes-benz.com", "stellvertreter": "Larissa Radimersky", "stv_email": "larissa.radimersky@mercedes-benz.com", "data": []},
    {"name": "EDV", "leiter": "Siebert", "leiter_email": "emanuel.siebert@mercedes-benz.com", "stellvertreter": "", "stv_email": "", "data": []},
    {"name": "Med.&Verbr.-stoff", "leiter": "Zeller", "leiter_email": "tobias_felix.zeller@mercedes-benz.com", "stellvertreter": "Breig", "stv_email": "bernd.breig@mercedes-benz.com", "data": []},
    {"name": "WD-Fzg & Schutzkleidg", "leiter": "Wunsch", "leiter_email": "Fabian.Wunsch@mercedes-benz.com", "stellvertreter": "", "stv_email": "", "data": []},
    {"name": "RTW & MPG", "leiter": "Göpfrich", "leiter_email": "Markus.Goepfrich@mercedes-benz.com", "stellvertreter": "", "stv_email": "", "data": []},
    {"name": "Organisation Ambulanz", "leiter": "Wunsch", "leiter_email": "Fabian.Wunsch@mercedes-benz.com", "stellvertreter": "", "stv_email": "", "data": []},
    {"name": "Ablauf&Prozessprobleme", "leiter": "Putschler", "leiter_email": "walter.putschler@mercedes-benz.com", "stellvertreter": "", "stv_email": "", "data": []},
    {"name": "BGF", "leiter": "Zieger-Buchta", "leiter_email": "katrin.zieger-buchta@mercedes-benz.com", "stellvertreter": "Müller-Horn", "stv_email": "Susanne.Mueller-Horn@mercedes-benz.com", "data": []},
    {"name": "Hygiene", "leiter": "Wunsch", "leiter_email": "Fabian.Wunsch@mercedes-benz.com", "stellvertreter": "", "stv_email": "", "data": []},
]


def set_cell(ws, row, col, value, font=None, fill=None, alignment=None, border=None, number_format=None):
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


def setup_print(ws, title, cols="A:L"):
    """Configure print settings for a worksheet."""
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.6, bottom=0.6, header=0.3, footer=0.3)
    ws.oddHeader.center.text = f"LARIS - {title}"
    ws.oddHeader.center.size = 10
    ws.oddFooter.left.text = "Werksärztlicher Dienst Mercedes-Benz"
    ws.oddFooter.right.text = "Seite &P von &N"
    ws.print_title_rows = "4:4"  # Repeat header row on each page


def add_conditional_formatting(ws, max_row=504):
    """Add all conditional formatting rules to a Meldungs-Sheet."""
    data_range = f"A5:L{max_row}"

    # 1) Kategorie auto-color (Column J = col 10)
    kat_range = f"J5:J{max_row}"
    ws.conditional_formatting.add(kat_range, CellIsRule(
        operator="equal", formula=['"A"'], stopIfTrue=True,
        fill=PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")))
    ws.conditional_formatting.add(kat_range, CellIsRule(
        operator="equal", formula=['"B"'], stopIfTrue=True,
        fill=PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")))
    ws.conditional_formatting.add(kat_range, CellIsRule(
        operator="equal", formula=['"C"'], stopIfTrue=True,
        fill=PatternFill(start_color="FFE0B2", end_color="FFE0B2", fill_type="solid")))
    ws.conditional_formatting.add(kat_range, CellIsRule(
        operator="equal", formula=['"D"'], stopIfTrue=True,
        fill=PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")))

    # 2) Zeilen-Status: Erledigt (H hat Wert) = ganze Zeile leicht grün
    ws.conditional_formatting.add(data_range, FormulaRule(
        formula=[f'AND($B5<>"",$H5<>"")'],
        stopIfTrue=False,
        fill=PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")))

    # 3) Zeilen-Status: Offen (B hat Wert, H leer) = ganze Zeile leicht gelb
    ws.conditional_formatting.add(data_range, FormulaRule(
        formula=[f'AND($B5<>"",$H5="")'],
        stopIfTrue=False,
        fill=PatternFill(start_color="FFF8E1", end_color="FFF8E1", fill_type="solid")))

    # 4) Überfällig: >14 Tage offen ohne Kenntnisnahme (B hat Wert, E leer, >14 Tage)
    ws.conditional_formatting.add(data_range, FormulaRule(
        formula=[f'AND($B5<>"",$E5="",(TODAY()-$B5)>14)'],
        stopIfTrue=True,
        fill=PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid"),
        font=Font(color="B71C1C")))

    # 5) Letzte 7 Tage: Neue Meldungen (Meldung am innerhalb 7 Tage)
    #    Blauer linker Rand als "Neu"-Marker
    ws.conditional_formatting.add(f"A5:A{max_row}", FormulaRule(
        formula=[f'AND($B5<>"",(TODAY()-$B5)<=7)'],
        stopIfTrue=False,
        fill=PatternFill(start_color="BBDEFB", end_color="BBDEFB", fill_type="solid"),
        font=Font(bold=True, color="1565C0")))

    # 6) Pflichtfeld-Warnung: Datum da aber "Meldung durch" fehlt → rot markieren
    ws.conditional_formatting.add(f"C5:C{max_row}", FormulaRule(
        formula=[f'AND($B5<>"",$C5="")'],
        stopIfTrue=True,
        fill=PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid")))

    # 7) Pflichtfeld-Warnung: Datum da aber "Thema" fehlt → rot markieren
    ws.conditional_formatting.add(f"D5:D{max_row}", FormulaRule(
        formula=[f'AND($B5<>"",$D5="")'],
        stopIfTrue=True,
        fill=PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid")))


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 1: Dashboard (Gesamtübersicht)
# ══════════════════════════════════════════════════════════════════════════════
ws_dash = wb.active
ws_dash.title = "Dashboard"
ws_dash.sheet_properties.tabColor = "28A745"

ws_dash.column_dimensions["A"].width = 3
ws_dash.column_dimensions["B"].width = 28
ws_dash.column_dimensions["C"].width = 14
ws_dash.column_dimensions["D"].width = 14
ws_dash.column_dimensions["E"].width = 14
ws_dash.column_dimensions["F"].width = 5
ws_dash.column_dimensions["G"].width = 10
ws_dash.column_dimensions["H"].width = 10
ws_dash.column_dimensions["I"].width = 10
ws_dash.column_dimensions["J"].width = 10
ws_dash.column_dimensions["K"].width = 5
ws_dash.column_dimensions["L"].width = 22

# Title
ws_dash.merge_cells("B1:J2")
set_cell(ws_dash, 1, 2, "LARIS Dashboard - Gesamtübersicht",
         font=TITLE_FONT, alignment=Alignment(horizontal="center", vertical="center"))
ws_dash.row_dimensions[1].height = 25
ws_dash.row_dimensions[2].height = 25

# Gesamt-Zähler row
ws_dash.row_dimensions[3].height = 8  # spacer

ws_dash.row_dimensions[4].height = 35
ws_dash.merge_cells("B4:C4")
set_cell(ws_dash, 4, 2, "Gesamt Meldungen:",
         font=Font(name="Calibri", size=13, bold=True, color="333333"),
         alignment=Alignment(horizontal="right", vertical="center"))
set_cell(ws_dash, 4, 4, None,
         font=Font(name="Calibri", size=22, bold=True, color="0F3460"),
         alignment=center_align)
# Will fill formula after creating sheets

ws_dash.merge_cells("E4:F4")
set_cell(ws_dash, 4, 5, "Erledigt:",
         font=Font(name="Calibri", size=13, bold=True, color="333333"),
         alignment=Alignment(horizontal="right", vertical="center"))
set_cell(ws_dash, 4, 7, None,
         font=Font(name="Calibri", size=22, bold=True, color="28A745"),
         alignment=center_align)

ws_dash.merge_cells("H4:I4")
set_cell(ws_dash, 4, 8, "Offen:",
         font=Font(name="Calibri", size=13, bold=True, color="333333"),
         alignment=Alignment(horizontal="right", vertical="center"))
set_cell(ws_dash, 4, 10, None,
         font=Font(name="Calibri", size=22, bold=True, color="DC3545"),
         alignment=center_align)

# ── Bereichsübersicht Tabelle ────────────────────────────────────────────────
row = 6
dash_headers = ["Bereich", "Meldungen", "Erledigt", "Offen", "", "Kat A", "Kat B", "Kat C", "Kat D", "", "Leiter"]
for c, h in enumerate(dash_headers, 2):
    set_cell(ws_dash, row, c, h, font=HEADER_FONT, fill=HEADER_BG, alignment=center_wrap, border=thin_border)
ws_dash.row_dimensions[row].height = 30

# Data rows with formulas referencing each Bereich sheet
for i, bereich in enumerate(BEREICHE):
    r = row + 1 + i
    name = bereich["name"]
    safe_name = f"'{name}'"  # Quote sheet name for formulas

    alt_fill = LIGHT_GRAY_BG if i % 2 == 0 else WHITE_BG
    bereich_cell = set_cell(ws_dash, r, 2, name, font=Font(name="Calibri", size=10, bold=True, color="0F3460", underline="single"), fill=alt_fill, border=thin_border)
    bereich_cell.hyperlink = f"#'{name}'!A1"

    # Meldungen = COUNTA of column B (Meldung am)
    set_cell(ws_dash, r, 3, None, font=COUNTER_FONT, fill=alt_fill, alignment=center_align, border=thin_border)
    ws_dash.cell(row=r, column=3).value = f"=COUNTA({safe_name}!B5:B504)"

    # Erledigt = COUNTA of column H (Erledigt am)
    set_cell(ws_dash, r, 4, None,
             font=Font(name="Calibri", size=14, bold=True, color="28A745"),
             fill=alt_fill, alignment=center_align, border=thin_border)
    ws_dash.cell(row=r, column=4).value = f"=COUNTA({safe_name}!H5:H504)"

    # Offen = Meldungen - Erledigt
    set_cell(ws_dash, r, 5, None,
             font=Font(name="Calibri", size=14, bold=True, color="DC3545"),
             fill=alt_fill, alignment=center_align, border=thin_border)
    ws_dash.cell(row=r, column=5).value = f"=C{r}-D{r}"

    # Spacer
    set_cell(ws_dash, r, 6, None, fill=alt_fill)

    # Kategorie A count
    set_cell(ws_dash, r, 7, None, font=NORMAL_FONT, fill=alt_fill, alignment=center_align, border=thin_border)
    ws_dash.cell(row=r, column=7).value = f'=COUNTIF({safe_name}!J5:J504,"A")'

    # Kategorie B count
    set_cell(ws_dash, r, 8, None, font=NORMAL_FONT, fill=alt_fill, alignment=center_align, border=thin_border)
    ws_dash.cell(row=r, column=8).value = f'=COUNTIF({safe_name}!J5:J504,"B")'

    # Kategorie C count
    set_cell(ws_dash, r, 9, None, font=NORMAL_FONT, fill=alt_fill, alignment=center_align, border=thin_border)
    ws_dash.cell(row=r, column=9).value = f'=COUNTIF({safe_name}!J5:J504,"C")'

    # Kategorie D count
    set_cell(ws_dash, r, 10, None, font=NORMAL_FONT, fill=alt_fill, alignment=center_align, border=thin_border)
    ws_dash.cell(row=r, column=10).value = f'=COUNTIF({safe_name}!J5:J504,"D")'

    # Spacer
    set_cell(ws_dash, r, 11, None, fill=alt_fill)

    # Leiter
    set_cell(ws_dash, r, 12, bereich["leiter"], font=NORMAL_FONT, fill=alt_fill, border=thin_border,
             alignment=center_align)

# Summen-Zeile
sum_row = row + 1 + len(BEREICHE)
set_cell(ws_dash, sum_row, 2, "GESAMT", font=Font(name="Calibri", size=12, bold=True, color="FFFFFF"),
         fill=HEADER_BG, alignment=center_align, border=thin_border)
for col in [3, 4, 5, 7, 8, 9, 10]:
    col_letter = get_column_letter(col)
    set_cell(ws_dash, sum_row, col, None,
             font=Font(name="Calibri", size=12, bold=True, color="FFFFFF"),
             fill=HEADER_BG, alignment=center_align, border=thin_border)
    ws_dash.cell(row=sum_row, column=col).value = f"=SUM({col_letter}{row+1}:{col_letter}{sum_row-1})"
set_cell(ws_dash, sum_row, 6, None, fill=HEADER_BG)
set_cell(ws_dash, sum_row, 11, None, fill=HEADER_BG)
set_cell(ws_dash, sum_row, 12, None, fill=HEADER_BG, border=thin_border)

# Fill Gesamt-Zähler from sum row
ws_dash["D4"] = f"=C{sum_row}"
ws_dash["G4"] = f"=D{sum_row}"
ws_dash["J4"] = f"=E{sum_row}"

# Conditional formatting on Offen column - red if > 0
offen_range = f"E{row+1}:E{sum_row-1}"
ws_dash.conditional_formatting.add(offen_range, CellIsRule(
    operator="greaterThan", formula=["0"], stopIfTrue=False,
    fill=PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid"),
    font=Font(bold=True, color="B71C1C")))

# Conditional formatting on Kat D - red if > 0
katd_range = f"J{row+1}:J{sum_row-1}"
ws_dash.conditional_formatting.add(katd_range, CellIsRule(
    operator="greaterThan", formula=["0"], stopIfTrue=False,
    fill=PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid"),
    font=Font(bold=True, color="B71C1C")))

# Kategorie columns auto-color
for col_letter, color in [("G", "D4EDDA"), ("H", "FFF3CD"), ("I", "FFE0B2"), ("J", "F8D7DA")]:
    rng = f"{col_letter}{row+1}:{col_letter}{sum_row-1}"
    ws_dash.conditional_formatting.add(rng, CellIsRule(
        operator="greaterThan", formula=["0"], stopIfTrue=False,
        fill=PatternFill(start_color=color, end_color=color, fill_type="solid")))

# ── Legende ──────────────────────────────────────────────────────────────────
leg_row = sum_row + 2
ws_dash.merge_cells(f"B{leg_row}:E{leg_row}")
set_cell(ws_dash, leg_row, 2, "Legende Farbmarkierungen in den Bereichs-Sheets:",
         font=Font(name="Calibri", size=11, bold=True, color="333333"))

legends = [
    ("BBDEFB", "Nr.-Spalte blau", "Neue Meldung (letzte 7 Tage)"),
    ("FFF8E1", "Zeile gelb", "Meldung offen (noch nicht erledigt)"),
    ("E8F5E9", "Zeile grün", "Meldung erledigt"),
    ("FFCDD2", "Zeile rot", "Überfällig (>14 Tage ohne Kenntnisnahme)"),
    ("FFCDD2", "Zelle rot", "Pflichtfeld fehlt (Name oder Thema leer)"),
    ("D4EDDA", "Kategorie A", "Unsichere Handlung"),
    ("FFF3CD", "Kategorie B", "Beinahefehler ohne Folgen"),
    ("FFE0B2", "Kategorie C", "Fehler mit leichten bis mittelschweren Folgen"),
    ("F8D7DA", "Kategorie D", "Fehler mit gravierenden Folgen"),
]

for j, (color, label, desc) in enumerate(legends):
    r = leg_row + 1 + j
    set_cell(ws_dash, r, 2, None,
             fill=PatternFill(start_color=color, end_color=color, fill_type="solid"),
             border=thin_border)
    set_cell(ws_dash, r, 3, label, font=BOLD_FONT, border=thin_border)
    ws_dash.merge_cells(f"D{r}:F{r}")
    set_cell(ws_dash, r, 4, desc, font=NORMAL_FONT, border=thin_border)

ws_dash.freeze_panes = "B6"

# Print setup for dashboard
ws_dash.page_setup.orientation = "landscape"
ws_dash.page_setup.paperSize = ws_dash.PAPERSIZE_A4
ws_dash.page_setup.fitToWidth = 1
ws_dash.page_setup.fitToHeight = 1
ws_dash.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 2: Info
# ══════════════════════════════════════════════════════════════════════════════
ws_info = wb.create_sheet(title="Info")
ws_info.sheet_properties.tabColor = "0F3460"

ws_info.column_dimensions["A"].width = 3
ws_info.column_dimensions["B"].width = 4
ws_info.column_dimensions["C"].width = 55
ws_info.column_dimensions["D"].width = 20
ws_info.column_dimensions["E"].width = 5
ws_info.column_dimensions["F"].width = 20

# Title
ws_info.merge_cells("B1:F2")
set_cell(ws_info, 1, 2, "LARIS - Lernende Abweichungs-, Risiko- und Informationssammlung",
         font=TITLE_FONT, alignment=Alignment(horizontal="center", vertical="center"))

ws_info.merge_cells("B3:F3")
set_cell(ws_info, 3, 2, "Werksärztlicher Dienst Mercedes-Benz Rastatt/Kuppenheim",
         font=Font(name="Calibri", size=11, italic=True, color="666666"),
         alignment=Alignment(horizontal="center"))

ws_info.merge_cells("B4:F4")
set_cell(ws_info, 4, 2, "Version 2026 - Ohne Makros & Blattschutz",
         font=Font(name="Calibri", size=10, bold=True, color="28A745"),
         alignment=Alignment(horizontal="center"))

# Purpose
row = 6
ws_info.merge_cells(f"B{row}:F{row}")
set_cell(ws_info, row, 2, "Die Liste dient zur Meldung von:", font=INFO_TITLE_FONT)

for i, item in enumerate(["Vorkommnissen", "fehlendem oder fehlerhaftem Material", "Administrationsbedarf im Bereich"]):
    r = row + 1 + i
    set_cell(ws_info, r, 2, "●", font=BULLET_FONT)
    ws_info.merge_cells(f"C{r}:F{r}")
    set_cell(ws_info, r, 3, item, font=NORMAL_FONT)

row = 10
ws_info.merge_cells(f"C{row}:F{row}")
set_cell(ws_info, row, 3, "an den jeweiligen Bereichsleiter!", font=BOLD_FONT)

row = 12
ws_info.merge_cells(f"B{row}:F{row}")
set_cell(ws_info, row, 2, "Informationen die alle betreffen müssen über E-Mail kommuniziert werden!",
         font=Font(name="Calibri", size=10, bold=True, color="DC3545"))

# Erstellen einer Meldung
row = 14
ws_info.merge_cells(f"B{row}:F{row}")
set_cell(ws_info, row, 2, "Erstellen einer Meldung", font=INFO_TITLE_FONT)

for i, txt in enumerate([
    'Unten den entsprechenden Reiter für den Bereich auswählen',
    'Im Feld "Meldung durch" den eigenen Namen auswählen (Dropdown)',
    'Datum im Feld "Meldung am" eintragen (Format: TT.MM.JJJJ)',
    'Info bei "Thema" eintragen',
    'Auf den grünen "Mail an Bereichsleiter"-Link klicken (öffnet Outlook automatisch)',
    'Speichern',
]):
    r = row + 1 + i
    set_cell(ws_info, r, 2, "●", font=BULLET_FONT)
    ws_info.merge_cells(f"C{r}:F{r}")
    set_cell(ws_info, r, 3, txt, font=NORMAL_FONT)

# Abarbeiten
row = 22
ws_info.merge_cells(f"B{row}:F{row}")
set_cell(ws_info, row, 2, "Abarbeiten einer Meldung durch den Bereichsleiter", font=INFO_TITLE_FONT)

for i, txt in enumerate([
    'Bei Kenntnisnahme im Feld "Kenntnisnahme durch" eigenen Namen auswählen und Datum eintragen',
    'Massnahmen im entsprechenden Feld dokumentieren',
    'Nach Erledigung im Feld "Erledigt durch" eigenen Namen auswählen und Datum eintragen',
    'Evtl. weitere Infos im entsprechenden Feld eintragen',
    'In Spalte "In REKO..." auswählen wie besprochen wurde',
]):
    r = row + 1 + i
    set_cell(ws_info, r, 2, "●", font=BULLET_FONT)
    ws_info.merge_cells(f"C{r}:F{r}")
    set_cell(ws_info, r, 3, txt, font=NORMAL_FONT)

# Hinweise
row = 29
ws_info.merge_cells(f"B{row}:F{row}")
set_cell(ws_info, row, 2, "Hinweise zur neuen Version", font=INFO_TITLE_FONT)

for i, txt in enumerate([
    "Kein Blattschutz - alle Zellen sind frei editierbar",
    "Keine Makros/ActiveX - kein Hängen, keine Sicherheitswarnungen",
    "Dropdown-Listen für Namen und Kategorien (Klick auf die Zelle)",
    "Datumsfelder manuell ausfüllen (Format: TT.MM.JJJJ HH:MM)",
    "Meldungs- und Erledigungszähler werden automatisch berechnet (Formeln)",
    "Mail-Versand per klickbarem Link (öffnet Outlook mit Empfänger + Betreff + BCC)",
    "Dashboard-Reiter zeigt Gesamtübersicht aller Bereiche auf einen Blick",
    "Farbige Zeilen: Grün=Erledigt, Gelb=Offen, Rot=Überfällig (>14 Tage)",
    "Neue Meldungen (letzte 7 Tage) werden in der Nr.-Spalte blau markiert",
    "Druckoptimiert: Querformat A4, Kopfzeile wiederholt sich automatisch",
    "Pflichtfelder: Fehlender Name/Thema wird rot markiert wenn Datum eingetragen",
    "Jeder Bereich hat eine eigene Tab-Farbe für schnelle Navigation",
    "Stellvertreter werden automatisch in CC der Mail genommen",
    "Dashboard-Links: Klick auf Bereichsname springt direkt zum Sheet",
]):
    r = row + 1 + i
    set_cell(ws_info, r, 2, "✓", font=Font(name="Calibri", size=10, bold=True, color="28A745"))
    ws_info.merge_cells(f"C{r}:F{r}")
    set_cell(ws_info, r, 3, txt, font=NORMAL_FONT)

# Bereichstabelle
row = 46
headers_info = ["Bereich", "", "Leiter", "", "Stellvertreter", ""]
for c, h in enumerate(headers_info, 2):
    set_cell(ws_info, row, c, h, font=HEADER_FONT, fill=HEADER_BG, alignment=center_align, border=thin_border)
ws_info.merge_cells(f"B{row}:C{row}")
ws_info.merge_cells(f"D{row}:E{row}")

for i, bereich in enumerate(BEREICHE):
    r = row + 1 + i
    ws_info.merge_cells(f"B{r}:C{r}")
    set_cell(ws_info, r, 2, bereich["name"], font=NORMAL_FONT, border=thin_border)
    ws_info.merge_cells(f"D{r}:E{r}")
    set_cell(ws_info, r, 4, bereich["leiter"], font=NORMAL_FONT, border=thin_border)
    set_cell(ws_info, r, 6, bereich.get("stellvertreter", ""), font=NORMAL_FONT, border=thin_border)


# ══════════════════════════════════════════════════════════════════════════════
# MELDUNGS-SHEETS (one per Bereich)
# ══════════════════════════════════════════════════════════════════════════════

HEADERS = [
    ("Nr.", 5),
    ("Meldung am", 18),
    ("Meldung durch", 16),
    ("Thema", 55),
    ("Kenntnisnahme am", 18),
    ("Kenntnisnahme durch", 17),
    ("Massnahme", 55),
    ("Erledigt am", 18),
    ("Erledigt durch", 15),
    ("Kategorie", 11),
    ("Info", 40),
    ("In REKO oder Persönlich besprochen?", 28),
]

KATEGORIE_LEGEND = (
    "A = unsichere Handlung\n"
    "B = Beinahefehler ohne Folgen\n"
    "C = Fehler mit leichten bis mittelschweren Folgen\n"
    "D = Fehler mit gravierenden Folgen"
)

MAX_DATA_ROW = 504

for b_idx, bereich in enumerate(BEREICHE):
    ws = wb.create_sheet(title=bereich["name"])
    ws.sheet_properties.tabColor = TAB_COLORS[b_idx % len(TAB_COLORS)]

    # Column widths
    for i, (_, width) in enumerate(HEADERS, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # ── Row 1: Bereich title + Dashboard link ───────────────────────────────
    ws.merge_cells("A1:K1")
    set_cell(ws, 1, 1, bereich["name"],
             font=Font(name="Calibri", size=16, bold=True, color="FFFFFF"),
             fill=PatternFill(start_color="0F3460", end_color="0F3460", fill_type="solid"),
             alignment=Alignment(horizontal="left", vertical="center"))
    ws.row_dimensions[1].height = 35

    dash_link = set_cell(ws, 1, 12, "← Dashboard",
        font=Font(name="Calibri", size=10, bold=True, color="FFFFFF", underline="single"),
        fill=PatternFill(start_color="0F3460", end_color="0F3460", fill_type="solid"),
        alignment=Alignment(horizontal="center", vertical="center"))
    dash_link.hyperlink = "#Dashboard!A1"

    # ── Row 2: Counters + Kategorie-Legende ──────────────────────────────────
    ws.row_dimensions[2].height = 55
    set_cell(ws, 2, 1, "Meldungen:", font=BOLD_FONT, alignment=Alignment(horizontal="right", vertical="center"))
    ws["B2"] = f'=COUNTA(B5:B{MAX_DATA_ROW})'
    set_cell(ws, 2, 2, None, font=COUNTER_FONT, alignment=center_align)

    set_cell(ws, 2, 3, "Erledigt:", font=BOLD_FONT, alignment=Alignment(horizontal="right", vertical="center"))
    ws["D2"] = f'=COUNTA(H5:H{MAX_DATA_ROW})'
    set_cell(ws, 2, 4, None, font=Font(name="Calibri", size=14, bold=True, color="28A745"), alignment=center_align)

    set_cell(ws, 2, 5, "Offen:", font=BOLD_FONT, alignment=Alignment(horizontal="right", vertical="center"))
    ws["F2"] = '=B2-D2'
    set_cell(ws, 2, 6, None, font=Font(name="Calibri", size=14, bold=True, color="DC3545"), alignment=center_align)

    ws.merge_cells("H2:L2")
    set_cell(ws, 2, 8, KATEGORIE_LEGEND,
             font=Font(name="Calibri", size=9, color="666666"),
             alignment=Alignment(wrap_text=True, vertical="top"),
             fill=PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid"))

    # ── Row 3: Mail-Link ─────────────────────────────────────────────────────
    ws.row_dimensions[3].height = 30
    leiter_email = bereich["leiter_email"]
    leiter_name = bereich["leiter"]
    bereich_name = bereich["name"]
    stv_email = bereich.get("stv_email", "")
    stv_name = bereich.get("stellvertreter", "")
    bcc_email = "walter.putschler@mercedes-benz.com"

    subject = urllib.parse.quote(f"LARIS Meldung - {bereich_name}")
    body = urllib.parse.quote(
        f"Neue LARIS-Meldung im Bereich: {bereich_name}\n\n"
        f"Bitte in LARIS prüfen und bearbeiten.\n\n"
        f"Mit freundlichen Grüßen")
    # Build mailto with optional CC for Stellvertreter
    mailto_params = f"subject={subject}&bcc={bcc_email}&body={body}"
    if stv_email:
        mailto_params = f"cc={stv_email}&{mailto_params}"
    mailto_url = f"mailto:{leiter_email}?{mailto_params}"

    ws.merge_cells("A3:G3")
    mail_label = f"Mail an: {leiter_name}"
    if stv_name:
        mail_label += f" (CC: {stv_name})"
    mail_cell = set_cell(ws, 3, 1,
        f"✉ {mail_label}",
        font=Font(name="Calibri", size=11, bold=True, color="FFFFFF", underline="single"),
        fill=PatternFill(start_color="28A745", end_color="28A745", fill_type="solid"),
        alignment=Alignment(horizontal="center", vertical="center"))
    mail_cell.hyperlink = mailto_url

    ws.merge_cells("H3:L3")
    # Dynamic mail reminder: shows count of unacknowledged entries
    reminder_cell = set_cell(ws, 3, 8, None,
        font=Font(name="Calibri", size=11, bold=True, color="333333"),
        alignment=Alignment(horizontal="center", vertical="center"))
    reminder_cell.value = (
        f'=IF(COUNTIFS(B5:B{MAX_DATA_ROW},"<>",E5:E{MAX_DATA_ROW},"")>0,'
        f'"⚠ "&COUNTIFS(B5:B{MAX_DATA_ROW},"<>",E5:E{MAX_DATA_ROW},"")&'
        f'" Meldung(en) ohne Kenntnisnahme - Mail senden!",'
        f'"✓ Alles bearbeitet | BCC: Putschler")'
    )

    # Conditional formatting: red when unacknowledged entries exist
    ws.conditional_formatting.add("H3", FormulaRule(
        formula=[f'COUNTIFS(B5:B{MAX_DATA_ROW},"<>",E5:E{MAX_DATA_ROW},"")>0'],
        stopIfTrue=True,
        fill=PatternFill(start_color="FF6F00", end_color="FF6F00", fill_type="solid"),
        font=Font(name="Calibri", size=11, bold=True, color="FFFFFF")))
    # Green when all done
    ws.conditional_formatting.add("H3", FormulaRule(
        formula=[f'COUNTIFS(B5:B{MAX_DATA_ROW},"<>",E5:E{MAX_DATA_ROW},"")=0'],
        stopIfTrue=True,
        fill=PatternFill(start_color="28A745", end_color="28A745", fill_type="solid"),
        font=Font(name="Calibri", size=11, bold=True, color="FFFFFF")))

    # ── Row 4: Headers ───────────────────────────────────────────────────────
    ws.row_dimensions[4].height = 32
    for col_idx, (header_text, _) in enumerate(HEADERS, 1):
        set_cell(ws, 4, col_idx, header_text,
                 font=HEADER_FONT, fill=HEADER_BG, alignment=center_wrap, border=thin_border)

    # ── Data Validations (reference named ranges in Hilfslisten) ─────────────
    dv_names = DataValidation(type="list", formula1="=NamenListe", allow_blank=True)
    dv_names.error = "Bitte einen Namen aus der Liste wählen"
    dv_names.errorTitle = "Ungültiger Name"
    dv_names.prompt = "Name auswählen"
    dv_names.promptTitle = "Mitarbeiter"
    ws.add_data_validation(dv_names)
    dv_names.add(f"C5:C{MAX_DATA_ROW}")
    dv_names.add(f"F5:F{MAX_DATA_ROW}")
    dv_names.add(f"I5:I{MAX_DATA_ROW}")

    dv_kat = DataValidation(type="list", formula1="=KatListe", allow_blank=True)
    dv_kat.error = "Bitte A, B, C oder D wählen"
    dv_kat.errorTitle = "Ungültige Kategorie"
    dv_kat.prompt = "Kategorie wählen (A-D)"
    dv_kat.promptTitle = "Fehlerkategorie"
    ws.add_data_validation(dv_kat)
    dv_kat.add(f"J5:J{MAX_DATA_ROW}")

    dv_reko = DataValidation(type="list", formula1="=REKOListe", allow_blank=True)
    dv_reko.error = "Bitte aus der Liste wählen"
    dv_reko.errorTitle = "Ungültige Auswahl"
    dv_reko.prompt = "Besprechungsart wählen"
    dv_reko.promptTitle = "REKO / Besprochen"
    ws.add_data_validation(dv_reko)
    dv_reko.add(f"L5:L{MAX_DATA_ROW}")

    # ── Existing Data ────────────────────────────────────────────────────────
    for row_idx, entry in enumerate(bereich["data"]):
        r = 5 + row_idx
        set_cell(ws, r, 1, row_idx + 1, font=NORMAL_FONT, alignment=center_align, border=thin_border)
        set_cell(ws, r, 2, entry["meldung_am"], font=NORMAL_FONT, alignment=center_align,
                 border=thin_border, number_format="DD.MM.YYYY HH:MM")
        set_cell(ws, r, 3, entry["meldung_durch"], font=NORMAL_FONT, alignment=center_align, border=thin_border)
        set_cell(ws, r, 4, entry["thema"], font=NORMAL_FONT, alignment=wrap_align, border=thin_border)
        set_cell(ws, r, 5, entry["kenntnisnahme_am"], font=NORMAL_FONT, alignment=center_align,
                 border=thin_border, number_format="DD.MM.YYYY HH:MM")
        set_cell(ws, r, 6, entry["kenntnisnahme_durch"], font=NORMAL_FONT, alignment=center_align, border=thin_border)
        set_cell(ws, r, 7, entry["massnahme"], font=NORMAL_FONT, alignment=wrap_align, border=thin_border)
        set_cell(ws, r, 8, entry["erledigt_am"], font=NORMAL_FONT, alignment=center_align,
                 border=thin_border, number_format="DD.MM.YYYY HH:MM")
        set_cell(ws, r, 9, entry["erledigt_durch"], font=NORMAL_FONT, alignment=center_align, border=thin_border)
        set_cell(ws, r, 10, entry.get("kategorie", ""), font=BOLD_FONT, alignment=center_align, border=thin_border)
        set_cell(ws, r, 11, entry.get("info", ""), font=NORMAL_FONT, alignment=wrap_align, border=thin_border)
        set_cell(ws, r, 12, entry.get("reko", ""), font=NORMAL_FONT, alignment=wrap_align, border=thin_border)
        ws.row_dimensions[r].height = max(30, 15 * (1 + len(entry.get("thema", "")) // 70))

    # ── Pre-format empty rows ────────────────────────────────────────────────
    next_data_row = 5 + len(bereich["data"])
    for r in range(next_data_row, next_data_row + 50):
        # Nr. formula
        set_cell(ws, r, 1, f'=IF(B{r}<>"",ROW()-4,"")', font=NORMAL_FONT, alignment=center_align, border=thin_border)
        for col in range(2, 13):
            set_cell(ws, r, col, None, font=NORMAL_FONT, border=thin_border)
            if col in (2, 5, 8):
                ws.cell(row=r, column=col).number_format = "DD.MM.YYYY HH:MM"
                ws.cell(row=r, column=col).alignment = center_align
            elif col in (3, 6, 9, 10):
                ws.cell(row=r, column=col).alignment = center_align
            else:
                ws.cell(row=r, column=col).alignment = wrap_align

    # ── Conditional Formatting ───────────────────────────────────────────────
    add_conditional_formatting(ws, MAX_DATA_ROW)

    # ── Freeze panes ────────────────────────────────────────────────────────
    ws.freeze_panes = "A5"

    # ── Print Setup ──────────────────────────────────────────────────────────
    setup_print(ws, bereich["name"])


# ══════════════════════════════════════════════════════════════════════════════
# SHEET: Fehlerkategorien
# ══════════════════════════════════════════════════════════════════════════════
ws_kat = wb.create_sheet(title="Fehlerkategorien")
ws_kat.sheet_properties.tabColor = "DC3545"

ws_kat.column_dimensions["A"].width = 12
ws_kat.column_dimensions["B"].width = 45
ws_kat.column_dimensions["C"].width = 50
ws_kat.column_dimensions["D"].width = 50

headers_kat = ["Kategorie", "Beschreibung", "Details", "Beispiele"]
for col, h in enumerate(headers_kat, 1):
    set_cell(ws_kat, 1, col, h, font=HEADER_FONT, fill=HEADER_BG, alignment=center_wrap, border=thin_border)
ws_kat.row_dimensions[1].height = 30

kat_data = [
    ("D", "Fehler mit gravierenden Folgen",
     "daraus resultierende schwere Gesundheitsschäden, wesentliche finanzielle Nachteile oder ein erheblicher Imageverlust",
     "gesetzl. vorgeschriebene Unterweisungen oder Überprüfungen liegen nicht vor, Behandlungsfehler mit Konsequenzen, Defekte von Notfallgeräten"),
    ("C", "Fehler mit leichten bis mittelschweren Folgen",
     "Fehler mit Gesundheitsschäden, Einsatzfähigkeit HSM gefährdet, Fehler führte zu externer Kundenbeschwerde, Vorgegebener Standard wird nicht durchgeführt",
     "Behandlungsfehler ohne Konsequenzen, Termin- oder Materialkonflikt nicht rechtzeitig erkannt, Ausgang von Medikamenten, NADOK Protokolle nicht vollständig ausgefüllt"),
    ("B", "Beinahefehler ohne Folgen",
     "ein Fehler wurde noch vor seinem wirksamen Eintreten erkannt, bzw. durch Gegenmaßnahmen können Folgen noch abgefangen werden",
     "Termin- oder Materialkonflikt gerade noch rechtzeitig erkannt und behoben"),
    ("A", "unsichere Handlung",
     "es wird unsauber gearbeitet, aber es entsteht kein Schaden",
     "Vorgaben werden nicht eingehalten, z.B. Eintrag der RTW Checks in FB 260"),
]

fills_kat = {
    "D": PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid"),
    "C": PatternFill(start_color="FFE0B2", end_color="FFE0B2", fill_type="solid"),
    "B": PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid"),
    "A": PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid"),
}

for i, (kat, desc, details, examples) in enumerate(kat_data):
    r = 2 + i
    fill = fills_kat[kat]
    set_cell(ws_kat, r, 1, kat, font=Font(name="Calibri", size=14, bold=True, color="333333"),
             fill=fill, alignment=center_align, border=thin_border)
    set_cell(ws_kat, r, 2, desc, font=BOLD_FONT, fill=fill, alignment=wrap_align, border=thin_border)
    set_cell(ws_kat, r, 3, details, font=NORMAL_FONT, fill=fill, alignment=wrap_align, border=thin_border)
    set_cell(ws_kat, r, 4, examples, font=NORMAL_FONT, fill=fill, alignment=wrap_align, border=thin_border)
    ws_kat.row_dimensions[r].height = 60

ws_kat.freeze_panes = "A2"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET: Hilfslisten
# ══════════════════════════════════════════════════════════════════════════════
ws_hilf = wb.create_sheet(title="Hilfslisten")
ws_hilf.sheet_properties.tabColor = "666666"

ws_hilf.column_dimensions["A"].width = 22
ws_hilf.column_dimensions["B"].width = 40
ws_hilf.column_dimensions["C"].width = 35
ws_hilf.column_dimensions["D"].width = 25

ws_hilf.merge_cells("A1:D1")
set_cell(ws_hilf, 1, 1, "Mitarbeiterliste & E-Mail-Adressen",
         font=Font(name="Calibri", size=14, bold=True, color="0F3460"),
         alignment=Alignment(horizontal="center", vertical="center"))
ws_hilf.row_dimensions[1].height = 30

for col, h in enumerate(["Name", "E-Mail", "Funktion/Bereich", "Notiz"], 1):
    set_cell(ws_hilf, 2, col, h, font=HEADER_FONT, fill=HEADER_BG, alignment=center_wrap, border=thin_border)

contacts = [
    ("Dr. Markus Frei", "markus.frei@mercedes-benz.com", "Arzt", ""),
    ("Dr. Sabine Schmidt", "sabine.m.schmidt@mercedes-benz.com", "Ärztin", ""),
    ("Dr. Lilla Vitez", "lilla.vitez@mercedes-benz.com", "Ärztin", ""),
    ("Bernd Breig", "bernd.breig@mercedes-benz.com", "Stv. Med.&Verbr.-stoff", ""),
    ("Elke Krempl", "elke.krempl@mercedes-benz.com", "Mitarbeiterin", ""),
    ("Emanuel Siebert", "emanuel.siebert@mercedes-benz.com", "Leiter EDV", ""),
    ("Walter Putschler", "walter.putschler@mercedes-benz.com", "Leiter Ablauf&Prozessprobleme", "BCC-Empfänger"),
    ("Benjamin Jochim", "benjamin.jochim@mercedes-benz.com", "Mitarbeiter", ""),
    ("Harry Zuber", "harry.zuber@mercedes-benz.com", "Mitarbeiter", ""),
    ("Markus Göpfrich", "Markus.Goepfrich@mercedes-benz.com", "Leiter RTW & MPG", "Name im Dropdown: Göpfrich"),
    ("Tobias Zeller", "tobias_felix.zeller@mercedes-benz.com", "Leiter Med.&Verbr.-stoff", ""),
    ("Emily Kim Schmidt", "emily_kim.schmidt@mercedes-benz.com", "Leiterin Sekretariat", ""),
    ("Katrin Zieger-Buchta", "katrin.zieger-buchta@mercedes-benz.com", "Leiterin BGF", ""),
    ("Susanne Müller-Horn", "Susanne.Mueller-Horn@mercedes-benz.com", "Stv. BGF", ""),
    ("Larissa Radimersky", "larissa.radimersky@mercedes-benz.com", "Stv. Sekretariat", ""),
    ("Fabian Wunsch", "Fabian.Wunsch@mercedes-benz.com", "Leiter WD-Fzg, Org. Ambulanz, Hygiene", ""),
]

for i, (name, email, role, note) in enumerate(contacts):
    r = 3 + i
    alt_fill = LIGHT_GRAY_BG if i % 2 == 0 else WHITE_BG
    set_cell(ws_hilf, r, 1, name, font=NORMAL_FONT, fill=alt_fill, border=thin_border)
    set_cell(ws_hilf, r, 2, email, font=NORMAL_FONT, fill=alt_fill, border=thin_border)
    set_cell(ws_hilf, r, 3, role, font=NORMAL_FONT, fill=alt_fill, border=thin_border, alignment=wrap_align)
    set_cell(ws_hilf, r, 4, note, font=NORMAL_FONT, fill=alt_fill, border=thin_border)

ws_hilf.freeze_panes = "A3"

# ── Dropdown-Listen (Spalte F: Namen, Spalte H: REKO) ────────────────────────
# These are used as data validation sources via cell references (avoids 255 char limit)
list_start_row = 22

set_cell(ws_hilf, list_start_row, 6, "Dropdown: Namen", font=BOLD_FONT)
sorted_names = sorted(set(NAMES))
for i, name in enumerate(sorted_names):
    set_cell(ws_hilf, list_start_row + 1 + i, 6, name, font=NORMAL_FONT)
NAMES_LAST_ROW = list_start_row + len(sorted_names)

set_cell(ws_hilf, list_start_row, 8, "Dropdown: REKO", font=BOLD_FONT)
for i, opt in enumerate(REKO_OPTIONS):
    set_cell(ws_hilf, list_start_row + 1 + i, 8, opt, font=NORMAL_FONT)
REKO_LAST_ROW = list_start_row + len(REKO_OPTIONS)

set_cell(ws_hilf, list_start_row, 10, "Dropdown: Kategorie", font=BOLD_FONT)
for i, kat in enumerate(["A", "B", "C", "D"]):
    set_cell(ws_hilf, list_start_row + 1 + i, 10, kat, font=NORMAL_FONT)
KAT_LAST_ROW = list_start_row + 4

ws_hilf.column_dimensions["F"].width = 22
ws_hilf.column_dimensions["H"].width = 30
ws_hilf.column_dimensions["J"].width = 18

# Store references for use in data validation (defined names)
from openpyxl.workbook.defined_name import DefinedName

# Named ranges for data validation
namen_ref = f"Hilfslisten!$F${list_start_row+1}:$F${NAMES_LAST_ROW}"
reko_ref = f"Hilfslisten!$H${list_start_row+1}:$H${REKO_LAST_ROW}"
kat_ref = f"Hilfslisten!$J${list_start_row+1}:$J${KAT_LAST_ROW}"

dn_namen = DefinedName("NamenListe", attr_text=namen_ref)
dn_reko = DefinedName("REKOListe", attr_text=reko_ref)
dn_kat = DefinedName("KatListe", attr_text=kat_ref)
wb.defined_names.add(dn_namen)
wb.defined_names.add(dn_reko)
wb.defined_names.add(dn_kat)


# ══════════════════════════════════════════════════════════════════════════════
# Save
# ══════════════════════════════════════════════════════════════════════════════
output_path = "/home/user/Ambulanzzeug/LARIS_2026.xlsx"
wb.save(output_path)
print(f"Saved: {output_path}")
print(f"Sheets: {wb.sheetnames}")
print("Done!")
