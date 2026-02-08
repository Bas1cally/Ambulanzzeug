#!/usr/bin/env python3
"""
Creates a fresh LARIS_2026.xlsx without macros, ActiveX, or sheet protection.
All dropdowns use native Excel data validation.
"""
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from copy import copy
import datetime

wb = openpyxl.Workbook()

# ── Colors & Styles ──────────────────────────────────────────────────────────
MERCEDES_DARK = "1A1A2E"
MERCEDES_ACCENT = "0F3460"
HEADER_BG = PatternFill(start_color="0F3460", end_color="0F3460", fill_type="solid")
HEADER_BG2 = PatternFill(start_color="16213E", end_color="16213E", fill_type="solid")
SUBHEADER_BG = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
LIGHT_GRAY_BG = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
WHITE_BG = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
GREEN_BG = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
YELLOW_BG = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
RED_BG = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
ORANGE_BG = PatternFill(start_color="FFE0B2", end_color="FFE0B2", fill_type="solid")
KATEGORIE_A_BG = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
KATEGORIE_B_BG = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
KATEGORIE_C_BG = PatternFill(start_color="FFE0B2", end_color="FFE0B2", fill_type="solid")
KATEGORIE_D_BG = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")

TITLE_FONT = Font(name="Calibri", size=18, bold=True, color="0F3460")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
SUBHEADER_FONT = Font(name="Calibri", size=10, bold=True, color="333333")
NORMAL_FONT = Font(name="Calibri", size=10, color="333333")
BOLD_FONT = Font(name="Calibri", size=10, bold=True, color="333333")
LINK_FONT = Font(name="Calibri", size=10, color="0F3460", underline="single")
BULLET_FONT = Font(name="Calibri", size=10, bold=True, color="0F3460")
HINT_FONT = Font(name="Calibri", size=9, italic=True, color="888888")
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
    "Radimersky", "Wunsch", "Goepfrich"
]

KATEGORIEN = ["A", "B", "C", "D"]

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
    "Besprochen mit  Zeller",
    "Besprochen mit  Schmidt",
    "Besprochen mit  Zieger-Buchta",
    "Besprochen mit  Müller-Horn",
    "Besprochen mit  Radimersky",
    "Besprochen mit  Wunsch",
]

# ── Bereiche (categories) with their leaders and existing data ───────────────
BEREICHE = [
    {
        "name": "Sekretariat",
        "leiter": "Emily Schmidt",
        "leiter_email": "emily_kim.schmidt@mercedes-benz.com",
        "stellvertreter": "Larissa Radimersky",
        "data": []
    },
    {
        "name": "EDV",
        "leiter": "Siebert",
        "leiter_email": "emanuel.siebert@mercedes-benz.com",
        "stellvertreter": "",
        "data": [
            {
                "meldung_am": datetime.datetime(2025, 8, 14, 15, 52, 22),
                "meldung_durch": "Dr.Frei",
                "thema": "Laris geht nicht richtig und Meldeliste muss aktualisiert werden und Mailbutton geht nicht und es zeigt die Meldung der Active-X Inhalt ist blockiert",
                "kenntnisnahme_am": datetime.datetime(2025, 8, 15, 10, 53, 31),
                "kenntnisnahme_durch": "Zuber",
                "massnahme": "Laris war nie kaputt, die Excel Anwendung muss für die Benutzung eingestellt werden im Bereich Active-X Inhalte",
                "erledigt_am": datetime.datetime(2025, 8, 15, 10, 53, 43),
                "erledigt_durch": "Zuber",
                "kategorie": "C",
                "info": "",
                "reko": "Nicht erforderlich"
            },
            {
                "meldung_am": datetime.datetime(2025, 9, 26, 8, 26, 9),
                "meldung_durch": "Dr.Frei",
                "thema": "Das Blatt Organisation Ambulanz in Laris hat einen Schreibschutz, so dass nichts eingetragen werden kann - bitte diesen entfernen",
                "kenntnisnahme_am": datetime.datetime(2025, 9, 26, 11, 2, 0),
                "kenntnisnahme_durch": "Siebert",
                "massnahme": "Rechtsklick auf die Kachel -> Schutz Pausieren",
                "erledigt_am": datetime.datetime(2025, 9, 26, 11, 2, 0),
                "erledigt_durch": "Siebert",
                "kategorie": "A",
                "info": "",
                "reko": "REKO"
            },
            {
                "meldung_am": datetime.datetime(2026, 1, 23, 9, 31, 45),
                "meldung_durch": "Dr.Schmidt",
                "thema": "Reko Protokoll funktioniert, kein stabiler Prozess",
                "kenntnisnahme_am": datetime.datetime(2026, 1, 27, 8, 0, 0),
                "kenntnisnahme_durch": "Siebert",
                "massnahme": "Speicher Geupdated, der Indexer muss noch ein Fehler behoben werden, in arbeit.",
                "erledigt_am": None,
                "erledigt_durch": "",
                "kategorie": "",
                "info": "",
                "reko": ""
            },
        ]
    },
    {
        "name": "Med.&Verbr.-stoff",
        "leiter": "Zeller",
        "leiter_email": "tobias_felix.zeller@mercedes-benz.com",
        "stellvertreter": "Breig",
        "data": [
            {
                "meldung_am": datetime.datetime(2025, 9, 1, 11, 18, 52),
                "meldung_durch": "Dr.Schmidt",
                "thema": "Antibiotika sind teilweise abgelaufen, nicht geordnet gelagert",
                "kenntnisnahme_am": None,
                "kenntnisnahme_durch": "",
                "massnahme": "Es wurde ein Verfalls Tabelle angelegt an der Innenseite des Schranks die Monatlich kontrolliert wird",
                "erledigt_am": datetime.datetime(2025, 9, 26, 10, 55, 15),
                "erledigt_durch": "Siebert",
                "kategorie": "C",
                "info": "",
                "reko": "Besprochen mit  Dr.Schmidt"
            },
            {
                "meldung_am": datetime.datetime(2025, 12, 15, 16, 41, 53),
                "meldung_durch": "Dr. Vitez",
                "thema": "Hepatitis A-Impfstoff (Havrix) in November abgelaufen",
                "kenntnisnahme_am": None,
                "kenntnisnahme_durch": "",
                "massnahme": "",
                "erledigt_am": None,
                "erledigt_durch": "",
                "kategorie": "C",
                "info": "",
                "reko": ""
            },
            {
                "meldung_am": datetime.datetime(2026, 1, 23, 9, 34, 45),
                "meldung_durch": "Dr.Schmidt",
                "thema": "Einträge (siehe 15.12.25) in diesem Bereich werden nicht rückgemeldet bzw. abgearbeitet",
                "kenntnisnahme_am": None,
                "kenntnisnahme_durch": "",
                "massnahme": "",
                "erledigt_am": None,
                "erledigt_durch": "",
                "kategorie": "",
                "info": "",
                "reko": ""
            },
        ]
    },
    {
        "name": "WD-Fzg & Schutzkleidg",
        "leiter": "Wunsch",
        "leiter_email": "Fabian.Wunsch@mercedes-benz.com",
        "stellvertreter": "",
        "data": [
            {
                "meldung_am": datetime.datetime(2025, 11, 13, 9, 5, 54),
                "meldung_durch": "Zuber",
                "thema": "Barduschwäsche nur Herren raus gestellt, Damen volle Kiste stehen lassen",
                "kenntnisnahme_am": None,
                "kenntnisnahme_durch": "",
                "massnahme": "",
                "erledigt_am": None,
                "erledigt_durch": "",
                "kategorie": "",
                "info": "",
                "reko": ""
            },
            {
                "meldung_am": datetime.datetime(2026, 2, 3, 8, 0, 0),
                "meldung_durch": "Putschler",
                "thema": "Test",
                "kenntnisnahme_am": None,
                "kenntnisnahme_durch": "",
                "massnahme": "Test",
                "erledigt_am": None,
                "erledigt_durch": "",
                "kategorie": "",
                "info": "",
                "reko": ""
            },
        ]
    },
    {
        "name": "RTW & MPG",
        "leiter": "Goepfrich",
        "leiter_email": "Markus.Goepfrich@mercedes-benz.com",
        "stellvertreter": "",
        "data": [
            {
                "meldung_am": datetime.datetime(2025, 7, 3, 4, 50, 44),
                "meldung_durch": "Breig",
                "thema": "OP-Leuchte in Ambulanz 1 defekt",
                "kenntnisnahme_am": datetime.datetime(2025, 8, 20, 17, 13, 22),
                "kenntnisnahme_durch": "Goepfrich",
                "massnahme": "Reperatur durch Reposition der Halterung (Dr. Frei)",
                "erledigt_am": datetime.datetime(2025, 8, 20, 17, 13, 13),
                "erledigt_durch": "Goepfrich",
                "kategorie": "A",
                "info": "",
                "reko": "Besprochen mit  Dr.Frei"
            },
            {
                "meldung_am": datetime.datetime(2025, 9, 10, 21, 56, 58),
                "meldung_durch": "Wunsch",
                "thema": "RR Manschette C3 aus dem RTW defekt, Leitung abgelöst durch MA der WF",
                "kenntnisnahme_am": datetime.datetime(2025, 9, 13, 6, 14, 30),
                "kenntnisnahme_durch": "Goepfrich",
                "massnahme": "Neuteil am 24.09. erhalten u. eingeräumt. Zwischenzeitlich Nutzung d. RR-Manschette C3 KU.",
                "erledigt_am": datetime.datetime(2025, 9, 24, 11, 24, 41),
                "erledigt_durch": "Goepfrich",
                "kategorie": "A",
                "info": "",
                "reko": "Besprochen mit  Dr.Schmidt"
            },
            {
                "meldung_am": datetime.datetime(2025, 9, 12, 4, 7, 28),
                "meldung_durch": "Zuber",
                "thema": "Fehlendes Sperrlager für defekte Medizinprodukte, Defekte RR Manschette vom 10.09. kann nicht entsprechend hinterlegt werden. Bitte Sperrlager definieren und kommunizieren.",
                "kenntnisnahme_am": datetime.datetime(2025, 9, 13, 6, 14, 32),
                "kenntnisnahme_durch": "Goepfrich",
                "massnahme": "Einrichtung eines Sperrlagers für defekte Medizinprodukte u. -geräte in Physikalische Therapie im 029/10 Schrank 2.",
                "erledigt_am": datetime.datetime(2025, 10, 24, 21, 17, 35),
                "erledigt_durch": "Goepfrich",
                "kategorie": "A",
                "info": "",
                "reko": "Besprochen mit  Zuber"
            },
            {
                "meldung_am": datetime.datetime(2025, 11, 6),
                "meldung_durch": "Krempl",
                "thema": "Cryo-Thermal Gerät im WD Kuppenheim lässt sich nicht hochfahren, Bildschirm bleibt schwarz",
                "kenntnisnahme_am": datetime.datetime(2025, 12, 4, 15, 23, 56),
                "kenntnisnahme_durch": "Goepfrich",
                "massnahme": "Reparatur durch Hersteller",
                "erledigt_am": datetime.datetime(2026, 1, 28, 4, 11, 48),
                "erledigt_durch": "Goepfrich",
                "kategorie": "A",
                "info": "",
                "reko": "Besprochen mit  Dr.Frei"
            },
            {
                "meldung_am": datetime.datetime(2025, 12, 4, 15, 24, 23),
                "meldung_durch": "Göpfrich",
                "thema": "Untersuchungsleuchte Ambulanz 1 defekt, Befestigung Lampenkopf gebrochen",
                "kenntnisnahme_am": datetime.datetime(2025, 12, 4, 15, 24, 46),
                "kenntnisnahme_durch": "Goepfrich",
                "massnahme": "",
                "erledigt_am": None,
                "erledigt_durch": "",
                "kategorie": "",
                "info": "",
                "reko": ""
            },
        ]
    },
    {
        "name": "Organisation Ambulanz",
        "leiter": "Wunsch",
        "leiter_email": "Fabian.Wunsch@mercedes-benz.com",
        "stellvertreter": "",
        "data": [
            {
                "meldung_am": datetime.datetime(2025, 8, 19, 10, 51, 38),
                "meldung_durch": "Putschler",
                "thema": "EDTA Röhrchen vom 18.8. (K'heim) wurden in Rastatt nach dem Zentrifugieren in den Kühlschrank gelegt",
                "kenntnisnahme_am": datetime.datetime(2025, 8, 19, 10, 58, 22),
                "kenntnisnahme_durch": "Putschler",
                "massnahme": "Rücksprache mit Labor; neue Verfahrensanweisung verschickt und visualisiert",
                "erledigt_am": datetime.datetime(2025, 8, 19, 11, 2, 26),
                "erledigt_durch": "Putschler",
                "kategorie": "A",
                "info": "",
                "reko": ""
            },
        ]
    },
    {
        "name": "Ablauf&Prozessprobleme",
        "leiter": "Putschler",
        "leiter_email": "walter.putschler@mercedes-benz.com",
        "stellvertreter": "",
        "data": [
            {
                "meldung_am": datetime.datetime(2025, 9, 1, 11, 7, 34),
                "meldung_durch": "Dr.Schmidt",
                "thema": "Laris Eintrag für Bereich bei EDV nicht möglich, obwohl Einträge in Bereichen zB Sekretariat/Mail und an Bereichsleiter möglich sind",
                "kenntnisnahme_am": datetime.datetime(2025, 9, 22),
                "kenntnisnahme_durch": "Putschler",
                "massnahme": "Info an Harry Zuber--> Bitte Fehler prüfen und beheben",
                "erledigt_am": None,
                "erledigt_durch": "",
                "kategorie": "A",
                "info": "",
                "reko": ""
            },
            {
                "meldung_am": datetime.datetime(2025, 9, 1, 11, 18, 10),
                "meldung_durch": "Dr.Schmidt",
                "thema": "Bereich Med. & Verbr.: Antibiotika sind teilweise abgelaufen, nicht geordnet gelagert; hier Laris Eintrag nicht möglich, obwohl Einträge in Bereichen zB Sekretariat/Mail und an Bereichsleiter möglich sind",
                "kenntnisnahme_am": datetime.datetime(2025, 9, 2),
                "kenntnisnahme_durch": "Putschler",
                "massnahme": "Bereichverantwortlicher hat Fehler abgestellt",
                "erledigt_am": datetime.datetime(2025, 9, 4),
                "erledigt_durch": "Siebert",
                "kategorie": "C",
                "info": "",
                "reko": ""
            },
            {
                "meldung_am": datetime.datetime(2025, 9, 23, 8, 54, 12),
                "meldung_durch": "Dr.Schmidt",
                "thema": "Dokument für Einverständniserklärung Blutabnahme wurden in falsche Akte gescannt",
                "kenntnisnahme_am": datetime.datetime(2025, 9, 23),
                "kenntnisnahme_durch": "Putschler",
                "massnahme": "Fehler durch Sekretariat erkannt und behoben",
                "erledigt_am": datetime.datetime(2025, 9, 23),
                "erledigt_durch": "Radimersky",
                "kategorie": "D",
                "info": "",
                "reko": ""
            },
            {
                "meldung_am": datetime.datetime(2025, 9, 23, 15, 56, 5),
                "meldung_durch": "Krempl",
                "thema": "Glucoseröhrchen wurden über Nacht nicht gekühlt, dafür aber nicht die EDTA-Röhrchen die am Vortag aus Kuppenheim gebracht wurden.",
                "kenntnisnahme_am": datetime.datetime(2025, 9, 24),
                "kenntnisnahme_durch": "Putschler",
                "massnahme": "Laborergebnis wird abgewartet; bei Abweichung werden die betreffenden Patienten erneut einbestellt",
                "erledigt_am": datetime.datetime(2025, 9, 24),
                "erledigt_durch": "Putschler",
                "kategorie": "C",
                "info": "Mit Mitarbeiter persönlich besprochen",
                "reko": ""
            },
            {
                "meldung_am": datetime.datetime(2025, 9, 26, 8, 14, 40),
                "meldung_durch": "Putschler",
                "thema": "Serumröhrchen von drei Patienten aus Kuppenheim wurden am Vortag nicht aus der Zentrifuge genommen und waren 16 Stunden ungekühlt",
                "kenntnisnahme_am": datetime.datetime(2025, 9, 26, 8, 14, 21),
                "kenntnisnahme_durch": "Putschler",
                "massnahme": "Serumröhrchen in den Kühlschrank getan; Rücksprache mit Arzt ob noch zu verwenden",
                "erledigt_am": None,
                "erledigt_durch": "",
                "kategorie": "C",
                "info": "",
                "reko": ""
            },
            {
                "meldung_am": datetime.datetime(2025, 10, 20, 7, 9, 16),
                "meldung_durch": "Zuber",
                "thema": "Testmeldung aus Qualitätsgründen",
                "kenntnisnahme_am": None,
                "kenntnisnahme_durch": "",
                "massnahme": "",
                "erledigt_am": None,
                "erledigt_durch": "",
                "kategorie": "",
                "info": "",
                "reko": ""
            },
        ]
    },
    {
        "name": "BGF",
        "leiter": "Zieger-Buchta",
        "leiter_email": "katrin.zieger-buchta@mercedes-benz.com",
        "stellvertreter": "Müller-Horn",
        "data": []
    },
    {
        "name": "Getriebewerk",
        "leiter": "Dr.Frei",
        "leiter_email": "markus.frei@mercedes-benz.com",
        "stellvertreter": "Dr.Schmidt",
        "data": []
    },
    {
        "name": "Hygiene",
        "leiter": "Wunsch",
        "leiter_email": "Fabian.Wunsch@mercedes-benz.com",
        "stellvertreter": "",
        "data": [
            {
                "meldung_am": datetime.datetime(2026, 2, 2),
                "meldung_durch": "Wunsch",
                "thema": "Hygiene Physikalische + Ruheraum, in letzter Zeit öfters schmutzige Liegen & Schwämme noch am Gerät",
                "kenntnisnahme_am": None,
                "kenntnisnahme_durch": "",
                "massnahme": "",
                "erledigt_am": None,
                "erledigt_durch": "",
                "kategorie": "A",
                "info": "",
                "reko": ""
            },
        ]
    },
]

# ── Helper to format datetime for display ────────────────────────────────────
DATE_FMT = "DD.MM.YYYY HH:MM"


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


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 1: Info
# ══════════════════════════════════════════════════════════════════════════════
ws_info = wb.active
ws_info.title = "Info"
ws_info.sheet_properties.tabColor = "0F3460"

# Column widths
ws_info.column_dimensions["A"].width = 3
ws_info.column_dimensions["B"].width = 4
ws_info.column_dimensions["C"].width = 55
ws_info.column_dimensions["D"].width = 20
ws_info.column_dimensions["E"].width = 5
ws_info.column_dimensions["F"].width = 20
ws_info.column_dimensions["G"].width = 5

# Title
ws_info.merge_cells("B1:F2")
set_cell(ws_info, 1, 2, "LARIS - Lernende Abweichungs-, Risiko- und Informationssammlung",
         font=TITLE_FONT, alignment=Alignment(horizontal="center", vertical="center"))

# Subtitle
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

items = ["Vorkommnissen", "fehlendem oder fehlerhaftem Material", "Administrationsbedarf im Bereich"]
for i, item in enumerate(items):
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

# ── Section: Erstellen einer Meldung ─────────────────────────────────────────
row = 14
ws_info.merge_cells(f"B{row}:F{row}")
set_cell(ws_info, row, 2, "Erstellen einer Meldung", font=INFO_TITLE_FONT)

instructions_create = [
    'Unten den entsprechenden Reiter für den Bereich auswählen',
    'Im Feld "Meldung durch" den eigenen Namen auswählen',
    'Datum im Feld "Meldung am" eintragen (Format: TT.MM.JJJJ)',
    'Info bei "Thema" eintragen',
    'Speichern',
]
for i, txt in enumerate(instructions_create):
    r = row + 1 + i
    set_cell(ws_info, r, 2, "●", font=BULLET_FONT)
    ws_info.merge_cells(f"C{r}:F{r}")
    set_cell(ws_info, r, 3, txt, font=NORMAL_FONT)

# ── Section: Abarbeiten ──────────────────────────────────────────────────────
row = 21
ws_info.merge_cells(f"B{row}:F{row}")
set_cell(ws_info, row, 2, "Abarbeiten einer Meldung durch den Bereichsleiter", font=INFO_TITLE_FONT)

instructions_process = [
    'Bei Kenntnisnahme im Feld "Kenntnisnahme durch" eigenen Namen auswählen und Datum eintragen',
    'Massnahmen im entsprechenden Feld dokumentieren',
    'Nach Erledigung im Feld "Erledigt durch" eigenen Namen auswählen und Datum eintragen',
    'Evtl. weitere Infos im entsprechenden Feld eintragen',
    'In Spalte "In REKO..." auswählen wie besprochen wurde',
]
for i, txt in enumerate(instructions_process):
    r = row + 1 + i
    set_cell(ws_info, r, 2, "●", font=BULLET_FONT)
    ws_info.merge_cells(f"C{r}:F{r}")
    set_cell(ws_info, r, 3, txt, font=NORMAL_FONT)

# ── Section: Hinweise ────────────────────────────────────────────────────────
row = 28
ws_info.merge_cells(f"B{row}:F{row}")
set_cell(ws_info, row, 2, "Hinweise zur neuen Version", font=INFO_TITLE_FONT)

hints = [
    "Kein Blattschutz - alle Zellen sind frei editierbar",
    "Keine Makros/ActiveX - kein Hängen, keine Sicherheitswarnungen",
    "Dropdown-Listen für Namen und Kategorien (Klick auf die Zelle)",
    "Datumsfelder manuell ausfüllen (Format: TT.MM.JJJJ HH:MM)",
    "Meldungs- und Erledigungszähler werden automatisch berechnet (Formeln)",
]
for i, txt in enumerate(hints):
    r = row + 1 + i
    set_cell(ws_info, r, 2, "✓", font=Font(name="Calibri", size=10, bold=True, color="28A745"))
    ws_info.merge_cells(f"C{r}:F{r}")
    set_cell(ws_info, r, 3, txt, font=NORMAL_FONT)

# ── Section: Bereichstabelle ─────────────────────────────────────────────────
row = 35
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

for bereich in BEREICHE:
    ws = wb.create_sheet(title=bereich["name"])
    ws.sheet_properties.tabColor = "0F3460"

    # Column widths
    for i, (_, width) in enumerate(HEADERS, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # ── Row 1: Bereich title ─────────────────────────────────────────────────
    ws.merge_cells("A1:L1")
    set_cell(ws, 1, 1, bereich["name"],
             font=Font(name="Calibri", size=16, bold=True, color="FFFFFF"),
             fill=PatternFill(start_color="0F3460", end_color="0F3460", fill_type="solid"),
             alignment=Alignment(horizontal="left", vertical="center"))
    ws.row_dimensions[1].height = 35

    # ── Row 2: Counters + Kategorie-Legende ──────────────────────────────────
    ws.row_dimensions[2].height = 25
    set_cell(ws, 2, 1, "Meldungen:", font=BOLD_FONT, alignment=Alignment(horizontal="right", vertical="center"))
    # COUNTIF formula for meldungen count (count non-empty cells in column B starting row 5)
    set_cell(ws, 2, 2, None, font=COUNTER_FONT, alignment=center_align)
    ws["B2"] = '=COUNTA(B5:B504)'

    set_cell(ws, 2, 3, "Erledigt:", font=BOLD_FONT, alignment=Alignment(horizontal="right", vertical="center"))
    set_cell(ws, 2, 4, None, font=Font(name="Calibri", size=14, bold=True, color="28A745"), alignment=center_align)
    ws["D2"] = '=COUNTA(H5:H504)'

    set_cell(ws, 2, 5, "Offen:", font=BOLD_FONT, alignment=Alignment(horizontal="right", vertical="center"))
    set_cell(ws, 2, 6, None, font=Font(name="Calibri", size=14, bold=True, color="DC3545"), alignment=center_align)
    ws["F2"] = '=B2-D2'

    # Kategorie legend
    ws.merge_cells("H2:L3")
    set_cell(ws, 2, 8, KATEGORIE_LEGEND,
             font=Font(name="Calibri", size=9, color="666666"),
             alignment=Alignment(wrap_text=True, vertical="top"),
             fill=PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid"))

    # ── Row 4: Headers ───────────────────────────────────────────────────────
    ws.row_dimensions[4].height = 32
    for col_idx, (header_text, _) in enumerate(HEADERS, 1):
        set_cell(ws, 4, col_idx, header_text,
                 font=HEADER_FONT, fill=HEADER_BG, alignment=center_wrap, border=thin_border)

    # ── Data Validations ─────────────────────────────────────────────────────
    # Names dropdown for columns C (Meldung durch), F (Kenntnisnahme durch), I (Erledigt durch)
    names_str = ",".join(sorted(set(NAMES)))
    dv_names = DataValidation(type="list", formula1=f'"{names_str}"', allow_blank=True)
    dv_names.error = "Bitte einen Namen aus der Liste wählen"
    dv_names.errorTitle = "Ungültiger Name"
    dv_names.prompt = "Name auswählen"
    dv_names.promptTitle = "Mitarbeiter"
    ws.add_data_validation(dv_names)
    dv_names.add("C5:C504")
    dv_names.add("F5:F504")
    dv_names.add("I5:I504")

    # Kategorie dropdown for column J
    dv_kat = DataValidation(type="list", formula1='"A,B,C,D"', allow_blank=True)
    dv_kat.error = "Bitte A, B, C oder D wählen"
    dv_kat.errorTitle = "Ungültige Kategorie"
    dv_kat.prompt = "Kategorie wählen (A-D)"
    dv_kat.promptTitle = "Fehlerkategorie"
    ws.add_data_validation(dv_kat)
    dv_kat.add("J5:J504")

    # REKO dropdown for column L
    reko_str = ",".join(REKO_OPTIONS)
    dv_reko = DataValidation(type="list", formula1=f'"{reko_str}"', allow_blank=True)
    dv_reko.error = "Bitte aus der Liste wählen"
    dv_reko.errorTitle = "Ungültige Auswahl"
    dv_reko.prompt = "Besprechungsart wählen"
    dv_reko.promptTitle = "REKO / Besprochen"
    ws.add_data_validation(dv_reko)
    dv_reko.add("L5:L504")

    # ── Existing Data ────────────────────────────────────────────────────────
    for row_idx, entry in enumerate(bereich["data"]):
        r = 5 + row_idx
        alt_fill = LIGHT_GRAY_BG if row_idx % 2 == 0 else WHITE_BG

        # Nr.
        set_cell(ws, r, 1, row_idx + 1, font=NORMAL_FONT, fill=alt_fill,
                 alignment=center_align, border=thin_border)

        # Meldung am
        set_cell(ws, r, 2, entry["meldung_am"], font=NORMAL_FONT, fill=alt_fill,
                 alignment=center_align, border=thin_border,
                 number_format="DD.MM.YYYY HH:MM")

        # Meldung durch
        set_cell(ws, r, 3, entry["meldung_durch"], font=NORMAL_FONT, fill=alt_fill,
                 alignment=center_align, border=thin_border)

        # Thema
        set_cell(ws, r, 4, entry["thema"], font=NORMAL_FONT, fill=alt_fill,
                 alignment=wrap_align, border=thin_border)

        # Kenntnisnahme am
        set_cell(ws, r, 5, entry["kenntnisnahme_am"], font=NORMAL_FONT, fill=alt_fill,
                 alignment=center_align, border=thin_border,
                 number_format="DD.MM.YYYY HH:MM")

        # Kenntnisnahme durch
        set_cell(ws, r, 6, entry["kenntnisnahme_durch"], font=NORMAL_FONT, fill=alt_fill,
                 alignment=center_align, border=thin_border)

        # Massnahme
        set_cell(ws, r, 7, entry["massnahme"], font=NORMAL_FONT, fill=alt_fill,
                 alignment=wrap_align, border=thin_border)

        # Erledigt am
        set_cell(ws, r, 8, entry["erledigt_am"], font=NORMAL_FONT, fill=alt_fill,
                 alignment=center_align, border=thin_border,
                 number_format="DD.MM.YYYY HH:MM")

        # Erledigt durch
        set_cell(ws, r, 9, entry["erledigt_durch"], font=NORMAL_FONT, fill=alt_fill,
                 alignment=center_align, border=thin_border)

        # Kategorie
        kat = entry.get("kategorie", "")
        kat_fill = alt_fill
        if kat == "A":
            kat_fill = KATEGORIE_A_BG
        elif kat == "B":
            kat_fill = KATEGORIE_B_BG
        elif kat == "C":
            kat_fill = KATEGORIE_C_BG
        elif kat == "D":
            kat_fill = KATEGORIE_D_BG
        set_cell(ws, r, 10, kat, font=BOLD_FONT, fill=kat_fill,
                 alignment=center_align, border=thin_border)

        # Info
        set_cell(ws, r, 11, entry.get("info", ""), font=NORMAL_FONT, fill=alt_fill,
                 alignment=wrap_align, border=thin_border)

        # REKO
        set_cell(ws, r, 12, entry.get("reko", ""), font=NORMAL_FONT, fill=alt_fill,
                 alignment=wrap_align, border=thin_border)

        # Set row height for readability
        ws.row_dimensions[r].height = max(30, 15 * (1 + len(entry.get("thema", "")) // 70))

    # ── Pre-format empty rows with Nr. formulas and light borders ────────────
    next_data_row = 5 + len(bereich["data"])
    for r in range(next_data_row, next_data_row + 50):
        alt_fill = LIGHT_GRAY_BG if (r - 5) % 2 == 0 else WHITE_BG
        # Nr. formula
        set_cell(ws, r, 1, f'=IF(B{r}<>"",ROW()-4,"")', font=NORMAL_FONT, fill=alt_fill,
                 alignment=center_align, border=thin_border)
        # Light borders on all columns
        for col in range(2, 13):
            set_cell(ws, r, col, None, font=NORMAL_FONT, fill=alt_fill, border=thin_border)
            if col in (2, 5, 8):  # Date columns
                ws.cell(row=r, column=col).number_format = "DD.MM.YYYY HH:MM"
                ws.cell(row=r, column=col).alignment = center_align
            elif col in (3, 6, 9, 10):  # Name/Kat columns
                ws.cell(row=r, column=col).alignment = center_align
            else:
                ws.cell(row=r, column=col).alignment = wrap_align

    # Freeze panes at row 5
    ws.freeze_panes = "A5"

    # Auto-filter
    ws.auto_filter.ref = f"A4:L{next_data_row + 49}"


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

fills_kat = {"D": KATEGORIE_D_BG, "C": KATEGORIE_C_BG, "B": KATEGORIE_B_BG, "A": KATEGORIE_A_BG}

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
# SHEET: Hilfslisten (Kontakte & Bereiche)
# ══════════════════════════════════════════════════════════════════════════════
ws_hilf = wb.create_sheet(title="Hilfslisten")
ws_hilf.sheet_properties.tabColor = "666666"

ws_hilf.column_dimensions["A"].width = 18
ws_hilf.column_dimensions["B"].width = 40
ws_hilf.column_dimensions["C"].width = 20
ws_hilf.column_dimensions["D"].width = 40

# Title
ws_hilf.merge_cells("A1:D1")
set_cell(ws_hilf, 1, 1, "Mitarbeiterliste & E-Mail-Adressen",
         font=Font(name="Calibri", size=14, bold=True, color="0F3460"),
         alignment=Alignment(horizontal="center", vertical="center"))
ws_hilf.row_dimensions[1].height = 30

headers_hilf = ["Name", "E-Mail", "Funktion/Bereich", "Notiz"]
for col, h in enumerate(headers_hilf, 1):
    set_cell(ws_hilf, 2, col, h, font=HEADER_FONT, fill=HEADER_BG, alignment=center_wrap, border=thin_border)

contacts = [
    ("Dr. Markus Frei", "markus.frei@mercedes-benz.com", "Arzt / Leiter Getriebewerk", ""),
    ("Dr. Sabine Schmidt", "sabine.m.schmidt@mercedes-benz.com", "Ärztin", ""),
    ("Dr. Lilla Vitez", "lilla.vitez@mercedes-benz.com", "Ärztin", ""),
    ("Bernd Breig", "bernd.breig@mercedes-benz.com", "Stv. Med.&Verbr.-stoff", ""),
    ("Elke Krempl", "elke.krempl@mercedes-benz.com", "Mitarbeiterin", ""),
    ("Emanuel Siebert", "emanuel.siebert@mercedes-benz.com", "Leiter EDV", ""),
    ("Walter Putschler", "walter.putschler@mercedes-benz.com", "Leiter Ablauf&Prozessprobleme", "Blindkopie-Empfänger"),
    ("Benjamin Jochim", "benjamin.jochim@mercedes-benz.com", "Mitarbeiter", ""),
    ("Harry Zuber", "harry.zuber@mercedes-benz.com", "Mitarbeiter", ""),
    ("Markus Göpfrich", "Markus.Goepfrich@mercedes-benz.com", "Leiter RTW & MPG", ""),
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


# ══════════════════════════════════════════════════════════════════════════════
# Save
# ══════════════════════════════════════════════════════════════════════════════
output_path = "/home/user/Ambulanzzeug/LARIS_2026.xlsx"
wb.save(output_path)
print(f"Saved: {output_path}")
print(f"Sheets: {wb.sheetnames}")
print("Done!")
