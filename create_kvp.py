#!/usr/bin/env python3
"""
Creates a fresh KVP_2026.xlsx - Kontinuierlicher Verbesserungsprozess.
Modern UX matching LARIS_2026 design language.
Includes Dashboard, conditional formatting, data validation, and archived year data.
"""
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.page import PageMargins
import datetime
import urllib.parse

wb = openpyxl.Workbook()

# ── Colors & Styles (matching LARIS design) ──────────────────────────────────
HEADER_BG = PatternFill(start_color="0F3460", end_color="0F3460", fill_type="solid")
LIGHT_GRAY_BG = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
WHITE_BG = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
GREEN_BG = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
RED_BG = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
YELLOW_BG = PatternFill(start_color="FFF8E1", end_color="FFF8E1", fill_type="solid")
ARCHIVE_BG = PatternFill(start_color="E8EAF6", end_color="E8EAF6", fill_type="solid")

TITLE_FONT = Font(name="Calibri", size=18, bold=True, color="0F3460")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
NORMAL_FONT = Font(name="Calibri", size=10, color="333333")
BOLD_FONT = Font(name="Calibri", size=10, bold=True, color="333333")
COUNTER_FONT = Font(name="Calibri", size=14, bold=True, color="0F3460")

thin_border = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)

wrap_align = Alignment(wrap_text=True, vertical="top")
center_align = Alignment(horizontal="center", vertical="center")
center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)

# ── Team members (same as LARIS) ─────────────────────────────────────────────
NAMES = [
    "Dr.Frei", "Dr.Schmidt", "Dr. Vitez", "Breig", "Krempl",
    "Siebert", "Putschler", "Jochim", "Zuber", "Göpfrich",
    "Zeller", "Schmidt", "Zieger-Buchta", "Müller-Horn",
    "Radimersky", "Wunsch",
]

# ── KVP Column structure ─────────────────────────────────────────────────────
HEADERS = [
    ("Nr.", 5),
    ("Datum", 13),
    ("Thema", 60),
    ("Ersparnis (Zeit, Geld, Kappa, u.s.w.)", 30),
    ("Einbringender", 18),
    ("Kommentar Bereichsverantwortlicher", 30),
    ("Kommentar vom Gremium", 25),
    ("Angenommen/\nAbgelehnt", 14),
    ("Verantwortlich\nfür Umsetzung", 18),
]

MAX_DATA_ROW = 204

# ── Archived data from previous years ────────────────────────────────────────
DATA_2025 = [
    {"datum": "2025-02-04", "thema": "Übersicht über aktuelle Themen HR/PHB im Besprechungszimmer", "ersparnis": "schneller Überblick über TOP-Themen und Verantwortlichkeiten", "einbringender": "Zieger-Buchta", "kommentar_bv": "", "kommentar_gremium": "soll grober Übersicht dienen, Topthemen", "status": "Angenommen", "verantwortlich": "alle"},
    {"datum": "2025-02-14", "thema": "Neue Desinfektionsspender vorgeschlagen: sauber am Auslass, kompatibel mit bestehenden Handpumpen.", "ersparnis": "Weniger Verschmutzung, nur ein Pumpenmodell, integrierte Auffangschale.", "einbringender": "Wunsch", "kommentar_bv": "Kosten bitte noch aufführen auch für evtl. Umbau, gemeinsame Lösung mit \"Altsystem mit Auffangschale\" möglich?", "kommentar_gremium": "wie aufwändig ist der Tausch der Spender? Anzahl? Aufwand?", "status": "", "verantwortlich": ""},
    {"datum": "2025-02-26", "thema": "WD-Hotline (22828) auf SI-Seite vermerken", "ersparnis": "Spart hin und Her rennen zwischen div. Telefonen wenn Hotline mehr genutzt wird.", "einbringender": "Zuber", "kommentar_bv": "", "kommentar_gremium": "wollt Ihr immer erreichbar sein?", "status": "Abgelehnt", "verantwortlich": ""},
    {"datum": "2025-02-18", "thema": "Kleidungsübersicht erstellt (Jacken, Polos)", "ersparnis": "Einfachere Neubestellung etc.", "einbringender": "Wunsch", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": "Umgesetzt"},
    {"datum": "2025-02-18", "thema": "Poloshirts mit User-ID", "ersparnis": "Einfachere Einsortierung und Zuordnung", "einbringender": "Wunsch", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": "Umgesetzt"},
    {"datum": "2025-04-14", "thema": "Poloshirts abgemeldet, Übersichtsliste von Bardusch", "ersparnis": "Spart monatlich Geld und bringt Übersicht", "einbringender": "Wunsch", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": "Umgesetzt"},
    {"datum": "2025-03-14", "thema": "Abklärung mit der Feuerwehr zur Etikettierung der Jacken, damit diese gewaschen werden können.", "ersparnis": "Hygienischer und ins bestehende Abwurfsystem der Hosen integrierbar.", "einbringender": "Wunsch", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": "in Umsetzung"},
    {"datum": "2025-07-22", "thema": "Digitalisierung des Medikamentenlagers: Jede Kanban-Karte wird mit einem QR-Code versehen. Beim Scannen öffnet sich automatisch eine vorgefertigte E-Mail im Standard-Mailprogramm, die direkt an den zuständigen Bereichsverantwortlichen adressiert ist.", "ersparnis": "Zeit gespart, Kosten gesenkt durch Reduzierung der Bestände, Lagerfläche optimiert, Fehler reduziert und Prozesse transparenter gestaltet.", "einbringender": "Siebert", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": "Manuel"},
    {"datum": "2025-07-30", "thema": "Erreichbarkeit von Ambulanz-Handy in SI Seite auflisten", "ersparnis": "Verbesserte Erreichbarkeit", "einbringender": "Zuber", "kommentar_bv": "Sabine geht auf Harry für Zusatzinfo zu", "kommentar_gremium": "Siehe Eintrag Nr. 3", "status": "Abgelehnt", "verantwortlich": ""},
    {"datum": "2025-07-30", "thema": "Abrechnung der Taxifahrten zu D-Arzt nach Betriebsunfällen. Abrechnung mit BGHM sollte möglich sein.", "ersparnis": "Geld", "einbringender": "Zuber", "kommentar_bv": "https://www.bghm-magazin.de/ausgaben/05-2022/krankentransport-nach-arbeitsunfall", "kommentar_gremium": "Prozess ist möglich, jedoch zu aufwendig, wurde mehrfach als neg. Business case geprüft", "status": "Abgelehnt", "verantwortlich": ""},
    {"datum": "2025-08-07", "thema": "Statt Fentanyl 0,5 mg/10 ml, Fentanyl 0,1 mg/2 ml bestellen. Patientensicherheit, Kosten, Haltbarkeit.", "ersparnis": "Geld (kostet ca. die Hälfte)", "einbringender": "Dr. Vitez", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": "nächste BTM Bestellung, Dr. Schmidt"},
    {"datum": "2025-09-02", "thema": "Analog zu Punkt 8 soll ins digitale Kanban-System noch Folgendes aufgenommen werden: Kaffee, Milch, Spültabs, Entkalker", "ersparnis": "Zeit & Geld gespart (keine teuren Spontankäufe)", "einbringender": "Zieger-Buchta", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": "Manuel"},
    {"datum": "", "thema": "Verlagerung des Standortes Gasflaschen-Schrank (Sauerstoff, Stickstoff) in den Außenbereich der Ambulanz Rastatt. Eventuell mit Heizung für den Winter um Vereisung zu vermeiden.", "ersparnis": "Aufwand, mit Gewinn Sicherheit", "einbringender": "Zuber", "kommentar_bv": "Dr. Frei: Idee prinzipiell gut, aber nicht neu. Wurde in der Vergangenheit aus optischen Gründen von der FP abgelehnt; Abstand zum Gebäude von 5m erforderlich", "kommentar_gremium": "", "status": "Abgelehnt", "verantwortlich": ""},
    {"datum": "2025-09-22", "thema": "Neue Türschilder für Eingangstür, und Sonstige Verwendung", "ersparnis": "Aushänge können schnell getauscht werden, sieht sehr professionell aus", "einbringender": "Wunsch", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": "Fabian"},
    {"datum": "2025-11-27", "thema": "Verschiedene Defibrillationselektroden für Corpuls3 sind im Einsatz schwer zu unterscheiden. Einheitliche Elektroden reduzieren Verwechslungsgefahr.", "ersparnis": "Geld, Zeit, höhere Anwendersicherheit", "einbringender": "Göpfrich", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": "Markus G."},
    {"datum": "2025-12-09", "thema": "Postkarten erstellen als Werbung für die Qualifizierungen", "ersparnis": "die richtigen Personen in die richtige Quali direkt einsteuern", "einbringender": "Zieger-Buchta", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": "Lara"},
    {"datum": "2026-02-06", "thema": "Ersatzbeschaffung für defektes Urisys 1100 geplant. Prüfung ergab: nur Netzteil defekt. Reparatur über eBay deutlich günstiger als Neubeschaffung.", "ersparnis": "Neues Netzteil günstiger als Ersatzbeschaffung Gerät, Zubehör bleibt nutzbar", "einbringender": "Göpfrich", "kommentar_bv": "", "kommentar_gremium": "", "status": "", "verantwortlich": ""},
]

DATA_2024 = [
    {"datum": "2024-02-09", "thema": "Einrichtung eines Übungsraumes zum Notfalltraining (Raum 0.033)", "ersparnis": "Bessere Übungsmöglichkeiten, weniger externer Aufwand", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": "Bernd, Walter"},
    {"datum": "2024-02-09", "thema": "Rückenschilder und Namensschilder passend für Softshelljacke", "ersparnis": "Einheitliches Erscheinungsbild", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": "Walter"},
    {"datum": "2024-03-22", "thema": "Umorganisation Räume im GZ", "ersparnis": "Bessere Raumnutzung", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": "Sanis"},
    {"datum": "2024-03-22", "thema": "Einheitliche Mützen für Winter", "ersparnis": "Einheitliches Erscheinungsbild", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": "Harry"},
    {"datum": "2024-03-22", "thema": "Handspiegel beschaffen für MA die AS-Brille abholen", "ersparnis": "Service für Mitarbeiter", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": "Larissa"},
    {"datum": "2024-03-22", "thema": "Einheitliche, aktuelle Fotos für WD MA im Social Intranet", "ersparnis": "Professionelles Auftreten", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "", "verantwortlich": ""},
    {"datum": "2024-05-22", "thema": "Wetterstation mit Luftdruck, Temperatur und Luftfeuchte in D1", "ersparnis": "Information für Einsatzplanung", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": ""},
    {"datum": "2024-08-15", "thema": "Täglicher RTW Check in digitale Checkliste umgewandelt", "ersparnis": "Zeit und Papier gespart", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": ""},
    {"datum": "2024-09-24", "thema": "RTW Alarme als Info in Teams automatisiert", "ersparnis": "Schnellere Information", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": ""},
    {"datum": "2024-10-24", "thema": "Diagnostic WC Mülleimer tauschen gegen Deckel mit Fußbedienung", "ersparnis": "Hygiene", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "", "verantwortlich": ""},
]

DATA_2023 = [
    {"datum": "2023-03-13", "thema": "Tablets/iPads für Anamnese-Bogen", "ersparnis": "Papierersparnis, Digitalisierung", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Abgelehnt", "verantwortlich": ""},
    {"datum": "2023-03-13", "thema": "Klingelknopf/Totmannmelder für Sani Nachtschicht", "ersparnis": "Sicherheit", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Abgelehnt", "verantwortlich": ""},
    {"datum": "2023-03-20", "thema": "Kühlschrank im Ruheraum abschalten", "ersparnis": "Strom", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": ""},
    {"datum": "2023-03-26", "thema": "Druckinfusionsmanschette als Einmalprodukt", "ersparnis": "Hygiene, Kosten", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": ""},
    {"datum": "2023-03-30", "thema": "Telefone in Ambulanzen neu strukturieren", "ersparnis": "Bessere Erreichbarkeit", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": ""},
    {"datum": "2023-03-31", "thema": "Große Corpuls3 Taschen anschaffen", "ersparnis": "Besserer Transport", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": ""},
    {"datum": "2023-03-31", "thema": "Therapie-Stammkabel-Verlängerung Corpuls3", "ersparnis": "Flexibilität", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Abgelehnt", "verantwortlich": ""},
    {"datum": "2023-05-03", "thema": "Beschriftung Notruftelefon erneuern", "ersparnis": "Klarheit", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": ""},
    {"datum": "2023-05-31", "thema": "Sichere Spritzenkennzeichnung nach ISO DIVI Standard", "ersparnis": "Patientensicherheit", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": ""},
    {"datum": "2023-07-07", "thema": "Briefkasten vor Arztzimmern", "ersparnis": "Organisation", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Abgelehnt", "verantwortlich": ""},
    {"datum": "2023-08-17", "thema": "USB-Umschalter für Dokumentenscanner", "ersparnis": "Effizienz", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": ""},
]

DATA_2022 = [
    {"datum": "2022-07-06", "thema": "Einwilligungserklärung ODIN archivieren - Datum auf Jahr 3000", "ersparnis": "Verwaltungsaufwand reduziert", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": ""},
    {"datum": "2022-07-20", "thema": "Nachtschichtersthelfer für EH-Auffrischungskurs", "ersparnis": "Bessere Ausbildung", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Abgelehnt", "verantwortlich": ""},
    {"datum": "2022-07-20", "thema": "Heimlich-Manöver-Weste für EH-Kurs", "ersparnis": "Besseres Training", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": ""},
    {"datum": "2022-07-20", "thema": "Reduktion der Endotrachealtuben auf Größen 8,5+7,5+6,5", "ersparnis": "Kosten, Übersicht", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": ""},
    {"datum": "2022-07-20", "thema": "Verzicht auf Endotrachealtuben/Larynxtuben für Kinder", "ersparnis": "Vereinfachung", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": ""},
    {"datum": "2022-07-20", "thema": "Notfalltasche mit Pneumothorax-Punktionsnadeln", "ersparnis": "Notfallbereitschaft", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": ""},
    {"datum": "2022-07-20", "thema": "Notfalltasche mit Quick Clot Gauze und NAR-S Rolled Gauze", "ersparnis": "Notfallbereitschaft", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": ""},
    {"datum": "", "thema": "Zahnrettungsboxen für RTW, Ambulanz, EH-Kurs", "ersparnis": "Notfallbereitschaft", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": ""},
    {"datum": "2022-10-18", "thema": "Halterung für Notfallrucksack am Notfallwagen", "ersparnis": "Schnellerer Zugriff", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": ""},
    {"datum": "2022-11-16", "thema": "Webcam in Kuppenheim für Teams-Patientenbeurteilung", "ersparnis": "Telemedizin", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Abgelehnt", "verantwortlich": ""},
    {"datum": "2022-11-28", "thema": "Einsparung der Filterpatronen für Kaffeemaschine", "ersparnis": "Geld", "einbringender": "Team", "kommentar_bv": "", "kommentar_gremium": "", "status": "Angenommen", "verantwortlich": ""},
]

ALL_YEARS = [
    ("2026", []),
    ("2025", DATA_2025),
    ("2024", DATA_2024),
    ("2023", DATA_2023),
    ("2022", DATA_2022),
]

TAB_COLORS_KVP = {
    "2026": "28A745",  # Green - active year
    "2025": "1976D2",  # Blue
    "2024": "6D4C41",  # Brown
    "2023": "9E9E9E",  # Gray
    "2022": "BDBDBD",  # Light gray
}


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


def setup_print(ws, title):
    """Configure print settings for a worksheet."""
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.6, bottom=0.6, header=0.3, footer=0.3)
    ws.oddHeader.center.text = f"KVP - {title}"
    ws.oddHeader.center.size = 10
    ws.oddFooter.left.text = "Werksärztlicher Dienst Mercedes-Benz"
    ws.oddFooter.right.text = "Seite &P von &N"
    ws.print_title_rows = "4:4"


def add_kvp_conditional_formatting(ws, max_row=204):
    """Add conditional formatting to a KVP year sheet."""
    data_range = f"A5:I{max_row}"

    # 1) Angenommen = whole row light green
    ws.conditional_formatting.add(data_range, FormulaRule(
        formula=[f'$H5="Angenommen"'],
        stopIfTrue=False,
        fill=PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")))

    # 2) Abgelehnt = whole row light red
    ws.conditional_formatting.add(data_range, FormulaRule(
        formula=[f'$H5="Abgelehnt"'],
        stopIfTrue=False,
        fill=PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")))

    # 3) Pending (has date but no status) = whole row light yellow
    ws.conditional_formatting.add(data_range, FormulaRule(
        formula=[f'AND($B5<>"",$H5="")'],
        stopIfTrue=False,
        fill=PatternFill(start_color="FFF8E1", end_color="FFF8E1", fill_type="solid")))

    # 4) Status cell bold color
    status_range = f"H5:H{max_row}"
    ws.conditional_formatting.add(status_range, CellIsRule(
        operator="equal", formula=['"Angenommen"'], stopIfTrue=False,
        font=Font(name="Calibri", size=10, bold=True, color="1B5E20")))
    ws.conditional_formatting.add(status_range, CellIsRule(
        operator="equal", formula=['"Abgelehnt"'], stopIfTrue=False,
        font=Font(name="Calibri", size=10, bold=True, color="B71C1C")))

    # 5) New entries (last 14 days) - blue Nr marker
    ws.conditional_formatting.add(f"A5:A{max_row}", FormulaRule(
        formula=[f'AND($B5<>"",(TODAY()-$B5)<=14)'],
        stopIfTrue=False,
        fill=PatternFill(start_color="BBDEFB", end_color="BBDEFB", fill_type="solid"),
        font=Font(bold=True, color="1565C0")))


def create_year_sheet(wb, year_name, data, is_active=False):
    """Create a KVP year sheet with data."""
    if is_active:
        ws = wb.active
        ws.title = year_name
    else:
        ws = wb.create_sheet(title=year_name)

    ws.sheet_properties.tabColor = TAB_COLORS_KVP.get(year_name, "999999")

    # Column widths
    for i, (_, width) in enumerate(HEADERS, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # ── Row 1: Title bar ──────────────────────────────────────────────────
    ws.merge_cells("A1:H1")
    set_cell(ws, 1, 1,
        f"KVP Themenliste - Werksärztlicher Dienst Rastatt/Kuppenheim - {year_name}",
        font=TITLE_FONT,
        fill=PatternFill(start_color="0F3460", end_color="0F3460", fill_type="solid"),
        alignment=Alignment(horizontal="left", vertical="center"))
    # Override font color to white on dark bg
    ws.cell(row=1, column=1).font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    ws.row_dimensions[1].height = 38

    # Dashboard link (col I)
    dash_link = set_cell(ws, 1, 9, "← Dashboard",
        font=Font(name="Calibri", size=10, bold=True, color="FFFFFF", underline="single"),
        fill=PatternFill(start_color="0F3460", end_color="0F3460", fill_type="solid"),
        alignment=Alignment(horizontal="center", vertical="center"))
    dash_link.hyperlink = "#Dashboard!A1"

    # ── Row 2: Counters ───────────────────────────────────────────────────
    ws.row_dimensions[2].height = 30
    set_cell(ws, 2, 1, "Vorschläge:", font=BOLD_FONT, alignment=Alignment(horizontal="right", vertical="center"))
    ws["B2"] = f'=COUNTA(B5:B{MAX_DATA_ROW})'
    set_cell(ws, 2, 2, None, font=COUNTER_FONT, alignment=center_align)

    set_cell(ws, 2, 3, "Angenommen:", font=BOLD_FONT, alignment=Alignment(horizontal="right", vertical="center"))
    ws["D2"] = f'=COUNTIF(H5:H{MAX_DATA_ROW},"Angenommen")'
    set_cell(ws, 2, 4, None, font=Font(name="Calibri", size=14, bold=True, color="28A745"), alignment=center_align)

    set_cell(ws, 2, 5, "Abgelehnt:", font=BOLD_FONT, alignment=Alignment(horizontal="right", vertical="center"))
    ws["F2"] = f'=COUNTIF(H5:H{MAX_DATA_ROW},"Abgelehnt")'
    set_cell(ws, 2, 6, None, font=Font(name="Calibri", size=14, bold=True, color="DC3545"), alignment=center_align)

    set_cell(ws, 2, 7, "Offen:", font=BOLD_FONT, alignment=Alignment(horizontal="right", vertical="center"))
    ws["H2"] = '=B2-D2-F2'
    set_cell(ws, 2, 8, None, font=Font(name="Calibri", size=14, bold=True, color="FF6F00"), alignment=center_align)

    # ── Row 3: Spacer ─────────────────────────────────────────────────────
    ws.row_dimensions[3].height = 6

    # ── Row 4: Headers ────────────────────────────────────────────────────
    ws.row_dimensions[4].height = 36
    for col_idx, (header_text, _) in enumerate(HEADERS, 1):
        set_cell(ws, 4, col_idx, header_text,
                 font=HEADER_FONT, fill=HEADER_BG, alignment=center_wrap, border=thin_border)

    # ── Data Validation: Status dropdown ──────────────────────────────────
    dv_status = DataValidation(type="list", formula1='"Angenommen,Abgelehnt"', allow_blank=True)
    dv_status.error = "Bitte Angenommen oder Abgelehnt wählen"
    dv_status.errorTitle = "Ungültiger Status"
    dv_status.prompt = "Status auswählen"
    dv_status.promptTitle = "Entscheidung"
    ws.add_data_validation(dv_status)
    dv_status.add(f"H5:H{MAX_DATA_ROW}")

    # ── Data Validation: Names dropdown ───────────────────────────────────
    dv_names = DataValidation(type="list", formula1="=KVPNamenListe", allow_blank=True)
    dv_names.prompt = "Name auswählen"
    dv_names.promptTitle = "Einbringender"
    ws.add_data_validation(dv_names)
    dv_names.add(f"E5:E{MAX_DATA_ROW}")

    # ── Write data ────────────────────────────────────────────────────────
    for row_idx, entry in enumerate(data):
        r = 5 + row_idx
        # Nr
        set_cell(ws, r, 1, row_idx + 1, font=BOLD_FONT, alignment=center_align, border=thin_border)

        # Datum - parse string to date
        datum_str = entry.get("datum", "")
        datum_val = None
        if datum_str:
            for fmt in ("%Y-%m-%d", "%d.%m.%Y"):
                try:
                    datum_val = datetime.datetime.strptime(datum_str, fmt).date()
                    break
                except ValueError:
                    continue
        set_cell(ws, r, 2, datum_val, font=NORMAL_FONT, alignment=center_align,
                 border=thin_border, number_format="DD.MM.YYYY")

        set_cell(ws, r, 3, entry.get("thema", ""), font=NORMAL_FONT, alignment=left_wrap, border=thin_border)
        set_cell(ws, r, 4, entry.get("ersparnis", ""), font=NORMAL_FONT, alignment=left_wrap, border=thin_border)
        set_cell(ws, r, 5, entry.get("einbringender", ""), font=NORMAL_FONT, alignment=center_align, border=thin_border)
        set_cell(ws, r, 6, entry.get("kommentar_bv", ""), font=NORMAL_FONT, alignment=left_wrap, border=thin_border)
        set_cell(ws, r, 7, entry.get("kommentar_gremium", ""), font=NORMAL_FONT, alignment=left_wrap, border=thin_border)
        set_cell(ws, r, 8, entry.get("status", ""), font=NORMAL_FONT, alignment=center_align, border=thin_border)
        set_cell(ws, r, 9, entry.get("verantwortlich", ""), font=NORMAL_FONT, alignment=center_align, border=thin_border)

        # Auto row height based on content
        thema_len = len(entry.get("thema", ""))
        ws.row_dimensions[r].height = max(28, min(100, 14 * (1 + thema_len // 55)))

    # ── Pre-format empty rows with auto-numbering formula ─────────────────
    next_row = 5 + len(data)
    for r in range(next_row, next_row + 50):
        set_cell(ws, r, 1, f'=IF(B{r}<>"",ROW()-4,"")', font=BOLD_FONT, alignment=center_align, border=thin_border)
        for col in range(2, 10):
            set_cell(ws, r, col, None, font=NORMAL_FONT, border=thin_border)
            if col == 2:
                ws.cell(row=r, column=col).number_format = "DD.MM.YYYY"
                ws.cell(row=r, column=col).alignment = center_align
            elif col in (5, 8, 9):
                ws.cell(row=r, column=col).alignment = center_align
            else:
                ws.cell(row=r, column=col).alignment = left_wrap

    # ── Conditional Formatting ────────────────────────────────────────────
    add_kvp_conditional_formatting(ws, MAX_DATA_ROW)

    # ── Freeze + Filter ──────────────────────────────────────────────────
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:I{next_row + 49}"

    # ── Print Setup ──────────────────────────────────────────────────────
    setup_print(ws, f"KVP {year_name}")

    return ws


# ══════════════════════════════════════════════════════════════════════════════
# Create all year sheets (2026 first as active, then archives)
# ══════════════════════════════════════════════════════════════════════════════
year_sheets = {}
for i, (year_name, data) in enumerate(ALL_YEARS):
    ws = create_year_sheet(wb, year_name, data, is_active=(i == 0))
    year_sheets[year_name] = ws


# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD (inserted at position 0)
# ══════════════════════════════════════════════════════════════════════════════
ws_dash = wb.create_sheet(title="Dashboard", index=0)
ws_dash.sheet_properties.tabColor = "0F3460"

ws_dash.column_dimensions["A"].width = 3
ws_dash.column_dimensions["B"].width = 14
ws_dash.column_dimensions["C"].width = 14
ws_dash.column_dimensions["D"].width = 14
ws_dash.column_dimensions["E"].width = 14
ws_dash.column_dimensions["F"].width = 14
ws_dash.column_dimensions["G"].width = 5
ws_dash.column_dimensions["H"].width = 16
ws_dash.column_dimensions["I"].width = 45

# Title
ws_dash.merge_cells("B1:I2")
set_cell(ws_dash, 1, 2, "KVP Dashboard - Kontinuierlicher Verbesserungsprozess",
         font=TITLE_FONT, alignment=Alignment(horizontal="center", vertical="center"))
ws_dash.row_dimensions[1].height = 25
ws_dash.row_dimensions[2].height = 25

ws_dash.merge_cells("B3:I3")
set_cell(ws_dash, 3, 2, "Werksärztlicher Dienst Mercedes-Benz Rastatt/Kuppenheim",
         font=Font(name="Calibri", size=11, italic=True, color="666666"),
         alignment=Alignment(horizontal="center"))

# Spacer
ws_dash.row_dimensions[4].height = 8

# ── Gesamt-Zähler ─────────────────────────────────────────────────────────
ws_dash.row_dimensions[5].height = 40

ws_dash.merge_cells("B5:C5")
set_cell(ws_dash, 5, 2, "Gesamt Vorschläge:",
         font=Font(name="Calibri", size=13, bold=True, color="333333"),
         alignment=Alignment(horizontal="right", vertical="center"))
set_cell(ws_dash, 5, 4, None,
         font=Font(name="Calibri", size=24, bold=True, color="0F3460"),
         alignment=center_align)

set_cell(ws_dash, 5, 5, "Angenommen:",
         font=Font(name="Calibri", size=13, bold=True, color="333333"),
         alignment=Alignment(horizontal="right", vertical="center"))
set_cell(ws_dash, 5, 6, None,
         font=Font(name="Calibri", size=24, bold=True, color="28A745"),
         alignment=center_align)

ws_dash.merge_cells("H5:I5")
set_cell(ws_dash, 5, 8, None,
         font=Font(name="Calibri", size=13, bold=True, color="28A745"),
         alignment=Alignment(horizontal="left", vertical="center"))

# Spacer
ws_dash.row_dimensions[6].height = 8

# ── Jahresübersicht Tabelle ───────────────────────────────────────────────
row = 7
dash_headers = ["Jahr", "Vorschläge", "Angenommen", "Abgelehnt", "Offen", "", "Annahmequote", "Top-Thema"]
for c, h in enumerate(dash_headers, 2):
    set_cell(ws_dash, row, c, h, font=HEADER_FONT, fill=HEADER_BG, alignment=center_wrap, border=thin_border)
ws_dash.row_dimensions[row].height = 30

for i, (year_name, _) in enumerate(ALL_YEARS):
    r = row + 1 + i
    safe_name = f"'{year_name}'"
    alt_fill = LIGHT_GRAY_BG if i % 2 == 0 else WHITE_BG

    # Year with hyperlink
    year_cell = set_cell(ws_dash, r, 2, year_name,
        font=Font(name="Calibri", size=12, bold=True, color="0F3460", underline="single"),
        fill=alt_fill, border=thin_border, alignment=center_align)
    year_cell.hyperlink = f"#{safe_name}!A1"

    # Vorschläge
    set_cell(ws_dash, r, 3, None, font=COUNTER_FONT, fill=alt_fill, alignment=center_align, border=thin_border)
    ws_dash.cell(row=r, column=3).value = f"=COUNTA({safe_name}!B5:B{MAX_DATA_ROW})"

    # Angenommen
    set_cell(ws_dash, r, 4, None,
             font=Font(name="Calibri", size=14, bold=True, color="28A745"),
             fill=alt_fill, alignment=center_align, border=thin_border)
    ws_dash.cell(row=r, column=4).value = f'=COUNTIF({safe_name}!H5:H{MAX_DATA_ROW},"Angenommen")'

    # Abgelehnt
    set_cell(ws_dash, r, 5, None,
             font=Font(name="Calibri", size=14, bold=True, color="DC3545"),
             fill=alt_fill, alignment=center_align, border=thin_border)
    ws_dash.cell(row=r, column=5).value = f'=COUNTIF({safe_name}!H5:H{MAX_DATA_ROW},"Abgelehnt")'

    # Offen
    set_cell(ws_dash, r, 6, None,
             font=Font(name="Calibri", size=14, bold=True, color="FF6F00"),
             fill=alt_fill, alignment=center_align, border=thin_border)
    ws_dash.cell(row=r, column=6).value = f"=C{r}-D{r}-E{r}"

    # Spacer
    set_cell(ws_dash, r, 7, None, fill=alt_fill)

    # Annahmequote
    set_cell(ws_dash, r, 8, None,
             font=Font(name="Calibri", size=12, bold=True, color="0F3460"),
             fill=alt_fill, alignment=center_align, border=thin_border)
    ws_dash.cell(row=r, column=8).value = f'=IF(C{r}>0,D{r}/C{r},"")'
    ws_dash.cell(row=r, column=8).number_format = "0%"

    # Top-Thema (first entry)
    set_cell(ws_dash, r, 9, None, font=NORMAL_FONT, fill=alt_fill, alignment=left_wrap, border=thin_border)
    ws_dash.cell(row=r, column=9).value = f"=IF({safe_name}!C5<>\"\",{safe_name}!C5,\"—\")"

    ws_dash.row_dimensions[r].height = 28

# Summen-Zeile
sum_row = row + 1 + len(ALL_YEARS)
set_cell(ws_dash, sum_row, 2, "GESAMT",
         font=Font(name="Calibri", size=12, bold=True, color="FFFFFF"),
         fill=HEADER_BG, alignment=center_align, border=thin_border)
for col in [3, 4, 5, 6]:
    col_letter = get_column_letter(col)
    set_cell(ws_dash, sum_row, col, None,
             font=Font(name="Calibri", size=12, bold=True, color="FFFFFF"),
             fill=HEADER_BG, alignment=center_align, border=thin_border)
    ws_dash.cell(row=sum_row, column=col).value = f"=SUM({col_letter}{row+1}:{col_letter}{sum_row-1})"
set_cell(ws_dash, sum_row, 7, None, fill=HEADER_BG)
# Gesamt-Quote
set_cell(ws_dash, sum_row, 8, None,
         font=Font(name="Calibri", size=12, bold=True, color="FFFFFF"),
         fill=HEADER_BG, alignment=center_align, border=thin_border)
ws_dash.cell(row=sum_row, column=8).value = f"=IF(C{sum_row}>0,D{sum_row}/C{sum_row},\"\")"
ws_dash.cell(row=sum_row, column=8).number_format = "0%"
set_cell(ws_dash, sum_row, 9, None, fill=HEADER_BG, border=thin_border)

# Fill top counters from sum row
ws_dash["D5"] = f"=C{sum_row}"
ws_dash["F5"] = f"=D{sum_row}"
ws_dash["H5"] = f'="Annahmequote: "&TEXT(H{sum_row},"0%")'

# Conditional formatting on Offen column
offen_range = f"F{row+1}:F{sum_row-1}"
ws_dash.conditional_formatting.add(offen_range, CellIsRule(
    operator="greaterThan", formula=["0"], stopIfTrue=False,
    fill=PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid"),
    font=Font(bold=True, color="FF6F00")))

# Annahmequote color coding
quote_range = f"H{row+1}:H{sum_row}"
ws_dash.conditional_formatting.add(quote_range, CellIsRule(
    operator="greaterThanOrEqual", formula=["0.7"], stopIfTrue=False,
    fill=PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")))
ws_dash.conditional_formatting.add(quote_range, CellIsRule(
    operator="lessThan", formula=["0.5"], stopIfTrue=False,
    fill=PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")))

# ── Legende ───────────────────────────────────────────────────────────────
leg_row = sum_row + 2
ws_dash.merge_cells(f"B{leg_row}:E{leg_row}")
set_cell(ws_dash, leg_row, 2, "Legende Farbmarkierungen in den Jahres-Sheets:",
         font=Font(name="Calibri", size=11, bold=True, color="333333"))

legends = [
    ("BBDEFB", "Nr. blau", "Neuer Vorschlag (letzte 14 Tage)"),
    ("D4EDDA", "Zeile grün", "Angenommen"),
    ("F8D7DA", "Zeile rot", "Abgelehnt"),
    ("FFF8E1", "Zeile gelb", "Offen (noch keine Entscheidung)"),
]

for j, (color, label, desc) in enumerate(legends):
    r = leg_row + 1 + j
    set_cell(ws_dash, r, 2, None,
             fill=PatternFill(start_color=color, end_color=color, fill_type="solid"),
             border=thin_border)
    set_cell(ws_dash, r, 3, label, font=BOLD_FONT, border=thin_border)
    ws_dash.merge_cells(f"D{r}:F{r}")
    set_cell(ws_dash, r, 4, desc, font=NORMAL_FONT, border=thin_border)

# ── Info-Bereich ──────────────────────────────────────────────────────────
info_row = leg_row + len(legends) + 2
ws_dash.merge_cells(f"B{info_row}:I{info_row}")
set_cell(ws_dash, info_row, 2, "Hinweise zur KVP-Liste 2026:",
         font=Font(name="Calibri", size=11, bold=True, color="333333"))

hints = [
    "Neuen KVP-Vorschlag im aktiven Jahres-Tab eintragen (Datum + Thema + Einbringender)",
    "Status wird vom Gremium auf \"Angenommen\" oder \"Abgelehnt\" gesetzt",
    "Dashboard zeigt automatisch Statistiken über alle Jahre",
    "Einbringender kann per Dropdown ausgewählt werden",
    "Archiv-Tabs (2022-2025) enthalten alle bisherigen Vorschläge",
    "Kein Blattschutz, keine Makros - alles frei editierbar",
]

for j, hint in enumerate(hints):
    r = info_row + 1 + j
    set_cell(ws_dash, r, 2, "✓", font=Font(name="Calibri", size=10, bold=True, color="28A745"))
    ws_dash.merge_cells(f"C{r}:I{r}")
    set_cell(ws_dash, r, 3, hint, font=NORMAL_FONT)

ws_dash.freeze_panes = "B7"

# Print setup for dashboard
ws_dash.page_setup.orientation = "landscape"
ws_dash.page_setup.paperSize = ws_dash.PAPERSIZE_A4
ws_dash.page_setup.fitToWidth = 1
ws_dash.page_setup.fitToHeight = 1
ws_dash.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)


# ══════════════════════════════════════════════════════════════════════════════
# Hilfslisten sheet (for Named Range dropdown source)
# ══════════════════════════════════════════════════════════════════════════════
ws_hilf = wb.create_sheet(title="Hilfslisten")
ws_hilf.sheet_properties.tabColor = "666666"

ws_hilf.column_dimensions["A"].width = 22

set_cell(ws_hilf, 1, 1, "Dropdown: Einbringender", font=BOLD_FONT)
sorted_names = sorted(set(NAMES))
for i, name in enumerate(sorted_names):
    set_cell(ws_hilf, 2 + i, 1, name, font=NORMAL_FONT)
NAMES_LAST_ROW = 1 + len(sorted_names)

# Named range for names dropdown
from openpyxl.workbook.defined_name import DefinedName
namen_ref = f"Hilfslisten!$A$2:$A${NAMES_LAST_ROW}"
dn_namen = DefinedName("KVPNamenListe", attr_text=namen_ref)
wb.defined_names.add(dn_namen)


# ══════════════════════════════════════════════════════════════════════════════
# Save
# ══════════════════════════════════════════════════════════════════════════════
output_path = "/home/user/Ambulanzzeug/KVP_2026.xlsx"
wb.save(output_path)
print(f"Saved: {output_path}")
print(f"Sheets: {wb.sheetnames}")
print("Done!")
