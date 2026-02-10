#!/usr/bin/env python3
"""
Creates Vorsorgeaufwand_2026.xlsx - Lookup tool for occupational health examinations.
Replaces ActiveX ToggleButtons + VBA macros with data validation + formulas.
"""
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.page import PageMargins
import openpyxl.worksheet.properties

wb = openpyxl.Workbook()

# ── Styles ────────────────────────────────────────────────────────────────────
HEADER_BG = PatternFill(start_color="0F3460", end_color="0F3460", fill_type="solid")
GREEN_BTN = PatternFill(start_color="28A745", end_color="28A745", fill_type="solid")
LIGHT_GRAY = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
WHITE_BG = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
YELLOW_BG = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
RESULT_BG = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
SELECTED_BG = PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid")

TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
HEADER_FONT = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
EXAM_FONT = Font(name="Calibri", size=10, color="333333")
EXAM_BOLD = Font(name="Calibri", size=10, bold=True, color="333333")
X_FONT = Font(name="Calibri", size=9, color="666666")
RESULT_FONT = Font(name="Calibri", size=11, bold=True, color="1B5E20")
RESULT_NUM = Font(name="Calibri", size=10, color="333333")
SELECT_FONT = Font(name="Calibri", size=9, bold=True, color="B71C1C")

thin_border = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)
medium_left = Border(
    left=Side(style="medium", color="333333"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)

center_align = Alignment(horizontal="center", vertical="center")
center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_center = Alignment(horizontal="left", vertical="center")
left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
rotated = Alignment(horizontal="center", vertical="bottom", text_rotation=90, wrap_text=True)


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
# G-Number definitions (42 examination types)
# ══════════════════════════════════════════════════════════════════════════════
G_NUMBERS = [
    ("DGC", "Daimler Gesundheits-Check"),
    ("G1.1-1.4", "Staubbelastung"),
    ("G2", "Blei oder seine Verbindungen"),
    ("G7", "Kohlenmonoxid"),
    ("G8", "Benzol"),
    ("G15", ""),
    ("G17", ""),
    ("G20", "Lärm"),
    ("G21", ""),
    ("G23", "Obstruktive Atemwegserkrankungen"),
    ("G24", "Hauterkrankungen"),
    ("G25 mit Peri", "Fahr-Steuer- und Überwachungstätigkeiten"),
    ("G25 o. Peri", ""),
    ("G26.1", "Atemschutzgeräte"),
    ("G26.2", "Atemschutzgeräte"),
    ("G26.3 / DE-ASG3", "Atemschutzgeräte"),
    ("G27", "Isocyanate"),
    ("G29", "Toluol und Xylol"),
    ("G30", "Hitzearbeiten"),
    ("G33", "Aromatische Nitro- und Aminoverbindungen"),
    ("G34", "Fluor und seine anorg. Verbindungen"),
    ("G35", "Arbeitsaufenthalt im Ausland"),
    ("G35 China", ""),
    ("G35 Südafrika", ""),
    ("G37", "Bildschirmarbeitsplätze"),
    ("G38", "Nickel oder seine Verbindungen"),
    ("G39", "Schweißrauche"),
    ("G40", "Krebserz./erbgutveränd. Gefahrstoffe"),
    ("G41", "Arbeiten mit Absturzgefahr"),
    ("G42", "Tätigkeiten mit Infektionsgefährdung"),
    ("G44", ""),
    ("G45", "Styrol"),
    ("G46", "Belastungen Muskel-/Skelettsystem"),
    ("Höhentest/Klimakammer", ""),
    ("Hörprüfer", ""),
    ("InfSchG (Gastro)", ""),
    ("O2-red.", ""),
    ("Rissprüfer", ""),
    ("Sichtprüfer", ""),
    ("3-Schicht", ""),
    ("AuS", ""),
    ("FEV1", "Fahrerlaubnisverordnung"),
]

# ══════════════════════════════════════════════════════════════════════════════
# Examination rows (28 procedures)
# ══════════════════════════════════════════════════════════════════════════════
EXAMINATIONS = [
    "01. Anamnese 2016",
    "02. Arbeitsanamnesen - Freitext",
    "03. Audiometrie",
    "04. Flüstersprache",
    "05. Lärm (spezielle Anamnese G20)",
    "06. Körperliche Untersuchung",
    "07. orthopädische Anamnese",
    "08. psychosom. Untersuchung",
    "09. Labor - Urin",
    "10. Labor Blut EP01",
    "10. Labor Blut EP02",
    "11. Sehtest",
    "12. Perimetrie",
    "13. Lungenfunktion",
    "14. Ruhe EKG",
    "15. Belastungs EKG",
    "16. Körperanalysewaage",
    "17. Vitalparameter (RR, cm, kg)",
    "18. Röntgen",
    "19. Impfen",
    "20. Sonographie (Abdomen)",
    "21. Neurol. Unters. Fokussierend",
    "22. Rhinoskopie",
    "23. G35 Inhalte, spezifisch",
    "24. Gesundheitsberatung",
    "25. Vollumfängliche Einwilligung",
    "26. ODIN-Bogen",
    "27. Biomonitoring",
]

# Matrix: for each examination (row), which G-numbers require it
# Keys = G-number short names, stored as set of column indices (0-based into G_NUMBERS)
# Using column letters for readability, mapped to 0-based index

def g_idx(name):
    """Get 0-based index of a G-number."""
    for i, (g, _) in enumerate(G_NUMBERS):
        if g == name:
            return i
    raise ValueError(f"Unknown G-number: {name}")

# Build the X-mark matrix: MATRIX[exam_row_0based] = set of g_indices
# Transcribed from the analysis
MATRIX = [
    # 01. Anamnese: all except Z(G37), AN(Sichtprüfer)
    {0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,25,26,27,28,29,30,31,32,33,34,35,36,37,39,40,41},
    # 02. Arbeitsanamnesen: all except AN(Sichtprüfer)
    {0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,39,40,41},
    # 03. Audiometrie: G20, G26.2, G26.3, G41, Hörprüfer, FEV1
    {g_idx("G20"), g_idx("G26.2"), g_idx("G26.3 / DE-ASG3"), g_idx("G41"), g_idx("Hörprüfer"), g_idx("FEV1")},
    # 04. Flüstersprache: G25 mit Peri, G25 o. Peri
    {g_idx("G25 mit Peri"), g_idx("G25 o. Peri")},
    # 05. Lärm G20: only G20
    {g_idx("G20")},
    # 06. Körperliche Untersuchung: most except G20, G33, AN(Sichtprüfer), AM(Rissprüfer)
    {0,1,2,3,4,5,6,8,9,10,11,12,13,14,15,16,17,18,19,21,22,23,24,25,26,27,28,29,30,31,32,33,35,36,39,40,41},
    # 07. orthopädische Anamnese: DGC, G41, G46
    {g_idx("DGC"), g_idx("G41"), g_idx("G46")},
    # 08. psychosom.: DGC only
    {g_idx("DGC")},
    # 09. Labor Urin: G2, G8, G15, G21, G23, G25mitPeri, G25oPeri, G26.1, G26.2, G26.3, G29, G30, G33, G35, G35China, G35Südafrika, G38, G39, G40, G41, G42, G45, Höhentest, FEV1
    {g_idx("G2"), g_idx("G8"), g_idx("G15"), g_idx("G21"), g_idx("G23"), g_idx("G25 mit Peri"), g_idx("G25 o. Peri"),
     g_idx("G26.1"), g_idx("G26.2"), g_idx("G26.3 / DE-ASG3"), g_idx("G29"), g_idx("G30"), g_idx("G33"),
     g_idx("G35"), g_idx("G35 China"), g_idx("G35 Südafrika"), g_idx("G38"), g_idx("G39"), g_idx("G40"),
     g_idx("G41"), g_idx("G42"), g_idx("G45"), g_idx("Höhentest/Klimakammer"), g_idx("FEV1")},
    # 10. Labor Blut EP01: G7, G26.2, G35, G35China, G35Südafrika, G42, G44, Höhentest, O2-red, FEV1
    {g_idx("G7"), g_idx("G26.2"), g_idx("G35"), g_idx("G35 China"), g_idx("G35 Südafrika"),
     g_idx("G42"), g_idx("G44"), g_idx("Höhentest/Klimakammer"), g_idx("O2-red."), g_idx("FEV1")},
    # 10. Labor Blut EP02: G2, G8, G15, G26.3, G27, G29, G33, G38, G40, G45
    {g_idx("G2"), g_idx("G8"), g_idx("G15"), g_idx("G26.3 / DE-ASG3"), g_idx("G27"), g_idx("G29"),
     g_idx("G33"), g_idx("G38"), g_idx("G40"), g_idx("G45")},
    # 11. Sehtest: G17, G25mitPeri, G25oPeri, G26.2, G26.3, G35China, G35Südafrika, G37, G41, Rissprüfer, Sichtprüfer, AuS, FEV1
    {g_idx("G17"), g_idx("G25 mit Peri"), g_idx("G25 o. Peri"), g_idx("G26.2"), g_idx("G26.3 / DE-ASG3"),
     g_idx("G35 China"), g_idx("G35 Südafrika"), g_idx("G37"), g_idx("G41"),
     g_idx("Rissprüfer"), g_idx("Sichtprüfer"), g_idx("AuS"), g_idx("FEV1")},
    # 12. Perimetrie: G25mitPeri, G41, FEV1
    {g_idx("G25 mit Peri"), g_idx("G41"), g_idx("FEV1")},
    # 13. Lungenfunktion: G1.1-1.4, G7, G15, G23, G26.2, G26.3, G27, G35, G35China, G35Südafrika, G38, G39, G40, G45, Höhentest, O2-red
    {g_idx("G1.1-1.4"), g_idx("G7"), g_idx("G15"), g_idx("G23"), g_idx("G26.2"), g_idx("G26.3 / DE-ASG3"),
     g_idx("G27"), g_idx("G35"), g_idx("G35 China"), g_idx("G35 Südafrika"), g_idx("G38"), g_idx("G39"),
     g_idx("G40"), g_idx("G45"), g_idx("Höhentest/Klimakammer"), g_idx("O2-red.")},
    # 14. Ruhe EKG: G7, G26.2, G26.3, G30, G35, G35China, G35Südafrika, G41, Höhentest, O2-red, AuS
    {g_idx("G7"), g_idx("G26.2"), g_idx("G26.3 / DE-ASG3"), g_idx("G30"),
     g_idx("G35"), g_idx("G35 China"), g_idx("G35 Südafrika"), g_idx("G41"),
     g_idx("Höhentest/Klimakammer"), g_idx("O2-red."), g_idx("AuS")},
    # 15. Belastungs EKG: G7, G26.3, G30, G35, G35China, G35Südafrika, Höhentest, O2-red
    {g_idx("G7"), g_idx("G26.3 / DE-ASG3"), g_idx("G30"),
     g_idx("G35"), g_idx("G35 China"), g_idx("G35 Südafrika"),
     g_idx("Höhentest/Klimakammer"), g_idx("O2-red.")},
    # 16. Körperanalysewaage: DGC only
    {g_idx("DGC")},
    # 17. Vitalparameter: same as Anamnese basically (all except Z, AN)
    {0,1,2,3,4,5,6,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,25,26,27,28,29,30,31,32,33,35,36,39,40,41},
    # 18. Röntgen: G1.1-1.4, G2, G15, G23, G26.3, G35China, G35Südafrika, G40, Höhentest
    {g_idx("G1.1-1.4"), g_idx("G2"), g_idx("G15"), g_idx("G23"), g_idx("G26.3 / DE-ASG3"),
     g_idx("G35 China"), g_idx("G35 Südafrika"), g_idx("G40"), g_idx("Höhentest/Klimakammer")},
    # 19. Impfen: G35, G35China, G35Südafrika, G42
    {g_idx("G35"), g_idx("G35 China"), g_idx("G35 Südafrika"), g_idx("G42")},
    # 20. Sonographie: G35China, G40
    {g_idx("G35 China"), g_idx("G40")},
    # 21. Neurol. Fokussierend: G2, G29, G41, G45
    {g_idx("G2"), g_idx("G29"), g_idx("G41"), g_idx("G45")},
    # 22. Rhinoskopie: G15
    {g_idx("G15")},
    # 23. G35 Inhalte: G35, G35China, G35Südafrika
    {g_idx("G35"), g_idx("G35 China"), g_idx("G35 Südafrika")},
    # 24. Gesundheitsberatung: ALL 42
    set(range(42)),
    # 25. Einwilligung: ALL 42
    set(range(42)),
    # 26. ODIN-Bogen: G1.1-1.4, G2, G8, G15, G27, G33, G38, G40, G42, G44
    {g_idx("G1.1-1.4"), g_idx("G2"), g_idx("G8"), g_idx("G15"), g_idx("G27"), g_idx("G33"),
     g_idx("G38"), g_idx("G40"), g_idx("G42"), g_idx("G44")},
    # 27. Biomonitoring: G2, G7, G8, G15, G27, G29, G38, G39, G40, G45
    {g_idx("G2"), g_idx("G7"), g_idx("G8"), g_idx("G15"), g_idx("G27"), g_idx("G29"),
     g_idx("G38"), g_idx("G39"), g_idx("G40"), g_idx("G45")},
]

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 1: Vorsorgeaufwand (Main Interface)
# ══════════════════════════════════════════════════════════════════════════════
ws = wb.active
ws.title = "Vorsorgeaufwand"
ws.sheet_properties.tabColor = "0F3460"

# Column widths
ws.column_dimensions["A"].width = 35  # Exam names
for i in range(len(G_NUMBERS)):
    col_letter = get_column_letter(i + 2)
    ws.column_dimensions[col_letter].width = 5.5

# Results column
RESULT_COL = len(G_NUMBERS) + 2  # Column after all G-numbers
RESULT_LETTER = get_column_letter(RESULT_COL)
ws.column_dimensions[RESULT_LETTER].width = 3  # Spacer
RESULT_COL2 = RESULT_COL + 1
RESULT_LETTER2 = get_column_letter(RESULT_COL2)
ws.column_dimensions[RESULT_LETTER2].width = 45  # Result text

# ── Row 1: Title bar ─────────────────────────────────────────────────────────
ws.row_dimensions[1].height = 35
last_g_col = len(G_NUMBERS) + 1
last_g_letter = get_column_letter(last_g_col)
ws.merge_cells(f"A1:{last_g_letter}1")
set_cell(ws, 1, 1,
    "Vorsorgeaufwand - Einzeluntersuchungen MEDAS 2026",
    font=Font(name="Calibri", size=16, bold=True, color="FFFFFF"),
    fill=HEADER_BG,
    alignment=Alignment(horizontal="left", vertical="center"))

# Results header
ws.merge_cells(f"{RESULT_LETTER}1:{RESULT_LETTER2}1")
set_cell(ws, 1, RESULT_COL,
    "Ergebnis: Benötigte Untersuchungen",
    font=Font(name="Calibri", size=14, bold=True, color="FFFFFF"),
    fill=PatternFill(start_color="28A745", end_color="28A745", fill_type="solid"),
    alignment=Alignment(horizontal="center", vertical="center"))

# ── Row 2: G-Number Selection Row ────────────────────────────────────────────
ws.row_dimensions[2].height = 25
set_cell(ws, 2, 1, '▼ "x" eingeben zum Auswählen →',
    font=Font(name="Calibri", size=11, bold=True, color="FFFFFF"),
    fill=PatternFill(start_color="28A745", end_color="28A745", fill_type="solid"),
    alignment=left_center)

for i, (g_name, g_desc) in enumerate(G_NUMBERS):
    col = i + 2
    set_cell(ws, 2, col, None,
        font=Font(name="Calibri", size=11, bold=True, color="B71C1C"),
        fill=WHITE_BG,
        alignment=center_align,
        border=thin_border)

# Results counter in row 2
ws.merge_cells(f"{RESULT_LETTER}2:{RESULT_LETTER2}2")
counter_cell = set_cell(ws, 2, RESULT_COL, None,
    font=Font(name="Calibri", size=12, bold=True, color="28A745"),
    fill=RESULT_BG,
    alignment=center_align)
counter_cell.value = (
    f'=IF(COUNTA({RESULT_LETTER2}4:{RESULT_LETTER2}31)>0,'
    f'"✓ "&COUNTA({RESULT_LETTER2}4:{RESULT_LETTER2}31)&" Untersuchung(en) erforderlich",'
    f'"← x eingeben bei den gewünschten G-Nummern")'
)

# ── Row 3: G-Number Labels (rotated) ─────────────────────────────────────────
ws.row_dimensions[3].height = 110
set_cell(ws, 3, 1, "G-Nummer →",
    font=Font(name="Calibri", size=10, bold=True, color="0F3460"),
    fill=PatternFill(start_color="E8EAF6", end_color="E8EAF6", fill_type="solid"),
    alignment=Alignment(horizontal="right", vertical="center"))

for i, (g_name, g_desc) in enumerate(G_NUMBERS):
    col = i + 2
    set_cell(ws, 3, col, g_name,
        font=Font(name="Calibri", size=8, bold=True, color="0F3460"),
        fill=PatternFill(start_color="E8EAF6", end_color="E8EAF6", fill_type="solid"),
        alignment=rotated,
        border=thin_border)

# Results header row 3 - shows selected G-numbers summary
set_cell(ws, 3, RESULT_COL, None)
set_cell(ws, 3, RESULT_COL2, "Untersuchung",
    font=Font(name="Calibri", size=10, bold=True, color="0F3460"),
    fill=PatternFill(start_color="E8EAF6", end_color="E8EAF6", fill_type="solid"),
    alignment=left_center,
    border=thin_border)

# NO data validation needed - user just types "x" freely

# Conditional formatting: any non-empty selection cell turns red
ws.conditional_formatting.add(f"B2:{last_g_letter}2", FormulaRule(
    formula=[f'B2<>""'],
    stopIfTrue=False,
    fill=PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid"),
    font=Font(name="Calibri", size=11, bold=True, color="B71C1C")))

# ── Row 4+: Examination Matrix ───────────────────────────────────────────────
for exam_idx, exam_name in enumerate(EXAMINATIONS):
    r = 4 + exam_idx
    alt_fill = LIGHT_GRAY if exam_idx % 2 == 0 else WHITE_BG

    # Exam name in column A
    set_cell(ws, r, 1, exam_name,
        font=EXAM_FONT,
        fill=alt_fill,
        alignment=left_wrap,
        border=medium_left)
    ws.row_dimensions[r].height = 20

    # X marks in G-number columns
    for g_i in MATRIX[exam_idx]:
        col = g_i + 2
        val = "x"
        # Special values
        if exam_idx == 15 and g_i in (g_idx("G35"), g_idx("G35 China"), g_idx("G35 Südafrika")):
            val = "45J."  # Belastungs EKG only if >45 years
        if exam_idx == 9 and g_i in (g_idx("G7"), g_idx("G26.2"), g_idx("G35"), g_idx("G35 China"), g_idx("G35 Südafrika"),
                                      g_idx("G42"), g_idx("G44"), g_idx("Höhentest/Klimakammer"), g_idx("O2-red."), g_idx("FEV1")):
            cell_fill = YELLOW_BG  # EP01 = yellow
        elif exam_idx == 10:
            cell_fill = YELLOW_BG  # EP02 = yellow
        else:
            cell_fill = alt_fill

        set_cell(ws, r, col, val,
            font=X_FONT,
            fill=cell_fill,
            alignment=center_align,
            border=thin_border)

    # Fill empty G-cells with border
    for g_i in range(len(G_NUMBERS)):
        col = g_i + 2
        if g_i not in MATRIX[exam_idx]:
            set_cell(ws, r, col, None,
                font=X_FONT,
                fill=alt_fill,
                alignment=center_align,
                border=thin_border)

    # ── Result formula in result column ───────────────────────────────────
    # Build OR formula: if ANY selected G-number has an x in this row
    set_cell(ws, r, RESULT_COL, None, fill=alt_fill)  # spacer

    # Build formula: check if any selected G-column (row 2 not empty) has a value in this row
    # =IF(SUMPRODUCT(($B$2<>"")*( B{r}<>""))>0, A{r}, "")
    result_cell = set_cell(ws, r, RESULT_COL2, None,
        font=RESULT_NUM,
        fill=alt_fill,
        alignment=left_wrap,
        border=thin_border)
    result_cell.value = (
        f'=IF(SUMPRODUCT(($B$2:${last_g_letter}$2<>"")*'
        f'(B{r}:{last_g_letter}{r}<>""))>0,A{r},"")'
    )

# ── EP01/EP02 priority override: EP02 supersedes EP01 ─────────────────────
# EP01 = exam_idx 9, row 13.  EP02 = exam_idx 10, row 14.
ep01_row = 4 + 9   # row 13
ep02_row = 4 + 10  # row 14
# EP01 result: only show if EP01 is triggered AND EP02 is NOT triggered
ws.cell(row=ep01_row, column=RESULT_COL2).value = (
    f'=IF(AND(SUMPRODUCT(($B$2:${last_g_letter}$2<>"")*'
    f'(B{ep01_row}:{last_g_letter}{ep01_row}<>""))>0,'
    f'SUMPRODUCT(($B$2:${last_g_letter}$2<>"")*'
    f'(B{ep02_row}:{last_g_letter}{ep02_row}<>""))=0),'
    f'A{ep01_row},"")'
)
# EP02 result: show normally (already correct), but append " (ersetzt EP01)" hint
ws.cell(row=ep02_row, column=RESULT_COL2).value = (
    f'=IF(SUMPRODUCT(($B$2:${last_g_letter}$2<>"")*'
    f'(B{ep02_row}:{last_g_letter}{ep02_row}<>""))>0,'
    f'A{ep02_row}&" (ersetzt EP01)","")'
)

# ── Selected G-Numbers summary in row 3 result area ──────────────────────────
# Use TEXTJOIN (Office 365) to show which G-numbers are selected
set_cell(ws, 3, RESULT_COL, None)
summary_cell = set_cell(ws, 3, RESULT_COL2, None,
    font=Font(name="Calibri", size=9, bold=True, color="0F3460"),
    fill=PatternFill(start_color="E8EAF6", end_color="E8EAF6", fill_type="solid"),
    alignment=left_wrap,
    border=thin_border)
summary_cell.value = (
    f'=IF(COUNTA(B2:{last_g_letter}2)>0,'
    f'"Gewählt: "&TEXTJOIN(", ",TRUE,IF(B2:{last_g_letter}2<>"",B3:{last_g_letter}3,"")),'
    f'"← Untersuchungen")'
)

# ── Special rows after the matrix ─────────────────────────────────────────────
# Row for EP01/EP02 note + reset hint
note_row = 4 + len(EXAMINATIONS) + 1
ws.merge_cells(f"A{note_row}:{last_g_letter}{note_row}")
set_cell(ws, note_row, 1,
    "Hinweis: Gelbe Zellen = EP01/EP02 (Labor Blut). Wenn EP02 erforderlich, ersetzt EP02 den EP01. "
    "Belastungs-EKG bei G35-Varianten nur ab 45 Jahre (\"45J.\").",
    font=Font(name="Calibri", size=9, italic=True, color="666666"),
    alignment=left_wrap)
ws.row_dimensions[note_row].height = 30

# Reset hint row
reset_row = note_row + 1
ws.merge_cells(f"A{reset_row}:{last_g_letter}{reset_row}")
set_cell(ws, reset_row, 1,
    "Zurücksetzen: Markieren Sie Zeile 2 (grüne Auswahlzeile), drücken Sie Entf um alle Auswahlen zu löschen.",
    font=Font(name="Calibri", size=9, bold=True, italic=True, color="28A745"),
    alignment=left_wrap)
ws.row_dimensions[reset_row].height = 22

# ── Freeze panes ──────────────────────────────────────────────────────────────
ws.freeze_panes = "B4"

# ── Print Setup ───────────────────────────────────────────────────────────────
ws.page_setup.orientation = "landscape"
ws.page_setup.paperSize = ws.PAPERSIZE_A4
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
ws.page_margins = PageMargins(left=0.3, right=0.3, top=0.5, bottom=0.5, header=0.2, footer=0.2)
ws.oddHeader.center.text = "Vorsorgeaufwand - Einzeluntersuchungen MEDAS"
ws.oddFooter.left.text = "Werksärztlicher Dienst Mercedes-Benz"
ws.oddFooter.right.text = "Seite &P von &N"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 2: Biomonitoring
# ══════════════════════════════════════════════════════════════════════════════
ws_bio = wb.create_sheet(title="Biomonitoring")
ws_bio.sheet_properties.tabColor = "FF6F00"

ws_bio.column_dimensions["A"].width = 12
ws_bio.column_dimensions["B"].width = 38
ws_bio.column_dimensions["C"].width = 35
ws_bio.column_dimensions["D"].width = 22
ws_bio.column_dimensions["E"].width = 30

# Header
bio_headers = ["G-Nr.", "Bereich", "Substanz", "Probenmaterial", "Bedingungen"]
for c, h in enumerate(bio_headers, 1):
    set_cell(ws_bio, 1, c, h,
        font=HEADER_FONT, fill=HEADER_BG, alignment=center_wrap, border=thin_border)
ws_bio.row_dimensions[1].height = 28

bio_data = [
    ("G7", "KC, Log., Finish, Abfahren, Pforte", "COHb", "Fingersensor", ""),
    ("G8", "Betankung, QM, Finish, Tankstelle", "Benzol", "", ""),
    ("G27", "OF - DL", "kein Wert hinterlegt", "", ""),
    ("G29", "OF", "kein Wert hinterlegt", "", ""),
    ("", "QM", "kein Wert hinterlegt", "", ""),
    ("G38", "OF", "Nickel im Urin", "", ""),
    ("G45", "", "Styrol", "", ""),
]

for i, (g_nr, bereich, substanz, probe, beding) in enumerate(bio_data):
    r = 2 + i
    alt_fill = LIGHT_GRAY if i % 2 == 0 else WHITE_BG
    set_cell(ws_bio, r, 1, g_nr, font=EXAM_BOLD, fill=alt_fill, alignment=center_align, border=thin_border)
    set_cell(ws_bio, r, 2, bereich, font=EXAM_FONT, fill=alt_fill, alignment=left_wrap, border=thin_border)
    set_cell(ws_bio, r, 3, substanz, font=EXAM_FONT, fill=alt_fill, alignment=left_wrap, border=thin_border)
    set_cell(ws_bio, r, 4, probe, font=EXAM_FONT, fill=alt_fill, alignment=center_align, border=thin_border)
    set_cell(ws_bio, r, 5, beding, font=EXAM_FONT, fill=alt_fill, alignment=left_wrap, border=thin_border)

ws_bio.freeze_panes = "A2"


# ══════════════════════════════════════════════════════════════════════════════
# Save
# ══════════════════════════════════════════════════════════════════════════════
output_path = "/home/user/Ambulanzzeug/Vorsorgeaufwand_2026.xlsx"
wb.save(output_path)
print(f"Saved: {output_path}")
print(f"Sheets: {wb.sheetnames}")
print("Done!")
