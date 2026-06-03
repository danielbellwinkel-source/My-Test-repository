#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from datetime import date

# ─── Farben ───────────────────────────────────────────────────────────────────
C_RED       = "E31E24"
C_ANTHRAZIT = "2D2D2D"
C_WHITE     = "FFFFFF"
C_LIGHT     = "F7F7F7"
C_MID       = "E8E8E8"
C_GREEN_BG  = "C8E6C9"
C_RED_BG    = "FFCDD2"
C_YELLOW_BG = "FFF9C4"
C_BLUE_BG   = "E3F2FD"
C_ORANGE_BG = "FFE0B2"
C_GRAY_BG   = "ECEFF1"
C_RED_TXT   = "B71C1C"
C_GREEN_TXT = "1B5E20"

def fill(color):
    return PatternFill(start_color=color, end_color=color, fill_type="solid")

def font(name="Calibri", size=10, bold=False, color="000000", italic=False):
    return Font(name=name, size=size, bold=bold, color=color, italic=italic)

def align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def border_thin():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

# ─── Workbook ─────────────────────────────────────────────────────────────────
wb = openpyxl.Workbook()

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 1: Lead-Tracking
# ══════════════════════════════════════════════════════════════════════════════
ws = wb.active
ws.title = "Lead-Tracking 2026"
ws.sheet_view.showGridLines = True

# ─── Header-Spalten ───────────────────────────────────────────────────────────
headers = [
    ("Lead-Nr",               10),
    ("Datum",                 12),
    ("Mitarbeiter",           12),
    ("Kunde",                 18),
    ("Kontakt",               18),
    ("Kategorie",             14),
    ("Favorit (Modell+Preis)",24),
    ("Alternative (M+P)",     22),
    ("Typ",                    7),
    ("Offener Punkt",         15),
    ("Einwand",               10),
    ("KR",                     5),
    ("Status",                22),
    ("Naechster Schritt",     16),
    ("Check-Datum",           13),
    ("Angebotswert EUR",      14),
    ("Abschluss",             11),
    ("Verlustgrund",          14),
    ("Notiz",                 28),
    ("A",                      4),
]

for col, (h, w) in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col, value=h)
    cell.fill  = fill(C_ANTHRAZIT)
    cell.font  = font(bold=True, color=C_WHITE, size=9)
    cell.alignment = align("center", wrap=True)
    cell.border = border_thin()
    ws.column_dimensions[get_column_letter(col)].width = w

ws.row_dimensions[1].height = 28
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

# ─── Beispieldaten ────────────────────────────────────────────────────────────
rows = [
    ["2026-0001", date(2026,5,15), "Daniel",  "Hr. Egger",    "0664 1234567",      "Laufband",    "Modell B Pro - 2190", "Modell B - 1690",      "A","—","—",    5,"Kauf abgeschlossen",        "—",          date(2026,5,15), 2190,"Ja", "—",   "Sofortkauf nach Test"],
    ["2026-0002", date(2026,5,16), "Jonas",   "Hr. Mayr",     "0676 2345678",      "Crosstrainer","Modell X - 1490",     "Modell Y - 990",       "B","Maße","E1", 4,"Heimcheck angelegt",        "Check-Anruf",date(2026,5,30), 1490,"Offen","—",  "Masse bis Freitag"],
    ["2026-0003", date(2026,5,16), "Angie",   "Fr. Kofler",   "fr.kofler@mail.at", "Ergometer",   "Modell E - 1290",     "Modell E light - 890", "C","Partner","E2",3,"Heimcheck angelegt",       "2. Besuch",  date(2026,5,29), 1290,"Offen","—",  "Mann kommt mit"],
    ["2026-0004", date(2026,5,18), "Norbert", "Hr. Stocker",  "0699 3456789",      "Kraftstation","Power-Rack - 2890",   "Kompakt-Station - 1990","B","Raum","E5",4,"Entscheidungscheck geplant","Check-Anruf",date(2026,5,28), 2890,"Offen","—",  "Garage wird gemessen"],
    ["2026-0005", date(2026,5,19), "Daniel",  "Hr. Pichler",  "0664 4567890",      "Laufband",    "Modell A - 1890",     "—",                    "A","Budget","E3",4,"Angebot konkret",           "Check-Anruf",date(2026,5,27), 1890,"Offen","—",  "Finanzierung angefragt"],
    ["2026-0006", date(2026,5,20), "Jonas",   "Fam. Schmid",  "0676 5678901",      "Kraftstation","Multi-Station - 3490","—",                    "E","Raum","E5", 3,"Warten auf Raum",           "Wiedervorlage",date(2026,9,15),3490,"Offen","—",  "Keller-Rohbau im Herbst"],
    ["2026-0007", date(2026,5,21), "Angie",   "Fr. Wieser",   "fr.wieser@mail.at", "Ergometer",   "Modell E - 990",      "—",                    "A","—","—",    5,"Kauf abgeschlossen",        "—",          date(2026,5,21),  990,"Ja", "—",   "Geschenk fuer Ehemann"],
    ["2026-0008", date(2026,5,22), "Norbert", "Hr. Lindner",  "0699 6789012",      "Laufband",    "Modell C - 1290",     "Modell C basic - 790", "D","Budget","E3",2,"Verloren",                  "—",          date(2026,5,26),    0,"Nein","Online","Online guenstiger gekauft"],
    ["2026-0009", date(2026,5,23), "Daniel",  "Fam. Huber",   "0664 7890123",      "Laufband",    "noch offen",          "—",                    "C","Nutzung","E6",2,"Beratung offen",           "Check-Anruf",date(2026,6,2),      0,"Offen","—",  "Noch unsicher welches Modell"],
    ["2026-0010", date(2026,5,24), "Jonas",   "Hr. Brunner",  "0676 8901234",      "Crosstrainer","Modell X - 1490",     "—",                    "B","Maße","E1", 4,"Heimcheck angelegt",        "Check-Anruf",date(2026,5,31), 1490,"Offen","—",  "Deckenhoehe pruefen"],
]

for r_idx, row_data in enumerate(rows, 2):
    bg = C_LIGHT if r_idx % 2 == 0 else C_WHITE
    for c_idx, val in enumerate(row_data, 1):
        cell = ws.cell(row=r_idx, column=c_idx, value=val)
        cell.fill   = fill(bg)
        cell.border = border_thin()
        if isinstance(val, date):
            cell.number_format = "DD.MM.YYYY"
            cell.alignment = align("center")
            cell.font = font(size=9)
        elif c_idx in (1, 3, 9, 11, 12, 16, 17):
            cell.alignment = align("center")
            cell.font = font(size=9)
        else:
            cell.alignment = align("left", wrap=True)
            cell.font = font(size=9)
    # Kaufreife bold wenn >=4
    kr_cell = ws.cell(row=r_idx, column=12)
    if isinstance(kr_cell.value, int) and kr_cell.value >= 4:
        kr_cell.font = font(size=9, bold=True)
    # Ampel-Formel Spalte T (20)
    t = ws.cell(row=r_idx, column=20)
    t.value = (
        f'=IF(OR(M{r_idx}="Kauf abgeschlossen",M{r_idx}="Verloren",'
        f'M{r_idx}="Warten auf Raum"),"o",'
        f'IF(O{r_idx}<TODAY(),"!",'
        f'IF(O{r_idx}<=TODAY()+1,"~","ok")))'
    )
    t.alignment = align("center")
    t.font = font(size=9, bold=True)
    t.fill = fill(bg)
    t.border = border_thin()
    ws.row_dimensions[r_idx].height = 20

# ─── Data Validations ─────────────────────────────────────────────────────────
def add_dv(ws, sqref, formula1):
    dv = DataValidation(type="list", formula1=formula1, allow_blank=True, showErrorMessage=False)
    dv.sqref = sqref
    ws.add_data_validation(dv)

add_dv(ws, "C2:C2000", '"Daniel,Jonas,Angie,Norbert"')
add_dv(ws, "F2:F2000", '"Laufband,Ergometer,Crosstrainer,Rudergeraet,Kraftstation,Sonstiges"')
add_dv(ws, "I2:I2000", '"A,B,C,D,E,F,G,H,I,J"')
add_dv(ws, "K2:K2000", '"E1,E2,E3,E4,E5,E6,E7,-"')
add_dv(ws, "L2:L2000", '"1,2,3,4,5"')
add_dv(ws, "M2:M2000", '"Beratung offen,Heimcheck angelegt,Entscheidungscheck geplant,Angebot konkret,Kauf abgeschlossen,Warten auf Raum,Verloren"')
add_dv(ws, "N2:N2000", '"Check-Anruf,2. Besuch,Wiedervorlage,-"')
add_dv(ws, "Q2:Q2000", '"Ja,Nein,Offen"')
add_dv(ws, "R2:R2000", '"Online,Preis,Kein Bedarf,Konkurrenz,Keine Rueckmeldung,-"')

# ─── Conditional Formatting – Status (M) ─────────────────────────────────────
status_rules = [
    ('M2="Kauf abgeschlossen"',            C_GREEN_BG,  C_GREEN_TXT, False),
    ('M2="Verloren"',                      C_RED_BG,    C_RED_TXT,   False),
    ('M2="Heimcheck angelegt"',            C_BLUE_BG,   "1A237E",    False),
    ('M2="Entscheidungscheck geplant"',    C_YELLOW_BG, "F57F17",    True),
    ('M2="Angebot konkret"',               C_ORANGE_BG, "BF360C",    True),
    ('M2="Warten auf Raum"',               C_GRAY_BG,   "546E7A",    False),
]
for formula, bg_c, txt_c, bold in status_rules:
    ws.conditional_formatting.add("M2:M2000", FormulaRule(
        formula=[formula],
        fill=fill(bg_c),
        font=font(size=9, bold=bold, color=txt_c)
    ))

# Conditional Formatting – Check-Datum (O) – overdue / soon
ws.conditional_formatting.add("O2:O2000", FormulaRule(
    formula=['AND(O2<TODAY(),O2<>"",NOT(OR(M2="Kauf abgeschlossen",M2="Verloren",M2="Warten auf Raum")))'],
    fill=fill(C_RED_BG), font=font(size=9, bold=True, color=C_RED_TXT)
))
ws.conditional_formatting.add("O2:O2000", FormulaRule(
    formula=['AND(O2<=TODAY()+1,O2>=TODAY(),NOT(OR(M2="Kauf abgeschlossen",M2="Verloren",M2="Warten auf Raum")))'],
    fill=fill(C_YELLOW_BG), font=font(size=9, bold=True, color="F57F17")
))

# Conditional Formatting – Ampel-Spalte T
ws.conditional_formatting.add("T2:T2000", FormulaRule(
    formula=['T2="!"'],
    fill=fill(C_RED_BG), font=font(size=9, bold=True, color=C_RED_TXT)
))
ws.conditional_formatting.add("T2:T2000", FormulaRule(
    formula=['T2="~"'],
    fill=fill(C_YELLOW_BG), font=font(size=9, bold=True, color="F57F17")
))
ws.conditional_formatting.add("T2:T2000", FormulaRule(
    formula=['T2="ok"'],
    fill=fill(C_GREEN_BG), font=font(size=9, bold=True, color=C_GREEN_TXT)
))
ws.conditional_formatting.add("T2:T2000", FormulaRule(
    formula=['T2="o"'],
    fill=fill(C_GRAY_BG), font=font(size=9, color="607D8B")
))

# Kaufreife-Ampel (L)
ws.conditional_formatting.add("L2:L2000", FormulaRule(
    formula=['L2=5'], fill=fill(C_GREEN_BG), font=font(size=9, bold=True, color=C_GREEN_TXT)
))
ws.conditional_formatting.add("L2:L2000", FormulaRule(
    formula=['L2=4'], fill=fill(C_ORANGE_BG), font=font(size=9, bold=True)
))
ws.conditional_formatting.add("L2:L2000", FormulaRule(
    formula=['L2<=2'], fill=fill(C_GRAY_BG), font=font(size=9, color="888888")
))

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 2: KPI-Dashboard
# ══════════════════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("KPI Dashboard")
ws2.sheet_view.showGridLines = False

col_w2 = [30, 16, 14, 14, 22]
col_names2 = ["", "Gesamt", "Daniel", "Jonas", "Angie + Norbert"]
for i, w in enumerate(col_w2, 1):
    ws2.column_dimensions[get_column_letter(i)].width = w

# Titel
for c in range(1, 6):
    ws2.cell(row=1, column=c).fill = fill(C_RED)
ws2.merge_cells("A1:E1")
t = ws2.cell(row=1, column=1, value="FITSHOP HALL/TIROL  ·  KPI DASHBOARD 2026")
t.font = font(size=14, bold=True, color=C_WHITE)
t.alignment = align("left", "center")
ws2.row_dimensions[1].height = 36

sub = ws2.cell(row=2, column=1, value="Werte werden automatisch aus Lead-Tracking 2026 berechnet  |  Ampel: ! = ueberfaellig  ~ = heute/morgen  ok = im Plan  o = abgeschlossen")
sub.font = font(size=8, italic=True, color="888888")
ws2.merge_cells("A2:E2")
ws2.row_dimensions[2].height = 18

# Spaltenheader
ws2.row_dimensions[4].height = 24
for c, h in enumerate(col_names2, 1):
    cell = ws2.cell(row=4, column=c, value=h)
    cell.fill = fill(C_ANTHRAZIT)
    cell.font = font(bold=True, color=C_WHITE, size=10)
    cell.alignment = align("center", "center")
    cell.border = border_thin()

TN = "'Lead-Tracking 2026'"

kpis2 = [
    # label, gesamt, daniel, jonas, angie+norbert, number_format, highlight_if_positive
    ("Leads gesamt",
     f'=COUNTA({TN}!A2:A5000)',
     f'=COUNTIF({TN}!C2:C5000,"Daniel")',
     f'=COUNTIF({TN}!C2:C5000,"Jonas")',
     f'=COUNTIFS({TN}!C2:C5000,"Angie")+COUNTIFS({TN}!C2:C5000,"Norbert")',
     "0", False),
    ("Abschluesse (Kauf Ja)",
     f'=COUNTIF({TN}!Q2:Q5000,"Ja")',
     f'=COUNTIFS({TN}!C2:C5000,"Daniel",{TN}!Q2:Q5000,"Ja")',
     f'=COUNTIFS({TN}!C2:C5000,"Jonas",{TN}!Q2:Q5000,"Ja")',
     f'=COUNTIFS({TN}!C2:C5000,"Angie",{TN}!Q2:Q5000,"Ja")+COUNTIFS({TN}!C2:C5000,"Norbert",{TN}!Q2:Q5000,"Ja")',
     "0", False),
    ("Abschlussquote",
     '=IF(B5=0,"—",B6/B5)',
     '=IF(C5=0,"—",C6/C5)',
     '=IF(D5=0,"—",D6/D5)',
     '=IF(E5=0,"—",E6/E5)',
     "0%", False),
    ("Offene Leads",
     f'=COUNTIF({TN}!Q2:Q5000,"Offen")',
     f'=COUNTIFS({TN}!C2:C5000,"Daniel",{TN}!Q2:Q5000,"Offen")',
     f'=COUNTIFS({TN}!C2:C5000,"Jonas",{TN}!Q2:Q5000,"Offen")',
     f'=COUNTIFS({TN}!C2:C5000,"Angie",{TN}!Q2:Q5000,"Offen")+COUNTIFS({TN}!C2:C5000,"Norbert",{TN}!Q2:Q5000,"Offen")',
     "0", False),
    ("Verluste",
     f'=COUNTIF({TN}!Q2:Q5000,"Nein")',
     f'=COUNTIFS({TN}!C2:C5000,"Daniel",{TN}!Q2:Q5000,"Nein")',
     f'=COUNTIFS({TN}!C2:C5000,"Jonas",{TN}!Q2:Q5000,"Nein")',
     f'=COUNTIFS({TN}!C2:C5000,"Angie",{TN}!Q2:Q5000,"Nein")+COUNTIFS({TN}!C2:C5000,"Norbert",{TN}!Q2:Q5000,"Nein")',
     "0", False),
    ("Heimchecks angelegt",
     f'=COUNTIF({TN}!M2:M5000,"Heimcheck angelegt")',
     f'=COUNTIFS({TN}!C2:C5000,"Daniel",{TN}!M2:M5000,"Heimcheck angelegt")',
     f'=COUNTIFS({TN}!C2:C5000,"Jonas",{TN}!M2:M5000,"Heimcheck angelegt")',
     f'=COUNTIFS({TN}!C2:C5000,"Angie",{TN}!M2:M5000,"Heimcheck angelegt")+COUNTIFS({TN}!C2:C5000,"Norbert",{TN}!M2:M5000,"Heimcheck angelegt")',
     "0", False),
    ("Umsatz realisiert (EUR)",
     f'=SUMIF({TN}!Q2:Q5000,"Ja",{TN}!P2:P5000)',
     f'=SUMIFS({TN}!P2:P5000,{TN}!C2:C5000,"Daniel",{TN}!Q2:Q5000,"Ja")',
     f'=SUMIFS({TN}!P2:P5000,{TN}!C2:C5000,"Jonas",{TN}!Q2:Q5000,"Ja")',
     f'=SUMIFS({TN}!P2:P5000,{TN}!C2:C5000,"Angie",{TN}!Q2:Q5000,"Ja")+SUMIFS({TN}!P2:P5000,{TN}!C2:C5000,"Norbert",{TN}!Q2:Q5000,"Ja")',
     '#,##0 "EUR"', False),
    ("Umsatz Pipeline offen (EUR)",
     f'=SUMIF({TN}!Q2:Q5000,"Offen",{TN}!P2:P5000)',
     f'=SUMIFS({TN}!P2:P5000,{TN}!C2:C5000,"Daniel",{TN}!Q2:Q5000,"Offen")',
     f'=SUMIFS({TN}!P2:P5000,{TN}!C2:C5000,"Jonas",{TN}!Q2:Q5000,"Offen")',
     f'=SUMIFS({TN}!P2:P5000,{TN}!C2:C5000,"Angie",{TN}!Q2:Q5000,"Offen")+SUMIFS({TN}!P2:P5000,{TN}!C2:C5000,"Norbert",{TN}!Q2:Q5000,"Offen")',
     '#,##0 "EUR"', False),
    ("Durchschn. Angebotswert (EUR)",
     f'=IFERROR(AVERAGEIF({TN}!P2:P5000,">"&0),"—")',
     '=IFERROR(AVERAGEIFS(\'Lead-Tracking 2026\'!P2:P5000,\'Lead-Tracking 2026\'!C2:C5000,"Daniel",\'Lead-Tracking 2026\'!P2:P5000,">"&0),"—")',
     '=IFERROR(AVERAGEIFS(\'Lead-Tracking 2026\'!P2:P5000,\'Lead-Tracking 2026\'!C2:C5000,"Jonas",\'Lead-Tracking 2026\'!P2:P5000,">"&0),"—")',
     '=IFERROR(AVERAGEIFS(\'Lead-Tracking 2026\'!P2:P5000,\'Lead-Tracking 2026\'!C2:C5000,"Angie",\'Lead-Tracking 2026\'!P2:P5000,">"&0)+AVERAGEIFS(\'Lead-Tracking 2026\'!P2:P5000,\'Lead-Tracking 2026\'!C2:C5000,"Norbert",\'Lead-Tracking 2026\'!P2:P5000,">"&0),"—")',
     '#,##0 "EUR"', False),
    ("Ueberfaellige Leads (!)",
     f'=COUNTIFS({TN}!O2:O5000,"<"&TODAY(),{TN}!Q2:Q5000,"Offen")',
     f'=COUNTIFS({TN}!C2:C5000,"Daniel",{TN}!O2:O5000,"<"&TODAY(),{TN}!Q2:Q5000,"Offen")',
     f'=COUNTIFS({TN}!C2:C5000,"Jonas",{TN}!O2:O5000,"<"&TODAY(),{TN}!Q2:Q5000,"Offen")',
     f'=COUNTIFS({TN}!C2:C5000,"Angie",{TN}!O2:O5000,"<"&TODAY(),{TN}!Q2:Q5000,"Offen")+COUNTIFS({TN}!C2:C5000,"Norbert",{TN}!O2:O5000,"<"&TODAY(),{TN}!Q2:Q5000,"Offen")',
     "0", True),
]

for i, (label, g, d, j, an, nfmt, highlight) in enumerate(kpis2, 5):
    bg = C_LIGHT if i % 2 == 0 else C_WHITE
    lbl_cell = ws2.cell(row=i, column=1, value=label)
    lbl_cell.fill = fill(bg)
    lbl_cell.font = font(size=10, bold=True)
    lbl_cell.alignment = align("left", "center")
    lbl_cell.border = border_thin()
    for c_idx, val in enumerate([g, d, j, an], 2):
        cell = ws2.cell(row=i, column=c_idx, value=val)
        cell.fill = fill(bg)
        cell.font = font(size=10)
        cell.alignment = align("center", "center")
        cell.border = border_thin()
        if nfmt != "0":
            cell.number_format = nfmt
    ws2.row_dimensions[i].height = 22

# Highlight Ueberfaellige row
ueberfaellig_row = 5 + len(kpis2) - 1
ws2.conditional_formatting.add(f"B{ueberfaellig_row}:E{ueberfaellig_row}", FormulaRule(
    formula=[f'B{ueberfaellig_row}>0'],
    fill=fill(C_RED_BG), font=font(size=10, bold=True, color=C_RED_TXT)
))
ws2.conditional_formatting.add(f"B7:E7", FormulaRule(
    formula=['ISNUMBER(B7)'], fill=fill(C_GREEN_BG), font=font(size=10, bold=True, color=C_GREEN_TXT)
))

# ─── Trennlinie + Verlustgrund-Analyse ────────────────────────────────────────
sep_row = 5 + len(kpis2) + 2
for c in range(1, 6):
    cell = ws2.cell(row=sep_row, column=c)
    cell.fill = fill(C_ANTHRAZIT)
    cell.font = font(bold=True, color=C_WHITE, size=10)
cell_lbl = ws2.cell(row=sep_row, column=1, value="VERLUSTANALYSE")
cell_lbl.alignment = align("left", "center")
ws2.merge_cells(f"A{sep_row}:E{sep_row}")
ws2.row_dimensions[sep_row].height = 22

verlust_gruende = ["Online", "Preis", "Kein Bedarf", "Konkurrenz", "Keine Rueckmeldung"]
for vi, vg in enumerate(verlust_gruende, sep_row + 1):
    bg = C_LIGHT if vi % 2 == 0 else C_WHITE
    ws2.cell(row=vi, column=1, value=vg).fill = fill(bg)
    ws2.cell(row=vi, column=1).font = font(size=10)
    ws2.cell(row=vi, column=1).alignment = align("left")
    ws2.cell(row=vi, column=1).border = border_thin()
    formula = f'=COUNTIF({TN}!R2:R5000,A{vi})'
    c2 = ws2.cell(row=vi, column=2, value=formula)
    c2.fill = fill(bg); c2.font = font(size=10); c2.alignment = align("center"); c2.border = border_thin()
    ws2.merge_cells(f"C{vi}:E{vi}")
    ws2.cell(row=vi, column=3).fill = fill(bg)
    ws2.row_dimensions[vi].height = 20

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 3: Legende & Referenz
# ══════════════════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("Legende")
ws3.sheet_view.showGridLines = False
ws3.column_dimensions["A"].width = 22
ws3.column_dimensions["B"].width = 38
ws3.column_dimensions["C"].width = 30
ws3.column_dimensions["D"].width = 5

# Titel
ws3.merge_cells("A1:C1")
t3 = ws3.cell(row=1, column=1, value="FITSHOP TRACKING SYSTEM – LEGENDE 2026")
t3.fill = fill(C_RED); t3.font = font(size=13, bold=True, color=C_WHITE)
t3.alignment = align("left", "center"); ws3.row_dimensions[1].height = 32

def section_header(ws, row, text, span_end="C"):
    ws.merge_cells(f"A{row}:{span_end}{row}")
    c = ws.cell(row=row, column=1, value=text)
    c.fill = fill(C_ANTHRAZIT); c.font = font(bold=True, color=C_WHITE, size=10)
    c.alignment = align("left", "center"); ws.row_dimensions[row].height = 22

def tbl_row(ws, row, col1, col2, col3="", alternate=False):
    bg = C_LIGHT if alternate else C_WHITE
    for c, val in enumerate([col1, col2, col3], 1):
        cell = ws.cell(row=row, column=c, value=val)
        cell.fill = fill(bg); cell.font = font(size=9)
        cell.alignment = align("left", "center", wrap=True)
        cell.border = border_thin()
    ws.row_dimensions[row].height = 18

# Kundentypen
section_header(ws3, 3, "KUNDENTYPEN (Spalte Typ)")
kundentypen = [
    ("A – Sofortkaeufer",     "Kauft heute, kein Heimcheck noetig",         "Direktabschluss"),
    ("B – Logischer Pruefer", "Will Fakten, muss Masse nehmen",             "Heimcheck vereinbaren"),
    ("C – Unsicherer",        "Braucht Bestaetigung oder Partner",          "Partner einladen / 2. Besuch"),
    ("D – Preisfokussierter", "Vergleicht online, Budget-sensibel",         "Mehrwert argumentieren"),
    ("E – Raumplaner",        "Raum noch nicht fertig",                     "Wiedervorlage in X Monaten"),
    ("F – Gesundheitskunde",  "Arzt-/Reha-Empfehlung",                     "Empfehlung aufgreifen"),
    ("G – Geschenkekaeufer",  "Kauft fuer jemand anderen",                  "Direkt zum Abschluss"),
    ("H – Wiederkaeeufer",    "Kennt Fitshop schon",                        "Upgrade / Ergaenzung"),
    ("I – Profi/Enthusiast",  "Sehr technisch, viel Ahnung",                "Auf Augenhoehe beraten"),
    ("J – B2B/Verein",        "Unternehmen, Hotel, Schule, Verein",         "Angebot + Termin"),
]
for i, (a, b, c) in enumerate(kundentypen, 4):
    tbl_row(ws3, i, a, b, c, i % 2 == 0)

# Einwandtypen
section_header(ws3, 15, "EINWANDTYPEN (Spalte Einwand)")
einwaende = [
    ("E1 – Masse unklar",        "Raum/Flaeche muss gemessen werden",      "Heimcheck vereinbaren"),
    ("E2 – Partner fehlt",       "Entscheidung nicht alleine moeglich",    "2. Besuch mit Partner"),
    ("E3 – Budget",              "Preis ist Huerden, Finanzierung pruefen","Finanzierung zeigen"),
    ("E4 – Qualitaet/Marke",     "Zweifel an Produkt oder Marke",          "Testfahrt, Garantie"),
    ("E5 – Raum nicht fertig",   "Umbau / Renovierung laeuft noch",        "Wiedervorlage Herbst"),
    ("E6 – Nutzungskomfort",     "Unsicher ob Geraet zum Training passt",  "Nutzungsanalyse vertiefen"),
    ("E7 – Keine Dringlichkeit", "Kein Zeitdruck, kein Anlass",            "Anlass schaffen / ankern"),
]
for i, (a, b, c) in enumerate(einwaende, 16):
    tbl_row(ws3, i, a, b, c, i % 2 == 0)

# Status-Legende
section_header(ws3, 24, "STATUS-AMPEL (Spalte A / Check-Datum)")
ampel_rows = [
    ("ok  (gruen)",  "Check-Datum liegt in > 2 Tagen – alles im Plan"),
    ("~   (gelb)",   "Check-Datum heute oder morgen – sofort handeln"),
    ("!   (rot)",    "Check-Datum ueberschritten – dringend nachfassen!"),
    ("o   (grau)",   "Abgeschlossen / Verloren / Wartet – kein Follow-up"),
]
ampel_fills = [C_GREEN_BG, C_YELLOW_BG, C_RED_BG, C_GRAY_BG]
for i, ((sym, txt), af) in enumerate(zip(ampel_rows, ampel_fills), 25):
    c1 = ws3.cell(row=i, column=1, value=sym)
    c1.fill = fill(af); c1.font = font(size=9, bold=True); c1.border = border_thin()
    c1.alignment = align("center", "center")
    c2 = ws3.cell(row=i, column=2, value=txt)
    c2.fill = fill(af); c2.font = font(size=9); c2.border = border_thin()
    ws3.merge_cells(f"B{i}:C{i}")
    ws3.row_dimensions[i].height = 18

# Status-Farbkodierung
section_header(ws3, 30, "STATUS-FARBKODIERUNG (Spalte Status)")
status_farben = [
    ("Kauf abgeschlossen", "Gruen",   "Abgeschlossen, kein Follow-up"),
    ("Heimcheck angelegt", "Blau",    "Naechster Schritt: Check-Anruf"),
    ("Entscheidungscheck geplant", "Gelb",   "Kurz vor Abschluss"),
    ("Angebot konkret",    "Orange",  "Angebot liegt vor, nachfassen"),
    ("Verloren",           "Rot",     "Lead verloren, Verlustgrund eintragen"),
    ("Warten auf Raum",    "Grau",    "Wiedervorlage in Monaten"),
]
sf_fills = [C_GREEN_BG, C_BLUE_BG, C_YELLOW_BG, C_ORANGE_BG, C_RED_BG, C_GRAY_BG]
for i, ((stat, farbe, erkl), sf) in enumerate(zip(status_farben, sf_fills), 31):
    c1 = ws3.cell(row=i, column=1, value=stat)
    c1.fill = fill(sf); c1.font = font(size=9, bold=True); c1.border = border_thin()
    c2 = ws3.cell(row=i, column=2, value=farbe)
    c2.fill = fill(sf); c2.font = font(size=9); c2.border = border_thin()
    c3 = ws3.cell(row=i, column=3, value=erkl)
    c3.fill = fill(sf); c3.font = font(size=9); c3.border = border_thin()
    ws3.row_dimensions[i].height = 18

# ─── Sheet-Reihenfolge ────────────────────────────────────────────────────────
wb.active = ws  # Lead-Tracking als erstes aktiv

# ─── Speichern ────────────────────────────────────────────────────────────────
path = "/home/user/My-Test-repository/fitshop-sop/Fitshop_Tracking_2026.xlsx"
wb.save(path)
print(f"Saved: {path}")
print(f"Sheets: {[s.title for s in wb.worksheets]}")
