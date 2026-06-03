#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

# ─── Farben ───────────────────────────────────────────────────────────────────
C_RED        = "E31E24"
C_RED_DARK   = "B71C1C"
C_RED_LIGHT  = "FFEBEE"
C_WHITE      = "FFFFFF"
C_OFFWHITE   = "FAFAFA"
C_ANTHRAZIT  = "2D2D2D"
C_GRAY_LIGHT = "F5F5F5"
C_GRAY_MID   = "E0E0E0"
C_GREEN_BG   = "C8E6C9"
C_GREEN_TXT  = "1B5E20"
C_RED_BG     = "FFCDD2"
C_RED_TXT    = "B71C1C"
C_YELLOW_BG  = "FFF9C4"
C_YELLOW_TXT = "F57F17"
C_BLUE_BG    = "E3F2FD"
C_BLUE_TXT   = "0D47A1"
C_ORANGE_BG  = "FFE0B2"
C_ORANGE_TXT = "BF360C"
C_GRAY_BG    = "ECEFF1"
C_GRAY_TXT   = "546E7A"

def fill(color):
    return PatternFill(start_color=color, end_color=color, fill_type="solid")

def font(name="Calibri", size=10, bold=False, color="000000", italic=False):
    return Font(name=name, size=size, bold=bold, color=color, italic=italic)

def align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def thin_border(color="D0D0D0"):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def header_border():
    s = Side(style="medium", color=C_RED)
    return Border(bottom=s)

def add_dv(ws, sqref, formula1):
    dv = DataValidation(type="list", formula1=formula1, allow_blank=True, showErrorMessage=False)
    dv.sqref = sqref
    ws.add_data_validation(dv)

# ─── Workbook ─────────────────────────────────────────────────────────────────
wb = openpyxl.Workbook()

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 1 – Lead-Tracking 2026
# ══════════════════════════════════════════════════════════════════════════════
ws = wb.active
ws.title = "Lead-Tracking 2026"
ws.sheet_view.showGridLines = True

# Zeile 1: Fitshop-Banner
ws.merge_cells("A1:T1")
banner = ws.cell(row=1, column=1, value="FITSHOP")
banner.fill = fill(C_RED)
banner.font = font(size=20, bold=True, color=C_WHITE, name="Calibri")
banner.alignment = align("left", "center")
ws.row_dimensions[1].height = 38

# Zeile 2: Untertitel
ws.merge_cells("A2:T2")
sub = ws.cell(row=2, column=1, value="LEAD-TRACKING 2026  ·  Hall/Tirol")
sub.fill = fill(C_ANTHRAZIT)
sub.font = font(size=10, bold=False, color="AAAAAA", italic=True)
sub.alignment = align("left", "center")
ws.row_dimensions[2].height = 18

# Zeile 3: Spaltenheader
headers = [
    ("Lead-Nr",              10),
    ("Datum",                12),
    ("Mitarbeiter",          12),
    ("Kunde",                18),
    ("Kontakt",              17),
    ("Kategorie",            14),
    ("Favorit (Modell+Preis)", 24),
    ("Alternative",          20),
    ("Typ",                   6),
    ("Offener Punkt",        15),
    ("Einwand",              10),
    ("KR",                    5),
    ("Status",               22),
    ("Naechster Schritt",    16),
    ("Check-Datum",          13),
    ("Angebotswert EUR",     15),
    ("Abschluss",            11),
    ("Verlustgrund",         14),
    ("Notiz",                28),
    ("A",                     4),
]

for col, (h, w) in enumerate(headers, 1):
    cell = ws.cell(row=3, column=col, value=h)
    cell.fill  = fill(C_ANTHRAZIT)
    cell.font  = font(bold=True, color=C_WHITE, size=9)
    cell.alignment = align("center", "center", wrap=True)
    cell.border = thin_border("404040")
    ws.column_dimensions[get_column_letter(col)].width = w

ws.row_dimensions[3].height = 26

# Ampel-Formel + leere Zeilen (4-203)
for r in range(4, 204):
    # Ampel-Formel in Spalte T
    t = ws.cell(row=r, column=20)
    t.value = (
        f'=IF(M{r}="","",IF(OR(M{r}="Kauf abgeschlossen",M{r}="Verloren",'
        f'M{r}="Warten auf Raum"),"o",'
        f'IF(O{r}="","",IF(O{r}<TODAY(),"!",'
        f'IF(O{r}<=TODAY()+1,"~","ok")))))'
    )
    t.alignment = align("center")
    t.font = font(size=9, bold=True)
    bg = C_GRAY_LIGHT if r % 2 == 0 else C_WHITE
    t.fill = fill(bg)
    t.border = thin_border()
    ws.row_dimensions[r].height = 20
    # Hellgrau/weiß abwechselnd für leere Zellen
    for c in range(1, 20):
        cell = ws.cell(row=r, column=c)
        cell.fill = fill(bg)
        cell.border = thin_border()
        cell.font = font(size=9)
        if c in (1, 3, 9, 11, 12, 15, 16, 17):
            cell.alignment = align("center")
        else:
            cell.alignment = align("left", wrap=True)

# Freeze ab Zeile 4 (unter Banner + Spaltenheader)
ws.freeze_panes = "A4"
ws.auto_filter.ref = f"A3:{get_column_letter(len(headers))}3"

# ─── Data Validations ─────────────────────────────────────────────────────────
add_dv(ws, "C4:C2000", '"Daniel,Jonas,Angie,Norbert"')
add_dv(ws, "F4:F2000", '"Laufband,Ergometer,Crosstrainer,Rudergeraet,Kraftstation,Sonstiges"')
add_dv(ws, "I4:I2000", '"A,B,C,D,E,F,G,H,I,J"')
add_dv(ws, "K4:K2000",
    '"E1 Masse,E2 Partner,E3 Budget,E4 Qualitaet,E5 Raum,E6 Nutzung,E7 Dringlichkeit,'
    'E8 Ausmessen,E9 Frau fragen,E10 Passt es rein,-"')
add_dv(ws, "L4:L2000", '"1,2,3,4,5"')
add_dv(ws, "M4:M2000",
    '"Beratung offen,Heimcheck angelegt,Entscheidungscheck geplant,'
    'Angebot konkret,Kauf abgeschlossen,Warten auf Raum,Verloren"')
add_dv(ws, "N4:N2000", '"Check-Anruf,2. Besuch,Wiedervorlage,-"')
add_dv(ws, "Q4:Q2000", '"Ja,Nein,Offen"')
add_dv(ws, "R4:R2000", '"Online,Preis,Kein Bedarf,Konkurrenz,Keine Rueckmeldung,-"')

# ─── Conditional Formatting ───────────────────────────────────────────────────
# Status (M)
for formula, bg_c, txt_c, bold in [
    ('M4="Kauf abgeschlossen"',          C_GREEN_BG,  C_GREEN_TXT,  False),
    ('M4="Verloren"',                    C_RED_BG,    C_RED_TXT,    False),
    ('M4="Heimcheck angelegt"',          C_BLUE_BG,   C_BLUE_TXT,   False),
    ('M4="Entscheidungscheck geplant"',  C_YELLOW_BG, C_YELLOW_TXT, True),
    ('M4="Angebot konkret"',             C_ORANGE_BG, C_ORANGE_TXT, True),
    ('M4="Warten auf Raum"',             C_GRAY_BG,   C_GRAY_TXT,   False),
]:
    ws.conditional_formatting.add("M4:M2000", FormulaRule(
        formula=[formula], fill=fill(bg_c), font=font(size=9, bold=bold, color=txt_c)))

# Check-Datum (O)
ws.conditional_formatting.add("O4:O2000", FormulaRule(
    formula=['AND(O4<TODAY(),O4<>"",NOT(OR(M4="Kauf abgeschlossen",M4="Verloren",M4="Warten auf Raum")))'],
    fill=fill(C_RED_BG), font=font(size=9, bold=True, color=C_RED_TXT)))
ws.conditional_formatting.add("O4:O2000", FormulaRule(
    formula=['AND(O4<=TODAY()+1,O4>=TODAY(),NOT(OR(M4="Kauf abgeschlossen",M4="Verloren",M4="Warten auf Raum")))'],
    fill=fill(C_YELLOW_BG), font=font(size=9, bold=True, color=C_YELLOW_TXT)))

# Ampel (T)
for sym, bg_c, txt_c in [("!", C_RED_BG, C_RED_TXT), ("~", C_YELLOW_BG, C_YELLOW_TXT),
                          ("ok", C_GREEN_BG, C_GREEN_TXT), ("o", C_GRAY_BG, C_GRAY_TXT)]:
    ws.conditional_formatting.add("T4:T2000", FormulaRule(
        formula=[f'T4="{sym}"'], fill=fill(bg_c), font=font(size=9, bold=True, color=txt_c)))

# Kaufreife (L)
ws.conditional_formatting.add("L4:L2000", FormulaRule(
    formula=['L4=5'], fill=fill(C_GREEN_BG), font=font(size=9, bold=True, color=C_GREEN_TXT)))
ws.conditional_formatting.add("L4:L2000", FormulaRule(
    formula=['L4=4'], fill=fill(C_ORANGE_BG), font=font(size=9, bold=True)))
ws.conditional_formatting.add("L4:L2000", FormulaRule(
    formula=['L4<=2'], fill=fill(C_GRAY_BG), font=font(size=9, color="888888")))

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 2 – KPI Dashboard
# ══════════════════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("KPI Dashboard")
ws2.sheet_view.showGridLines = False

# Spaltenbreiten: Label + Gesamt + Daniel + Jonas + Angie + Norbert
col_widths2 = [30, 14, 14, 14, 14, 14]
col_labels2 = ["", "Gesamt", "Daniel", "Jonas", "Angie", "Norbert"]
for i, w in enumerate(col_widths2, 1):
    ws2.column_dimensions[get_column_letter(i)].width = w

# Banner
ws2.merge_cells("A1:F1")
b2 = ws2.cell(row=1, column=1, value="FITSHOP")
b2.fill = fill(C_RED); b2.font = font(size=20, bold=True, color=C_WHITE)
b2.alignment = align("left", "center"); ws2.row_dimensions[1].height = 38

ws2.merge_cells("A2:F2")
s2 = ws2.cell(row=2, column=1, value="KPI DASHBOARD 2026  ·  Hall/Tirol")
s2.fill = fill(C_ANTHRAZIT); s2.font = font(size=10, color="AAAAAA", italic=True)
s2.alignment = align("left", "center"); ws2.row_dimensions[2].height = 18

ws2.merge_cells("A3:F3")
h3 = ws2.cell(row=3, column=1,
    value="Werte werden automatisch aus dem Lead-Tracking berechnet. Seite aktualisieren mit Strg+Alt+F9")
h3.font = font(size=8, italic=True, color="999999")
h3.alignment = align("left", "center"); ws2.row_dimensions[3].height = 16

# Spaltenheader (Zeile 4)
for c, lbl in enumerate(col_labels2, 1):
    cell = ws2.cell(row=4, column=c, value=lbl)
    cell.fill = fill(C_RED) if c > 1 else fill(C_ANTHRAZIT)
    cell.font = font(bold=True, color=C_WHITE, size=10)
    cell.alignment = align("center", "center")
    cell.border = thin_border("CC0000" if c > 1 else "404040")
ws2.row_dimensions[4].height = 24

TN = "'Lead-Tracking 2026'"

def cf(col, name):
    return f'=COUNTIFS({TN}!C4:C2000,"{name}",{TN}!{col}4:{col}2000,"Ja")'

kpis2 = [
    # (label, gesamt_formula, daniel, jonas, angie, norbert, number_format, red_if_positive)
    ("Leads gesamt",
     f'=COUNTA({TN}!A4:A2000)',
     f'=COUNTIF({TN}!C4:C2000,"Daniel")',
     f'=COUNTIF({TN}!C4:C2000,"Jonas")',
     f'=COUNTIF({TN}!C4:C2000,"Angie")',
     f'=COUNTIF({TN}!C4:C2000,"Norbert")',
     "0", False),
    ("Abschluesse (Kauf Ja)",
     f'=COUNTIF({TN}!Q4:Q2000,"Ja")',
     f'=COUNTIFS({TN}!C4:C2000,"Daniel",{TN}!Q4:Q2000,"Ja")',
     f'=COUNTIFS({TN}!C4:C2000,"Jonas",{TN}!Q4:Q2000,"Ja")',
     f'=COUNTIFS({TN}!C4:C2000,"Angie",{TN}!Q4:Q2000,"Ja")',
     f'=COUNTIFS({TN}!C4:C2000,"Norbert",{TN}!Q4:Q2000,"Ja")',
     "0", False),
    ("Abschlussquote",
     '=IF(B5=0,"—",B6/B5)',
     '=IF(C5=0,"—",C6/C5)',
     '=IF(D5=0,"—",D6/D5)',
     '=IF(E5=0,"—",E6/E5)',
     '=IF(F5=0,"—",F6/F5)',
     "0%", False),
    ("Offene Leads",
     f'=COUNTIF({TN}!Q4:Q2000,"Offen")',
     f'=COUNTIFS({TN}!C4:C2000,"Daniel",{TN}!Q4:Q2000,"Offen")',
     f'=COUNTIFS({TN}!C4:C2000,"Jonas",{TN}!Q4:Q2000,"Offen")',
     f'=COUNTIFS({TN}!C4:C2000,"Angie",{TN}!Q4:Q2000,"Offen")',
     f'=COUNTIFS({TN}!C4:C2000,"Norbert",{TN}!Q4:Q2000,"Offen")',
     "0", False),
    ("Verluste",
     f'=COUNTIF({TN}!Q4:Q2000,"Nein")',
     f'=COUNTIFS({TN}!C4:C2000,"Daniel",{TN}!Q4:Q2000,"Nein")',
     f'=COUNTIFS({TN}!C4:C2000,"Jonas",{TN}!Q4:Q2000,"Nein")',
     f'=COUNTIFS({TN}!C4:C2000,"Angie",{TN}!Q4:Q2000,"Nein")',
     f'=COUNTIFS({TN}!C4:C2000,"Norbert",{TN}!Q4:Q2000,"Nein")',
     "0", False),
    ("Heimchecks angelegt",
     f'=COUNTIF({TN}!M4:M2000,"Heimcheck angelegt")',
     f'=COUNTIFS({TN}!C4:C2000,"Daniel",{TN}!M4:M2000,"Heimcheck angelegt")',
     f'=COUNTIFS({TN}!C4:C2000,"Jonas",{TN}!M4:M2000,"Heimcheck angelegt")',
     f'=COUNTIFS({TN}!C4:C2000,"Angie",{TN}!M4:M2000,"Heimcheck angelegt")',
     f'=COUNTIFS({TN}!C4:C2000,"Norbert",{TN}!M4:M2000,"Heimcheck angelegt")',
     "0", False),
    ("Umsatz realisiert (EUR)",
     f'=SUMIF({TN}!Q4:Q2000,"Ja",{TN}!P4:P2000)',
     f'=SUMIFS({TN}!P4:P2000,{TN}!C4:C2000,"Daniel",{TN}!Q4:Q2000,"Ja")',
     f'=SUMIFS({TN}!P4:P2000,{TN}!C4:C2000,"Jonas",{TN}!Q4:Q2000,"Ja")',
     f'=SUMIFS({TN}!P4:P2000,{TN}!C4:C2000,"Angie",{TN}!Q4:Q2000,"Ja")',
     f'=SUMIFS({TN}!P4:P2000,{TN}!C4:C2000,"Norbert",{TN}!Q4:Q2000,"Ja")',
     '#,##0 "EUR"', False),
    ("Umsatz Pipeline offen (EUR)",
     f'=SUMIF({TN}!Q4:Q2000,"Offen",{TN}!P4:P2000)',
     f'=SUMIFS({TN}!P4:P2000,{TN}!C4:C2000,"Daniel",{TN}!Q4:Q2000,"Offen")',
     f'=SUMIFS({TN}!P4:P2000,{TN}!C4:C2000,"Jonas",{TN}!Q4:Q2000,"Offen")',
     f'=SUMIFS({TN}!P4:P2000,{TN}!C4:C2000,"Angie",{TN}!Q4:Q2000,"Offen")',
     f'=SUMIFS({TN}!P4:P2000,{TN}!C4:C2000,"Norbert",{TN}!Q4:Q2000,"Offen")',
     '#,##0 "EUR"', False),
    ("Durchschn. Angebotswert (EUR)",
     f'=IFERROR(AVERAGEIF({TN}!P4:P2000,">"&0),"—")',
     f'=IFERROR(AVERAGEIFS({TN}!P4:P2000,{TN}!C4:C2000,"Daniel",{TN}!P4:P2000,">"&0),"—")',
     f'=IFERROR(AVERAGEIFS({TN}!P4:P2000,{TN}!C4:C2000,"Jonas",{TN}!P4:P2000,">"&0),"—")',
     f'=IFERROR(AVERAGEIFS({TN}!P4:P2000,{TN}!C4:C2000,"Angie",{TN}!P4:P2000,">"&0),"—")',
     f'=IFERROR(AVERAGEIFS({TN}!P4:P2000,{TN}!C4:C2000,"Norbert",{TN}!P4:P2000,">"&0),"—")',
     '#,##0 "EUR"', False),
    ("Ueberfaellige Leads (!)",
     f'=COUNTIFS({TN}!O4:O2000,"<"&TODAY(),{TN}!Q4:Q2000,"Offen")',
     f'=COUNTIFS({TN}!C4:C2000,"Daniel",{TN}!O4:O2000,"<"&TODAY(),{TN}!Q4:Q2000,"Offen")',
     f'=COUNTIFS({TN}!C4:C2000,"Jonas",{TN}!O4:O2000,"<"&TODAY(),{TN}!Q4:Q2000,"Offen")',
     f'=COUNTIFS({TN}!C4:C2000,"Angie",{TN}!O4:O2000,"<"&TODAY(),{TN}!Q4:Q2000,"Offen")',
     f'=COUNTIFS({TN}!C4:C2000,"Norbert",{TN}!O4:O2000,"<"&TODAY(),{TN}!Q4:Q2000,"Offen")',
     "0", True),
]

for i, (label, g, d, j, a, n, nfmt, red_flag) in enumerate(kpis2, 5):
    bg = C_GRAY_LIGHT if i % 2 == 0 else C_WHITE
    lbl = ws2.cell(row=i, column=1, value=label)
    lbl.fill = fill(bg); lbl.font = font(size=10, bold=True)
    lbl.alignment = align("left", "center"); lbl.border = thin_border()
    for c_idx, val in enumerate([g, d, j, a, n], 2):
        cell = ws2.cell(row=i, column=c_idx, value=val)
        cell.fill = fill(bg); cell.font = font(size=10)
        cell.alignment = align("center", "center"); cell.border = thin_border()
        if nfmt != "0":
            cell.number_format = nfmt
    ws2.row_dimensions[i].height = 22

# Überfällige Leads rot markieren (letzte Zeile = 5 + len(kpis2) - 1)
last_kpi_row = 5 + len(kpis2) - 1
ws2.conditional_formatting.add(f"B{last_kpi_row}:F{last_kpi_row}", FormulaRule(
    formula=[f'B{last_kpi_row}>0'],
    fill=fill(C_RED_BG), font=font(size=10, bold=True, color=C_RED_TXT)))
# Abschlussquote grün wenn > 0
ws2.conditional_formatting.add("B7:F7", FormulaRule(
    formula=['ISNUMBER(B7)'], fill=fill(C_GREEN_BG), font=font(size=10, bold=True, color=C_GREEN_TXT)))

# Verlustanalyse-Abschnitt
sep = last_kpi_row + 2
ws2.merge_cells(f"A{sep}:F{sep}")
seph = ws2.cell(row=sep, column=1, value="VERLUSTANALYSE")
seph.fill = fill(C_RED); seph.font = font(bold=True, color=C_WHITE, size=10)
seph.alignment = align("left", "center"); ws2.row_dimensions[sep].height = 22

for vi, vg in enumerate(["Online","Preis","Kein Bedarf","Konkurrenz","Keine Rueckmeldung"], sep+1):
    bg = C_GRAY_LIGHT if vi % 2 == 0 else C_WHITE
    c1 = ws2.cell(row=vi, column=1, value=vg)
    c1.fill = fill(bg); c1.font = font(size=10); c1.border = thin_border()
    c2 = ws2.cell(row=vi, column=2,
        value=f'=COUNTIF({TN}!R4:R2000,A{vi})')
    c2.fill = fill(bg); c2.font = font(size=10, bold=True)
    c2.alignment = align("center"); c2.border = thin_border()
    for c_idx in range(3, 7):
        ws2.cell(row=vi, column=c_idx).fill = fill(bg)
    ws2.row_dimensions[vi].height = 20

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 3 – Legende
# ══════════════════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("Legende")
ws3.sheet_view.showGridLines = False
for i, w in enumerate([22, 40, 30], 1):
    ws3.column_dimensions[get_column_letter(i)].width = w

# Banner
ws3.merge_cells("A1:C1")
b3 = ws3.cell(row=1, column=1, value="FITSHOP")
b3.fill = fill(C_RED); b3.font = font(size=20, bold=True, color=C_WHITE)
b3.alignment = align("left", "center"); ws3.row_dimensions[1].height = 38

ws3.merge_cells("A2:C2")
s3 = ws3.cell(row=2, column=1, value="LEGENDE & REFERENZ 2026  ·  Hall/Tirol")
s3.fill = fill(C_ANTHRAZIT); s3.font = font(size=10, color="AAAAAA", italic=True)
s3.alignment = align("left", "center"); ws3.row_dimensions[2].height = 18

def section_hdr(ws, row, text):
    ws.merge_cells(f"A{row}:C{row}")
    c = ws.cell(row=row, column=1, value=text)
    c.fill = fill(C_RED); c.font = font(bold=True, color=C_WHITE, size=10)
    c.alignment = align("left", "center"); ws.row_dimensions[row].height = 22

def col_hdr(ws, row, c1, c2, c3):
    for ci, val in enumerate([c1, c2, c3], 1):
        cell = ws.cell(row=row, column=ci, value=val)
        cell.fill = fill(C_ANTHRAZIT); cell.font = font(bold=True, color=C_WHITE, size=9)
        cell.alignment = align("center", "center"); cell.border = thin_border("404040")
    ws.row_dimensions[row].height = 20

def tbl(ws, row, col1, col2, col3="", alt=False):
    bg = C_GRAY_LIGHT if alt else C_WHITE
    for ci, val in enumerate([col1, col2, col3], 1):
        cell = ws.cell(row=row, column=ci, value=val)
        cell.fill = fill(bg); cell.font = font(size=9)
        cell.alignment = align("left", "center", wrap=True)
        cell.border = thin_border()
    ws.row_dimensions[row].height = 18

# Kundentypen
section_hdr(ws3, 4, "KUNDENTYPEN (Dropdown: Typ)")
col_hdr(ws3, 5, "Kuerzel", "Beschreibung", "Strategie")
for i, (a, b, c) in enumerate([
    ("A – Sofortkaeufer",     "Kauft heute, kein Heimcheck",           "Direktabschluss"),
    ("B – Logischer Pruefer", "Will Fakten, muss Masse nehmen",        "Heimcheck vereinbaren"),
    ("C – Unsicherer",        "Braucht Bestaetigung oder Partner",     "Partner einladen"),
    ("D – Preisfokussierter", "Vergleicht online, Budget-sensibel",    "Mehrwert zeigen"),
    ("E – Raumplaner",        "Raum noch nicht fertig",                "Wiedervorlage in X Mon."),
    ("F – Gesundheitskunde",  "Arzt-/Reha-Empfehlung",                "Empfehlung aufgreifen"),
    ("G – Geschenkekaeufer",  "Kauft fuer jemand anderen",             "Direkt zum Abschluss"),
    ("H – Wiederkaeeufer",    "Kennt Fitshop schon",                   "Upgrade / Ergaenzung"),
    ("I – Profi/Enthusiast",  "Sehr technisch, viel Ahnung",           "Auf Augenhoehe beraten"),
    ("J – B2B/Verein",        "Unternehmen, Hotel, Schule, Verein",    "Angebot + Termin"),
], 6):
    tbl(ws3, i, a, b, c, i % 2 == 0)

# Einwandtypen (E1–E10)
section_hdr(ws3, 17, "EINWANDTYPEN (Dropdown: Einwand)")
col_hdr(ws3, 18, "Code", "Kunden-Formulierung / Bedeutung", "Naechster Schritt")
for i, (a, b, c) in enumerate([
    ("E1 – Masse",         "Masse unklar / muss noch ausmessen",         "Heimcheck vereinbaren"),
    ("E2 – Partner",       "Partner fehlt / muss noch mit Frau reden",   "2. Besuch mit Partner"),
    ("E3 – Budget",        "Preis zu hoch / muss noch schauen",          "Finanzierung zeigen"),
    ("E4 – Qualitaet",     "Zweifel an Produkt oder Marke",              "Testfahrt, Garantie"),
    ("E5 – Raum",          "Raum nicht fertig / muss noch schauen obs reinpasst", "Wiedervorlage"),
    ("E6 – Nutzung",       "Unsicher ob Geraet zum Training passt",      "Nutzungsanalyse"),
    ("E7 – Dringlichkeit", "Kein Zeitdruck, kein Anlass",                "Anlass schaffen"),
    ("E8 – Ausmessen",     "Muss noch ausmessen (= Masse, Heimcheck)",   "Heimcheck Termin"),
    ("E9 – Frau fragen",   "Muss noch mit Frau/Partner reden",           "2. Besuch einladen"),
    ("E10 – Passt es rein","Muss noch schauen obs reinpasst (Masse/Raum)","Heimcheck + Masse"),
], 19):
    tbl(ws3, i, a, b, c, i % 2 == 0)

# Ampel-Legende
section_hdr(ws3, 30, "AMPEL-LEGENDE (Spalte A)")
for i, (sym, txt, bg_c) in enumerate([
    ("ok", "Check-Datum liegt in > 2 Tagen – alles im Plan", C_GREEN_BG),
    ("~",  "Check-Datum heute oder morgen – sofort handeln", C_YELLOW_BG),
    ("!",  "Check-Datum ueberschritten – dringend nachfassen!", C_RED_BG),
    ("o",  "Abgeschlossen / Verloren / Wartet – kein Follow-up", C_GRAY_BG),
], 31):
    c1 = ws3.cell(row=i, column=1, value=sym)
    c1.fill = fill(bg_c); c1.font = font(size=10, bold=True); c1.border = thin_border()
    c1.alignment = align("center", "center")
    ws3.merge_cells(f"B{i}:C{i}")
    c2 = ws3.cell(row=i, column=2, value=txt)
    c2.fill = fill(bg_c); c2.font = font(size=9); c2.border = thin_border()
    ws3.row_dimensions[i].height = 18

# Status-Farbkodierung
section_hdr(ws3, 36, "STATUS-FARBKODIERUNG")
for i, (stat, erkl, bg_c) in enumerate([
    ("Kauf abgeschlossen",          "Abgeschlossen, kein Follow-up",         C_GREEN_BG),
    ("Heimcheck angelegt",          "Naechster Schritt: Check-Anruf",        C_BLUE_BG),
    ("Entscheidungscheck geplant",  "Kurz vor Abschluss – nachfassen",       C_YELLOW_BG),
    ("Angebot konkret",             "Angebot liegt vor",                     C_ORANGE_BG),
    ("Verloren",                    "Lead verloren – Verlustgrund eintragen", C_RED_BG),
    ("Warten auf Raum",             "Wiedervorlage in Monaten",              C_GRAY_BG),
], 37):
    c1 = ws3.cell(row=i, column=1, value=stat)
    c1.fill = fill(bg_c); c1.font = font(size=9, bold=True); c1.border = thin_border()
    ws3.merge_cells(f"B{i}:C{i}")
    c2 = ws3.cell(row=i, column=2, value=erkl)
    c2.fill = fill(bg_c); c2.font = font(size=9); c2.border = thin_border()
    ws3.row_dimensions[i].height = 18

# ─── Speichern ────────────────────────────────────────────────────────────────
wb.active = ws
path = "/home/user/My-Test-repository/fitshop-sop/Fitshop_Tracking_2026.xlsx"
wb.save(path)
print(f"Gespeichert: {path}")
print(f"Sheets: {[s.title for s in wb.worksheets]}")
