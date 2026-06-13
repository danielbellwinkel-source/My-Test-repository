"""
Sales Tracking Dashboard Builder
Erstellt das 'Dashboard Daten' Sheet mit:
- Umsatzentwicklung 2026 (A1:F13 + H1:H13)
- Performance-Index Tabelle (J1:N13)
- Zwei Liniendiagramme
"""

import shutil
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, Reference
from openpyxl.utils import get_column_letter

SRC = "/root/.claude/uploads/92562e3e-b707-519e-8307-08a3ba26ba1f/cc99fa36-Sales_Tracking_NB.xlsx"
DEST = "/home/user/My-Test-repository/Sales_Tracking_NB_Dashboard.xlsx"

shutil.copy2(SRC, DEST)
wb = openpyxl.load_workbook(DEST)

# ─── Farbpalette (passt zum Fitshop-Stil) ────────────────────────────────────
C_DARK   = "1F2937"   # Dunkelgrau Hintergrund
C_HEADER = "374151"   # Header-Zeile
C_BLUE   = "2563EB"   # Akzent Blau
C_GREEN  = "16A34A"   # Grün
C_WHITE  = "FFFFFF"
C_LGRAY  = "F3F4F6"   # Helles Grau für Alt-Zeilen
C_BORDER = "D1D5DB"

def hdr_fill(color=C_DARK):
    return PatternFill("solid", fgColor=color)

def thin_border():
    s = Side(style="thin", color=C_BORDER)
    return Border(left=s, right=s, top=s, bottom=s)

def style_header(cell, text, bg=C_DARK, fg=C_WHITE, bold=True, center=True):
    cell.value = text
    cell.font = Font(bold=bold, color=fg, name="Calibri", size=10)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center" if center else "left",
                                vertical="center", wrap_text=False)
    cell.border = thin_border()

def style_data(cell, value=None, number_format=None, alt=False):
    if value is not None:
        cell.value = value
    cell.font = Font(name="Calibri", size=10)
    cell.fill = PatternFill("solid", fgColor=C_LGRAY if alt else C_WHITE)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border()
    if number_format:
        cell.number_format = number_format

# ─── Sheet anlegen ────────────────────────────────────────────────────────────
if "Dashboard Daten" in wb.sheetnames:
    del wb["Dashboard Daten"]

ws = wb.create_sheet("Dashboard Daten", 2)
ws.sheet_view.showGridLines = False

# Tab-Farbe
ws.sheet_properties.tabColor = "2563EB"

# ─── Spaltenbreiten ───────────────────────────────────────────────────────────
col_widths = {
    "A": 10, "B": 14, "C": 12, "D": 12, "E": 12, "F": 12,
    "G": 3,  "H": 10, "I": 3,
    "J": 10, "K": 12, "L": 12, "M": 12, "N": 12,
}
for col, w in col_widths.items():
    ws.column_dimensions[col].width = w

# Zeilenhöhe für Header
ws.row_dimensions[1].height = 22

# ─── Überschriften-Bereiche ───────────────────────────────────────────────────
# Bereich 1: Umsatzentwicklung
ws.merge_cells("A1:F1")
style_header(ws["A1"], "UMSATZENTWICKLUNG 2026 – Monatlicher Umsatz pro Mitarbeiter",
             bg=C_BLUE, fg=C_WHITE, center=True)

# Bereich 2: Performance-Index
ws.merge_cells("J1:N1")
style_header(ws["J1"], "PERFORMANCE-INDEX 2026 – Score pro Mitarbeiter (0–100)",
             bg=C_DARK, fg=C_WHITE, center=True)

ws.row_dimensions[2].height = 18

# Spaltenheader Umsatz
for col, label in zip(["A","B","C","D","E","F"],
                       ["Monat","Gesamt","Daniel","Jonas","Angie","Norbert"]):
    style_header(ws[f"{col}2"], label, bg=C_HEADER, fg=C_WHITE)

# Spaltenheader Performance
for col, label in zip(["J","K","L","M","N"],
                       ["Monat","Daniel","Jonas","Angie","Norbert"]):
    style_header(ws[f"{col}2"], label, bg=C_HEADER, fg=C_WHITE)

# MonatNr-Header (versteckt/Hilfspalette)
ws["H2"] = "MonatNr"
ws["H2"].font = Font(color="9CA3AF", size=9, italic=True, name="Calibri")
ws["H2"].alignment = Alignment(horizontal="center")

# ─── Monate & MonatNr ─────────────────────────────────────────────────────────
MONTHS_DE = ["Jan","Feb","Mär","Apr","Mai","Jun",
             "Jul","Aug","Sep","Okt","Nov","Dez"]
EMPLOYEES  = ["Daniel","Jonas","Angie","Norbert"]

for i, (mon, mnum) in enumerate(zip(MONTHS_DE, range(1, 13))):
    row = i + 3          # Daten ab Zeile 3
    alt = (i % 2 == 0)  # Alternating

    # MonatNr (Hilfsspalte, kleine graue Schrift)
    ws[f"H{row}"] = mnum
    ws[f"H{row}"].font = Font(color="9CA3AF", size=9, name="Calibri")
    ws[f"H{row}"].alignment = Alignment(horizontal="center")

    # ── Monatslabel (A) ──────────────────────────────────────────────────────
    style_data(ws[f"A{row}"], mon, alt=alt)
    ws[f"A{row}"].font = Font(bold=True, name="Calibri", size=10)

    # ── Gesamt-Umsatz (B) ────────────────────────────────────────────────────
    ws[f"B{row}"] = (
        f"=SUMPRODUCT("
        f"(YEAR('Lead-Tracking 2026'!$B$4:$B$2000)=2026)*"
        f"(MONTH('Lead-Tracking 2026'!$B$4:$B$2000)=$H{row})*"
        f"('Lead-Tracking 2026'!$Q$4:$Q$2000=\"Ja\")*"
        f"('Lead-Tracking 2026'!$P$4:$P$2000))"
    )
    style_data(ws[f"B{row}"], number_format='#,##0.00 "€"', alt=alt)

    # ── Umsatz pro Mitarbeiter (C–F) ─────────────────────────────────────────
    for col, emp in zip(["C","D","E","F"], EMPLOYEES):
        ws[f"{col}{row}"] = (
            f"=SUMPRODUCT("
            f"(YEAR('Lead-Tracking 2026'!$B$4:$B$2000)=2026)*"
            f"(MONTH('Lead-Tracking 2026'!$B$4:$B$2000)=$H{row})*"
            f"('Lead-Tracking 2026'!$C$4:$C$2000=\"{emp}\")*"
            f"('Lead-Tracking 2026'!$Q$4:$Q$2000=\"Ja\")*"
            f"('Lead-Tracking 2026'!$P$4:$P$2000))"
        )
        style_data(ws[f"{col}{row}"], number_format='#,##0.00 "€"', alt=alt)

    # ── Performance-Index (J–N) ───────────────────────────────────────────────
    # Monatslabel
    style_data(ws[f"J{row}"], mon, alt=alt)
    ws[f"J{row}"].font = Font(bold=True, name="Calibri", size=10)

    # Score-Formel pro Mitarbeiter
    # Score = Abschlussquote*50 + (Umsatz_Emp/Leads/1000)*30 + (KR_Avg/5)*20
    umsatz_cols = {"Daniel":"C","Jonas":"D","Angie":"E","Norbert":"F"}

    for col, emp in zip(["K","L","M","N"], EMPLOYEES):
        uc = umsatz_cols[emp]

        leads = (
            f"SUMPRODUCT((YEAR('Lead-Tracking 2026'!$B$4:$B$2000)=2026)*"
            f"(MONTH('Lead-Tracking 2026'!$B$4:$B$2000)=$H{row})*"
            f"('Lead-Tracking 2026'!$C$4:$C$2000=\"{emp}\"))"
        )
        abschl = (
            f"SUMPRODUCT((YEAR('Lead-Tracking 2026'!$B$4:$B$2000)=2026)*"
            f"(MONTH('Lead-Tracking 2026'!$B$4:$B$2000)=$H{row})*"
            f"('Lead-Tracking 2026'!$C$4:$C$2000=\"{emp}\")*"
            f"('Lead-Tracking 2026'!$Q$4:$Q$2000=\"Ja\"))"
        )
        kr_sum = (
            f"SUMPRODUCT((YEAR('Lead-Tracking 2026'!$B$4:$B$2000)=2026)*"
            f"(MONTH('Lead-Tracking 2026'!$B$4:$B$2000)=$H{row})*"
            f"('Lead-Tracking 2026'!$C$4:$C$2000=\"{emp}\")*"
            f"('Lead-Tracking 2026'!$L$4:$L$2000>0)*"
            f"'Lead-Tracking 2026'!$L$4:$L$2000)"
        )
        kr_cnt = (
            f"SUMPRODUCT((YEAR('Lead-Tracking 2026'!$B$4:$B$2000)=2026)*"
            f"(MONTH('Lead-Tracking 2026'!$B$4:$B$2000)=$H{row})*"
            f"('Lead-Tracking 2026'!$C$4:$C$2000=\"{emp}\")*"
            f"('Lead-Tracking 2026'!$L$4:$L$2000>0))"
        )

        formula = (
            f"=IFERROR("
            f"({abschl})/({leads})*50"
            f"+({uc}{row}/({leads})/1000*30)"
            f"+(({kr_sum})/({kr_cnt})/5*20)"
            f",\"\")"
        )
        ws[f"{col}{row}"] = formula
        style_data(ws[f"{col}{row}"], number_format='0.0', alt=alt)

# ─── Summenzeile Umsatz ───────────────────────────────────────────────────────
SUM_ROW = 16
ws.row_dimensions[SUM_ROW].height = 20
for col, label in zip(["A","B","C","D","E","F"],
                       ["Gesamt", "=SUM(B3:B14)", "=SUM(C3:C14)",
                        "=SUM(D3:D14)", "=SUM(E3:E14)", "=SUM(F3:F14)"]):
    ws[f"{col}{SUM_ROW}"] = label
    ws[f"{col}{SUM_ROW}"].font = Font(bold=True, name="Calibri", size=10, color=C_WHITE)
    ws[f"{col}{SUM_ROW}"].fill = PatternFill("solid", fgColor=C_BLUE)
    ws[f"{col}{SUM_ROW}"].alignment = Alignment(horizontal="center", vertical="center")
    ws[f"{col}{SUM_ROW}"].border = thin_border()
    if col != "A":
        ws[f"{col}{SUM_ROW}"].number_format = '#,##0.00 "€"'

# ─── Diagramm 1: Umsatzentwicklung ───────────────────────────────────────────
chart1 = LineChart()
chart1.title = "Umsatzentwicklung 2026"
chart1.style = 2
chart1.grouping = "standard"
chart1.y_axis.title = "Umsatz (EUR)"
chart1.x_axis.title = "Monat"
chart1.legend.position = "b"
chart1.width = 22
chart1.height = 14

# Datenreihen: Gesamt + 4 Mitarbeiter (Spalten B–F, Zeilen 2–14 mit Header)
data1 = Reference(ws, min_col=2, max_col=6, min_row=2, max_row=14)
chart1.add_data(data1, titles_from_data=True)

# Kategorien (Monate)
cats1 = Reference(ws, min_col=1, min_row=3, max_row=14)
chart1.set_categories(cats1)

# Linien dicker machen
for i, series in enumerate(chart1.series):
    series.graphicalProperties.line.width = 25000  # ~2.5pt in EMU
    series.smooth = True

ws.add_chart(chart1, "A18")

# ─── Diagramm 2: Performance-Entwicklung ─────────────────────────────────────
chart2 = LineChart()
chart2.title = "Performance-Entwicklung Mitarbeiter"
chart2.style = 2
chart2.grouping = "standard"
chart2.y_axis.title = "Score (0–100)"
chart2.x_axis.title = "Monat"
chart2.legend.position = "b"
chart2.width = 22
chart2.height = 14

# Datenreihen: 4 Mitarbeiter (Spalten K–N, Zeilen 2–14 mit Header)
data2 = Reference(ws, min_col=11, max_col=14, min_row=2, max_row=14)
chart2.add_data(data2, titles_from_data=True)

cats2 = Reference(ws, min_col=10, min_row=3, max_row=14)
chart2.set_categories(cats2)

for series in chart2.series:
    series.graphicalProperties.line.width = 25000
    series.smooth = True

ws.add_chart(chart2, "J18")

# ─── Hinweis-Zeile ───────────────────────────────────────────────────────────
ws["A15"] = "Hinweis: Daten werden automatisch aus 'Lead-Tracking 2026' berechnet. Aktualisieren: Strg+Alt+F9"
ws["A15"].font = Font(italic=True, color="6B7280", size=9, name="Calibri")
ws.merge_cells("A15:N15")

# ─── Speichern ───────────────────────────────────────────────────────────────
wb.save(DEST)
print(f"Fertig! Gespeichert: {DEST}")
print("Sheets:", wb.sheetnames)
