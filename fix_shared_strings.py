# -*- coding: utf-8 -*-
"""Wandelt inlineStr-Zellen in eine echte sharedStrings-Tabelle um.
Grund: dieser openpyxl-Build schreibt t='inlineStr' ohne sharedStrings.xml,
woran Excel for Mac beim Oeffnen abstuerzt. Excel erwartet Shared Strings.
Chirurgischer Eingriff: nur String-Speicherung wird geaendert, alles andere bleibt.
"""
import zipfile, re, shutil, os, sys

SRC = "Training_2.xlsx"
TMP = "Training_2_fixed.xlsx"

z = zipfile.ZipFile(SRC)
names = z.namelist()
sheet_names = [n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n)]

# Shared-String-Tabelle aufbauen
strings = []          # geordnete Liste eindeutiger (escaped) Texte
index = {}            # text -> idx
total_refs = 0

def get_index(text):
    global total_refs
    total_refs += 1
    if text in index:
        return index[text]
    idx = len(strings)
    index[text] = idx
    strings.append(text)
    return idx

# Regex fuer gepaarte inlineStr-Zellen:  <c ... t="inlineStr" ...><is>...</is></c>
cell_pair = re.compile(r'<c\b([^>]*?)\st="inlineStr"([^>]*?)>(<is>.*?</is>)</c>', re.S)
# leere selbstschliessende inlineStr-Zelle: <c ... t="inlineStr" ... />
cell_self = re.compile(r'<c\b([^>]*?)\st="inlineStr"([^>]*?)/>', re.S)
# Text aus <is> extrahieren
t_inner = re.compile(r'<t(?:\s[^>]*)?>(.*?)</t>', re.S)

new_sheets = {}
for n in sheet_names:
    d = z.read(n).decode("utf-8")

    def repl_pair(m):
        pre, post, isblock = m.group(1), m.group(2), m.group(3)
        tm = t_inner.search(isblock)
        text = tm.group(1) if tm else ""
        idx = get_index(text)
        return f'<c{pre} t="s"{post}><v>{idx}</v></c>'

    d = cell_pair.sub(repl_pair, d)

    def repl_self(m):
        pre, post = m.group(1), m.group(2)
        # leere Zelle ohne Typ (gueltige Leerzelle)
        return f'<c{pre}{post}/>'

    d = cell_self.sub(repl_self, d)
    new_sheets[n] = d.encode("utf-8")

# sharedStrings.xml bauen
def si_for(text):
    # xml:space erhalten, falls fuehrende/abschliessende Leerzeichen oder Zeilenumbrueche
    raw = text
    if raw != raw.strip() or "\n" in raw or "\r" in raw or "\t" in raw:
        return f'<si><t xml:space="preserve">{raw}</t></si>'
    return f'<si><t>{raw}</t></si>'

sst = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
       '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
       f'count="{total_refs}" uniqueCount="{len(strings)}">'
       + "".join(si_for(s) for s in strings) + "</sst>").encode("utf-8")

# [Content_Types].xml: Override fuer sharedStrings ergaenzen
ct = z.read("[Content_Types].xml").decode("utf-8")
if "sharedStrings" not in ct:
    ct = ct.replace("</Types>",
        '<Override PartName="/xl/sharedStrings.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" /></Types>')

# xl/_rels/workbook.xml.rels: Relationship fuer sharedStrings ergaenzen
rels_name = "xl/_rels/workbook.xml.rels"
rels = z.read(rels_name).decode("utf-8")
existing_ids = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels)]
new_id = (max(existing_ids) + 1) if existing_ids else 1
if "sharedStrings.xml" not in rels:
    rels = rels.replace("</Relationships>",
        f'<Relationship Id="rId{new_id}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
        'Target="sharedStrings.xml" /></Relationships>')

# Neues Zip schreiben
with zipfile.ZipFile(TMP, "w", zipfile.ZIP_DEFLATED) as out:
    for item in z.infolist():
        n = item.filename
        if n in new_sheets:
            out.writestr(item, new_sheets[n])
        elif n == "[Content_Types].xml":
            out.writestr(item, ct.encode("utf-8"))
        elif n == rels_name:
            out.writestr(item, rels.encode("utf-8"))
        else:
            out.writestr(item, z.read(n))
    # sharedStrings.xml hinzufuegen
    out.writestr("xl/sharedStrings.xml", sst)

z.close()
shutil.move(TMP, SRC)

print(f"Shared strings: {len(strings)} eindeutig, {total_refs} Referenzen.")
print("inlineStr -> sharedStrings konvertiert. Datei aktualisiert:", SRC)
