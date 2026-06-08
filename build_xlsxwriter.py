# -*- coding: utf-8 -*-
"""Baut Training_2.xlsx mit XlsxWriter (Excel-konformer Output, robust fuer Excel for Mac).
- echte sharedStrings, saubere Tabellen-XML
- Theme-Font auf Arial gepatcht -> alles Arial
- INDEX/MATCH statt XLOOKUP (maximale Kompatibilitaet)
- fullCalcOnLoad erzwungen
"""
import xlsxwriter, datetime, zipfile, shutil, re, os

PATH = "/home/user/My-Test-repository/Training_2.xlsx"
TODAY = datetime.datetime(2026, 6, 8)

C_DARK="#1F3864"; C_HEAD="#305496"; C_SUB="#D9E1F2"; C_BAND="#F2F6FC"; C_BLOCK="#FFE699"
C_GREEN="#C6EFCE"; C_GREENT="#006100"; C_YELLOW="#FFEB9C"; C_ORANGE="#FFD580"
C_RED="#FFC7CE"; C_REDT="#9C0006"

wb = xlsxwriter.Workbook(PATH, {"in_memory": True})
wb.set_calc_mode("auto")

_cache = {}
def fmt(**kw):
    kw.setdefault("font_name", "Arial")
    key = tuple(sorted((k, str(v)) for k, v in kw.items()))
    if key not in _cache:
        _cache[key] = wb.add_format(kw)
    return _cache[key]

F_TITLE  = fmt(bold=1, font_size=14, font_color="white", bg_color=C_DARK, align="left", valign="vcenter")
F_SEC    = fmt(bold=1, font_size=11, font_color="#1F3864", bg_color=C_SUB, align="left", valign="vcenter")
F_H      = fmt(bold=1, font_size=10, font_color="white", bg_color=C_HEAD, align="center", valign="vcenter", text_wrap=1, border=1)
F_BODY   = fmt(font_size=10, border=1, text_wrap=1, valign="top")
F_BODYC  = fmt(font_size=10, border=1, text_wrap=1, valign="top", align="center")
F_BOLD   = fmt(bold=1, font_size=10, border=1, text_wrap=1, valign="top")
F_BOLDC  = fmt(bold=1, font_size=10, border=1, text_wrap=1, valign="top", align="center")
F_IT     = fmt(font_size=9, italic=1, text_wrap=1, valign="top")
F_NOTE   = fmt(font_size=8, italic=1, valign="top", text_wrap=1)
F_BLOCK  = fmt(bold=1, font_size=12, bg_color=C_BLOCK, align="center", valign="vcenter", border=1)
F_LINK   = fmt(font_size=9, font_color="#0563C1", border=1, text_wrap=1, valign="top", underline=1)
F_DATE   = fmt(font_size=10, num_format="yyyy-mm-dd")
def numf(spec, **extra): return fmt(font_size=10, border=1, valign="top", align="center", num_format=spec, **extra)
F_RED   = fmt(bg_color=C_RED, font_color=C_REDT)
F_ORANGE= fmt(bg_color=C_ORANGE)
F_GREEN = fmt(bg_color=C_GREEN, font_color=C_GREENT)
F_YELLOW= fmt(bg_color=C_YELLOW)
def gate(color, tcolor=None):
    d = dict(font_size=10, border=1, text_wrap=1, valign="top", bg_color=color)
    if tcolor: d["font_color"] = tcolor
    return fmt(**d)

def put(ws, row, col, val, f=None):           # 1-indexed
    r, c = row - 1, col - 1
    if isinstance(val, str) and val.startswith("="):
        ws.write_formula(r, c, val, f)
    elif val is None:
        if f: ws.write_blank(r, c, None, f)
    else:
        ws.write(r, c, val, f)

def title(ws, a1, text):
    ws.merge_range(a1, text, F_TITLE)
def section(ws, a1, text):
    ws.merge_range(a1, text, F_SEC)
def headers(ws, row, hs, start=1):
    for i, h in enumerate(hs):
        put(ws, row, start + i, h, F_H)

# ============================================================ Library data
LIBRARY = [
 ("90","neunzig grad,90,90 grad,90 grad liegestuetz,ninety,90 degree push up,neunzig","90 Degree Push-up","Skill90","Brust/Schulter/Planche","N","N","N"),
 ("OAHS_line","oahs,oahs line,einarmiger handstand,one arm handstand,weight shift,gewichtsverlagerung,oahs holds","OAHS Line / Weight-Shift / Holds","SkillOAHS","Schulter/Core/Balance","N","J","J"),
 ("OAHS_shapes","oahs shapes,straddle,half straddle,diamond,legs together,oahs formen,shape density","OAHS Shape-Density","SkillOAHS","Schulter/Core/Balance","N","J","J"),
 ("HS_line","handstand,handstand line,hs line,modern line,balance,handstand balance,linie","Handstand Line / Balance","SkillOAHS","Schulter/Core/Balance","N","J","N"),
 ("OAHS_flag_supported","supported flag,pre flag,supported pre flag,flag supported,unterstuetzter flag,pre flag lean","Supported Pre-Flag Lean","SkillFlag","Schulter/Side-Body","N","J","J"),
 ("flag_sideline","sideline,sideline hold,mini flag,flag hold,flag,seitlicher flag","Sideline / Mini-Flag Hold","SkillFlag","Schulter/Side-Body","N","J","J"),
 ("pseudo_planche_pushup","pseudo planche,pseudo planche push up,ppp,pseudo planche liegestuetz","Pseudo-Planche Push-up","Planche","Schulter/Brust","N","N","N"),
 ("planche_lean","planche lean,lean,planche neigung,plankenlean","Planche Lean","Planche","Schulter/Core","N","J","N"),
 ("planche_hold","planche hold,tuck planche,advanced tuck,adv tuck,straddle planche,planche halten","Planche Hold (Tuck/Adv/Straddle)","Planche","Schulter/Core","N","J","N"),
 ("WPU","weighted pull up,wpu,gewichtetes klimmziehen,klimmzug mit gewicht,weighted pullup","Weighted Pull-up","Pull","Lat/Biceps","J","N","N"),
 ("pullup","pull up,pullup,klimmzug,klimmziehen","Pull-up","Pull","Lat/Biceps","J","N","N"),
 ("chin_up","chin up,chinup,kinnzug,unterhand klimmzug","Chin-up","Pull","Lat/Biceps","J","N","N"),
 ("lat_pulldown","lat pulldown,pulldown,latzug,latziehen","Lat Pulldown","Pull","Lat","J","N","N"),
 ("chest_row","chest supported row,seal row,brustgestuetztes rudern,rudern gestuetzt","Chest-Supported / Seal Row","Pull","Oberer Ruecken","J","N","N"),
 ("oa_row","one arm row,einarmiges rudern,oa row,unilateral row","One-Arm Row","Pull","Oberer Ruecken","J","N","J"),
 ("rear_delt_fly","rear delt fly,reverse fly,hintere schulter,rear delt","Rear-Delt Fly","Pull","Hintere Schulter","J","N","N"),
 ("face_pull","face pull,facepull,gesichtszug","Face Pull","Pull","Hintere Schulter/Rotatoren","J","N","N"),
 ("scap_serratus","scap,serratus,scapula,schulterblatt,scap pull,serratus push","Scap / Serratus","Mobility","Schulterblatt","N","N","N"),
 ("ext_rotation","external rotation,aussenrotation,ar,rotatorenmanschette,cuff","External Rotation","Mobility","Rotatorenmanschette","J","N","J"),
 ("wall_hspu","wall hspu,hspu,handstand push up,handstand liegestuetz,wand hspu","Wall HSPU","Push","Schulter/Trizeps","J","N","N"),
 ("pike_press","pike press,pike push up,db press,schulterdruecken,overhead press,ohp,pike","Pike / DB Vertical Press","Push","Schulter","J","N","N"),
 ("ring_dips","ring dips,dips,barren,ringdips","Ring Dips / Dips","Push","Brust/Trizeps","J","N","N"),
 ("deficit_pushup","deficit push up,deficit pushup,liegestuetz deficit,push up,liegestuetz","Deficit Push-up","Push","Brust/Trizeps","N","N","N"),
 ("lateral_raise","lateral raise,seitheben,lat raise,seitliches heben","Lateral Raise","Push","Seitliche Schulter","J","N","N"),
 ("triceps","triceps,trizeps,trizepsdruecken,pushdown","Triceps","Push","Trizeps","J","N","N"),
 ("biceps","biceps,bizeps,curl,bizeps curl","Biceps","Pull","Bizeps","J","N","N"),
 ("bss","bulgarian,bulgarian split squat,bss,split squat,ausfallschritt,ffe split squat,front foot elevated","Bulgarian / FFE Split Squat","Legs","Quad/Glute","J","N","J"),
 ("leg_press","leg press,beinpresse,hack squat,hack,belt squat,guertelkniebeuge","Leg Press / Hack / Belt Squat","Legs","Quad/Glute","J","N","N"),
 ("ham_curl","hamstring curl,beinbeuger,leg curl,beincurl","Hamstring Curl","Legs","Hamstrings","J","N","N"),
 ("hip_thrust","hip thrust,hueftstoss,glute bridge,hueftheben","Hip Thrust","Legs","Glute","J","N","N"),
 ("back_ext_iso","back extension,rueckenstrecker,back ext iso,hyperextension,iso hold ruecken","Back-Extension Iso","Legs","Posteriore Kette/LWS","N","J","N"),
 ("calves","calves,waden,wadenheben,calf raise","Calves","Legs","Waden","J","N","N"),
 ("hinge_pattern","hinge,hinge pattern,dowel hinge,hueftbeuge,rdl pattern,deadlift pattern,scharnier","Hinge Pattern (Dowel/Rebuild)","Legs","Posteriore Kette/LWS","J","N","N"),
 ("pallof","pallof,pallof press,anti rotation","Pallof Press","Core","Anti-Rotation","J","N","J"),
 ("dead_bug","dead bug,deadbug,toter kaefer","Dead Bug","Core","Anti-Extension","N","N","N"),
 ("side_plank","side plank,seitstuetz,seitlicher unterarmstuetz","Side Plank","Core","Lateral/Anti-Lateral","N","J","J"),
 ("hollow","hollow,hollow hold,hollow body,compression,kompression,hohlkreuz hold","Hollow / Compression Core","Core","Anteriore Kette","N","J","N"),
 ("bws_mobility","bws,bws mobility,brustwirbelsaeule,thoracic,t spine,extension rotation,ribs down,atmung,downshift","BWS / Thoracic Mobility","Mobility","BWS/Atmung","N","J","N"),
 ("wrist_prep","wrist prep,handgelenk,handgelenk prep,handgelenke,wrist warm up,wrist","Wrist Prep","Wrist","Handgelenke","N","N","N"),
 ("wrist_strength","wrist strength,handgelenk kraft,wrist curl,handgelenk kraeftigung","Wrist Strengthening","Wrist","Handgelenke","J","N","N"),
]
DAYTYPES = ["T1","T2","T3","T4","T5","Rest"]
LIB_HR = 3
LIB_FIRST = LIB_HR + 1                 # 4
LIB_LAST = LIB_HR + len(LIBRARY)       # 43

# ============================================================ 1) Dashboard
dash = wb.add_worksheet("B6_Dashboard")
dash.set_column("A:A", 34); dash.set_column("B:B", 12); dash.set_column("C:C", 16)
dash.set_column("D:D", 26); dash.set_column("E:E", 40); dash.set_column("F:I", 14)
title(dash, "A1:F1", "B6 — Performance-Block Dashboard (Hand Balancing / Calisthenics)")
put(dash, 3, 1, "Block =", fmt(bold=1, align="right", valign="vcenter"))
put(dash, 3, 2, "B6", F_BLOCK)
dash.merge_range("C3:F3", "Folgeblock: diese Zelle auf B7 setzen, Sheets klonen, Prefix anpassen.", F_BODY)

section(dash, "A5:F5", "Situationsbewertung")
situ = [
 "Fortgeschrittener, autodidaktischer Hand-Balancer (1,90 m, ~80,8 kg). Skills von Canes auf den Boden uebertragen.",
 "Erster Performance-Block nach 6-Monats-Diaet (90 -> ~81 kg). Maintenance, KEIN Defizit, begrenzte Recovery.",
 "OAHS: links stabil, rechts limitiert (Schulter-Mobility/Koordination). Unilaterale Arbeit leicht zugunsten rechts.",
 "90 Push-up: 1 saubere Rep meist gut, 2. Rep sehr schwer. Ziel = 3 saubere Reps am Stueck.",
 "Ruecken (a) LWS/sakral: alte Reizung nach RDL/Deadlift, am Ausheilen -> Hinge-Rebuild ueber Pain-Gates.",
 "Ruecken (b) BWS: Einklemm-Gefuehl mittig/oben, ohne Ausstrahlen -> Mobility-/Tissue-Baustein, monitoren, KEINE Diagnose.",
 "Prinzip: Skill-Qualitaet + Gelenk-/Rueckengesundheit schlagen Volumen und Fatigue-PRs.",
]
r = 6
for s in situ:
    put(dash, r, 1, "•", F_BODYC); dash.merge_range(r-1, 1, r-1, 5, s, F_BODY); r += 1

r += 1
section(dash, f"A{r}:F{r}", "Zielhierarchie (skill-first; keine Wettkaempfe -> Progress-Maximierung, Gesundheit als Constraint)")
goals = [
 "1) 90 Push-up -> 3 saubere Reps am Stueck (Hauptziel).",
 "2) OAHS Shapes + Flag-Pfad (Straddle -> Shapes; supported Pre-Flag -> Sideline-Hold -> reduzierter Support -> seltene freie Versuche).",
 "3) Modern-Handstand-Linie und weitere Shapes (Flieger/Figa offen, Fernziel, nicht erzwingen).",
 "4) Kraftunterbau (Planche/Push/Pull) als Mittel zum Zweck fuer 1-3.",
 "5) Hypertrophie nur soweit sie Skills oder Schultergesundheit stuetzt.",
]
r += 1
for g in goals:
    dash.merge_range(r-1, 0, r-1, 5, g, F_BODY); r += 1

r += 1
section(dash, f"A{r}:G{r}", "Wochen-Themen & Volumen-Leitplanken (effektive harte Saetze/Woche)")
r += 1
headers(dash, r, ["Woche / Thema","Push*","Pull/Upper-Back","Beine","90-Touches","OAHS-Bloecke","Planche-Dosen"])
VOL = [
 ("W1 Re-Entry / Baseline","12-14","14-16","8-10","3 (1 schwer / 2 leicht)","3","2"),
 ("W2 Build","14-16","16-18","10-12","3","3","2-3"),
 ("W3 Peak","15-16","17-18","12-14","3 (1 schwer / 2 assist)","3","3"),
 ("W4 Consolidate / Test","9-12","10-12","6-8","2 (Quality-Test)","beste Holds","1"),
]
r += 1
for row in VOL:
    for i, val in enumerate(row):
        put(dash, r, 1+i, val, F_BODY if i == 0 else F_BODYC)
    r += 1
dash.merge_range(r-1, 0, r-1, 6,
    "*Push zaehlt 90, Planche, Dips, HSPU, Vertikalpress zusammen. Taegliche Skill-Mikrodosis (6x/Woche) zaehlt NICHT ins harte Volumen.", F_NOTE)
r += 2

section(dash, f"A{r}:E{r}", "Live-Metriken (Formeln aus Log / Weekly_Summary)")
r += 1
live_hr = r
headers(dash, r, ["Metrik","Wert","Target","Interpretation","Action"])
r += 1
def live(label, valf, target, interpf, action, nf="0.0"):
    global r
    put(dash, r, 1, label, F_BOLD)
    vcell = f"B{r}"
    put(dash, r, 2, valf, numf(nf))
    put(dash, r, 3, target, F_BODYC)
    put(dash, r, 4, interpf.format(v=vcell), F_BODY)
    put(dash, r, 5, action, F_BODY)
    r += 1
live("Geloggte Zeilen gesamt", "=COUNTA(tblLog[Date])", "wachsend",
     '=IF({v}=0,"noch leer",IF({v}<10,"wenig Daten","laeuft"))', "Taeglich 1 Zeile/Uebung anhaengen.", "0")
live("Durchschnitt TopRPE (Block)", "=IFERROR(AVERAGE(tblLog[TopRPE]),0)", "7,0-8,5",
     '=IF({v}=0,"-",IF({v}>8.7,"zu hoch",IF({v}<6.5,"Luft nach oben","im Ziel")))',
     "Wenn dauerhaft >8,5: Akzessorik kuerzen, Skill frisch halten.")
live("Durchschnitt Pain (Block)", "=IFERROR(AVERAGE(tblLog[Pain_0_10]),0)", "<2,5",
     '=IF({v}<2.5,"OK",IF({v}<3.5,"beobachten","regredieren"))', "Pain >3 oder Morgen schlechter -> Stufe runter (Rules).")
live("Pain-Flags >=4", '=COUNTIF(tblLog[Pain_0_10],">=4")', "0",
     '=IF({v}=0,"OK","STOP/abklaeren")', "Scharf/neurologisch -> Uebung stoppen, ggf. abklaeren.", "0")
live("Saubere 90-Reps gesamt", '=SUMIFS(tblLog[CleanReps],tblLog[ExerciseID],"90")', "Trend hoch",
     '=IF({v}=0,"-","Fortschritt loggen")', "Ziel: Rep 1 submaximal -> Weg zu 3 sauberen Reps.", "0")
live("OAHS Best-Hold links (s)", '=MAXIFS(tblLog[BestHold_s],tblLog[Category],"SkillOAHS",tblLog[Side],"L")', "Trend hoch",
     '=IF({v}=0,"-",IF({v}>=5,"stabil","aufbauen"))', "Nur saubere Holds zaehlen.", "0")
live("OAHS Best-Hold rechts (s)", '=MAXIFS(tblLog[BestHold_s],tblLog[Category],"SkillOAHS",tblLog[Side],"R")', "naeher an links",
     '=IF({v}=0,"-",IF({v}>=MAXIFS(tblLog[BestHold_s],tblLog[Category],"SkillOAHS",tblLog[Side],"L")*0.7,"Asymmetrie ok","rechts priorisieren"))',
     "Rechts mehr Reps/Holds, Overhead/BWS-Mobility.", "0")
live("Flag Best-Hold links (s)", '=MAXIFS(tblLog[BestHold_s],tblLog[Category],"SkillFlag",tblLog[Side],"L")', "Trend hoch",
     '=IF({v}=0,"-","Support reduzieren wenn sauber")', "Support nur reduzieren bei sauberem Re-Entry.", "0")
live("Flag Best-Hold rechts (s)", '=MAXIFS(tblLog[BestHold_s],tblLog[Category],"SkillFlag",tblLog[Side],"R")', "Trend hoch",
     '=IF({v}=0,"-","Support reduzieren wenn sauber")', "Kontrollierter Return Pflicht.", "0")
# CF auf Pain/Pain-Flags Werte
dash.conditional_format(f"B{live_hr+3}", {"type":"cell","criteria":">=","value":2.5,"format":F_ORANGE})
dash.conditional_format(f"B{live_hr+4}", {"type":"cell","criteria":">=","value":1,"format":F_RED})

# ============================================================ 2) Schedule
sch = wb.add_worksheet("B6_Schedule")
for col, w in {"A:A":18,"B:B":7,"C:C":10,"D:D":30,"E:E":34,"F:F":40,"G:G":34,"H:H":26,"I:I":34,"J:J":34}.items():
    sch.set_column(col, w)
title(sch, "A1:J1", "B6 — Schedule (4 Wochen x 5 Tage + taegliche Skill-Mikrodosis)")
sch.merge_range("A2:J2", "Jede Einheit startet mit Skill-Mikrodosis (8-15 min): Handgelenke -> HS-Line/Balance -> OAHS-Line/Weight-Shift. "
                "Keine Max-Versuche an Nicht-Skill-Tagen. Skill nie ans Ende ermuedeter Sessions.", F_IT)
headers(sch, 3, ["Woche","Tag","Day Type","Main Focus","Skill-Touchpoint","Strength","Back/Hinge-Stress","Placement","Progression-Target","Notes"])
DAY_BASE = {
 "T1": ("90 Heavy + OAHS Line + Vertikalkraft","OAHS Line/Weight-Shift + 2-3 saubere Holds/Seite",
        "90 Cluster (Singles->Double) + Pseudo-Planche/Planche-Hold + Weighted Pull-up + Scap/AR","niedrig (kein Hinge)","nach Ruhe-/Low-Tag (hoechster CNS-Tag)"),
 "T2": ("OAHS Shapes + Pull-Volumen","OAHS Shape-Density (Straddle->Half/Diamond), nur saubere Holds",
        "Chest-supported/Seal Row + Pull-up/Pulldown + Rear-Delt/Face Pull + Hollow/Compression","niedrig (kein LWS-Reiz)","Pull-Schwerpunkt, Skill frisch"),
 "T3": ("Beine (wirbelsaeulen-freundlich) + Core + BWS-Mobility","nur Line, keine Max-Versuche",
        "BSS/FFE Split Squat + Leg Press/Hack/Belt + Ham Curl + Calves + Pallof/Dead Bug/Side Plank",
        "Gate-gesteuert: Hinge mit Dowel -> entscheidet den Tag; KEIN schwerer Hinge","BWS-Mobility-Block fix (Extension/Rotation, ribs down, Atmung)"),
 "T4": ("OAHS Flag/Transitions + 90 Light/Assisted","Laengeres Warm-up + Side-Body-Aktivierung",
        "Supported Pre-Flag-Lean + Sideline/Mini-Flag + 90 leicht/Negativ+Pause + Vertikalpress + OA/Chest Row","niedrig (ribs down, kein LWS-Arch)","Support nur reduzieren wenn Re-Entry sauber"),
 "T5": ("Upper Hypertrophy (getrimmt) + 3. 90-Mikrodosis","90 Technik/Negativ-Mikro (low volume, sauber)",
        "Push-Pump (Ring-Dips/Maschine/Deficit) + Pull-Pump (Row+Pulldown/Chin) + Schultern + Arme opt.","niedrig","BWS-Mobility + Lat/Pec/Forearm + Atmung-Downshift"),
}
WEEK_THEME = {1:"Re-Entry/Baseline",2:"Build",3:"Peak",4:"Consolidate/Test"}
WEEK_PROG = {
 1:{"T1":"Baseline: 90 4 Saetze x1, RPE<=8; Holds Qualitaet finden","T2":"Shape-Baseline, saubere Straddle-Holds zaehlen","T3":"Hinge-Gate testen (Level), Lasten konservativ","T4":"Support hoch, Flag-Geometrie lernen","T5":"Pump leicht, Technik 90 Negativ"},
 2:{"T1":"90 5 Saetze, Cluster 1->Double anpeilen, RPE 8-8,5","T2":"Shape-Density +1 Block, Uebergang Half-Straddle","T3":"Hinge ggf. Level +1, Beinlast +","T4":"Support leicht reduzieren wenn sauber","T5":"Pump +1-2 Saetze, 90 Negativ +Pause"},
 3:{"T1":"90 Peak: saubere Double in 1-2 Saetzen, RPE<=8,5 (kein Failure)","T2":"beste Shapes konsolidieren, Diamond/Legs-together-Prep","T3":"Beinlast Peak, Hinge halten (nicht steigern wenn Pain)","T4":"Mini-Flag-Hold laenger, seltene reduzierte Supports","T5":"Pump Peak getrimmt, 90 Technik scharf"},
 4:{"T1":"Quality-Test: saubere Reps/Holds beweisen, Volumen runter","T2":"beste Holds testen, kein Grind","T3":"Deload Beine, Mobility-Fokus","T4":"beste Flag-Holds dokumentieren","T5":"Mini-Dosis, Regeneration"},
}
r = 4
for wk in [1,2,3,4]:
    for dt in ["T1","T2","T3","T4","T5"]:
        focus, skill, strength, back, place = DAY_BASE[dt]
        put(sch, r, 1, f"W{wk} {WEEK_THEME[wk]}", F_BODYC)
        put(sch, r, 2, dt, F_BOLDC)
        put(sch, r, 3, dt, F_BODYC)
        put(sch, r, 4, focus, F_BODY); put(sch, r, 5, skill, F_BODY); put(sch, r, 6, strength, F_BODY)
        put(sch, r, 7, back, F_BODY); put(sch, r, 8, place, F_BODY); put(sch, r, 9, WEEK_PROG[wk][dt], F_BODY)
        put(sch, r, 10, "Skill-Mikrodosis zuerst; bei Qualitaetsverlust 90 -> erst Push-Akzessorik kuerzen", F_BODY)
        r += 1
sch.freeze_panes("A4")

# ============================================================ 3) Plan_Detail
plan = wb.add_worksheet("B6_Plan_Detail")
for col, w in {"A:A":8,"B:B":14,"C:C":40,"D:D":16,"E:E":12,"F:F":14,"G:G":14,"H:H":14,"I:I":9,"J:J":40,"K:K":22,"L:L":30}.items():
    plan.set_column(col, w)
title(plan, "A1:L1", "B6 — Plan Detail (Uebungsbloecke je Tag T1-T5)")
plan.merge_range("A2:L2", "W1-W4 = Saetze x Reps/Holds je Woche. RPE/Quality und Gates beachten. Skill immer zuerst, nie ermuedet.", F_IT)
headers(plan, 3, ["Day Type","Block","Uebung/Drill","W1","W2","W3","W4","RPE/Quality","Rest","Progression/Gate","Back-Rule","Notes"])
PLAN = [
 ("ALLE","Warm-up","Wrist Prep + Handgelenks-Kraeftigung","1 Block","1","1","1","Qual 4-5","-","fix vor jeder Einheit","neutral","wegen hoher Handstand-Last"),
 ("ALLE","Skill-Mikrodosis","HS Line/Balance","8-15 min","kurz","kurz","kurz","Qual 4-5","-","Frequenz > Intensitaet","ribs down","keine Max-Versuche an Nicht-Skill-Tagen"),
 ("ALLE","Skill-Mikrodosis","OAHS Line / Weight-Shift","kurz","kurz","kurz","kurz","Qual 4-5","-","taeglich, nur sauber","neutral","rechts leicht bevorzugen"),
 ("T1","Skill","OAHS Line + 2-3 saubere Holds/Seite","3/Seite","3","3","2 (Test)","Qual 4-5","60-90s","nur saubere Holds zaehlen","neutral","rechts ggf. +1 Hold"),
 ("T1","T1 Haupt","90 Heavy — Cluster Singles->Double","4x1","5x1","2x2 / Rest 1","Quality-Test","RPE 8-8,5","2-3 min","immer 1 Rep VOR Formverlust stoppen; nie Failure default","kein Arch, ribs down","Weg zu 3: Kraftdach + Cluster + Hebel"),
 ("T1","Planche","Pseudo-Planche Push-up ODER Adv-Tuck/Straddle Hold","3x6-8 / 3x8-12s","3","3","1","RPE 7-8","2 min","Straight-Arm-Fokus; Lean/Hold steigern","neutral","Schluessel-Assistance fuer 90"),
 ("T1","Vertikal-Pull","Weighted Pull-up","4x4","4-5x3-5","5x3","3x3","RPE 7-8","2-3 min","Last +, Reps sauber","neutral","Zug-Kraftbasis"),
 ("T1","Schulter-Hygiene","Scap/Serratus + External Rotation","2x10-15","2","2","1","RPE 6","45-60s","Qualitaet, kein Ego","neutral","Schultergesundheit"),
 ("T2","Skill","OAHS Shape-Density (Straddle->Half/Diamond)","3 Bloecke","3","3","beste Holds","Qual 4-5","60-90s","nur saubere Holds; Progression der Form","neutral","rechts priorisieren"),
 ("T2","Haupt-Zug","Chest-supported / Seal Row","3x10-12","4x10-12","4x8-10","2x10","RPE 7-8","90s","Last/Reps +","brustgestuetzt, kein LWS-Reiz","oberer Ruecken"),
 ("T2","Vertikal-Zug","Pull-up / Lat Pulldown Volumen","3x8-10","4x8-10","4x6-8","2x8","RPE 7-8","90s","Volumen steuern","neutral","Lat-Volumen"),
 ("T2","Schulter","Rear-Delt Fly + Face Pull","2-3x12-15","3","3","2","RPE 6-7","45-60s","Qualitaet","neutral","hintere Schulter/Rotatoren"),
 ("T2","Core","Hollow / Compression","3x20-30s","3","3","2","Qual 4","45s","Spannung halten","KEIN LWS-Reiz","Compression fuer Shapes"),
 ("T3","Status/Gate","Hinge-Pattern mit Dowel (Status-Check)","2-3 Reps","2-3","2-3","2-3","Pain-Gate","-","Gate entscheidet Tag (Rules)","Pain >3 -> regredieren","LWS + BWS Status"),
 ("T3","Bein uni","Bulgarian / FFE Split Squat","3x8-10/Seite","3x8-10","4x8","2x10","RPE 7-8","90s","Last +, Tiefe sauber","wirbelsaeulen-neutral","Quad/Glute"),
 ("T3","Bein bi","Leg Press / Hack / Belt Squat","3x10-12","3x10-12","4x10","2x12","RPE 7-8","2 min","Last +","kein axialer Grind","Quad/Glute"),
 ("T3","Post. Kette","Hamstring Curl + (gated) leichte post. Kette","3x10-12","3","3","2","RPE 7","75s","Hip-Thrust nur schmerzfrei; KEIN schwerer Hinge","Gate","Hamstrings/Glute"),
 ("T3","Bein","Calves","3x12-15","3","3","2","RPE 7","45s","Last/Reps +","neutral","Waden"),
 ("T3","BWS-Mobility","BWS Extension/Rotation, ribs down, Atmung","1 Block","1","1","1","Qual 4-5","-","fix; Protokoll in Rules","Einklemm-Thema monitoren","kein Ausstrahlen = ok"),
 ("T3","Core","Pallof + Dead Bug + Side Plank","2-3 je","3","3","2","Qual 4","45-60s","Anti-Ext/Anti-Rot/Anti-Lat","Anti-Flexion","LWS-schonend"),
 ("T4","Warm-up","Laenger + Side-Body-Aktivierung","1 Block","1","1","1","Qual 4-5","-","vor Flag noetig","neutral","Verletzungsschutz"),
 ("T4","Flag","Supported Pre-Flag-Lean","4 Entries/Seite","4","4-5","beste","Qual 4-5","90s","Support nur reduzieren wenn Re-Entry sauber","kontrolliert","Flag-Pfad"),
 ("T4","Flag","Sideline / Mini-Flag-Hold + kontrollierter Return","3/Seite","3-4","4","beste","Qual 4-5","90s","Hold-Zeit/Support progressiv","kontrollierter Return","seltene freie Versuche spaeter"),
 ("T4","90 Light","90 leicht/assistiert ODER Negativ+Pause","3-4x2-3","4x2-3","4x2","2x2","RPE 6-7","2 min","sauberes Volumen, KEIN Failure","ribs down","2. Wochendosis 90"),
 ("T4","Vertikal-Press","Wall-HSPU / Pike / DB","3x5-8","3x6-8","4x6","2x6","RPE 7-8","90s","ribs down, kein LWS-Arch","kein Arch","Vertikalkraft"),
 ("T4","Zug-Balance","One-Arm / Chest-supported Row","3x10/Seite","3","3","2","RPE 7","75s","rechts ggf. +","gestuetzt","Zug-Balance"),
 ("T5","90 Mikro","90 Technik/Negativ-Mikro","3x2","3x2-3","3x2","2x2","RPE 6-7 / Qual 4-5","2 min","3. Wochendosis, low volume, sauber","neutral","Technik scharf halten"),
 ("T5","Push-Pump","Ring-Dips / Maschine / Deficit Push-up","3x10-12","3-4x10-12","4x10","2x12","RPE 7-8","75s","getrimmt; bei Skill-Abfall zuerst kuerzen","neutral","Push-Hypertrophie"),
 ("T5","Pull-Pump","Row + Pulldown/Chin","3x10-12","4x10-12","4x10","2x12","RPE 7-8","75s","komplettiert Wochen-Zug","neutral","Pull-Hypertrophie"),
 ("T5","Schulter","Lateral Raise + Rear-Delt","2-3x12-15","3","3","2","RPE 6-7","45s","getrimmt","neutral","Schultern"),
 ("T5","Arme opt.","Biceps / Triceps","2x10-15","2-3","2","1","RPE 7","45s","optional, nur wenn Recovery ok","neutral","Arme"),
 ("T5","Regen","BWS-Mobility + Lat/Pec/Forearm + Atmung-Downshift","1 Block","1","1","1","Qual 4-5","-","Downshift Parasympathikus","BWS-Pflege","Recovery"),
]
band_color = {"ALLE":"#E2EFDA","T1":"#FCE4D6","T2":"#DDEBF7","T3":"#FFF2CC","T4":"#E2DFF5","T5":"#F2F2F2"}
r = 4
for row in PLAN:
    dtc = row[0]
    f_dt = fmt(bold=1, font_size=10, border=1, valign="top", align="center", bg_color=band_color[dtc])
    for i, val in enumerate(row):
        f = f_dt if i == 0 else (F_BODYC if i in (3,4,5,6,7,8) else F_BODY)
        put(plan, r, 1+i, val, f)
    r += 1
plan.freeze_panes("C4")

# ============================================================ 4) Exercise_Library
lib = wb.add_worksheet("B6_Exercise_Library")
for col, w in {"A:A":22,"B:B":46,"C:C":30,"D:D":12,"E:E":24,"F:H":11,"K:K":12}.items():
    lib.set_column(col, w)
title(lib, "A1:H1", "B6 — Exercise Library (Referenz fuer Voice-Logging)")
lib.merge_range("A2:H2", "Aliase Komma-getrennt, deutsch + englisch. Das Logging-Modell mappt gesprochene Begriffe "
                "ueber diese Tabelle auf die ExerciseID. TracksLoad/Hold/Side = J/N.", F_IT)
LIB_HEADERS = ["ExerciseID","Aliases","DisplayName","Category","MuscleTag","TracksLoad","TracksHold","TracksSide"]
lib.add_table(LIB_HR-1, 0, LIB_LAST-1, 7, {
    "name": "tblLib", "style": "Table Style Medium 9",
    "columns": [{"header": h} for h in LIB_HEADERS]})
for ri, row in enumerate(LIBRARY):
    rr = LIB_FIRST + ri
    for ci, val in enumerate(row):
        lib.write(rr-1, ci, val)   # plain -> Arial via Theme + Table-Banding
# DayTypes-Hilfsliste in Spalte K
lib.write(LIB_HR-1, 10, "DayTypes", F_H)
for i, d in enumerate(DAYTYPES):
    lib.write(LIB_HR + i, 10, d, F_BODYC)
lib.freeze_panes("A4")

# ============================================================ 5) Training_Log
log = wb.add_worksheet("B6_Training_Log")
LOG_HEADERS = ["Date","DayType","Week","ExerciseID","Category","Side","Sets","Reps_or_Scheme",
               "Load_kg","TopRPE","Pain_0_10","Quality_0_5","BestHold_s","CleanReps","Notes","Key"]
logw = {"A:A":12,"B:B":9,"C:C":7,"D:D":20,"E:E":12,"F:F":7,"G:G":7,"H:H":15,"I:I":9,"J:J":8,"K:K":10,"L:L":11,"M:M":11,"N:N":10,"O:O":42,"P:P":10}
for col, w in logw.items():
    log.set_column(col, w)
cat_formula = ('=IFERROR(INDEX(B6_Exercise_Library!$D$%d:$D$%d,'
               'MATCH([@ExerciseID],B6_Exercise_Library!$A$%d:$A$%d,0)),"")' % (LIB_FIRST, LIB_LAST, LIB_FIRST, LIB_LAST))
log_cols = []
for h in LOG_HEADERS:
    if h == "Category":
        log_cols.append({"header": h, "formula": cat_formula})
    elif h == "Key":
        log_cols.append({"header": h, "formula": '=[@DayType]&"_W"&[@Week]'})
    else:
        log_cols.append({"header": h})
# Log-Daten: (Date,DayType,Week,ExerciseID,Category=None,Side,Sets,Reps,Load,TopRPE,Pain,Quality,BestHold,CleanReps,Notes,Key=None)
N = None
LOG_DATA = [
 (TODAY,"T1",1,"90",N,"NA",5,"1",0,8.5,1,N,N,5,"Demo-Zeile: 90 Grad schwer, je 1 saubere Rep, Handgelenk gut",N),
 # Session 2026-06-08 (T5, Woche 1) — per Voice geloggt
 (TODAY,"T5",1,"OAHS_line",N,"L",3,N,N,N,N,4,N,N,"OAHS, vor allem links, sauber",N),
 (TODAY,"T5",1,"OAHS_line",N,"R",N,N,N,N,N,3,N,N,"OAHS rechts, 3 von 5 Qualitaet",N),
 (TODAY,"T5",1,"ring_dips",N,"NA",3,"10",0,N,N,N,N,N,N,N),
 (TODAY,"T5",1,"lat_pulldown",N,"NA",2,"10",65,N,N,N,N,N,N,N),
 (TODAY,"T5",1,"chest_row",N,"NA",2,"10",75,N,N,N,N,N,"Cable Rudern; Satz 2: 70 kg x10",N),
 (TODAY,"T5",1,"lateral_raise",N,"NA",3,"12",8,N,N,N,N,N,"Kurzhanteln",N),
 (TODAY,"T5",1,"rear_delt_fly",N,"NA",3,"12",25,N,N,N,N,N,"Reverse Butterfly, Supersatz",N),
 (TODAY,"T5",1,"scap_serratus",N,"NA",3,"12",1,N,N,N,N,N,"Y-Raises, Supersatz",N),
 (TODAY,"T5",1,"triceps",N,"NA",3,"12",16.5,N,N,N,N,N,"Kabelzug",N),
 (TODAY,"T5",1,"biceps",N,"NA",3,"10",23.5,N,N,N,N,N,"Kabelzug",N),
]
LOG_LAST = 1 + len(LOG_DATA)   # Headerzeile + Datenzeilen (1-basiert)
log.add_table("A1:P%d" % LOG_LAST, {"name": "tblLog", "style": "Table Style Medium 2", "columns": log_cols})
# Datenzeilen (Demo + geloggte Sessions). Spalten 4 (Category) & 15 (Key) = Tabellenformel -> NICHT schreiben.
def write_logrow(ridx, row):
    for ci, val in enumerate(row):
        if ci in (4, 15) or val is None or val == "":
            continue
        if ci == 0:
            log.write_datetime(ridx, 0, val, F_DATE)
        else:
            log.write(ridx, ci, val)
for i, row in enumerate(LOG_DATA):
    write_logrow(1 + i, row)
log.freeze_panes("A2")
# Datenvalidierung (Zeilen 2..1000)
log.data_validation("B2:B1000", {"validate":"list","source":"=DayTypes"})
log.data_validation("C2:C1000", {"validate":"integer","criteria":"between","minimum":1,"maximum":4})
log.data_validation("D2:D1000", {"validate":"list","source":"=ExerciseIDs"})
log.data_validation("F2:F1000", {"validate":"list","source":["L","R","NA"]})
log.data_validation("J2:J1000", {"validate":"decimal","criteria":"between","minimum":0,"maximum":10})
log.data_validation("K2:K1000", {"validate":"integer","criteria":"between","minimum":0,"maximum":10})
log.data_validation("L2:L1000", {"validate":"integer","criteria":"between","minimum":0,"maximum":5})
# Bedingte Formatierung
log.conditional_format("K2:K1000", {"type":"cell","criteria":">=","value":4,"format":F_RED})
log.conditional_format("K2:K1000", {"type":"cell","criteria":"==","value":3,"format":F_ORANGE})
log.conditional_format("K2:K1000", {"type":"cell","criteria":"between","minimum":0,"maximum":2,"format":F_GREEN})
log.conditional_format("J2:J1000", {"type":"cell","criteria":">","value":8.5,"format":F_YELLOW})

# ============================================================ 6) Nutrition_Targets
nut = wb.add_worksheet("B6_Nutrition_Targets")
for col, w in {"A:A":24,"B:B":12,"C:C":12,"D:D":12,"E:E":10,"F:F":10,"G:G":46}.items():
    nut.set_column(col, w)
title(nut, "A1:G1", "B6 — Nutrition Targets (Maintenance, KEIN Defizit)")
section(nut, "A3:C3", "Eingaben / Konstanten")
put(nut, 4, 1, "Aktuelles Gewicht (kg)", F_BOLD); put(nut, 4, 2, 80.8, numf("0.0"))
put(nut, 5, 1, "Zielkorridor (kg)", F_BOLD); put(nut, 5, 2, "80,5 - 81,0", F_BODYC)
put(nut, 6, 1, "Protein g/kg", F_BOLD); put(nut, 6, 2, 2.2, numf("0.0"))
put(nut, 7, 1, "Fett g/kg (Minimum)", F_BOLD); put(nut, 7, 2, 0.9, numf("0.0"))
put(nut, 8, 1, "Maintenance Basis (kcal)", F_BOLD); put(nut, 8, 2, 2700, numf("0"))
put(nut, 9, 1, "Protein-Ziel (g) [Formel]", F_BOLD); put(nut, 9, 2, "=ROUND(BW*B6,0)", numf("0"))
put(nut, 10, 1, "Fett-Minimum (g) [Formel]", F_BOLD); put(nut, 10, 2, "=ROUND(BW*B7,0)", numf("0"))
nut.merge_range("C4:G4", "CF: ausserhalb 80,5-81,0 wird rot markiert.", F_BODY)
nut.merge_range("C6:G6", "Protein 2,1-2,3 g/kg (~170-185 g). Fett >=0,8 g/kg. Rest = Carbs um Training.", F_BODY)
section(nut, "A12:G12", "Day-Types: kcal & Makros (Carbs per Formel = Rest)")
headers(nut, 13, ["Day Type","kcal-Delta","kcal-Ziel","Protein (g)","Fett (g)","Carbs (g)","Hinweis"])
NUT_DAYS = [
 ("T1 (CNS/heavy)", 250, "Mini-Refeed-Option: mehr Carbs vor schwerem Skill-Tag"),
 ("T2", 150, "Pull-Volumen, moderate Carbs"),
 ("T3 (Beine)", 200, "Beinlast, Carbs hoch"),
 ("T4 (Flag)", 150, "Skill-lastig, Carbs moderat-hoch"),
 ("T5 (Pump)", 100, "Hypertrophie getrimmt"),
 ("Rest", -150, "Ruhetag, Carbs runter, Protein halten"),
]
r = 14
for dt, delta, hint in NUT_DAYS:
    put(nut, r, 1, dt, F_BOLD)
    put(nut, r, 2, delta, numf("+0;-0"))
    put(nut, r, 3, f"=$B$8+B{r}", numf("0"))
    put(nut, r, 4, "=$B$9", numf("0"))
    put(nut, r, 5, "=$B$10", numf("0"))
    put(nut, r, 6, f"=ROUND((C{r}-D{r}*4-E{r}*9)/4,0)", numf("0"))
    put(nut, r, 7, hint, F_BODY)
    r += 1
section(nut, f"A{r+1}:G{r+1}", "Anpass-Trigger")
r += 2
for t in [
 "BW < 80,2 kg UND Performance runter  ->  +100 bis +150 kcal (zuerst Carbs).",
 "BW > 82 kg schnell UND Taille schlechter  ->  -100 bis -150 kcal.",
 "Vor schweren Skill-Tagen (T1/T4): optionaler Mini-Refeed (Carbs +).",
 "KEIN Defizit in diesem Block — Ziel ist Performance auf Maintenance.",
]:
    nut.merge_range(r-1, 0, r-1, 6, t, F_BODY); r += 1
nut.conditional_format("B4", {"type":"cell","criteria":"not between","minimum":80.5,"maximum":81.0,"format":F_RED})

# ============================================================ 7) Rules
rules = wb.add_worksheet("B6_Rules")
for col, w in {"A:A":26,"B:B":40,"C:C":34,"D:D":30,"E:E":30}.items():
    rules.set_column(col, w)
title(rules, "A1:E1", "B6 — Rules (Autoregulation, Gates, Hinge-Rebuild, BWS-Protokoll, Referenzen)")
r = 3
section(rules, f"A{r}:E{r}", "Vier-Stufen-Gates (Basis: Pain 0-10, Quality 0-5, RPE)")
r += 1
headers(rules, r, ["Bereich","GREEN","YELLOW","ORANGE","RED"]); r += 1
GATES = [
 ("General","Pain 0-2, Quality 4-5, RPE planmaessig -> Last/Volumen leicht hoch","Pain 3 ODER Quality 3 -> Last halten, Technik priorisieren","Pain 3 + Morgen schlechter -> 1 Stufe regredieren, Volumen kuerzen","Pain >=4 / scharf / neurologisch -> Uebung STOP, ggf. abklaeren"),
 ("90 Push-up","Rep 1 sauber, kein Schmerz -> Cluster/Last steigern","2. Rep Formverlust -> bei Singles bleiben, Hebel/Kraftdach","Schulter sperrt/Beschwerden -> auf Negativ+Pause zurueck","Schmerz Schulter/Ellbogen scharf -> stoppen"),
 ("OAHS / Flag","saubere Holds, Return kontrolliert -> Support reduzieren","wackelig/rechts sperrt -> Support halten, mehr Reps rechts","Re-Entry unsauber -> Support erhoehen, Geometrie ueben","Schulter scharf / Kontrollverlust -> stoppen"),
 ("Back / Hinge","Pain 0-1, Pattern sauber -> Hinge-Level +1 (Rules-Leiter)","Pain 2 -> Level halten, Volumen niedrig","Pain 3 ODER Morgen schlechter -> Level -1, kein Hinge-Load","Pain >=4 / Ausstrahlen / Taubheit -> STOP, abklaeren"),
]
for g in GATES:
    put(rules, r, 1, g[0], F_BOLD)
    put(rules, r, 2, g[1], gate(C_GREEN, C_GREENT)); put(rules, r, 3, g[2], gate(C_YELLOW))
    put(rules, r, 4, g[3], gate(C_ORANGE)); put(rules, r, 5, g[4], gate(C_RED, C_REDT))
    r += 1
r += 1
section(rules, f"A{r}:E{r}", "Skill-Progressionsregeln"); r += 1
for s in [
 "Hoch nur bei GREEN ueber min. 2 Sessions: Last/Hold/Support eine Stufe.",
 "Runter sofort bei ORANGE/RED oder wenn naechster Morgen schlechter.",
 "90: Weg zu 3 Reps = Kraftdach anheben (Rep 1 submax) + Cluster-Strength-Endurance + sauberer Hebel. Nie Failure default.",
 "OAHS Flag-Pfad: supported Pre-Flag -> Sideline-Hold -> reduzierter Support -> seltene freie Versuche. Support nur senken wenn Re-Entry sauber.",
 "Schwaechere (rechte) Seite leicht bevorzugen: mehr Reps/Holds, Overhead-/BWS-Mobility.",
 "Skill immer zuerst und frisch; bei Qualitaetsverlust ZUERST Push-Akzessorik kuerzen, nicht den Skill.",
]:
    rules.merge_range(r-1, 0, r-1, 4, "• " + s, F_BODY); r += 1
r += 1
section(rules, f"A{r}:E{r}", "Hinge-Rebuild-Leiter (Level 0-3) — KEIN schwerer / Barbell-Hinge in diesem Block"); r += 1
headers(rules, r, ["Level","Allowed when","Exercises","Prescription","Move up when"]); r += 1
HINGE = [
 ("0","Pain bei Dowel-Hinge 0-1, ADL schmerzfrei","Dowel Hip-Hinge, Hip-Airplane leicht, Back-Ext Iso (kurz)","2-3x5-8 Pattern, Iso 2-3x10-20s, taeglich moeglich","2 Sessions GREEN, Morgen ok"),
 ("1","Level 0 GREEN stabil","Hip-Thrust leicht, 45° Back-Ext (BW), Bird-Dog","2-3x8-12, RPE<=6, schmerzfrei","2 Sessions GREEN, kein Naechst-Morgen-Reiz"),
 ("2","Level 1 GREEN","Hip-Thrust moderat, KB Deadlift LEICHT (Pattern), Ham Curl betont","3x8-10, RPE 6-7","2 Sessions GREEN"),
 ("3","Level 2 GREEN (Ende Block / B7)","Trap-Bar / RDL LEICHT-moderat (erst naechster Block)","2-3x6-8, RPE<=7 — NICHT in B6","Naechster Block, schmerzfrei"),
]
for h in HINGE:
    put(rules, r, 1, h[0], F_BOLDC)
    for i in range(1, 5): put(rules, r, 1+i, h[i], F_BODY)
    r += 1
rules.merge_range(r-1, 0, r-1, 4, "WICHTIG: In B6 maximal Level 2. Kein schwerer Barbell-Hinge, keine Max-Grinds als Default. "
                  "Posteriore Kette ueber Ham-Curl, Back-Ext-Iso, optional Hip-Thrust (nur schmerzfrei).",
                  fmt(bold=1, font_size=9, font_color=C_REDT, border=1, text_wrap=1, valign="top"))
r += 2
section(rules, f"A{r}:E{r}", "BWS-Mobility-Protokoll (Einklemm-Thema, KEINE Diagnose — nur Training/Mobility)"); r += 1
for b in [
 "Extension: Foam-Roller BWS-Extension 2x8-10; Quadruped T-Spine Extension.",
 "Rotation: Open-Book / Quadruped Rotation 2x6-8/Seite, langsam.",
 "Rib-Position: 'ribs down', Bauch-Atmung im 90/90, 3-4 Atemzuege; kein Flaring unter Overhead-Last.",
 "Atmung/Downshift: 5 min nasale Ausatem-betonte Atmung am Sessionende (Parasympathikus).",
 "Platzierung: voll an T3 & T5, Kurzversion im taeglichen Warm-up. Monitoren via Pain-Spalte im Log.",
]:
    rules.merge_range(r-1, 0, r-1, 4, "• " + b, F_BODY); r += 1
r += 1
section(rules, f"A{r}:E{r}", "Referenzen (Quelle | URL | wofuer)"); r += 1
headers(rules, r, ["Quelle","URL","Wofuer","",""]); r += 1
REFS = [
 ("Schoenfeld u.a. (PMC8884877)","https://pmc.ncbi.nlm.nih.gov/articles/PMC8884877/","Hypertrophie-Volumen ~12-20 Saetze/Muskel (individuell)"),
 ("Load/Reps Kraft vs. Hypertrophie (PMC7927075)","https://pmc.ncbi.nlm.nih.gov/articles/PMC7927075/","Last-/Rep-Bereiche Kraft vs. Hypertrophie"),
 ("Handstand Factory","https://handstandfactory.com/one-arm-shapes/","OAHS One-Arm-Shapes (Prereqs, fortgeschritten)"),
 ("Berg Movement","https://www.bergmovement.com/calisthenics-blog/one-arm-handstand-drills-and-progressions-beginner","OAHS Drills/Progressionen"),
 ("Berg Movement","https://www.bergmovement.com/calisthenics-blog/90-degree-push-up-tutorial","90 Push-up Tutorial/Definition"),
 ("YouTube","https://www.youtube.com/watch?v=hLKoKIAp6Eg","90 Progressionen (Negativ/Band/Momentum)"),
 ("The Barbell Physio","https://thebarbellphysio.com/returning-to-deadlifts-after-back-pain/","Return-to-Deadlift nach Rueckenschmerz (Hinge zuerst)"),
]
for src, url, wof in REFS:
    put(rules, r, 1, src, F_BODY)
    rules.write_url(r-1, 1, url, F_LINK, url)
    rules.merge_range(r-1, 2, r-1, 4, wof, F_BODY)
    r += 1

# ============================================================ 8) Weekly_Summary
wsum = wb.add_worksheet("B6_Weekly_Summary")
for col, w in {"A:A":40,"B:E":10,"F:F":14}.items():
    wsum.set_column(col, w)
title(wsum, "A1:F1", "B6 — Weekly Summary (automatisch aus B6_Training_Log)")
wsum.merge_range("A2:F2", "Alle Werte per Formel ueber strukturierte Referenzen. Angehaengte Log-Zeilen fliessen automatisch ein.", F_IT)
headers(wsum, 3, ["Metrik","W1","W2","W3","W4","Block gesamt"])
def push_w(flt):
    return "=" + "+".join(f'SUMIFS(tblLog[Sets],tblLog[Category],"{c}"{flt})' for c in ("Skill90","Planche","Push"))
SUMROWS = []
def mrow(label, wf, gf, nf="0.0"):
    SUMROWS.append((label, wf, gf, nf))
mrow("Push — effektive Saetze (90/Planche/Push)", lambda w: push_w(f",tblLog[Week],{w}"), push_w(""))
mrow("Pull — effektive Saetze", lambda w: f'=SUMIFS(tblLog[Sets],tblLog[Category],"Pull",tblLog[Week],{w})', '=SUMIFS(tblLog[Sets],tblLog[Category],"Pull")')
mrow("Legs — effektive Saetze", lambda w: f'=SUMIFS(tblLog[Sets],tblLog[Category],"Legs",tblLog[Week],{w})', '=SUMIFS(tblLog[Sets],tblLog[Category],"Legs")')
mrow("90 — Touches (Eintraege)", lambda w: f'=COUNTIFS(tblLog[ExerciseID],"90",tblLog[Week],{w})', '=COUNTIFS(tblLog[ExerciseID],"90")', "0")
mrow("90 — saubere Reps gesamt", lambda w: f'=SUMIFS(tblLog[CleanReps],tblLog[ExerciseID],"90",tblLog[Week],{w})', '=SUMIFS(tblLog[CleanReps],tblLog[ExerciseID],"90")', "0")
mrow("OAHS — Skill-Bloecke (Eintraege)", lambda w: f'=COUNTIFS(tblLog[Category],"SkillOAHS",tblLog[Week],{w})', '=COUNTIFS(tblLog[Category],"SkillOAHS")', "0")
mrow("Planche — Dosen (Eintraege)", lambda w: f'=COUNTIFS(tblLog[Category],"Planche",tblLog[Week],{w})', '=COUNTIFS(tblLog[Category],"Planche")', "0")
mrow("OAHS Best-Hold links (s)", lambda w: f'=MAXIFS(tblLog[BestHold_s],tblLog[Category],"SkillOAHS",tblLog[Side],"L",tblLog[Week],{w})', '=MAXIFS(tblLog[BestHold_s],tblLog[Category],"SkillOAHS",tblLog[Side],"L")', "0")
mrow("OAHS Best-Hold rechts (s)", lambda w: f'=MAXIFS(tblLog[BestHold_s],tblLog[Category],"SkillOAHS",tblLog[Side],"R",tblLog[Week],{w})', '=MAXIFS(tblLog[BestHold_s],tblLog[Category],"SkillOAHS",tblLog[Side],"R")', "0")
mrow("Flag Best-Hold links (s)", lambda w: f'=MAXIFS(tblLog[BestHold_s],tblLog[Category],"SkillFlag",tblLog[Side],"L",tblLog[Week],{w})', '=MAXIFS(tblLog[BestHold_s],tblLog[Category],"SkillFlag",tblLog[Side],"L")', "0")
mrow("Flag Best-Hold rechts (s)", lambda w: f'=MAXIFS(tblLog[BestHold_s],tblLog[Category],"SkillFlag",tblLog[Side],"R",tblLog[Week],{w})', '=MAXIFS(tblLog[BestHold_s],tblLog[Category],"SkillFlag",tblLog[Side],"R")', "0")
mrow("Core — Saetze", lambda w: f'=SUMIFS(tblLog[Sets],tblLog[Category],"Core",tblLog[Week],{w})', '=SUMIFS(tblLog[Sets],tblLog[Category],"Core")')
mrow("Durchschnitt TopRPE", lambda w: f'=IFERROR(AVERAGEIFS(tblLog[TopRPE],tblLog[Week],{w}),0)', '=IFERROR(AVERAGE(tblLog[TopRPE]),0)')
mrow("Durchschnitt Pain", lambda w: f'=IFERROR(AVERAGEIFS(tblLog[Pain_0_10],tblLog[Week],{w}),0)', '=IFERROR(AVERAGE(tblLog[Pain_0_10]),0)')
mrow("Pain-Flags >=4", lambda w: f'=COUNTIFS(tblLog[Pain_0_10],">=4",tblLog[Week],{w})', '=COUNTIF(tblLog[Pain_0_10],">=4")', "0")
mrow("Geloggte Zeilen", lambda w: f'=COUNTIFS(tblLog[Week],{w})', '=COUNTA(tblLog[Date])', "0")
r = 4
for label, wf, gf, nf in SUMROWS:
    put(wsum, r, 1, label, F_BOLD)
    for wi, wk in enumerate([1,2,3,4]):
        put(wsum, r, 2+wi, wf(wk), numf(nf))
    put(wsum, r, 6, gf, numf(nf, bold=True))
    r += 1
wsum.freeze_panes("B4")

# ============================================================ 9) Log_Guide
guide = wb.add_worksheet("Log_Guide")
for col, w in {"A:A":8,"B:B":52,"C:C":40,"D:D":40}.items():
    guide.set_column(col, w)
title(guide, "A1:D1", "Log_Guide — Bedien- & Voice-Anleitung")
r = 3
section(guide, f"A{r}:D{r}", "Schema B6_Training_Log (genaue Reihenfolge)"); r += 1
put(guide, r, 1, "Spalte", F_H); put(guide, r, 2, "Typ / Validierung", F_H)
guide.merge_range(r-1, 2, r-1, 3, "Bedeutung", F_H); r += 1
SCHEMA = [
 ("Date","echtes Datum","Trainingstag ('heute' = aktuelles Datum)"),
 ("DayType","Dropdown T1-T5, Rest","Session-Typ"),
 ("Week","1-4","Mesozyklus-Woche"),
 ("ExerciseID","Dropdown aus Library","Uebung/Skill"),
 ("Category","FORMEL (INDEX/MATCH)","Auto-Kategorie — NICHT manuell"),
 ("Side","Dropdown L,R,NA","Seite (OAHS/unilateral), sonst NA"),
 ("Sets","Zahl","Anzahl Saetze/Bloecke"),
 ("Reps_or_Scheme","Text","z.B. 8 oder 5/3/1 oder Holds-Anzahl"),
 ("Load_kg","Zahl","Zusatzlast (BW-Uebungen = 0)"),
 ("TopRPE","0-10 (0,5er)","hoechste RPE des Eintrags"),
 ("Pain_0_10","0-10","Schmerz waehrend/dazu"),
 ("Quality_0_5","0-5","Skill-Qualitaet (bei Skill)"),
 ("BestHold_s","Zahl","bester Hold in Sekunden (bei Holds)"),
 ("CleanReps","Zahl","saubere Reps (v.a. 90/Skill)"),
 ("Notes","Text","freie Voice-Notiz"),
 ("Key","FORMEL =DayType&\"_W\"&Week","fuer Rollups — NICHT manuell"),
]
for sp, ty, be in SCHEMA:
    put(guide, r, 1, sp, F_BOLD); put(guide, r, 2, ty, F_BODY)
    guide.merge_range(r-1, 2, r-1, 3, be, F_BODY); r += 1
r += 1
section(guide, f"A{r}:D{r}", "Regeln fuer das Logging-Modell (Klartext)"); r += 1
for g in [
 "Append-only: neue Eintraege IMMER unten an B6_Training_Log anhaengen, nichts ueberschreiben.",
 "Alias -> ExerciseID ueber B6_Exercise_Library mappen. Bei nicht eindeutigem Begriff EINE kurze Rueckfrage, sonst nicht nachfragen.",
 "Nicht zutreffende Felder = NA oder leer (z.B. BestHold_s bei einer Kraftuebung).",
 "Category und Key werden per Formel gefuellt (Tabellen-Spaltenformel zieht automatisch) — NICHT manuell eintragen.",
 "Datum als echtes Datum; 'heute' = aktuelles Datum.",
 "Eine Zeile pro Uebung pro Session. OAHS/Flag mit beiden Seiten -> zwei Zeilen (Side=L und Side=R).",
]:
    guide.merge_range(r-1, 0, r-1, 3, "• " + g, F_BODY); r += 1
r += 1
section(guide, f"A{r}:D{r}", "Beispiel-Diktate -> exakte Zeile(n)"); r += 1
put(guide, r, 1, "#", F_H); put(guide, r, 2, "Diktat", F_H)
guide.merge_range(r-1, 2, r-1, 3, "Resultierende Zeile(n)", F_H); r += 1
EX = [
 ("1","Heute T1, Woche 2. 90 Grad, 5 Saetze, je 1 saubere Wiederholung, RPE 8,5, Ruecken 1, Handgelenk gut.",
  "Date=heute, DayType=T1, Week=2, ExerciseID=90, Side=NA, Sets=5, Reps_or_Scheme=1, Load_kg=0, TopRPE=8.5, Pain_0_10=1, CleanReps=5, Notes=Handgelenk gut  (Category/Key auto)"),
 ("2","OAHS Holds, links 3 Holds bester 7 Sekunden Qualitaet 4, rechts 2 Holds bester 4 Sekunden Qualitaet 3.",
  "ZWEI Zeilen, ExerciseID=OAHS_line: (Side=L, Sets=3, BestHold_s=7, Quality_0_5=4) und (Side=R, Sets=2, BestHold_s=4, Quality_0_5=3)"),
 ("3","Weighted Pull-up 4 Saetze 4 Wiederholungen plus 20 Kilo RPE 7.",
  "ExerciseID=WPU, Category(auto)=Pull, Side=NA, Sets=4, Reps_or_Scheme=4, Load_kg=20, TopRPE=7"),
 ("4","Supported Flag links 4 Entries Qualitaet 4 Ruecken 2.",
  "ExerciseID=OAHS_flag_supported, Side=L, Sets=4, Quality_0_5=4, Pain_0_10=2"),
 ("5","Pseudo Planche Push-up 3 mal 8 RPE 8.",
  "ExerciseID=pseudo_planche_pushup, Category(auto)=Planche, Sets=3, Reps_or_Scheme=8, TopRPE=8"),
 ("6","T3 Woche 1, Hinge mit Dowel 3 Reps Ruecken 1, dann Beinpresse 3 mal 12 mit 120 Kilo RPE 7.",
  "ZWEI Zeilen: (ExerciseID=hinge_pattern, Sets=3, Reps_or_Scheme=3, Pain_0_10=1) und (ExerciseID=leg_press, Sets=3, Reps_or_Scheme=12, Load_kg=120, TopRPE=7)"),
]
for n, dik, res in EX:
    put(guide, r, 1, n, F_BOLDC); put(guide, r, 2, dik, F_BODY)
    guide.merge_range(r-1, 2, r-1, 3, res, F_BODY); r += 1
r += 1
section(guide, f"A{r}:D{r}", "Kurze Bedienanleitung"); r += 1
for h in [
 "1) Taeglich: pro Uebung EINE Zeile unten in B6_Training_Log anhaengen (Dropdowns nutzen). Skill zuerst loggen.",
 "2) Voice: diktiere DayType, Woche, Uebung (Alias reicht), Saetze/Reps/Last, RPE, Pain, ggf. Quality/Hold/Seite. Das Modell mappt Alias->ExerciseID und haengt an.",
 "3) Category & Key fuellen sich automatisch (Tabellenformel). Niemals ueberschreiben.",
 "4) B6_Weekly_Summary und B6_Dashboard aktualisieren sich automatisch (Pain/RPE-Ampeln, Volumen, Best-Holds).",
 "5) PLAN anpassen: in B6_Plan_Detail / B6_Schedule. TRACKING anpassen: nur in B6_Training_Log.",
 "6) Gates & Hinge-Leiter in B6_Rules befolgen. Ernaehrung in B6_Nutrition_Targets (Maintenance).",
 "7) Folgeblock B7: B6_-Sheets klonen, BlockCode (B6_Dashboard!B3) auf B7 setzen, Prefix anpassen.",
]:
    guide.merge_range(r-1, 0, r-1, 3, h, F_BODY); r += 1

# ============================================================ Benannte Bereiche
wb.define_name("BlockCode",   "=B6_Dashboard!$B$3")
wb.define_name("BW",          "=B6_Nutrition_Targets!$B$4")
wb.define_name("ExerciseIDs", f"=B6_Exercise_Library!$A${LIB_FIRST}:$A${LIB_LAST}")
wb.define_name("DayTypes",    f"=B6_Exercise_Library!$K${LIB_FIRST}:$K${LIB_HR+len(DAYTYPES)}")

wb.close()

# ---- Theme-Font auf Arial patchen + fullCalcOnLoad erzwingen ----
tmp = PATH + ".tmp"
zin = zipfile.ZipFile(PATH)
with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for it in zin.infolist():
        data = zin.read(it.filename)
        if it.filename == "xl/theme/theme1.xml":
            t = data.decode("utf-8")
            t = t.replace('typeface="Calibri Light"', 'typeface="Arial"')
            t = t.replace('typeface="Calibri"', 'typeface="Arial"')
            data = t.encode("utf-8")
        elif it.filename == "xl/workbook.xml":
            t = data.decode("utf-8")
            if "<calcPr" in t:
                t = re.sub(r"<calcPr[^>]*/>", '<calcPr calcId="124519" fullCalcOnLoad="1"/>', t)
            else:
                t = t.replace("</workbook>", '<calcPr calcId="124519" fullCalcOnLoad="1"/></workbook>')
            data = t.encode("utf-8")
        zout.writestr(it, data)
zin.close()
shutil.move(tmp, PATH)
print("OK: Training_2.xlsx mit XlsxWriter gebaut (Theme=Arial, fullCalcOnLoad).")
