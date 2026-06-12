# -*- coding: utf-8 -*-
"""Baut Training_2_Sheets.xlsx — Google-Sheets-NATIVE Version.
Keine Excel-Tabellen / strukturierten Referenzen. Stattdessen feste Bereiche
($..$2:$..$1000), INDEX/MATCH, SUMIFS/COUNTIFS/AVERAGEIFS/MAXIFS — funktioniert
in Google Sheets UND Excel. In Google Sheets importieren -> Dashboard rechnet live.
"""
import xlsxwriter, zipfile, shutil, re, datetime
import training_data as T

PATH = "/home/user/My-Test-repository/Training_2_Sheets.xlsx"
LF, LL = T.LIB_FIRST, T.LIB_LAST          # 4, 43
def Rcol(c): return f"'B6_Training_Log'!${c}$2:${c}$1000"

C_DARK="#1F3864"; C_HEAD="#305496"; C_SUB="#D9E1F2"; C_BLOCK="#FFE699"
C_GREEN="#C6EFCE"; C_GREENT="#006100"; C_YELLOW="#FFEB9C"; C_ORANGE="#FFD580"; C_RED="#FFC7CE"; C_REDT="#9C0006"

wb = xlsxwriter.Workbook(PATH, {"in_memory": True})
wb.set_calc_mode("auto")
_cache = {}
def fmt(**kw):
    kw.setdefault("font_name", "Arial")
    key = tuple(sorted((k, str(v)) for k, v in kw.items()))
    if key not in _cache: _cache[key] = wb.add_format(kw)
    return _cache[key]
F_TITLE = fmt(bold=1, font_size=14, font_color="white", bg_color=C_DARK, align="left", valign="vcenter")
F_SEC   = fmt(bold=1, font_size=11, font_color="#1F3864", bg_color=C_SUB, align="left", valign="vcenter")
F_H     = fmt(bold=1, font_size=10, font_color="white", bg_color=C_HEAD, align="center", valign="vcenter", text_wrap=1, border=1)
F_BODY  = fmt(font_size=10, border=1, text_wrap=1, valign="top")
F_BODYC = fmt(font_size=10, border=1, text_wrap=1, valign="top", align="center")
F_BOLD  = fmt(bold=1, font_size=10, border=1, text_wrap=1, valign="top")
F_BOLDC = fmt(bold=1, font_size=10, border=1, text_wrap=1, valign="top", align="center")
F_IT    = fmt(font_size=9, italic=1, text_wrap=1, valign="top")
F_NOTE  = fmt(font_size=8, italic=1, valign="top", text_wrap=1)
F_BLOCK = fmt(bold=1, font_size=12, bg_color=C_BLOCK, align="center", valign="vcenter", border=1)
F_LINK  = fmt(font_size=9, font_color="#0563C1", border=1, text_wrap=1, valign="top", underline=1)
F_DATE  = fmt(font_size=10, num_format="yyyy-mm-dd", border=1)
def numf(spec, **extra): return fmt(font_size=10, border=1, valign="top", align="center", num_format=spec, **extra)
F_RED=fmt(bg_color=C_RED, font_color=C_REDT); F_ORANGE=fmt(bg_color=C_ORANGE)
F_GREEN=fmt(bg_color=C_GREEN, font_color=C_GREENT); F_YELLOW=fmt(bg_color=C_YELLOW)
def gate(color, tcolor=None):
    d = dict(font_size=10, border=1, text_wrap=1, valign="top", bg_color=color)
    if tcolor: d["font_color"]=tcolor
    return fmt(**d)
def put(ws, row, col, val, f=None):
    r,c = row-1, col-1
    if isinstance(val,str) and val.startswith("="): ws.write_formula(r,c,val,f)
    elif val is None:
        if f: ws.write_blank(r,c,None,f)
    else: ws.write(r,c,val,f)
def title(ws,a1,t): ws.merge_range(a1,t,F_TITLE)
def section(ws,a1,t): ws.merge_range(a1,t,F_SEC)
def headers(ws,row,hs,start=1):
    for i,h in enumerate(hs): put(ws,row,start+i,h,F_H)

# ---------------- Dashboard ----------------
dash = wb.add_worksheet("B6_Dashboard")
dash.set_column("A:A",34); dash.set_column("B:B",12); dash.set_column("C:C",16); dash.set_column("D:D",26); dash.set_column("E:E",40); dash.set_column("F:I",14)
title(dash,"A1:F1","B6 — Performance-Block Dashboard (Google-Sheets-Version)")
put(dash,3,1,"Block =",fmt(bold=1,align="right",valign="vcenter")); put(dash,3,2,"B6",F_BLOCK)
dash.merge_range("C3:F3","Folgeblock: diese Zelle auf B7 setzen, Sheets klonen, Prefix anpassen.",F_BODY)
section(dash,"A5:F5","Situationsbewertung")
r=6
for s in T.SITU:
    put(dash,r,1,"•",F_BODYC); dash.merge_range(r-1,1,r-1,5,s,F_BODY); r+=1
r+=1; section(dash,f"A{r}:F{r}","Zielhierarchie (skill-first; Gesundheit als Constraint)"); r+=1
for g in T.GOALS:
    dash.merge_range(r-1,0,r-1,5,g,F_BODY); r+=1
r+=1; section(dash,f"A{r}:G{r}","Wochen-Themen & Volumen-Leitplanken (effektive harte Saetze/Woche)"); r+=1
headers(dash,r,["Woche / Thema","Push*","Pull/Upper-Back","Beine","90-Touches","OAHS-Bloecke","Planche-Dosen"]);
for row in T.VOL:
    r+=1
    for i,val in enumerate(row): put(dash,r,1+i,val,F_BODY if i==0 else F_BODYC)
r+=1; dash.merge_range(r-1,0,r-1,6,"*Push zaehlt 90, Planche, Dips, HSPU, Vertikalpress zusammen. Taegliche Skill-Mikrodosis zaehlt NICHT ins harte Volumen.",F_NOTE)
r+=2; section(dash,f"A{r}:E{r}","Live-Metriken (Formeln aus Log / Weekly_Summary)"); r+=1
live_hr=r; headers(dash,r,["Metrik","Wert","Target","Interpretation","Action"]); r+=1
def live(label,valf,target,interpf,action,nf="0.0"):
    global r
    put(dash,r,1,label,F_BOLD); v=f"B{r}"
    put(dash,r,2,valf,numf(nf)); put(dash,r,3,target,F_BODYC)
    put(dash,r,4,interpf.format(v=v),F_BODY); put(dash,r,5,action,F_BODY); r+=1
live("Geloggte Zeilen gesamt", f"=COUNTA({Rcol('A')})","wachsend",'=IF({v}=0,"noch leer",IF({v}<10,"wenig Daten","laeuft"))',"Taeglich 1 Zeile/Uebung anhaengen.","0")
live("Durchschnitt TopRPE (Block)", f"=IFERROR(AVERAGE({Rcol('J')}),0)","7,0-8,5",'=IF({v}=0,"-",IF({v}>8.7,"zu hoch",IF({v}<6.5,"Luft nach oben","im Ziel")))',"Wenn dauerhaft >8,5: Akzessorik kuerzen, Skill frisch halten.")
live("Durchschnitt Pain (Block)", f"=IFERROR(AVERAGE({Rcol('K')}),0)","<2,5",'=IF({v}<2.5,"OK",IF({v}<3.5,"beobachten","regredieren"))',"Pain >3 oder Morgen schlechter -> Stufe runter (Rules).")
live("Pain-Flags >=4", f'=COUNTIF({Rcol("K")},">=4")',"0",'=IF({v}=0,"OK","STOP/abklaeren")',"Scharf/neurologisch -> Uebung stoppen.","0")
live("Saubere 90-Reps gesamt", f'=SUMIFS({Rcol("N")},{Rcol("D")},"90")',"Trend hoch",'=IF({v}=0,"-","Fortschritt loggen")',"Ziel: Rep 1 submaximal -> Weg zu 3 sauberen Reps.","0")
live("OAHS Best-Hold links (s)", f'=MAXIFS({Rcol("M")},{Rcol("E")},"SkillOAHS",{Rcol("F")},"L")',"Trend hoch",'=IF({v}=0,"-",IF({v}>=5,"stabil","aufbauen"))',"Nur saubere Holds zaehlen.","0")
live("OAHS Best-Hold rechts (s)", f'=MAXIFS({Rcol("M")},{Rcol("E")},"SkillOAHS",{Rcol("F")},"R")',"naeher an links",'=IF({v}=0,"-",IF({v}>=MAXIFS('+Rcol("M")+','+Rcol("E")+',"SkillOAHS",'+Rcol("F")+',"L")*0.7,"Asymmetrie ok","rechts priorisieren"))',"Rechts mehr Reps/Holds, Overhead/BWS-Mobility.","0")
live("Flag Best-Hold links (s)", f'=MAXIFS({Rcol("M")},{Rcol("E")},"SkillFlag",{Rcol("F")},"L")',"Trend hoch",'=IF({v}=0,"-","Support reduzieren wenn sauber")',"Support nur bei sauberem Re-Entry senken.","0")
live("Flag Best-Hold rechts (s)", f'=MAXIFS({Rcol("M")},{Rcol("E")},"SkillFlag",{Rcol("F")},"R")',"Trend hoch",'=IF({v}=0,"-","Support reduzieren wenn sauber")',"Kontrollierter Return Pflicht.","0")
dash.conditional_format(f"B{live_hr+3}", {"type":"cell","criteria":">=","value":2.5,"format":F_ORANGE})
dash.conditional_format(f"B{live_hr+4}", {"type":"cell","criteria":">=","value":1,"format":F_RED})

# ---------------- Schedule ----------------
sch = wb.add_worksheet("B6_Schedule")
for col,w in {"A:A":18,"B:B":7,"C:C":10,"D:D":30,"E:E":34,"F:F":40,"G:G":34,"H:H":26,"I:I":34,"J:J":34}.items(): sch.set_column(col,w)
title(sch,"A1:J1","B6 — Schedule (4 Wochen x 5 Tage + taegliche Skill-Mikrodosis)")
sch.merge_range("A2:J2","Jede Einheit startet mit Skill-Mikrodosis (8-15 min). Keine Max-Versuche an Nicht-Skill-Tagen. Skill nie ans Ende ermuedeter Sessions.",F_IT)
sch.merge_range("A3:J3","Empfohlener Wochen-Rhythmus (5 Einheiten + 2 Ruhetage):  T1 - T2 - REST - T3 - T4 - T5 - REST.  "
                "T1 immer NACH einem Ruhe-/Low-Tag (hoechster CNS-Tag). Nicht 4 harte Tage am Stueck. "
                "Bei akkumulierter Fatigue/Schmerz (DOMS, Handgelenk) zusaetzlichen Ruhetag einschieben.",
                fmt(bold=1,font_size=10,bg_color="#FFF2CC",border=1,text_wrap=1,valign="vcenter"))
headers(sch,4,["Woche","Tag","Day Type","Main Focus","Skill-Touchpoint","Strength","Back/Hinge-Stress","Placement","Progression-Target","Notes"])
r=5
for wk in [1,2,3,4]:
    for dt in ["T1","T2","T3","T4","T5"]:
        focus,skill,strength,back,place = T.DAY_BASE[dt]
        put(sch,r,1,f"W{wk} {T.WEEK_THEME[wk]}",F_BODYC); put(sch,r,2,dt,F_BOLDC); put(sch,r,3,dt,F_BODYC)
        put(sch,r,4,focus,F_BODY); put(sch,r,5,skill,F_BODY); put(sch,r,6,strength,F_BODY)
        put(sch,r,7,back,F_BODY); put(sch,r,8,place,F_BODY); put(sch,r,9,T.WEEK_PROG[wk][dt],F_BODY)
        put(sch,r,10,"Skill-Mikrodosis zuerst; bei Qualitaetsverlust 90 -> erst Push-Akzessorik kuerzen",F_BODY); r+=1
sch.freeze_panes("A5")

# ---------------- Plan_Detail ----------------
plan = wb.add_worksheet("B6_Plan_Detail")
for col,w in {"A:A":8,"B:B":14,"C:C":40,"D:D":16,"E:E":12,"F:F":14,"G:G":14,"H:H":14,"I:I":9,"J:J":40,"K:K":22,"L:L":30}.items(): plan.set_column(col,w)
title(plan,"A1:L1","B6 — Plan Detail (Uebungsbloecke je Tag T1-T5)")
plan.merge_range("A2:L2","W1-W4 = Saetze x Reps/Holds je Woche. Skill immer zuerst, nie ermuedet.",F_IT)
headers(plan,3,["Day Type","Block","Uebung/Drill","W1","W2","W3","W4","RPE/Quality","Rest","Progression/Gate","Back-Rule","Notes"])
r=4
for row in T.PLAN:
    f_dt = fmt(bold=1,font_size=10,border=1,valign="top",align="center",bg_color=T.BAND_COLOR[row[0]])
    for i,val in enumerate(row):
        f = f_dt if i==0 else (F_BODYC if i in (3,4,5,6,7,8) else F_BODY)
        put(plan,r,1+i,val,f)
    r+=1
plan.freeze_panes("C4")

# ---------------- Exercise_Library ----------------
lib = wb.add_worksheet("B6_Exercise_Library")
for col,w in {"A:A":22,"B:B":46,"C:C":30,"D:D":12,"E:E":24,"F:H":11,"K:K":12}.items(): lib.set_column(col,w)
title(lib,"A1:H1","B6 — Exercise Library (Referenz fuer Voice-Logging)")
lib.merge_range("A2:H2","Aliase Komma-getrennt, deutsch + englisch. Logging mappt gesprochene Begriffe auf ExerciseID. TracksLoad/Hold/Side = J/N.",F_IT)
headers(lib,T.LIB_HR,["ExerciseID","Aliases","DisplayName","Category","MuscleTag","TracksLoad","TracksHold","TracksSide"])
for ri,row in enumerate(T.LIBRARY):
    rr=LF+ri
    for ci,val in enumerate(row):
        put(lib,rr,1+ci,val,F_BODY if ci in (1,2,4) else F_BODYC)
put(lib,T.LIB_HR,11,"DayTypes",F_H)
for i,d in enumerate(T.DAYTYPES): put(lib,T.LIB_HR+1+i,11,d,F_BODYC)
lib.freeze_panes("A4")

# ---------------- Training_Log (plain ranges, INDEX/MATCH) ----------------
log = wb.add_worksheet("B6_Training_Log")
LOG_HEADERS = ["Date","DayType","Week","ExerciseID","Category","Side","Sets","Reps_or_Scheme","Load_kg","TopRPE","Pain_0_10","Quality_0_5","BestHold_s","CleanReps","Notes","Key"]
logw = {"A:A":12,"B:B":9,"C:C":7,"D:D":20,"E:E":12,"F:F":7,"G:G":7,"H:H":15,"I:I":9,"J:J":8,"K:K":10,"L:L":11,"M:M":11,"N:N":10,"O:O":42,"P:P":10}
for col,w in logw.items(): log.set_column(col,w)
headers(log,1,LOG_HEADERS)
def write_logrow(ridx, row):           # ridx 0-based row index (>=1)
    rno = ridx+1
    for ci,val in enumerate(row):
        if ci==4:   # Category -> INDEX/MATCH
            log.write_formula(ridx,4, f"=IFERROR(INDEX('B6_Exercise_Library'!$D${LF}:$D${LL},MATCH($D{rno},'B6_Exercise_Library'!$A${LF}:$A${LL},0)),\"\")", F_BODYC); continue
        if ci==15:  # Key
            log.write_formula(ridx,15, f'=IF($B{rno}="","",$B{rno}&"_W"&$C{rno})', F_BODYC); continue
        if val is None or val=="":
            continue
        if ci==0: log.write_datetime(ridx,0,val,F_DATE)
        else: log.write(ridx,ci,val,F_BODYC if ci not in (7,14) else F_BODY)
for i,row in enumerate(T.LOG_DATA):
    write_logrow(1+i, row)
log.freeze_panes("A2")
log.data_validation("B2:B1000",{"validate":"list","source":"=DayTypes"})
log.data_validation("C2:C1000",{"validate":"integer","criteria":"between","minimum":1,"maximum":4})
log.data_validation("D2:D1000",{"validate":"list","source":"=ExerciseIDs"})
log.data_validation("F2:F1000",{"validate":"list","source":["L","R","NA"]})
log.data_validation("J2:J1000",{"validate":"decimal","criteria":"between","minimum":0,"maximum":10})
log.data_validation("K2:K1000",{"validate":"integer","criteria":"between","minimum":0,"maximum":10})
log.data_validation("L2:L1000",{"validate":"integer","criteria":"between","minimum":0,"maximum":5})
log.conditional_format("K2:K1000",{"type":"cell","criteria":">=","value":4,"format":F_RED})
log.conditional_format("K2:K1000",{"type":"cell","criteria":"==","value":3,"format":F_ORANGE})
log.conditional_format("K2:K1000",{"type":"cell","criteria":"between","minimum":0,"maximum":2,"format":F_GREEN})
log.conditional_format("J2:J1000",{"type":"cell","criteria":">","value":8.5,"format":F_YELLOW})

# ---------------- Nutrition_Targets ----------------
nut = wb.add_worksheet("B6_Nutrition_Targets")
for col,w in {"A:A":24,"B:B":12,"C:C":12,"D:D":12,"E:E":10,"F:F":10,"G:G":46}.items(): nut.set_column(col,w)
title(nut,"A1:G1","B6 — Nutrition Targets (Maintenance, KEIN Defizit)")
section(nut,"A3:C3","Eingaben / Konstanten")
put(nut,4,1,"Aktuelles Gewicht (kg)",F_BOLD); put(nut,4,2,80.8,numf("0.0"))
put(nut,5,1,"Zielkorridor (kg)",F_BOLD); put(nut,5,2,"80,5 - 81,0",F_BODYC)
put(nut,6,1,"Protein g/kg",F_BOLD); put(nut,6,2,2.2,numf("0.0"))
put(nut,7,1,"Fett g/kg (Minimum)",F_BOLD); put(nut,7,2,0.9,numf("0.0"))
put(nut,8,1,"Maintenance Basis (kcal)",F_BOLD); put(nut,8,2,2700,numf("0"))
put(nut,9,1,"Protein-Ziel (g) [Formel]",F_BOLD); put(nut,9,2,"=ROUND(BW*B6,0)",numf("0"))
put(nut,10,1,"Fett-Minimum (g) [Formel]",F_BOLD); put(nut,10,2,"=ROUND(BW*B7,0)",numf("0"))
nut.merge_range("C4:G4","CF: ausserhalb 80,5-81,0 wird rot markiert.",F_BODY)
nut.merge_range("C6:G6","Protein 2,1-2,3 g/kg (~170-185 g). Fett >=0,8 g/kg. Rest = Carbs um Training.",F_BODY)
section(nut,"A12:G12","Day-Types: kcal & Makros (Carbs per Formel = Rest)")
headers(nut,13,["Day Type","kcal-Delta","kcal-Ziel","Protein (g)","Fett (g)","Carbs (g)","Hinweis"])
r=14
for dt,delta,hint in T.NUT_DAYS:
    put(nut,r,1,dt,F_BOLD); put(nut,r,2,delta,numf("+0;-0"))
    put(nut,r,3,f"=$B$8+B{r}",numf("0")); put(nut,r,4,"=$B$9",numf("0")); put(nut,r,5,"=$B$10",numf("0"))
    put(nut,r,6,f"=ROUND((C{r}-D{r}*4-E{r}*9)/4,0)",numf("0")); put(nut,r,7,hint,F_BODY); r+=1
section(nut,f"A{r+1}:G{r+1}","Anpass-Trigger"); r+=2
for t in T.TRIG:
    nut.merge_range(r-1,0,r-1,6,t,F_BODY); r+=1
nut.conditional_format("B4",{"type":"cell","criteria":"not between","minimum":80.5,"maximum":81.0,"format":F_RED})

# ---------------- Rules ----------------
rules = wb.add_worksheet("B6_Rules")
for col,w in {"A:A":26,"B:B":40,"C:C":34,"D:D":30,"E:E":30}.items(): rules.set_column(col,w)
title(rules,"A1:E1","B6 — Rules (Gates, Hinge-Rebuild, BWS-Protokoll, Referenzen)")
r=3; section(rules,f"A{r}:E{r}","Vier-Stufen-Gates (Pain 0-10, Quality 0-5, RPE)"); r+=1
headers(rules,r,["Bereich","GREEN","YELLOW","ORANGE","RED"]); r+=1
for g in T.GATES:
    put(rules,r,1,g[0],F_BOLD); put(rules,r,2,g[1],gate(C_GREEN,C_GREENT)); put(rules,r,3,g[2],gate(C_YELLOW)); put(rules,r,4,g[3],gate(C_ORANGE)); put(rules,r,5,g[4],gate(C_RED,C_REDT)); r+=1
r+=1; section(rules,f"A{r}:E{r}","Skill-Progressionsregeln"); r+=1
for s in T.SKILLR:
    rules.merge_range(r-1,0,r-1,4,"• "+s,F_BODY); r+=1
r+=1; section(rules,f"A{r}:E{r}","Hinge-Rebuild-Leiter (Level 0-3) — KEIN schwerer / Barbell-Hinge in diesem Block"); r+=1
headers(rules,r,["Level","Allowed when","Exercises","Prescription","Move up when"]); r+=1
for h in T.HINGE:
    put(rules,r,1,h[0],F_BOLDC)
    for i in range(1,5): put(rules,r,1+i,h[i],F_BODY)
    r+=1
rules.merge_range(r-1,0,r-1,4,"WICHTIG: In B6 maximal Level 2. Kein schwerer Barbell-Hinge, keine Max-Grinds als Default.",fmt(bold=1,font_size=9,font_color=C_REDT,border=1,text_wrap=1,valign="top")); r+=2
section(rules,f"A{r}:E{r}","BWS-Mobility-Protokoll (Einklemm-Thema, KEINE Diagnose)"); r+=1
for b in T.BWS:
    rules.merge_range(r-1,0,r-1,4,"• "+b,F_BODY); r+=1
r+=1; section(rules,f"A{r}:E{r}","Referenzen (Quelle | URL | wofuer)"); r+=1
headers(rules,r,["Quelle","URL","Wofuer","",""]); r+=1
for src,url,wof in T.REFS:
    put(rules,r,1,src,F_BODY); rules.write_url(r-1,1,url,F_LINK,url); rules.merge_range(r-1,2,r-1,4,wof,F_BODY); r+=1

# ---------------- Weekly_Summary (range formulas) ----------------
wsum = wb.add_worksheet("B6_Weekly_Summary")
for col,w in {"A:A":40,"B:E":10,"F:F":14}.items(): wsum.set_column(col,w)
title(wsum,"A1:F1","B6 — Weekly Summary (automatisch aus B6_Training_Log)")
wsum.merge_range("A2:F2","Alle Werte per Formel ueber feste Bereiche. Angehaengte Log-Zeilen fliessen automatisch ein.",F_IT)
headers(wsum,3,["Metrik","W1","W2","W3","W4","Block gesamt"])
def push_w(flt): return "=" + "+".join(f'SUMIFS({Rcol("G")},{Rcol("E")},"{c}"{flt})' for c in ("Skill90","Planche","Push"))
ROWS=[]
def mrow(label,wf,gf,nf="0.0"): ROWS.append((label,wf,gf,nf))
mrow("Push — effektive Saetze (90/Planche/Push)", lambda w: push_w(f",{Rcol('C')},{w}"), push_w(""))
mrow("Pull — effektive Saetze", lambda w: f'=SUMIFS({Rcol("G")},{Rcol("E")},"Pull",{Rcol("C")},{w})', f'=SUMIFS({Rcol("G")},{Rcol("E")},"Pull")')
mrow("Legs — effektive Saetze", lambda w: f'=SUMIFS({Rcol("G")},{Rcol("E")},"Legs",{Rcol("C")},{w})', f'=SUMIFS({Rcol("G")},{Rcol("E")},"Legs")')
mrow("90 — Touches (Eintraege)", lambda w: f'=COUNTIFS({Rcol("D")},"90",{Rcol("C")},{w})', f'=COUNTIF({Rcol("D")},"90")',"0")
mrow("90 — saubere Reps gesamt", lambda w: f'=SUMIFS({Rcol("N")},{Rcol("D")},"90",{Rcol("C")},{w})', f'=SUMIFS({Rcol("N")},{Rcol("D")},"90")',"0")
mrow("OAHS — Skill-Bloecke (Eintraege)", lambda w: f'=COUNTIFS({Rcol("E")},"SkillOAHS",{Rcol("C")},{w})', f'=COUNTIF({Rcol("E")},"SkillOAHS")',"0")
mrow("Planche — Dosen (Eintraege)", lambda w: f'=COUNTIFS({Rcol("E")},"Planche",{Rcol("C")},{w})', f'=COUNTIF({Rcol("E")},"Planche")',"0")
mrow("OAHS Best-Hold links (s)", lambda w: f'=MAXIFS({Rcol("M")},{Rcol("E")},"SkillOAHS",{Rcol("F")},"L",{Rcol("C")},{w})', f'=MAXIFS({Rcol("M")},{Rcol("E")},"SkillOAHS",{Rcol("F")},"L")',"0")
mrow("OAHS Best-Hold rechts (s)", lambda w: f'=MAXIFS({Rcol("M")},{Rcol("E")},"SkillOAHS",{Rcol("F")},"R",{Rcol("C")},{w})', f'=MAXIFS({Rcol("M")},{Rcol("E")},"SkillOAHS",{Rcol("F")},"R")',"0")
mrow("Flag Best-Hold links (s)", lambda w: f'=MAXIFS({Rcol("M")},{Rcol("E")},"SkillFlag",{Rcol("F")},"L",{Rcol("C")},{w})', f'=MAXIFS({Rcol("M")},{Rcol("E")},"SkillFlag",{Rcol("F")},"L")',"0")
mrow("Flag Best-Hold rechts (s)", lambda w: f'=MAXIFS({Rcol("M")},{Rcol("E")},"SkillFlag",{Rcol("F")},"R",{Rcol("C")},{w})', f'=MAXIFS({Rcol("M")},{Rcol("E")},"SkillFlag",{Rcol("F")},"R")',"0")
mrow("Core — Saetze", lambda w: f'=SUMIFS({Rcol("G")},{Rcol("E")},"Core",{Rcol("C")},{w})', f'=SUMIFS({Rcol("G")},{Rcol("E")},"Core")')
mrow("Durchschnitt TopRPE", lambda w: f'=IFERROR(AVERAGEIFS({Rcol("J")},{Rcol("C")},{w}),0)', f'=IFERROR(AVERAGE({Rcol("J")}),0)')
mrow("Durchschnitt Pain", lambda w: f'=IFERROR(AVERAGEIFS({Rcol("K")},{Rcol("C")},{w}),0)', f'=IFERROR(AVERAGE({Rcol("K")}),0)')
mrow("Pain-Flags >=4", lambda w: f'=COUNTIFS({Rcol("K")},">=4",{Rcol("C")},{w})', f'=COUNTIF({Rcol("K")},">=4")',"0")
mrow("Geloggte Zeilen", lambda w: f'=COUNTIFS({Rcol("C")},{w})', f'=COUNTA({Rcol("A")})',"0")
r=4
for label,wf,gf,nf in ROWS:
    put(wsum,r,1,label,F_BOLD)
    for wi,wk in enumerate([1,2,3,4]): put(wsum,r,2+wi,wf(wk),numf(nf))
    put(wsum,r,6,gf,numf(nf,bold=True)); r+=1
wsum.freeze_panes("B4")

# ---------------- Log_Guide ----------------
guide = wb.add_worksheet("Log_Guide")
for col,w in {"A:A":8,"B:B":52,"C:C":40,"D:D":40}.items(): guide.set_column(col,w)
title(guide,"A1:D1","Log_Guide — Bedien- & Voice-Anleitung (Google-Sheets-Version)")
r=3; section(guide,f"A{r}:D{r}","Schema B6_Training_Log (genaue Reihenfolge)"); r+=1
put(guide,r,1,"Spalte",F_H); put(guide,r,2,"Typ / Validierung",F_H); guide.merge_range(r-1,2,r-1,3,"Bedeutung",F_H); r+=1
for sp,ty,be in T.SCHEMA:
    put(guide,r,1,sp,F_BOLD); put(guide,r,2,ty,F_BODY); guide.merge_range(r-1,2,r-1,3,be,F_BODY); r+=1
r+=1; section(guide,f"A{r}:D{r}","Regeln fuer das Logging-Modell"); r+=1
for g in T.GRULES:
    guide.merge_range(r-1,0,r-1,3,"• "+g,F_BODY); r+=1
r+=1; section(guide,f"A{r}:D{r}","Beispiel-Diktate -> exakte Zeile(n)"); r+=1
put(guide,r,1,"#",F_H); put(guide,r,2,"Diktat",F_H); guide.merge_range(r-1,2,r-1,3,"Resultierende Zeile(n)",F_H); r+=1
for n,dik,res in T.EX:
    put(guide,r,1,n,F_BOLDC); put(guide,r,2,dik,F_BODY); guide.merge_range(r-1,2,r-1,3,res,F_BODY); r+=1
r+=1; section(guide,f"A{r}:D{r}","Kurze Bedienanleitung"); r+=1
for h in T.HOWTO:
    guide.merge_range(r-1,0,r-1,3,h,F_BODY); r+=1

wb.define_name("BlockCode","=B6_Dashboard!$B$3")
wb.define_name("BW","=B6_Nutrition_Targets!$B$4")
wb.define_name("ExerciseIDs", f"=B6_Exercise_Library!$A${LF}:$A${LL}")
wb.define_name("DayTypes", f"=B6_Exercise_Library!$K${LF}:$K${T.LIB_HR+len(T.DAYTYPES)}")
wb.close()

# Theme-Font Arial + fullCalcOnLoad
tmp = PATH+".tmp"; zin = zipfile.ZipFile(PATH)
with zipfile.ZipFile(tmp,"w",zipfile.ZIP_DEFLATED) as zout:
    for it in zin.infolist():
        data = zin.read(it.filename)
        if it.filename=="xl/theme/theme1.xml":
            t=data.decode(); t=t.replace('typeface="Calibri Light"','typeface="Arial"').replace('typeface="Calibri"','typeface="Arial"'); data=t.encode()
        elif it.filename=="xl/workbook.xml":
            t=data.decode()
            t = re.sub(r"<calcPr[^>]*/>", '<calcPr calcId="124519" fullCalcOnLoad="1"/>', t) if "<calcPr" in t else t.replace("</workbook>",'<calcPr calcId="124519" fullCalcOnLoad="1"/></workbook>')
            data=t.encode()
        zout.writestr(it,data)
zin.close(); shutil.move(tmp,PATH)
print("OK: Training_2_Sheets.xlsx gebaut (Google-Sheets-nativ).")
