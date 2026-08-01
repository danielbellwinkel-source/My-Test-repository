# -*- coding: utf-8 -*-
"""Baut Training_3_B7.xlsx — Block B7 (Advanced OAHS), Google-Sheets-nativ.
Neu ggue B6: Recovery_1_5-Spalte, B7_Skill_Progression + B7_Warmup Sheets,
RPE-Cap 7 / getrimmtes Volumen, Shape-Logging-IDs. Prefix B7_."""
import xlsxwriter, zipfile, shutil, re, datetime
import training_data as B
import training_data_b7 as T7

PATH = "/home/user/My-Test-repository/Training_3_B7.xlsx"
LF, LL = T7.LIB_FIRST, T7.LIB_LAST
def Rcol(c): return f"'B7_Training_Log'!${c}$2:${c}$1000"
D_DEMO = datetime.datetime(2026, 7, 1)
N = None

C_DARK="#1F3864"; C_HEAD="#305496"; C_SUB="#D9E1F2"; C_BLOCK="#FFE699"
C_GREEN="#C6EFCE"; C_GREENT="#006100"; C_YELLOW="#FFEB9C"; C_ORANGE="#FFD580"; C_RED="#FFC7CE"; C_REDT="#9C0006"

wb = xlsxwriter.Workbook(PATH, {"in_memory": True}); wb.set_calc_mode("auto")
_cache = {}
def fmt(**kw):
    kw.setdefault("font_name","Arial")
    key=tuple(sorted((k,str(v)) for k,v in kw.items()))
    if key not in _cache: _cache[key]=wb.add_format(kw)
    return _cache[key]
F_TITLE=fmt(bold=1,font_size=14,font_color="white",bg_color=C_DARK,align="left",valign="vcenter")
F_SEC=fmt(bold=1,font_size=11,font_color="#1F3864",bg_color=C_SUB,align="left",valign="vcenter")
F_H=fmt(bold=1,font_size=10,font_color="white",bg_color=C_HEAD,align="center",valign="vcenter",text_wrap=1,border=1)
F_BODY=fmt(font_size=10,border=1,text_wrap=1,valign="top")
F_BODYC=fmt(font_size=10,border=1,text_wrap=1,valign="top",align="center")
F_BOLD=fmt(bold=1,font_size=10,border=1,text_wrap=1,valign="top")
F_BOLDC=fmt(bold=1,font_size=10,border=1,text_wrap=1,valign="top",align="center")
F_IT=fmt(font_size=9,italic=1,text_wrap=1,valign="top")
F_NOTE=fmt(font_size=8,italic=1,valign="top",text_wrap=1)
F_BLOCK=fmt(bold=1,font_size=12,bg_color=C_BLOCK,align="center",valign="vcenter",border=1)
F_LINK=fmt(font_size=9,font_color="#0563C1",border=1,text_wrap=1,valign="top",underline=1)
F_DATE=fmt(font_size=10,num_format="yyyy-mm-dd",border=1)
def numf(spec,**e): return fmt(font_size=10,border=1,valign="top",align="center",num_format=spec,**e)
F_RED=fmt(bg_color=C_RED,font_color=C_REDT); F_ORANGE=fmt(bg_color=C_ORANGE)
F_GREEN=fmt(bg_color=C_GREEN,font_color=C_GREENT); F_YELLOW=fmt(bg_color=C_YELLOW)
def gate(color,tcolor=None):
    d=dict(font_size=10,border=1,text_wrap=1,valign="top",bg_color=color)
    if tcolor: d["font_color"]=tcolor
    return fmt(**d)
def put(ws,row,col,val,f=None):
    r,c=row-1,col-1
    if isinstance(val,str) and val.startswith("="): ws.write_formula(r,c,val,f)
    elif val is None:
        if f: ws.write_blank(r,c,None,f)
    else: ws.write(r,c,val,f)
def title(ws,a1,t): ws.merge_range(a1,t,F_TITLE)
def section(ws,a1,t): ws.merge_range(a1,t,F_SEC)
def headers(ws,row,hs,start=1):
    for i,h in enumerate(hs): put(ws,row,start+i,h,F_H)

# ---------------- B7_Dashboard ----------------
dash=wb.add_worksheet("B7_Dashboard")
dash.set_column("A:A",34);dash.set_column("B:B",12);dash.set_column("C:C",16);dash.set_column("D:D",26);dash.set_column("E:E",40);dash.set_column("F:I",14)
title(dash,"A1:F1","B7 — Performance-Block Dashboard (Advanced OAHS)")
put(dash,3,1,"Block =",fmt(bold=1,align="right",valign="vcenter")); put(dash,3,2,"B7",F_BLOCK)
dash.merge_range("C3:F3","Folgeblock: diese Zelle auf B8 setzen, Sheets klonen, Prefix anpassen.",F_BODY)
section(dash,"A5:F5","Situationsbewertung & B6-Learnings")
r=6
for s in T7.SITU:
    put(dash,r,1,"•",F_BODYC); dash.merge_range(r-1,1,r-1,5,s,F_BODY); r+=1
r+=1; section(dash,f"A{r}:F{r}","Zielhierarchie B7"); r+=1
for g in T7.GOALS:
    dash.merge_range(r-1,0,r-1,5,g,F_BODY); r+=1
r+=1; section(dash,f"A{r}:G{r}","Volumen-Leitplanken B7 (~15-20% unter B6, Akzessorik-RPE-Cap 7)"); r+=1
headers(dash,r,["Woche / Thema","Push*","Pull/Upper-Back","Beine","90-Touches","OAHS-Bloecke","Planche-Dosen"])
for row in T7.VOL:
    r+=1
    for i,val in enumerate(row): put(dash,r,1+i,val,F_BODY if i==0 else F_BODYC)
r+=1; dash.merge_range(r-1,0,r-1,6,"*Push = 90+Planche+Dips+HSPU+Vertikalpress. Taegliche Skill-Mikrodosis zaehlt NICHT. Skill-Fokus = Progressionsarbeit (siehe B7_Skill_Progression).",F_NOTE)
r+=2; section(dash,f"A{r}:E{r}","Live-Metriken (Formeln aus Log / Weekly_Summary)"); r+=1
live_hr=r; headers(dash,r,["Metrik","Wert","Target","Interpretation","Action"]); r+=1
def live(label,valf,target,interpf,action,nf="0.0"):
    global r
    put(dash,r,1,label,F_BOLD); v=f"B{r}"
    put(dash,r,2,valf,numf(nf)); put(dash,r,3,target,F_BODYC)
    put(dash,r,4,interpf.format(v=v),F_BODY); put(dash,r,5,action,F_BODY); r+=1
live("Geloggte Zeilen gesamt",f"=COUNTA({Rcol('A')})","wachsend",'=IF({v}=0,"noch leer",IF({v}<10,"wenig Daten","laeuft"))',"Taeglich loggen, inkl. Recovery.","0")
live("Durchschnitt Recovery (1-5)",f"=IFERROR(AVERAGE({Rcol('O')}),0)",">=3",'=IF({v}=0,"-",IF({v}>=3.3,"gut",IF({v}>=2.5,"grenzwertig","zu tief -> Volumen/Rest")))',"Dauerhaft <3 -> Volumen runter / Rest rein.")
live("Durchschnitt TopRPE (Akzessorik-Check)",f"=IFERROR(AVERAGE({Rcol('J')}),0)","<=7,5",'=IF({v}=0,"-",IF({v}>7.8,"RPE-Creep!",IF({v}<6.5,"sehr locker","ok")))',"Akzessorik-Cap 7 einhalten - B6-Fehler vermeiden.")
live("Durchschnitt Pain",f"=IFERROR(AVERAGE({Rcol('K')}),0)","<2,5",'=IF({v}<2.5,"OK",IF({v}<3.5,"beobachten","regredieren"))',"Handgelenk monitoren; >Wochen -> Physio.")
live("Saubere 90-Reps gesamt",f'=SUMIFS({Rcol("N")},{Rcol("D")},"90")',"Trend hoch",'=IF({v}=0,"-","Fortschritt")',"Nur FRISCH maxen (Recovery>=3).","0")
live("Straddle Best-Hold links (s)",f'=MAXIFS({Rcol("M")},{Rcol("D")},"oahs_straddle",{Rcol("F")},"L")',"8-10s",'=IF({v}=0,"-",IF({v}>=8,"Tier-A links ok","aufbauen"))',"Konsolidieren.","0")
live("Straddle Best-Hold rechts (s)",f'=MAXIFS({Rcol("M")},{Rcol("D")},"oahs_straddle",{Rcol("F")},"R")',"an links angleichen",'=IF({v}=0,"-",IF({v}>=8,"rechts stabil","rechts priorisieren"))',"RECHTS +1 Satz.","0")
live("Diamond Best-Hold links (s)",f'=MAXIFS({Rcol("M")},{Rcol("D")},"oahs_diamond",{Rcol("F")},"L")',"8-10s",'=IF({v}=0,"-",IF({v}>=8,"ok","aufbauen"))',"Tier A.","0")
live("Figa/Flag Best-Hold (s, bester)",f'=MAX(MAXIFS({Rcol("M")},{Rcol("D")},"oahs_figa"),MAXIFS({Rcol("M")},{Rcol("D")},"oahs_flag"))',"Entries -> Hold",'=IF({v}=0,"noch offen","Fernziel laeuft")',"Erst Entries, dann Zeit.","0")
dash.conditional_format(f"B{live_hr+2}",{"type":"cell","criteria":"<=","value":2.5,"format":F_RED})
dash.conditional_format(f"B{live_hr+2}",{"type":"cell","criteria":">=","value":3.3,"format":F_GREEN})
dash.conditional_format(f"B{live_hr+4}",{"type":"cell","criteria":">=","value":2.5,"format":F_ORANGE})

# ---------------- B7_Schedule ----------------
sch=wb.add_worksheet("B7_Schedule")
for col,w in {"A:A":18,"B:B":7,"C:C":10,"D:D":30,"E:E":34,"F:F":40,"G:G":34,"H:H":26,"I:I":34,"J:J":34}.items(): sch.set_column(col,w)
title(sch,"A1:J1","B7 — Schedule (4 Wochen x 5 Tage + taegliche Skill-Mikrodosis)")
sch.merge_range("A2:J2","Skill-Fokus zuerst & frisch. Akzessorik RPE-Cap 7. Recovery 1-5 taeglich loggen -> steuert die Intensitaet.",F_IT)
sch.merge_range("A3:J3","Empfohlener Wochen-Rhythmus: T1 - T2 - REST - T3 - T4 - T5 - REST. T1 nach Ruhetag. Bei Recovery <3 im Schnitt: Extra-Rest / Volumen runter.",
                fmt(bold=1,font_size=10,bg_color="#FFF2CC",border=1,text_wrap=1,valign="vcenter"))
headers(sch,4,["Woche","Tag","Day Type","Main Focus","Skill-Touchpoint","Strength","Back/Hinge-Stress","Placement","Progression-Target","Notes"])
r=5
for wk in [1,2,3,4]:
    for dt in ["T1","T2","T3","T4","T5"]:
        focus,skill,strength,back,place=B.DAY_BASE[dt]
        put(sch,r,1,f"W{wk} {B.WEEK_THEME[wk]}",F_BODYC); put(sch,r,2,dt,F_BOLDC); put(sch,r,3,dt,F_BODYC)
        put(sch,r,4,focus,F_BODY); put(sch,r,5,skill,F_BODY); put(sch,r,6,strength,F_BODY)
        put(sch,r,7,back,F_BODY); put(sch,r,8,place,F_BODY); put(sch,r,9,B.WEEK_PROG[wk][dt],F_BODY)
        put(sch,r,10,"Skill-Fokus frisch zuerst; Akzessorik RPE<=7; Recovery loggen",F_BODY); r+=1
sch.freeze_panes("A5")

# ---------------- B7_Plan_Detail ----------------
plan=wb.add_worksheet("B7_Plan_Detail")
for col,w in {"A:A":8,"B:B":16,"C:C":44,"D:D":18,"E:E":12,"F:F":14,"G:G":14,"H:H":16,"I:I":9,"J:J":42,"K:K":22,"L:L":30}.items(): plan.set_column(col,w)
title(plan,"A1:L1","B7 — Plan Detail (RPE-Cap 7 auf Akzessorik, Skill als Progressionssystem)")
plan.merge_range("A2:L2","Skill-Fokus zuerst & frisch, Metrik loggen (siehe B7_Skill_Progression). Akzessorik max RPE 7. Recovery <=2/5 -> Skill nur Technik, Akzessorik -1 Satz.",F_IT)
headers(plan,3,["Day Type","Block","Uebung/Drill","W1","W2","W3","W4","RPE/Quality","Rest","Progression/Gate","Back-Rule","Notes"])
r=4
for row in T7.PLAN:
    f_dt=fmt(bold=1,font_size=10,border=1,valign="top",align="center",bg_color=B.BAND_COLOR.get(row[0],"#FFFFFF"))
    for i,val in enumerate(row):
        f=f_dt if i==0 else (F_BODYC if i in (3,4,5,6,7,8) else F_BODY)
        put(plan,r,1+i,val,f)
    r+=1
plan.freeze_panes("C4")

# ---------------- B7_Skill_Progression (NEU) ----------------
sk=wb.add_worksheet("B7_Skill_Progression")
for col,w in {"A:A":26,"B:B":40,"C:C":34,"D:D":34,"E:E":40}.items(): sk.set_column(col,w)
title(sk,"A1:E1","B7 — Skill-Progression OAHS (Advanced: Shapes -> Transitions -> Figa/Flag)")
sk.merge_range("A2:E2","Prinzip: pro Skill EINE messbare Metrik; auf der Stufe bleiben bis Kriterium erfuellt, DANN aufsteigen. 'Time under balance' ist die Waehrung. Kein 'irgendwie mal stehen'.",F_IT)
section(sk,"A4:E4","Progressions-Leiter (dein Einstieg: Tier A, links besitzt / rechts konsolidieren)")
headers(sk,5,["Tier","Inhalt / Shapes","Metrik","Aktueller Stand","Aufstieg wenn"])
r=6
for tier,inhalt,metrik,stand,auf in T7.SKILL_LADDER:
    put(sk,r,1,tier,F_BOLD); put(sk,r,2,inhalt,F_BODY); put(sk,r,3,metrik,F_BODY)
    put(sk,r,4,stand,gate("#FFF2CC")); put(sk,r,5,auf,gate(C_GREEN,C_GREENT)); r+=1
r+=1; section(sk,f"A{r}:E{r}","Regeln"); r+=1
for reg in T7.SKILL_RULES:
    sk.merge_range(r-1,0,r-1,4,"• "+reg,F_BODY); r+=1
r+=1; section(sk,f"A{r}:E{r}","Wie loggen (damit Progress im Dashboard sichtbar wird)"); r+=1
sk.merge_range(r-1,0,r-1,4,"Pro Shape EINE Zeile im B7_Training_Log: ExerciseID + Side (L/R) + BestHold_s + Quality_0_5. Verfuegbare Shape-IDs:",F_BODY); r+=1
shape_ids="oahs_straddle · oahs_one_leg_bent · oahs_diamond · oahs_half_straddle · oahs_legs_together · oahs_tuck · oahs_twist · oahs_transition · oahs_figa · oahs_flag"
sk.merge_range(r-1,0,r-1,4,shape_ids,fmt(font_size=10,border=1,text_wrap=1,valign="top",bg_color="#F2F6FC")); r+=1
sk.merge_range(r-1,0,r-1,4,"Beispiel-Diktat: 'T1 Woche 1, Straddle links 4 Versuche bester 7 Sekunden Qualitaet 4, rechts 3 Versuche bester 4 Sekunden Qualitaet 3' -> zwei Zeilen oahs_straddle (L/R).",F_IT); r+=1

# ---------------- B7_Warmup (NEU) ----------------
wu=wb.add_worksheet("B7_Warmup")
for col,w in {"A:A":30,"B:B":16,"C:C":80}.items(): wu.set_column(col,w)
title(wu,"A1:C1","B7 — Warm-up + Taegliche HS-Mikrodosis (recherchiert, festes Protokoll)")
wu.merge_range("A2:C2","Handgelenke sind der Engpass: Handstand = volles Koerpergewicht durch ~90 Grad Handgelenks-Extension. Kurz & progressiv aufwaermen, NICHT ermuedend.",F_IT)
section(wu,"A3:C3","1) Handgelenk-Warm-up (<=5-8 min)")
headers(wu,4,["Block","Dauer","Inhalt"])
r=5
for block,dauer,inhalt in T7.WARMUP:
    put(wu,r,1,block,F_BOLD); put(wu,r,2,dauer,F_BODYC); put(wu,r,3,inhalt,F_BODY); r+=1
r+=1; section(wu,f"A{r}:C{r}","2) Taegliche HS-Mikrodosis (8-15 min, VOR jeder Einheit) - FIXES Protokoll"); r+=1
headers(wu,r,["Block","Dauer","Inhalt"]); r+=1
for block,dauer,inhalt in T7.MICRODOSE:
    put(wu,r,1,block,F_BOLD); put(wu,r,2,dauer,F_BODYC); put(wu,r,3,inhalt,F_BODY); r+=1
r+=1
for rule in T7.MICRODOSE_RULES:
    wu.merge_range(r-1,0,r-1,2,"• "+rule,F_BODY); r+=1
r+=1; section(wu,f"A{r}:C{r}","Referenzen"); r+=1
headers(wu,r,["Quelle","URL",""]); r+=1
for src,url in list(T7.WARMUP_REFS)+list(T7.MICRODOSE_REFS):
    put(wu,r,1,src,F_BODY); wu.write_url(r-1,1,url,F_LINK,url); put(wu,r,3,"",F_BODY); r+=1

# ---------------- B7_Exercise_Library ----------------
lib=wb.add_worksheet("B7_Exercise_Library")
for col,w in {"A:A":22,"B:B":46,"C:C":34,"D:D":12,"E:E":24,"F:H":11,"K:K":12}.items(): lib.set_column(col,w)
title(lib,"A1:H1","B7 — Exercise Library (inkl. OAHS-Shapes/Transitions/Figa/Flag)")
lib.merge_range("A2:H2","Aliase deutsch+englisch. Shape-IDs fuer Skill-Progression siehe B7_Skill_Progression.",F_IT)
headers(lib,T7.LIB_HR,["ExerciseID","Aliases","DisplayName","Category","MuscleTag","TracksLoad","TracksHold","TracksSide"])
for ri,row in enumerate(T7.LIBRARY):
    rr=LF+ri
    for ci,val in enumerate(row): put(lib,rr,1+ci,val,F_BODY if ci in (1,2,4) else F_BODYC)
put(lib,T7.LIB_HR,11,"DayTypes",F_H)
for i,d in enumerate(B.DAYTYPES): put(lib,T7.LIB_HR+1+i,11,d,F_BODYC)
lib.freeze_panes("A4")

# ---------------- B7_Training_Log (mit Recovery_1_5) ----------------
log=wb.add_worksheet("B7_Training_Log")
LOG_HEADERS=["Date","DayType","Week","ExerciseID","Category","Side","Sets","Reps_or_Scheme","Load_kg","TopRPE","Pain_0_10","Quality_0_5","BestHold_s","CleanReps","Recovery_1_5","Notes","Key"]
logw={"A:A":12,"B:B":9,"C:C":7,"D:D":20,"E:E":12,"F:F":7,"G:G":7,"H:H":16,"I:I":9,"J:J":8,"K:K":10,"L:L":11,"M:M":11,"N:N":10,"O:O":11,"P:P":42,"Q:Q":10}
for col,w in logw.items(): log.set_column(col,w)
headers(log,1,LOG_HEADERS)
D1 = datetime.datetime(2026, 7, 8)   # B7 W1 T1 (Recovery-Gate 2/5)
D2 = datetime.datetime(2026, 7, 9)   # B7 W1 T2 (Recovery 3/5)
D3 = datetime.datetime(2026, 7, 11)  # B7 W1 T3 (Recovery 4/5)
D4 = datetime.datetime(2026, 7, 12)  # B7 W1 Rest statt T4 (Symptome)
D5 = datetime.datetime(2026, 7, 14)  # B7 W1 T4 (Recovery 4, daheim)
D6  = datetime.datetime(2026, 7, 18)  # Urlaub Session 1 (Rings/Parallettes T5)
D6a = datetime.datetime(2026, 7, 20)  # Urlaub Strand-Session 2
D6b = datetime.datetime(2026, 7, 22)  # Urlaub Strand-Session 3
D7  = datetime.datetime(2026, 7, 28)  # B7 W2 T1 (Recovery 3, Wieder-Einstieg)
D8  = datetime.datetime(2026, 8, 1)   # B7 W2 T2 (Recovery 5)
D9  = datetime.datetime(2026, 8, 2)   # B7 W2 T3 (Recovery ~4)
D10 = datetime.datetime(2026, 8, 4)   # B7 W2 T4 (Recovery 3->2, Trash)
D11 = datetime.datetime(2026, 8, 5)   # B7 W2 T5 (Recovery ~4-5)
D12 = datetime.datetime(2026, 8, 7)   # B7 W3 T1 (Recovery 3)
DEMO=[
 (D1,"T1",1,"oahs_straddle",N,"L",3,N,0,N,N,4,6,N,2,"Recovery-Gate 2/5; links meist >5s, sauber",N),
 (D1,"T1",1,"oahs_straddle",N,"R",3,N,0,N,N,3,5,N,2,"rechts 2-5s, inkonsistent",N),
 (D1,"T1",1,"oahs_one_leg_bent",N,"L",3,N,0,N,N,4,6,N,2,"links >5s",N),
 (D1,"T1",1,"oahs_one_leg_bent",N,"R",3,N,0,N,N,3,4,N,2,"rechts 2-5s",N),
 (D1,"T1",1,"oahs_diamond",N,"L",3,N,0,N,N,4,5,N,2,"links ~5s",N),
 (D1,"T1",1,"oahs_diamond",N,"R",3,N,0,N,N,3,4,N,2,"rechts 2-5s",N),
 (D1,"T1",1,"90",N,"NA",3,"2 Negativ",0,7,N,4,N,N,2,"Negativ-Version (Recovery-Gate); RPE 6-7, Parallettes",N),
 (D1,"T1",1,"WPU",N,"NA",2,"4",20,6,N,N,N,N,2,"Straps, leicht (Gate); 20 kg",N),
 (D1,"T1",1,"planche_hold",N,"NA",2,N,0,6,N,N,8,N,2,"Tuck Planche 2x8s",N),
 (D1,"T1",1,"ext_rotation",N,"NA",2,"12",N,6,N,N,N,N,2,"External Rotation, leichtes Band",N),
 (D1,"T1",1,"scap_serratus",N,"NA",2,"12",0,6,N,N,N,N,2,"Scap Pull-ups; linke Hand seitlich spuerbar (kein Schmerz)",N),
 # --- B7 W1 T2 (Recovery 3/5) ---
 (D2,"T2",1,"oahs_straddle",N,"L",3,N,0,N,N,4,8,N,3,"8s 2/3 links; Schulter bereits bei 8s-Holds muede",N),
 (D2,"T2",1,"oahs_straddle",N,"R",3,N,0,N,N,4,8,N,3,"8s 2/3 rechts - verbessert!",N),
 (D2,"T2",1,"oahs_one_leg_bent",N,"L",3,N,0,N,N,4,8,N,3,"8s 2/3 links",N),
 (D2,"T2",1,"oahs_one_leg_bent",N,"R",3,N,0,N,N,3,8,N,3,"8s 1/3 rechts",N),
 (D2,"T2",1,"oahs_diamond",N,"L",3,N,0,N,N,3,6,N,3,"kein 8s, 4-6s, letzter Satz muede",N),
 (D2,"T2",1,"oahs_diamond",N,"R",3,N,0,N,N,3,6,N,3,"kein 8s, 4-6s, muede",N),
 (D2,"T2",1,"chest_row",N,"NA",3,"10",90,7,N,N,N,N,3,"Seated/Chest Row",N),
 (D2,"T2",1,"lat_pulldown",N,"NA",3,"8",70,7,N,N,N,N,3,N,N),
 (D2,"T2",1,"rear_delt_fly",N,"NA",2,"15",25,6,N,N,N,N,3,"Reverse Butterfly",N),
 (D2,"T2",1,"face_pull",N,"NA",2,"15",25,7,N,N,N,N,3,N,N),
 (D2,"T2",1,"hollow",N,"NA",3,"30s",N,8,N,N,30,N,3,"3x30s",N),
 # --- B7 W1 T3 (Recovery 4/5) ---
 (D3,"T3",1,"HS_line",N,"NA",3,N,0,N,N,4,20,N,4,"20+s Line-Holds",N),
 (D3,"T3",1,"oahs_legs_together",N,"NA",2,N,0,N,N,N,N,N,4,"Legs-together Weight-Shifts, 2/Seite (Line-Arbeit)",N),
 (D3,"T3",1,"oahs_legs_together",N,"L",3,N,0,N,N,3,8,N,4,"SUPPORTED (2 Finger), 8s - KEIN freier Hold",N),
 (D3,"T3",1,"oahs_legs_together",N,"R",3,N,0,N,N,3,8,N,4,"SUPPORTED (2 Finger), 8s - KEIN freier Hold",N),
 (D3,"T3",1,"bss",N,"NA",3,"8,7,7",40,8,N,N,N,N,4,"RPE eher 8 (ueber Cap 7)",N),
 (D3,"T3",1,"leg_press",N,"NA",3,"12",70,7,N,N,N,N,4,N,N),
 (D3,"T3",1,"ham_curl",N,"NA",3,"12,10,10",50,8,N,N,N,N,4,"RPE 8 (ueber Cap 7)",N),
 (D3,"T3",1,"calves",N,"NA",3,"15",70,7,N,N,N,N,4,N,N),
 # --- B7 W1: T4 ausgelassen -> Rest wegen Symptomen ---
 (D4,"Rest",1,"rest",N,"NA",N,N,N,N,N,N,N,N,2,"Symptome: Heuschnupfen, Kopfschmerz (unueblich), Kreislauf/zittrig, Muskelzuckungen, Stress -> Rest statt T4",N),
 # --- B7 W1 T4 (Recovery 4/5, daheim wegen Zeit) ---
 (D5,"T4",1,"oahs_flag",N,"NA",4,N,0,N,N,3,3,N,4,"Straddle: OAHS-Holds + Return zum Center; sonst kurzer Straddle Flag Lean Hold",N),
 (D5,"T4",1,"oahs_flag",N,"NA",3,N,0,N,N,3,3,N,4,"Diamond Flag Lean Hold",N),
 (D5,"T4",1,"90",N,"NA",3,"3,2,2 Negativ",0,7,N,4,N,N,4,"Negative, Parallettes",N),
 (D5,"T4",1,"pike_press",N,"NA",3,"7,6,6",0,7,N,N,N,N,4,"Pike Push-ups feet elevated",N),
 (D5,"T4",1,"biceps",N,"NA",3,N,N,7,N,N,N,N,4,"Blaues Band (Pull-Ersatz daheim); Row entfiel (kein Equipment)",N),
 # --- Urlaub Session 1 (T5, Rings + tiefe Parallettes) ---
 (D6,"T5",1,"90",N,"NA",3,"2",0,7,N,N,N,N,4,"Urlaub Rings/Parallettes; RPE 6-7",N),
 (D6,"T5",1,"ring_dips",N,"NA",3,"11",0,7,N,N,N,N,4,"Ring Dips 3x11 (Urlaub)",N),
 # --- Urlaub: 2 improvisierte Strand-Sessions (Trimm-dich-Pfad) ---
 (D6a,"T5",1,"deficit_pushup",N,"NA",N,N,0,N,N,N,N,N,N,"Urlaub Strand-Session 2: Push-up-Varianten + Pull-ups + Handstaende (improvisiert, keine genauen Zahlen)",N),
 (D6b,"T5",1,"deficit_pushup",N,"NA",N,N,0,N,N,N,N,N,N,"Urlaub Strand-Session 3: Push-up-Varianten + Pull-ups + Handstaende (improvisiert)",N),
 # --- B7 W2 T1 (Recovery 3, zurueck aus Urlaub, Wieder-Einstieg) ---
 (D7,"T1",2,"oahs_straddle",N,"L",2,N,0,N,N,4,6,N,3,"W2 Re-Entry, rusty; gute Saetze >5s links",N),
 (D7,"T1",2,"oahs_straddle",N,"R",2,N,0,N,N,3,5,N,3,"paar solide + einige trash",N),
 (D7,"T1",2,"oahs_diamond",N,"L",2,N,0,N,N,3,5,N,3,"links solide",N),
 (D7,"T1",2,"oahs_diamond",N,"R",1,N,0,N,N,2,4,N,3,"Qualitaet nahm ab -> gestoppt, andere Shapes weggelassen",N),
 (D7,"T1",2,"90",N,"NA",4,"1",0,9,N,N,N,4,3,"4x1 voll; S1 8-9 (schwer hoch), S2 7-8 (Momentum gut), S3-4 ~8 (Momentum nicht ideal)",N),
 (D7,"T1",2,"pseudo_planche_pushup",N,"NA",3,"6",0,7,N,N,N,N,3,N,N),
 (D7,"T1",2,"WPU",N,"NA",4,"3",35,8,N,N,N,N,3,"Straps; fuehlt sich ueber Bloecke eher schwaecher an (war mal 40 kg) -> Double-Progression",N),
 (D7,"T1",2,"HS_line",N,"NA",2,N,0,N,N,N,N,N,3,"Endurance-Versuch improvisiert (Sprossenwand), nicht smooth -> Protokoll auf C2W-Holds umgestellt",N),
 # --- B7 W2 T2 (Recovery 5, nach 2 Extra-Rest wegen drohender Krankheit; Balance noch etwas off) ---
 (D8,"T2",2,"oahs_straddle",N,"L",3,N,0,N,N,4,5,N,5,"alle 3 Saetze 5s+ (Balance noch off nach Krankheits-Andeutung)",N),
 (D8,"T2",2,"oahs_straddle",N,"R",3,N,0,N,N,3,3,N,5,"~3s+",N),
 (D8,"T2",2,"oahs_diamond",N,"L",3,N,0,N,N,3,5,N,5,"2 Saetze 5s+, 1 Satz 3s",N),
 (D8,"T2",2,"oahs_diamond",N,"R",3,N,0,N,N,2,3,N,5,"alle ~2-3s",N),
 (D8,"T2",2,"oahs_one_leg_bent",N,"L",2,N,0,N,N,3,5,N,5,"2 Versuche >5s",N),
 (D8,"T2",2,"oahs_one_leg_bent",N,"R",2,N,0,N,N,2,5,N,5,"1x 5s + 1 trash; 3. Satz beidseitig trash -> gestoppt (Fatigue)",N),
 (D8,"T2",2,"chest_row",N,"NA",3,"10",100,7,N,N,N,N,5,"korrigiert: 100 kg (vorher 90 falsch gerechnet) -> keine echte Steigerung, nur korrektes Gewicht",N),
 (D8,"T2",2,"lat_pulldown",N,"NA",3,"8",70,7,N,N,N,N,5,N,N),
 (D8,"T2",2,"face_pull",N,"NA",3,"12",30,7,N,N,N,N,5,"haette ~3 mehr geschafft (eher RPE 7)",N),
 (D8,"T2",2,"rear_delt_fly",N,"NA",2,"12",2,7,N,N,N,N,5,"DB bent-over 2 kg (Maschine besetzt)",N),
 (D8,"T2",2,"hollow",N,"NA",3,"30s",N,7,N,N,30,N,5,"3x30s",N),
 # --- B7 W2 T3 (Recovery ~4/5, Beine + starke Endurance) ---
 (D9,"T3",2,"OAHS_line",N,"NA",1,N,0,N,N,N,N,N,4,"Taegliche Mikrodosis A-C gemacht (Fingerdruck-Kalibrierung + Weight-Shift)",N),
 (D9,"T3",2,"bss",N,"NA",3,"8",40,7,N,N,N,N,4,"40 kg jetzt @ RPE 7 (vorher 40 = RPE 8) -> Rebound",N),
 (D9,"T3",2,"leg_press",N,"NA",3,"12",80,8,N,N,N,N,4,"80 kg (75 gabs nicht) -> RPE 8, leicht ueber Cap; naechstes Mal 80x10",N),
 (D9,"T3",2,"ham_curl",N,"NA",3,"10",45,7,N,N,N,N,4,N,N),
 (D9,"T3",2,"calves",N,"NA",3,"15",80,7,N,N,N,N,4,N,N),
 (D9,"T3",2,"pallof",N,"NA",2,N,N,N,N,N,N,N,4,"2 Runden Core-Zirkel (Anti-Bewegung)",N),
 (D9,"T3",2,"HS_line",N,"NA",2,N,0,N,N,4,50,N,4,"Handstand-Endurance FREISTEHEND 2x50s (stark!)",N),
 # --- B7 W2 T4 (Recovery 3 -> auf 2 abgerutscht; Trash-Einheit) ---
 (D10,"T4",2,"oahs_flag",N,"L",4,N,0,N,N,3,3,N,2,"Straddle: unsupported Lean+Return + Freeze; links 2-3s Freeze manchmal; ein paar Fallouts",N),
 (D10,"T4",2,"oahs_flag",N,"R",4,N,0,N,N,2,1,N,2,"rechts max 1s Freeze+Return, froh es zu schaffen; Fallouts",N),
 (D10,"T4",2,"90",N,"NA",3,"2 Negativ",0,9,N,1,N,0,2,"Trash-Tag; Balance mies (untypisch), S2 fast Kontrollverlust; nur Negative, abgebrochen damit RPE nicht explodiert",N),
 (D10,"T4",2,"pike_press",N,"NA",3,"6",0,7,N,N,N,N,2,"ging noch (Balance/Last reduziert)",N),
 (D10,"T4",2,"chest_row",N,"NA",2,"10",100,7,N,N,N,N,2,"Chest-supported Row",N),
 # --- B7 W2 T5 (Recovery ~4-5, starker Rebound) ---
 (D11,"T5",2,"oahs_legs_together",N,"NA",2,N,0,N,N,3,N,N,4,"Mikrodosis: Legs-together OHNE Noodle-Support, ging gut (Tier-B-Signal)",N),
 (D11,"T5",2,"90",N,"NA",3,"2 Negativ",0,7,N,N,N,N,4,"kontrolliert, kein Failure (Recovery ~4-5)",N),
 (D11,"T5",2,"ring_dips",N,"NA",3,"12",0,6,N,N,N,N,4,"leicht (RPE 5-6) -> Progressionspotenzial",N),
 (D11,"T5",2,"deficit_pushup",N,"NA",2,"10",0,7,N,N,N,N,4,"hohe Parallettes, Fuesse erhoeht",N),
 (D11,"T5",2,"chest_row",N,"NA",3,"10",80,7,N,N,N,N,4,"Kabelzug (besserer Pump)",N),
 (D11,"T5",2,"lat_pulldown",N,"NA",3,"10,9,8",70,7,N,N,N,N,4,N,N),
 (D11,"T5",2,"lateral_raise",N,"NA",3,"12",N,7,N,N,N,N,4,"Gewicht nicht genannt (~8 kg?)",N),
 (D11,"T5",2,"rear_delt_fly",N,"NA",3,"12",30,7,N,N,N,N,4,"Maschine",N),
 (D11,"T5",2,"biceps",N,"NA",2,"12,10",25.25,7,N,N,N,N,4,"Kabel 23,5+1,75",N),
 (D11,"T5",2,"triceps",N,"NA",2,"12",18.25,7,N,N,N,N,4,"Kabel 16,5+1,75",N),
 # --- B7 W3 T1 (Recovery 3; Balance off, Kraft fuehlt schwaecher - evtl. BW hoch nach Urlaub+AYCE) ---
 (D12,"T1",3,"oahs_straddle",N,"L",3,N,0,N,N,2,4,N,3,"1-2 Holds 4s+ aber sehr shaky; schwaechstes Level seit langem",N),
 (D12,"T1",3,"oahs_straddle",N,"R",3,N,0,N,N,1,1,N,3,"max ~1s Freeze dann Fall",N),
 (D12,"T1",3,"oahs_diamond",N,"L",3,N,0,N,N,2,4,N,3,"shaky",N),
 (D12,"T1",3,"oahs_diamond",N,"R",3,N,0,N,N,1,1,N,3,N,N),
 (D12,"T1",3,"oahs_one_leg_bent",N,"L",2,N,0,N,N,2,4,N,3,N,N),
 (D12,"T1",3,"oahs_one_leg_bent",N,"R",2,N,0,N,N,1,1,N,3,N,N),
 (D12,"T1",3,"90",N,"NA",5,"3 Singles ok + 2 Fails",0,8,N,N,N,3,3,"Weg rauf Kampf, Momentum noetig; fuehlte sich nach WENIGER Kraft an (evtl. schwerer/BW hoch), nicht Ansteuerung",N),
 (D12,"T1",3,"planche_hold",N,"NA",3,N,0,7,N,N,8,N,3,"Tuck Planche 8/8/6 s",N),
 (D12,"T1",3,"WPU",N,"NA",3,"3",35,8,N,N,N,N,3,"Wide grip; recht schwach @RPE8 (evtl. BW hoch)",N),
 (D12,"T1",3,"HS_line",N,"NA",3,N,0,N,N,4,60,N,3,"Endurance freistehend Parallettes 3x60s - steigt weiter (2x50 -> 3x60)!",N),
]
def write_logrow(ridx,row):
    rno=ridx+1
    for ci,val in enumerate(row):
        if ci==4:
            log.write_formula(ridx,4,f"=IFERROR(INDEX('B7_Exercise_Library'!$D${LF}:$D${LL},MATCH($D{rno},'B7_Exercise_Library'!$A${LF}:$A${LL},0)),\"\")",F_BODYC); continue
        if ci==16:
            log.write_formula(ridx,16,f'=IF($B{rno}="","",$B{rno}&"_W"&$C{rno})',F_BODYC); continue
        if val is None or val=="": continue
        if ci==0: log.write_datetime(ridx,0,val,F_DATE)
        else: log.write(ridx,ci,val,F_BODYC if ci not in (7,15) else F_BODY)
for i,row in enumerate(DEMO): write_logrow(1+i,row)
log.freeze_panes("A2")
log.data_validation("B2:B1000",{"validate":"list","source":"=DayTypes"})
log.data_validation("C2:C1000",{"validate":"integer","criteria":"between","minimum":1,"maximum":4})
log.data_validation("D2:D1000",{"validate":"list","source":"=ExerciseIDs"})
log.data_validation("F2:F1000",{"validate":"list","source":["L","R","NA"]})
log.data_validation("J2:J1000",{"validate":"decimal","criteria":"between","minimum":0,"maximum":10})
log.data_validation("K2:K1000",{"validate":"integer","criteria":"between","minimum":0,"maximum":10})
log.data_validation("L2:L1000",{"validate":"integer","criteria":"between","minimum":0,"maximum":5})
log.data_validation("O2:O1000",{"validate":"integer","criteria":"between","minimum":1,"maximum":5})
log.conditional_format("K2:K1000",{"type":"cell","criteria":">=","value":4,"format":F_RED})
log.conditional_format("K2:K1000",{"type":"cell","criteria":"==","value":3,"format":F_ORANGE})
log.conditional_format("K2:K1000",{"type":"cell","criteria":"between","minimum":0,"maximum":2,"format":F_GREEN})
log.conditional_format("J2:J1000",{"type":"cell","criteria":">","value":8.5,"format":F_YELLOW})
log.conditional_format("O2:O1000",{"type":"cell","criteria":"<=","value":2,"format":F_RED})
log.conditional_format("O2:O1000",{"type":"cell","criteria":">=","value":4,"format":F_GREEN})

# ---------------- B7_Nutrition_Targets ----------------
nut=wb.add_worksheet("B7_Nutrition_Targets")
for col,w in {"A:A":24,"B:B":12,"C:C":12,"D:D":12,"E:E":10,"F:F":10,"G:G":46}.items(): nut.set_column(col,w)
title(nut,"A1:G1","B7 — Nutrition Targets (Maintenance, KEIN Defizit)")
section(nut,"A3:C3","Eingaben / Konstanten")
put(nut,4,1,"Aktuelles Gewicht (kg)",F_BOLD); put(nut,4,2,81.0,numf("0.0"))
put(nut,5,1,"Zielkorridor (kg)",F_BOLD); put(nut,5,2,"80,5 - 81,5",F_BODYC)
put(nut,6,1,"Protein g/kg",F_BOLD); put(nut,6,2,2.2,numf("0.0"))
put(nut,7,1,"Fett g/kg (Minimum)",F_BOLD); put(nut,7,2,0.9,numf("0.0"))
put(nut,8,1,"Maintenance Basis (kcal)",F_BOLD); put(nut,8,2,2700,numf("0"))
put(nut,9,1,"Protein-Ziel (g) [Formel]",F_BOLD); put(nut,9,2,"=ROUND(BW*B6,0)",numf("0"))
put(nut,10,1,"Fett-Minimum (g) [Formel]",F_BOLD); put(nut,10,2,"=ROUND(BW*B7,0)",numf("0"))
nut.merge_range("C4:G4","CF: ausserhalb 80,5-81,5 wird rot markiert. An grossen Aktiv-Tagen (Hike/Volleyball) Aktivitaet drauf rechnen.",F_BODY)
nut.merge_range("C6:G6","Protein 2,1-2,3 g/kg (~170-185 g). Fett >=0,8 g/kg. Rest Carbs um Training.",F_BODY)
section(nut,"A12:G12","Day-Types: kcal & Makros (Carbs per Formel = Rest)")
headers(nut,13,["Day Type","kcal-Delta","kcal-Ziel","Protein (g)","Fett (g)","Carbs (g)","Hinweis"])
r=14
for dt,delta,hint in B.NUT_DAYS:
    put(nut,r,1,dt,F_BOLD); put(nut,r,2,delta,numf("+0;-0"))
    put(nut,r,3,f"=$B$8+B{r}",numf("0")); put(nut,r,4,"=$B$9",numf("0")); put(nut,r,5,"=$B$10",numf("0"))
    put(nut,r,6,f"=ROUND((C{r}-D{r}*4-E{r}*9)/4,0)",numf("0")); put(nut,r,7,hint,F_BODY); r+=1
section(nut,f"A{r+1}:G{r+1}","Anpass-Trigger"); r+=2
for t in B.TRIG:
    nut.merge_range(r-1,0,r-1,6,t,F_BODY); r+=1
nut.conditional_format("B4",{"type":"cell","criteria":"not between","minimum":80.5,"maximum":81.5,"format":F_RED})

# ---------------- B7_Rules ----------------
rules=wb.add_worksheet("B7_Rules")
for col,w in {"A:A":26,"B:B":40,"C:C":34,"D:D":30,"E:E":30}.items(): rules.set_column(col,w)
title(rules,"A1:E1","B7 — Rules (Gates inkl. Wrist, Hinge-Rebuild, BWS, Referenzen)")
r=3; section(rules,f"A{r}:E{r}","Vier-Stufen-Gates (Pain 0-10, Quality 0-5, RPE, Recovery 1-5)"); r+=1
headers(rules,r,["Bereich","GREEN","YELLOW","ORANGE","RED"]); r+=1
GATES_B7=list(B.GATES)+[
 ("Recovery (1-5)","4-5 -> planmaessig, Skill-Test moeglich","3 -> Volumen halten, Akzessorik RPE 6-7","2 -> Skill nur Technik, Akzessorik -1 Satz & RPE 6","1 -> nur Mobility/leicht oder Rest"),
]
for g in GATES_B7:
    put(rules,r,1,g[0],F_BOLD); put(rules,r,2,g[1],gate(C_GREEN,C_GREENT)); put(rules,r,3,g[2],gate(C_YELLOW)); put(rules,r,4,g[3],gate(C_ORANGE)); put(rules,r,5,g[4],gate(C_RED,C_REDT)); r+=1
r+=1; section(rules,f"A{r}:E{r}","Skill-Progressionsregeln (siehe B7_Skill_Progression)"); r+=1
for s in T7.SKILL_RULES:
    rules.merge_range(r-1,0,r-1,4,"• "+s,F_BODY); r+=1
r+=1; section(rules,f"A{r}:E{r}","Hinge-Rebuild-Leiter (Level 0-3) — KEIN schwerer Barbell-Hinge"); r+=1
headers(rules,r,["Level","Allowed when","Exercises","Prescription","Move up when"]); r+=1
for h in B.HINGE:
    put(rules,r,1,h[0],F_BOLDC)
    for i in range(1,5): put(rules,r,1+i,h[i],F_BODY)
    r+=1
r+=1; section(rules,f"A{r}:E{r}","BWS-Mobility-Protokoll (KEINE Diagnose)"); r+=1
for b in B.BWS:
    rules.merge_range(r-1,0,r-1,4,"• "+b,F_BODY); r+=1
r+=1; section(rules,f"A{r}:E{r}","Referenzen"); r+=1
headers(rules,r,["Quelle","URL","Wofuer","",""]); r+=1
for src,url,wof in B.REFS:
    put(rules,r,1,src,F_BODY); rules.write_url(r-1,1,url,F_LINK,url); rules.merge_range(r-1,2,r-1,4,wof,F_BODY); r+=1

# ---------------- B7_Weekly_Summary ----------------
wsum=wb.add_worksheet("B7_Weekly_Summary")
for col,w in {"A:A":42,"B:E":10,"F:F":14}.items(): wsum.set_column(col,w)
title(wsum,"A1:F1","B7 — Weekly Summary (automatisch aus B7_Training_Log)")
wsum.merge_range("A2:F2","Alle Werte per Formel ueber feste Bereiche. Angehaengte Zeilen fliessen automatisch ein.",F_IT)
headers(wsum,3,["Metrik","W1","W2","W3","W4","Block gesamt"])
def push_w(flt): return "="+"+".join(f'SUMIFS({Rcol("G")},{Rcol("E")},"{c}"{flt})' for c in ("Skill90","Planche","Push"))
ROWS=[]
def mrow(label,wf,gf,nf="0.0"): ROWS.append((label,wf,gf,nf))
mrow("Push — effektive Saetze",lambda w:push_w(f",{Rcol('C')},{w}"),push_w(""))
mrow("Pull — effektive Saetze",lambda w:f'=SUMIFS({Rcol("G")},{Rcol("E")},"Pull",{Rcol("C")},{w})',f'=SUMIFS({Rcol("G")},{Rcol("E")},"Pull")')
mrow("Legs — effektive Saetze",lambda w:f'=SUMIFS({Rcol("G")},{Rcol("E")},"Legs",{Rcol("C")},{w})',f'=SUMIFS({Rcol("G")},{Rcol("E")},"Legs")')
mrow("90 — saubere Reps",lambda w:f'=SUMIFS({Rcol("N")},{Rcol("D")},"90",{Rcol("C")},{w})',f'=SUMIFS({Rcol("N")},{Rcol("D")},"90")',"0")
mrow("OAHS — Skill-Bloecke",lambda w:f'=COUNTIFS({Rcol("E")},"SkillOAHS",{Rcol("C")},{w})',f'=COUNTIF({Rcol("E")},"SkillOAHS")',"0")
mrow("Straddle Best-Hold L (s)",lambda w:f'=MAXIFS({Rcol("M")},{Rcol("D")},"oahs_straddle",{Rcol("F")},"L",{Rcol("C")},{w})',f'=MAXIFS({Rcol("M")},{Rcol("D")},"oahs_straddle",{Rcol("F")},"L")',"0")
mrow("Straddle Best-Hold R (s)",lambda w:f'=MAXIFS({Rcol("M")},{Rcol("D")},"oahs_straddle",{Rcol("F")},"R",{Rcol("C")},{w})',f'=MAXIFS({Rcol("M")},{Rcol("D")},"oahs_straddle",{Rcol("F")},"R")',"0")
mrow("Diamond Best-Hold L (s)",lambda w:f'=MAXIFS({Rcol("M")},{Rcol("D")},"oahs_diamond",{Rcol("F")},"L",{Rcol("C")},{w})',f'=MAXIFS({Rcol("M")},{Rcol("D")},"oahs_diamond",{Rcol("F")},"L")',"0")
mrow("Flag/Figa Best-Hold (s)",lambda w:f'=MAX(MAXIFS({Rcol("M")},{Rcol("D")},"oahs_flag",{Rcol("C")},{w}),MAXIFS({Rcol("M")},{Rcol("D")},"oahs_figa",{Rcol("C")},{w}))',f'=MAX(MAXIFS({Rcol("M")},{Rcol("D")},"oahs_flag"),MAXIFS({Rcol("M")},{Rcol("D")},"oahs_figa"))',"0")
mrow("Durchschnitt TopRPE",lambda w:f'=IFERROR(AVERAGEIFS({Rcol("J")},{Rcol("C")},{w}),0)',f'=IFERROR(AVERAGE({Rcol("J")}),0)')
mrow("Durchschnitt Recovery (1-5)",lambda w:f'=IFERROR(AVERAGEIFS({Rcol("O")},{Rcol("C")},{w}),0)',f'=IFERROR(AVERAGE({Rcol("O")}),0)')
mrow("Durchschnitt Pain",lambda w:f'=IFERROR(AVERAGEIFS({Rcol("K")},{Rcol("C")},{w}),0)',f'=IFERROR(AVERAGE({Rcol("K")}),0)')
mrow("Pain-Flags >=4",lambda w:f'=COUNTIFS({Rcol("K")},">=4",{Rcol("C")},{w})',f'=COUNTIF({Rcol("K")},">=4")',"0")
mrow("Geloggte Zeilen",lambda w:f'=COUNTIFS({Rcol("C")},{w})',f'=COUNTA({Rcol("A")})',"0")
r=4
for label,wf,gf,nf in ROWS:
    put(wsum,r,1,label,F_BOLD)
    for wi,wk in enumerate([1,2,3,4]): put(wsum,r,2+wi,wf(wk),numf(nf))
    put(wsum,r,6,gf,numf(nf,bold=True)); r+=1
wsum.freeze_panes("B4")

# ---------------- Log_Guide ----------------
guide=wb.add_worksheet("Log_Guide")
for col,w in {"A:A":8,"B:B":52,"C:C":40,"D:D":40}.items(): guide.set_column(col,w)
title(guide,"A1:D1","Log_Guide B7 — Bedien- & Voice-Anleitung")
r=3; section(guide,f"A{r}:D{r}","Schema B7_Training_Log (NEU: Spalte Recovery_1_5)"); r+=1
put(guide,r,1,"Spalte",F_H); put(guide,r,2,"Typ",F_H); guide.merge_range(r-1,2,r-1,3,"Bedeutung",F_H); r+=1
SCHEMA_B7=list(B.SCHEMA)
SCHEMA_B7.insert(14,("Recovery_1_5","1-5 (Dropdown)","Tages-Recovery: Schlaf/Spannung/Motivation. Steuert Intensitaet (Gate)."))
for sp,ty,be in SCHEMA_B7:
    put(guide,r,1,sp,F_BOLD); put(guide,r,2,ty,F_BODY); guide.merge_range(r-1,2,r-1,3,be,F_BODY); r+=1
r+=1; section(guide,f"A{r}:D{r}","Regeln fuer das Logging-Modell"); r+=1
for g in list(B.GRULES)+["Recovery 1-5 pro Session eintragen (erste Zeile reicht). Shapes einzeln loggen (oahs_straddle etc.) mit Side+BestHold_s+Quality."]:
    guide.merge_range(r-1,0,r-1,3,"• "+g,F_BODY); r+=1
r+=1; section(guide,f"A{r}:D{r}","Kurz-Anleitung"); r+=1
for h in ["1) Recovery (1-5) + Warm-up (B7_Warmup) zuerst.","2) Skill-Fokus frisch: aktuelle Stufe (B7_Skill_Progression), Shapes einzeln loggen (Side/Hold/Quality).","3) 90 nur bei Recovery>=3 schwer; Akzessorik RPE-Cap 7.","4) Dashboard/Weekly_Summary rechnen automatisch (inkl. Recovery, Best-Holds pro Shape).","5) Aufstieg nur per Kriterium in B7_Skill_Progression; rechts +1 Satz."]:
    guide.merge_range(r-1,0,r-1,3,h,F_BODY); r+=1

wb.define_name("BlockCode","=B7_Dashboard!$B$3")
wb.define_name("BW","=B7_Nutrition_Targets!$B$4")
wb.define_name("ExerciseIDs",f"=B7_Exercise_Library!$A${LF}:$A${LL}")
wb.define_name("DayTypes",f"=B7_Exercise_Library!$K${LF}:$K${T7.LIB_HR+len(B.DAYTYPES)}")
wb.close()

tmp=PATH+".tmp"; zin=zipfile.ZipFile(PATH)
with zipfile.ZipFile(tmp,"w",zipfile.ZIP_DEFLATED) as zout:
    for it in zin.infolist():
        data=zin.read(it.filename)
        if it.filename=="xl/theme/theme1.xml":
            t=data.decode(); t=t.replace('typeface="Calibri Light"','typeface="Arial"').replace('typeface="Calibri"','typeface="Arial"'); data=t.encode()
        elif it.filename=="xl/workbook.xml":
            t=data.decode()
            t=re.sub(r"<calcPr[^>]*/>",'<calcPr calcId="124519" fullCalcOnLoad="1"/>',t) if "<calcPr" in t else t.replace("</workbook>",'<calcPr calcId="124519" fullCalcOnLoad="1"/></workbook>')
            data=t.encode()
        zout.writestr(it,data)
zin.close(); shutil.move(tmp,PATH)
print("OK: Training_3_B7.xlsx gebaut.")
