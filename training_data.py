# -*- coding: utf-8 -*-
"""Gemeinsame Daten fuer die Workbook-Builder (Single Source of Truth)."""
import datetime

TODAY = datetime.datetime(2026, 6, 8)
D_T1  = datetime.datetime(2026, 6, 10)
N = None

# (ID, Aliases, DisplayName, Category, MuscleTag, Load, Hold, Side)
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
LIB_FIRST = LIB_HR + 1
LIB_LAST = LIB_HR + len(LIBRARY)

SITU = [
 "Fortgeschrittener, autodidaktischer Hand-Balancer (1,90 m, ~80,8 kg). Skills von Canes auf den Boden uebertragen.",
 "Erster Performance-Block nach 6-Monats-Diaet (90 -> ~81 kg). Maintenance, KEIN Defizit, begrenzte Recovery.",
 "OAHS: links stabil, rechts limitiert (Schulter-Mobility/Koordination). Unilaterale Arbeit leicht zugunsten rechts.",
 "90 Push-up: 1 saubere Rep meist gut, 2. Rep sehr schwer. Ziel = 3 saubere Reps am Stueck.",
 "Ruecken (a) LWS/sakral: alte Reizung nach RDL/Deadlift, am Ausheilen -> Hinge-Rebuild ueber Pain-Gates.",
 "Ruecken (b) BWS: Einklemm-Gefuehl mittig/oben, ohne Ausstrahlen -> Mobility-/Tissue-Baustein, monitoren, KEINE Diagnose.",
 "Prinzip: Skill-Qualitaet + Gelenk-/Rueckengesundheit schlagen Volumen und Fatigue-PRs.",
]
GOALS = [
 "1) 90 Push-up -> 3 saubere Reps am Stueck (Hauptziel).",
 "2) OAHS Shapes + Flag-Pfad (Straddle -> Shapes; supported Pre-Flag -> Sideline-Hold -> reduzierter Support -> seltene freie Versuche).",
 "3) Modern-Handstand-Linie und weitere Shapes (Flieger/Figa offen, Fernziel, nicht erzwingen).",
 "4) Kraftunterbau (Planche/Push/Pull) als Mittel zum Zweck fuer 1-3.",
 "5) Hypertrophie nur soweit sie Skills oder Schultergesundheit stuetzt.",
]
VOL = [
 ("W1 Re-Entry / Baseline","12-14","14-16","8-10","3 (1 schwer / 2 leicht)","3","2"),
 ("W2 Build","14-16","16-18","10-12","3","3","2-3"),
 ("W3 Peak","15-16","17-18","12-14","3 (1 schwer / 2 assist)","3","3"),
 ("W4 Consolidate / Test","9-12","10-12","6-8","2 (Quality-Test)","beste Holds","1"),
]
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
BAND_COLOR = {"ALLE":"#E2EFDA","T1":"#FCE4D6","T2":"#DDEBF7","T3":"#FFF2CC","T4":"#E2DFF5","T5":"#F2F2F2"}
NUT_DAYS = [
 ("T1 (CNS/heavy)", 250, "Mini-Refeed-Option: mehr Carbs vor schwerem Skill-Tag"),
 ("T2", 150, "Pull-Volumen, moderate Carbs"),
 ("T3 (Beine)", 200, "Beinlast, Carbs hoch"),
 ("T4 (Flag)", 150, "Skill-lastig, Carbs moderat-hoch"),
 ("T5 (Pump)", 100, "Hypertrophie getrimmt"),
 ("Rest", -150, "Ruhetag, Carbs runter, Protein halten"),
]
TRIG = [
 "BW < 80,2 kg UND Performance runter  ->  +100 bis +150 kcal (zuerst Carbs).",
 "BW > 82 kg schnell UND Taille schlechter  ->  -100 bis -150 kcal.",
 "Vor schweren Skill-Tagen (T1/T4): optionaler Mini-Refeed (Carbs +).",
 "KEIN Defizit in diesem Block — Ziel ist Performance auf Maintenance.",
]
GATES = [
 ("General","Pain 0-2, Quality 4-5, RPE planmaessig -> Last/Volumen leicht hoch","Pain 3 ODER Quality 3 -> Last halten, Technik priorisieren","Pain 3 + Morgen schlechter -> 1 Stufe regredieren, Volumen kuerzen","Pain >=4 / scharf / neurologisch -> Uebung STOP, ggf. abklaeren"),
 ("90 Push-up","Rep 1 sauber, kein Schmerz -> Cluster/Last steigern","2. Rep Formverlust -> bei Singles bleiben, Hebel/Kraftdach","Schulter sperrt/Beschwerden -> auf Negativ+Pause zurueck","Schmerz Schulter/Ellbogen scharf -> stoppen"),
 ("OAHS / Flag","saubere Holds, Return kontrolliert -> Support reduzieren","wackelig/rechts sperrt -> Support halten, mehr Reps rechts","Re-Entry unsauber -> Support erhoehen, Geometrie ueben","Schulter scharf / Kontrollverlust -> stoppen"),
 ("Back / Hinge","Pain 0-1, Pattern sauber -> Hinge-Level +1 (Rules-Leiter)","Pain 2 -> Level halten, Volumen niedrig","Pain 3 ODER Morgen schlechter -> Level -1, kein Hinge-Load","Pain >=4 / Ausstrahlen / Taubheit -> STOP, abklaeren"),
 ("Wrist / Handgelenk","Pain 0-2, kein Ziehen -> normal laden, Wrist-Prep","Leichtes Ziehen (z.B. ulnar/Pisiform) -> auf Canes/Parallettes, Volumen runter, Hand neu setzen","Ziehen nimmt zu / Morgen schlechter -> OAHS-Last raus, nur Prep/Mobility","Scharf / Taubheit / Kribbeln -> stoppen, abklaeren"),
]
SKILLR = [
 "Hoch nur bei GREEN ueber min. 2 Sessions: Last/Hold/Support eine Stufe.",
 "Runter sofort bei ORANGE/RED oder wenn naechster Morgen schlechter.",
 "90: Weg zu 3 Reps = Kraftdach anheben (Rep 1 submax) + Cluster-Strength-Endurance + sauberer Hebel. Nie Failure default.",
 "OAHS Flag-Pfad: supported Pre-Flag -> Sideline-Hold -> reduzierter Support -> seltene freie Versuche. Support nur senken wenn Re-Entry sauber.",
 "Schwaechere (rechte) Seite leicht bevorzugen: mehr Reps/Holds, Overhead-/BWS-Mobility.",
 "Skill immer zuerst und frisch; bei Qualitaetsverlust ZUERST Push-Akzessorik kuerzen, nicht den Skill.",
]
HINGE = [
 ("0","Pain bei Dowel-Hinge 0-1, ADL schmerzfrei","Dowel Hip-Hinge, Hip-Airplane leicht, Back-Ext Iso (kurz)","2-3x5-8 Pattern, Iso 2-3x10-20s, taeglich moeglich","2 Sessions GREEN, Morgen ok"),
 ("1","Level 0 GREEN stabil","Hip-Thrust leicht, 45 Grad Back-Ext (BW), Bird-Dog","2-3x8-12, RPE<=6, schmerzfrei","2 Sessions GREEN, kein Naechst-Morgen-Reiz"),
 ("2","Level 1 GREEN","Hip-Thrust moderat, KB Deadlift LEICHT (Pattern), Ham Curl betont","3x8-10, RPE 6-7","2 Sessions GREEN"),
 ("3","Level 2 GREEN (Ende Block / B7)","Trap-Bar / RDL LEICHT-moderat (erst naechster Block)","2-3x6-8, RPE<=7 — NICHT in B6","Naechster Block, schmerzfrei"),
]
BWS = [
 "Extension: Foam-Roller BWS-Extension 2x8-10; Quadruped T-Spine Extension.",
 "Rotation: Open-Book / Quadruped Rotation 2x6-8/Seite, langsam.",
 "Rib-Position: 'ribs down', Bauch-Atmung im 90/90, 3-4 Atemzuege; kein Flaring unter Overhead-Last.",
 "Atmung/Downshift: 5 min nasale Ausatem-betonte Atmung am Sessionende (Parasympathikus).",
 "Platzierung: voll an T3 & T5, Kurzversion im taeglichen Warm-up. Monitoren via Pain-Spalte im Log.",
]
REFS = [
 ("Schoenfeld u.a. (PMC8884877)","https://pmc.ncbi.nlm.nih.gov/articles/PMC8884877/","Hypertrophie-Volumen ~12-20 Saetze/Muskel (individuell)"),
 ("Load/Reps Kraft vs. Hypertrophie (PMC7927075)","https://pmc.ncbi.nlm.nih.gov/articles/PMC7927075/","Last-/Rep-Bereiche Kraft vs. Hypertrophie"),
 ("Handstand Factory","https://handstandfactory.com/one-arm-shapes/","OAHS One-Arm-Shapes (Prereqs, fortgeschritten)"),
 ("Berg Movement","https://www.bergmovement.com/calisthenics-blog/one-arm-handstand-drills-and-progressions-beginner","OAHS Drills/Progressionen"),
 ("Berg Movement","https://www.bergmovement.com/calisthenics-blog/90-degree-push-up-tutorial","90 Push-up Tutorial/Definition"),
 ("YouTube","https://www.youtube.com/watch?v=hLKoKIAp6Eg","90 Progressionen (Negativ/Band/Momentum)"),
 ("The Barbell Physio","https://thebarbellphysio.com/returning-to-deadlifts-after-back-pain/","Return-to-Deadlift nach Rueckenschmerz (Hinge zuerst)"),
]
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
GRULES = [
 "Append-only: neue Eintraege IMMER unten anhaengen, nichts ueberschreiben.",
 "Alias -> ExerciseID ueber B6_Exercise_Library mappen. Bei nicht eindeutigem Begriff EINE kurze Rueckfrage, sonst nicht nachfragen.",
 "Nicht zutreffende Felder = NA oder leer (z.B. BestHold_s bei einer Kraftuebung).",
 "Category und Key werden per Formel/Skript gefuellt — NICHT manuell eintragen.",
 "Datum als echtes Datum; 'heute' = aktuelles Datum.",
 "Eine Zeile pro Uebung pro Session. OAHS/Flag mit beiden Seiten -> zwei Zeilen (Side=L und Side=R).",
]
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
HOWTO = [
 "1) Taeglich: pro Uebung EINE Zeile unten in B6_Training_Log anhaengen (Dropdowns nutzen). Skill zuerst loggen.",
 "2) Voice: diktiere DayType, Woche, Uebung (Alias reicht), Saetze/Reps/Last, RPE, Pain, ggf. Quality/Hold/Seite.",
 "3) Category & Key fuellen sich per Formel/Skript. Niemals ueberschreiben.",
 "4) B6_Weekly_Summary und B6_Dashboard aktualisieren sich automatisch.",
 "5) PLAN anpassen: in B6_Plan_Detail / B6_Schedule. TRACKING anpassen: nur in B6_Training_Log.",
 "6) Gates & Hinge-Leiter in B6_Rules befolgen. Ernaehrung in B6_Nutrition_Targets (Maintenance).",
 "7) Folgeblock B7: Sheets klonen, BlockCode (B6_Dashboard!B3) auf B7 setzen, Prefix anpassen.",
]
# Log-Daten: (Date,DayType,Week,ExerciseID,Cat=N,Side,Sets,Reps,Load,TopRPE,Pain,Quality,BestHold,CleanReps,Notes,Key=N)
LOG_DATA = [
 (TODAY,"T1",1,"90",N,"NA",5,"1",0,8.5,1,N,N,5,"Demo-Zeile: 90 Grad schwer, je 1 saubere Rep, Handgelenk gut",N),
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
 (D_T1,"T1",1,"OAHS_line",N,"L",2,N,0,8,1,5,N,N,"OAHS Line, sehr stabile Holds links, solide",N),
 (D_T1,"T1",1,"OAHS_line",N,"R",2,N,0,8,1,4,N,N,"OAHS Line rechts, stabil",N),
 (D_T1,"T1",1,"OAHS_shapes",N,"NA",3,N,0,8,1,2,N,N,"Diamond Shape, eher schwach",N),
 (D_T1,"T1",1,"90",N,"NA",4,"1",0,9,1,3,N,2,"4x1 Cluster, lange Pausen; keine Kraft Bottom Position; Rep2+3 sauber (RPE7-8), Rep1+4 (8-9) unsauber",N),
 (D_T1,"T1",1,"planche_hold",N,"NA",4,N,0,8,1,2,8,N,"1x Advanced Tuck versucht (nicht moeglich) + 3x8s Tuck Planche; schwaecher als sonst",N),
 (D_T1,"T1",1,"WPU",N,"NA",4,"1x3@40, 3x4@35",40,9,1,N,N,N,"Satz1: 40 kg x3; Satz2-4: 35 kg x4; RPE 8-9",N),
]
