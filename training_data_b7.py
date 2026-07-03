# -*- coding: utf-8 -*-
"""B7-spezifische Daten (Advanced OAHS). Baut auf training_data auf."""
import training_data as B

# --- Library: Basis + advanced OAHS-Shapes/Transitions/Figa/Flag (fuers Shape-Logging) ---
SHAPE_ROWS = [
 ("oahs_straddle","straddle,straddle oahs,graetsche,straddle one arm","OAHS Straddle","SkillOAHS","Schulter/Core/Balance","N","J","J"),
 ("oahs_one_leg_bent","one leg bent,olb,ein bein gebeugt,oahs olb","OAHS One-Leg-Bent","SkillOAHS","Schulter/Core/Balance","N","J","J"),
 ("oahs_diamond","diamond,diamond oahs,raute","OAHS Diamond","SkillOAHS","Schulter/Core/Balance","N","J","J"),
 ("oahs_half_straddle","half straddle,halbe graetsche,half straddle oahs","OAHS Half-Straddle","SkillOAHS","Schulter/Core/Balance","N","J","J"),
 ("oahs_legs_together","legs together,beine zusammen,lt,legs together oahs","OAHS Legs-Together","SkillOAHS","Schulter/Core/Balance","N","J","J"),
 ("oahs_tuck","tuck oahs,gehockt,tuck shape","OAHS Tuck","SkillOAHS","Schulter/Core/Balance","N","J","J"),
 ("oahs_twist","twist,twist oahs,gedreht,twisting one arm","OAHS Twist","SkillOAHS","Schulter/Core/Balance","N","J","J"),
 ("oahs_transition","transition,transitions,shape wechsel,shape change,uebergang","OAHS Shape-Transition","SkillOAHS","Schulter/Core/Balance","N","J","J"),
 ("oahs_figa","figa,t shape,arm anlegen,arm along body,stacked oahs","OAHS Figa (T-Shape / Arm-anlegen)","SkillFlag","Schulter/Side-Body","N","J","J"),
 ("oahs_flag","one arm flag,oahs flag,handstand flag,flag oahs","OAHS Flag","SkillFlag","Schulter/Side-Body","N","J","J"),
]
LIBRARY = B.LIBRARY + SHAPE_ROWS
LIB_HR = 3
LIB_FIRST = 4
LIB_LAST = LIB_HR + len(LIBRARY)

SITU = [
 "OAHS-Level (korrigiert): OAHS wird BEHERRSCHT. Links frisch/aufgewaermt zuverlaessig Straddle, One-Leg-Bent, Diamond 5s+. Rechts manchmal aehnlich, aber inkonsistenter.",
 "Ziel: alle Shapes + Transitions (auch schwere) + Arm-anlegen/Figa (T-Shape) + OAHS Flag (Figa-Familie).",
 "B6-Learning 1: chronisch ~10-15% unter-erholt (Maintenance nach Diaet + aussergewoehnliche Hitze -> flacherer Schlaf).",
 "B6-Learning 2: RPE-Creep auf Akzessorik (oft 8-9 statt 7) + Salvage-Negative bei RPE 9 -> Skill-Qualitaet litt zuerst.",
 "B6-Learning 3: 90-Bottom-Ansteuerung kollabiert unter Fatigue (nicht Kraft, sondern muedes CNS). Frisch: 3 saubere Reps @RPE7 moeglich.",
 "B6-Learning 4: OAHS hatte Frequenz aber KEIN Progressionssystem -> Plateau. Fix: messbare Metrik + Aufstiegs-Kriterien (siehe Skill_Progression).",
 "Gewicht bei ~81 kg gehalten (Diaet-Ausfuehrung sauber). Handgelenk: von Pisiform in linke Handmuskulatur gewandert, besser aber nicht weg -> Warm-up-Protokoll + ggf. Physio.",
]
GOALS = [
 "1) OAHS Basic-Shapes beidseitig konsolidieren (8-10s) + RECHTS an links angleichen (Konsistenz/Trefferquote).",
 "2) Shapes erweitern (Half-Straddle -> Legs-Together -> Tuck) + Transitions SYSTEMATISCH (bisher vernachlaessigt).",
 "3) Advanced/Ziele: Twist, Arm-anlegen/Figa (T-Shape), OAHS Flag.",
 "4) 90 Push-up -> 3 saubere Reps (nur FRISCH maxen), Cluster.",
 "5) Kraft/Hypertrophie nur als Stuetze - RPE-diszipliniert (Cap 7 auf Akzessorik).",
]
# Volumen-Leitplanken B7 (~15-20% unter B6, RPE-Cap 7)
VOL = [
 ("W1 Re-Entry","10-12","12-14","8-10","3 (1 Skill-Fokus / 2 leicht)","3 (Progression!)","2"),
 ("W2 Build","12-14","14-15","10-12","3","3","2-3"),
 ("W3 Peak (nur wenn Recovery haelt)","13-14","14-16","10-12","3","3","3"),
 ("W4 Consolidate/Test","8-10","9-11","6-8","2 (Quality-Test)","beste Shapes/Holds","1"),
]

# Skill-Progressions-Leiter (Advanced): (Tier, Inhalt/Shapes, Metrik, Aktueller Stand, Aufstieg wenn)
SKILL_LADDER = [
 ("A — Basis besitzen",
  "Straddle · One-Leg-Bent · Diamond",
  "Hold-Sekunden + Trefferquote (saubere/gesamte Versuche) pro Seite",
  "LINKS: 5s+ zuverlaessig (besitzt). RECHTS: inkonsistent",
  "beide Seiten 8-10s clean + Trefferquote >=80% ueber 3 Sessions"),
 ("B — Shapes erweitern",
  "Half-Straddle -> Legs-Together -> Tuck",
  "Hold-Sekunden + Trefferquote pro Seite, pro Shape",
  "offen",
  "je Shape 8-10s clean beidseitig, dann naechste Stufe"),
 ("C — Transitions (eigener Strang!)",
  "Shape<->Shape (z.B. Straddle<->Diamond, Straddle->Legs-Together), saubere Entries/Exits",
  "saubere Wechsel pro Seite (Anzahl) + Kontrolle",
  "offen — bisher vernachlaessigt (laut Recherche der Schluessel zu Progress)",
  "1 sauberer Wechsel zuverlaessig -> Ketten aus 2-3 Shapes"),
 ("D — Advanced / Ziele",
  "Twist · Arm-anlegen / Figa (T-Shape) · OAHS Flag (Figa-Familie)",
  "Entries (Anzahl) -> Hold-Sekunden + Kontrolle pro Seite",
  "Fernziel",
  "erste saubere Entries -> Haltezeit aufbauen (wie Tier A)"),
]
SKILL_RULES = [
 "PRINZIP: pro Skill EINE messbare Metrik. Auf der aktuellen Stufe bleiben bis Kriterium erfuellt, DANN Stufe +1. 'Time under balance' ist die Waehrung.",
 "RECHTS-Regel: schwaechere/inkonsistente Seite bekommt +1 Satz bzw. mehr Versuche. Stufe erst hoch, wenn RECHTS ein (leicht niedrigeres) Kriterium erfuellt -> Asymmetrie nicht vergroessern.",
 "FRISCH ueben: Skill-Fokus zu Beginn (T1/T2/T4), NICHT am muedem Ende. 20-30 Qualitaetsversuche pro Fokus-Slot, Metrik loggen.",
 "Transitions NICHT vernachlaessigen: 'Balance ist so gut wie Ein-/Ausgang'. Ab Tier B parallel Transitions ueben.",
 "Loggen: pro Shape eine Zeile mit ExerciseID (oahs_straddle, oahs_diamond, ...), Side, BestHold_s, Quality_0_5. So wird Progress im Weekly_Summary/Dashboard sichtbar.",
 "Recovery-Gate: Recovery <=2/5 -> Skill nur Technik/leicht (kein Stufen-Test), Akzessorik -1 Satz & RPE 6.",
]
# Warm-up / Handgelenk-Protokoll (recherchiert)
WARMUP = [
 ("Grundregel","<=5-8 min gesamt","Laenger erzeugt Vor-Ermuedung. Nie bis zur Ermuedung direkt vor dem Skill."),
 ("1) Durchblutung/Mobility","2-3 min","Handgelenks-Kreise 2x10 · Unterarm-Selbstmassage 30-60s · Ausschuetteln · leichtes Handgelenks-Rocking"),
 ("2) Graduelles Laden (alle Winkel)","2-3 min","Palm-Pushups 2x10 · Finger-Pushups 1-2x8 · Rocking auf Haenden Extension 2x10 · Radial/Ulnar-Rocks 2x8 je Richtung · Handruecken-Load leicht"),
 ("3) DEIN Thema (ulnar/Handmuskulatur links)","separat","Leichte Isometrien 3x15-20s + langsame exzentrische Wrist-Curls (leicht) 2x10 als KONDITIONIERUNG (nicht bis Ermuedung). Optional an Nicht-Skill-Tagen als Reha."),
 ("Handposition","laufend","Finger spreizen + Boden krallen · leichte Aussenrotation zum Entlasten der gereizten Seite · Parallettes wenn Boden provoziert"),
 ("Schmerz-Regel","immer","scharf/stechend/neurologisch = STOPP. Dumpf/Zug beim Aufwaermen = ok, langsam reingehen. Bleibt es ueber Wochen -> Physio abklaeren."),
]
WARMUP_REFS = [
 ("Dani Winks — Wrist Warm-Up for Handstands","https://www.daniwinksflexibility.com/bendy-blog/wrist-warm-up-for-handstands"),
 ("Berg Movement — Preventing Wrist Injury (9 steps)","https://www.bergmovement.com/calisthenics-blog/preventing-wrist-injury"),
 ("Handstand Factory — Grip/Forearm & Wrist Conditioning","https://handstandfactory.com/grip/"),
 ("Handstand Factory — One Arm Shapes","https://handstandfactory.com/one-arm-shapes/"),
 ("Handstand Factory — Handstand Flag (Podcast S1E33)","https://handstandfactory.com/podcast/s1ep33-the-handstand-flag/"),
]

# B7-Plan (RPE-Cap 7 auf Akzessorik, Volumen getrimmt, Skill als Progressionssystem)
PLAN = [
 ("ALLE","Recovery-Check","Recovery 1-5 loggen (Schlaf/Spannung/Motivation)","1","1","1","1","-","-","<=2/5 -> Skill nur Technik; Akzessorik -1 Satz & RPE 6","-","steuert den Tag"),
 ("ALLE","Warm-up","Handgelenk-Protokoll (siehe B7_Warmup)","<=8 min","<=8 min","<=8 min","<=8 min","kein Schmerz","-","ulnar/Hand-Fokus links; Parallettes","neutral","nie muede vor Skill"),
 ("ALLE","Skill-Mikrodosis","HS Line/Balance (2-arm, Parallettes)","5-8 min","5-8 min","5-8 min","kurz","Qual 4-5","-","Frequenz","ribs down","taeglich"),
 ("ALLE","Skill-Mikrodosis","OAHS aktuelle Stufe LEICHT (Line/Shape)","kurz","kurz","kurz","kurz","Qual 4-5","-","niedrige Intensitaet, reine Frequenz","neutral","rechts leicht bevorzugen"),
 ("T1","Skill-Fokus OAHS","Basic-Shapes konsolidieren (Straddle/OLB/Diamond) - 20-30 Qualitaetsversuche","3-4/Seite","4/Seite","4/Seite","2/Seite (Test)","Qual 4-5","60-90s","Metrik loggen (Hold-Sek/Trefferquote); Aufstieg per Skill_Progression","neutral","RECHTS +1 Satz; nur frisch"),
 ("T1","T1 Haupt","90 Heavy - Cluster Singles->Double (FRISCH)","4x1","5x1","2x2 / Rest 1","Quality-Test","RPE 8-8,5","2-3 min","nur bei Recovery >=3/5 schwer; sonst Technik/Negativ","kein Arch, ribs down","Weg zu 3 Reps"),
 ("T1","Planche","Pseudo-Planche Push-up ODER Tuck/Straddle Planche Hold","3x6-8","3","3","1","RPE <=7","2 min","Straight-Arm-Fokus","neutral","Schluessel-Assistenz 90"),
 ("T1","Vertikal-Pull","Weighted Pull-up (Straps ok)","3x4","4x3-4","4x3","2x3","RPE <=7","2-3 min","Last nur bei sauberen Reps +","neutral","Unterarm schonen -> Straps"),
 ("T1","Schulter-Hygiene","Scap/Serratus + Aussenrotation","2x12","2","2","1","RPE 6","45-60s","Qualitaet","neutral","Schultergesundheit"),
 ("T2","Skill-Fokus OAHS","Shapes erweitern (Half-Straddle/Legs-Together/Tuck) ODER Transitions","3 Bloecke","3","3","beste","Qual 4-5","60-90s","naechste Stufe nur wenn aktuelle beidseitig 8-10s; Transitions ab Tier B","neutral","Transitions bewusst ueben"),
 ("T2","Haupt-Zug","Chest-supported / Seal Row","3x10-12","3x10-12","3x8-10","2x10","RPE <=7","90s","Reps sauber","brustgestuetzt","oberer Ruecken"),
 ("T2","Vertikal-Zug","Lat Pulldown / Pull-up","3x8-10","3x8-10","3x6-8","2x8","RPE <=7","90s","kontrolliert","neutral","Lat"),
 ("T2","Schulter","Rear-Delt Fly + Face Pull","2x12-15","2-3","2","2","RPE 6-7","45-60s","Qualitaet","neutral","hintere Schulter"),
 ("T2","Core","Hollow / Compression","3x20-30s","3","3","2","Qual 4","45s","fuer Shapes","kein LWS-Reiz","Compression"),
 ("T3","Recovery/Gate","Hinge Dowel + Status","2x5","2x5","2x5","2x5","Pain-Gate","-","Gate entscheidet Tag; kein schwerer Hinge","Pain>3 regredieren","LWS/BWS"),
 ("T3","Bein uni","Bulgarian / FFE Split Squat","3x8/Seite","3x8","3x8","2x10","RPE <=7","90s","autoreguliert (Muskelkater!)","neutral","Quad/Glute"),
 ("T3","Bein bi","Leg Press / Hack","3x10-12","3x10","3x10","2x12","RPE <=7","2 min","Last + wenn erholt","kein Grind","Quad/Glute"),
 ("T3","Post. Kette","Hamstring Curl","3x10-12","3","3","2","RPE <=7","75s","kein schwerer Hinge","Gate","Hamstrings"),
 ("T3","Bein","Calves","3x12-15","3","3","2","RPE 7","45s","-","neutral","Waden"),
 ("T3","BWS-Mobility","BWS Ext/Rotation, ribs down, Atmung","1 Block","1","1","1","Qual 4-5","-","fix; Protokoll in Rules","Einklemm monitoren","-"),
 ("T3","Core","Pallof + Dead Bug + Side Plank","2 je","2-3","2-3","2","Qual 4","45-60s","Anti-Ext/Rot/Lat","Anti-Flexion","LWS-schonend"),
 ("T4","Warm-up+","Laenger + Side-Body-Aktivierung","1 Block","1","1","1","Qual 4-5","-","vor Flag/Figa noetig","neutral","Verletzungsschutz"),
 ("T4","Skill Flag/Figa","Unsupported Flag-Lean + 1-Arm-Release + Figa-Entries","4/Seite","4-5","4-5","beste","Qual 4-5","90s","Kontrolle VOR Haltezeit; Metrik loggen (Entries/Hold)","kontrolliert","Fernziel-Pfad"),
 ("T4","Skill Flag","Straddle Flag Hold (1-3s freeze) + Sideline","3/Seite","3-4","4","beste","Qual 4-5","90s","Support nur senken wenn Re-Entry sauber","kontrollierter Return","seltene freie Versuche"),
 ("T4","90 Light","90 Technik/Negativ+Pause","3x2-3","4x2-3","4x2","2x2","RPE 6-7","2 min","sauber, kein Failure","ribs down","2. 90-Dosis"),
 ("T4","Vertikal-Press","Wall-HSPU / Pike / DB","3x6-8","3x6-8","3x6","2x6","RPE <=7","90s","ribs down","kein Arch","Vertikalkraft"),
 ("T4","Zug-Balance","One-Arm / Chest-supported Row","2x10/Seite","2-3","2-3","2","RPE 7","75s","rechts +","gestuetzt","Balance"),
 ("T5","90 Mikro","90 Technik/Negativ-Mikro","3x2","3x2-3","3x2","2x2","RPE 6-7","2 min","3. Dosis, sauber","neutral","Technik"),
 ("T5","Push-Pump","Ring-Dips + Deficit Push-up","3x10-12 + 2x12-15","gleich","3x10 + 2x12","2x12","RPE <=7","75-90s","getrimmt","neutral","Push"),
 ("T5","Pull-Pump","Row + Pulldown/Chin","3x10 + 3x10","gleich","3x10","2x12","RPE <=7","90s","komplettiert Zug","neutral","Pull"),
 ("T5","Schultern","Lateral Raise + Rear-Delt","2-3x12-15","3","3","2","RPE 6-7","45s","getrimmt","neutral","Schultern"),
 ("T5","Arme opt.","Biceps / Triceps","1-2x10-12","2","1-2","1","RPE 7","45s","optional, Unterarm schonen","neutral","Arme"),
 ("T5","Regen","BWS + Lat/Pec/Unterarm + Atmung","1 Block","1","1","1","Qual 4-5","-","Downshift + Handgelenk-Reha","BWS-Pflege","Recovery"),
]
