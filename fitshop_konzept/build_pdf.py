#!/usr/bin/env python3
"""
Rendert das FitShop Sales-Enablement-Konzept als professionelles PDF.

Bearbeiten:
  - Inhalt/Texte: direkt in der Variable HTML weiter unten.
  - Beispiel-ROI-Zahlen: im Block ROI (gleich hier oben) – diese fliessen in die
    Beispielrechnung ein und sind bewusst leicht austauschbar.
  - Quell-Narrativ in Markdown: FitShop_SalesEnablement_Konzept.md

Bauen:
  python3 build_pdf.py
"""

from pathlib import Path
from weasyprint import HTML

BASE = Path(__file__).resolve().parent
FONTS = BASE / "fonts"
OUT = Path("/mnt/user-data/outputs/FitShop_SalesEnablement_Konzept.pdf")

# --- Editierbare Stammdaten -------------------------------------------------
META = {
    "name": "Daniel Bellwinkel",
    "datum": "15. Juni 2026",
    "rolle": "Sales Enablement &amp; Vertriebsentwicklung – Pilot Österreich",
    "pfad": "Sales Enablement Manager Österreich",
}

# Beispiel-ROI – ausschliesslich illustrativ. Mit echten Filialzahlen ersetzen.
ROI = {
    "beratungen": "200",      # X: Beratungen / Monat
    "lift": "3",              # Y: Conversion-Lift in Prozentpunkten
    "bon": "1.500",           # Z: Ø-Bonwert in €
    "ergebnis": "9.000",      # N: gerundetes Beispielergebnis in € / Monat
}
# ---------------------------------------------------------------------------

CSS = """
@font-face { font-family:'Inter'; src:url('FONT/Inter-Regular.ttf'); font-weight:400; }
@font-face { font-family:'Inter'; src:url('FONT/Inter-Medium.ttf'); font-weight:500; }
@font-face { font-family:'Inter'; src:url('FONT/Inter-SemiBold.ttf'); font-weight:600; }
@font-face { font-family:'Inter'; src:url('FONT/Inter-Bold.ttf'); font-weight:700; }
@font-face { font-family:'Fraunces'; src:url('FONT/Fraunces-Medium.ttf'); font-weight:500; }
@font-face { font-family:'Fraunces'; src:url('FONT/Fraunces-SemiBold.ttf'); font-weight:600; }
@font-face { font-family:'Fraunces'; src:url('FONT/Fraunces-Bold.ttf'); font-weight:700; }

:root{
  --ink:#15191B; --ink2:#3A4246; --muted:#6B7479;
  --paper:#FFFFFF; --panel:#F4F6F5; --line:#DDE2E0;
  --accent:#0E5A55; --accent-deep:#0A413D; --wash:#E6F0EE; --wash-line:#CFE0DC;
  --flag:#9A6B12; --flag-wash:#FAF2E0; --flag-line:#ECDBB5;
}

@page{
  size:A4; margin:15mm 16mm 14mm 16mm;
  @bottom-left{
    content:"Daniel Bellwinkel · 15. Juni 2026";
    font-family:'Inter'; font-size:7pt; color:#8A9290;
    border-top:0.5pt solid #E2E6E4; padding-top:3pt; vertical-align:top;
  }
  @bottom-center{
    content:"Interner Konzeptvorschlag — Vertraulich";
    font-family:'Inter'; font-size:7pt; color:#8A9290; letter-spacing:0.04em;
    border-top:0.5pt solid #E2E6E4; padding-top:3pt; vertical-align:top;
  }
  @bottom-right{
    content:counter(page) " / " counter(pages);
    font-family:'Inter'; font-size:7pt; color:#8A9290;
    border-top:0.5pt solid #E2E6E4; padding-top:3pt; vertical-align:top;
  }
}

*{ box-sizing:border-box; }
html{ font-family:'Inter'; color:var(--ink); font-size:9.5pt; line-height:1.46; }
body{ margin:0; }
p{ margin:0 0 6pt; color:var(--ink2); }
strong{ color:var(--ink); font-weight:600; }
em{ font-style:italic; }

/* ---- Masthead ---- */
.masthead{ margin-bottom:11pt; }
.topbar{ display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:9pt; }
.kicker{ font-size:7.4pt; font-weight:600; letter-spacing:0.18em; text-transform:uppercase;
         color:var(--accent); }
.kicker .path{ color:var(--muted); font-weight:500; }
.conf{ font-size:6.8pt; font-weight:600; letter-spacing:0.13em; text-transform:uppercase;
       color:var(--flag); border:0.8pt solid var(--flag-line); background:var(--flag-wash);
       padding:3pt 7pt; border-radius:2pt; white-space:nowrap; }
h1{ font-family:'Fraunces'; font-weight:600; font-size:25pt; line-height:1.04;
    letter-spacing:-0.01em; margin:0 0 3pt; color:var(--ink); }
h1 .accent{ color:var(--accent); }
.meta{ font-size:7.8pt; color:var(--muted); margin-top:7pt; }
.meta b{ color:var(--ink2); font-weight:600; }
.meta .sep{ color:var(--line); margin:0 7pt; }

/* signature: measurement tick-rule */
.tickrule{ height:6px; margin:11pt 0 0; border-top:1.4pt solid var(--accent);
  background-image:repeating-linear-gradient(90deg, var(--accent) 0 0.8px, transparent 0.8px 9px);
  background-size:100% 4px; background-repeat:repeat-x; background-position:top left; }

/* ---- Sections ---- */
section{ margin-top:10pt; }
.eyebrow{ font-size:7.6pt; font-weight:600; letter-spacing:0.13em; text-transform:uppercase;
  color:var(--accent); margin-bottom:5pt; }
.eyebrow .num{ color:var(--ink); }
.eyebrow .num::after{ content:""; display:inline-block; width:14pt; height:0.8pt;
  background:var(--line); vertical-align:middle; margin:0 7pt; }
h2{ font-family:'Fraunces'; font-weight:600; font-size:13.5pt; line-height:1.15;
    margin:0 0 6pt; color:var(--ink); letter-spacing:-0.005em; }

ul{ margin:0 0 7pt; padding:0; list-style:none; }
li{ position:relative; padding-left:13pt; margin-bottom:3.5pt; color:var(--ink2); }
li::before{ content:""; position:absolute; left:0; top:6.5pt; width:4pt; height:1.6pt;
  background:var(--accent); }
li strong{ color:var(--ink); }

.lead{ font-family:'Fraunces'; font-weight:500; font-size:11pt; line-height:1.4;
  color:var(--ink); margin:0 0 8pt; }

.onesentence{ background:var(--panel); border-left:2.4pt solid var(--accent);
  padding:9pt 12pt; margin:0 0 8pt; font-size:9.8pt; }
.onesentence .tag{ display:block; font-size:7pt; font-weight:600; letter-spacing:0.12em;
  text-transform:uppercase; color:var(--accent); margin-bottom:3pt; }

/* ---- ROI example box ---- */
.example{ background:var(--wash); border:0.8pt solid var(--wash-line);
  border-left:3pt solid var(--accent); border-radius:3pt; padding:11pt 13pt 10pt;
  margin:9pt 0 6pt; }
.example .flag{ display:inline-block; font-size:6.8pt; font-weight:700; letter-spacing:0.12em;
  text-transform:uppercase; color:var(--flag); background:var(--flag-wash);
  border:0.8pt solid var(--flag-line); padding:2.5pt 7pt; border-radius:2pt; margin-bottom:7pt; }
.formula{ font-size:9.6pt; color:var(--ink2); margin:0 0 4pt; }
.formula .var{ font-weight:600; color:var(--accent-deep); }
.formula .res{ font-family:'Fraunces'; font-weight:600; font-size:11.5pt; color:var(--accent-deep); }
.example .num-ex{ font-size:9pt; color:var(--ink2); margin:0 0 6pt; }
.example .note{ font-size:7.6pt; font-style:italic; color:var(--muted); margin:0;
  padding-top:6pt; border-top:0.6pt solid var(--wash-line); }

/* proof point */
.proof{ display:flex; gap:11pt; align-items:flex-start; background:var(--panel);
  border:0.8pt solid var(--line); border-radius:3pt; padding:10pt 12pt; margin-top:8pt; }
.proof .award{ font-family:'Fraunces'; font-weight:600; font-size:8pt; letter-spacing:0.08em;
  text-transform:uppercase; color:var(--accent); writing-mode:vertical-rl; transform:rotate(180deg);
  border-right:0.8pt solid var(--line); padding-right:9pt; }
.proof p{ margin:0; font-size:9pt; }
.proof .stats{ margin-top:5pt; font-size:8pt; color:var(--muted); }
.proof .stats b{ color:var(--accent-deep); font-family:'Fraunces'; font-weight:600; }

/* ---- KPI two columns ---- */
.cols{ display:flex; gap:11pt; margin-top:3pt; }
.card{ flex:1; border:0.8pt solid var(--line); border-radius:3pt; padding:10pt 12pt;
  background:var(--paper); }
.card.lead-card{ border-top:2.4pt solid var(--accent); }
.card.lag-card{ border-top:2.4pt solid var(--ink2); }
.card h3{ font-family:'Inter'; font-size:7.4pt; font-weight:700; letter-spacing:0.1em;
  text-transform:uppercase; margin:0 0 2pt; color:var(--ink); }
.card .sub{ font-size:7.2pt; color:var(--muted); margin:0 0 6pt; }
.card ul{ margin:0; }
.card li{ font-size:8.6pt; margin-bottom:3pt; }
.honest{ font-size:8.8pt; color:var(--ink2); margin-top:8pt;
  background:var(--panel); border-radius:3pt; padding:8pt 11pt; }

/* ---- Phases ---- */
.phase{ display:flex; gap:11pt; align-items:flex-start; padding:8pt 0; }
.phase + .gate{ margin:0 0 0 13pt; }
.gate{ font-size:7.2pt; font-weight:600; letter-spacing:0.1em; text-transform:uppercase;
  color:var(--accent); padding:1pt 0 1pt 17pt; position:relative; }
.gate::before{ content:""; position:absolute; left:7pt; top:-3pt; bottom:-3pt; width:1pt;
  background:repeating-linear-gradient(var(--accent) 0 2pt, transparent 2pt 4pt); }
.pnum{ flex:0 0 auto; width:26pt; height:26pt; border-radius:50%; background:var(--accent);
  color:#fff; font-family:'Fraunces'; font-weight:600; font-size:11pt; text-align:center;
  line-height:26pt; }
.phase .body{ flex:1; }
.phase h3{ font-family:'Inter'; font-size:9.4pt; font-weight:600; margin:1pt 0 2pt; color:var(--ink); }
.phase h3 span{ color:var(--muted); font-weight:500; font-size:8pt; }
.phase p{ margin:0; font-size:8.8pt; }

/* ---- generic 2-col list ---- */
.split{ display:flex; gap:18pt; }
.split ul{ flex:1; }

.ask{ background:var(--accent-deep); color:#EAF2F0; border-radius:4pt; padding:12pt 14pt;
  margin-top:10pt; }
.ask .eyebrow{ color:#7FC4BC; }
.ask h2{ color:#fff; margin-bottom:4pt; }
.ask p{ color:#D2E2DF; margin:0; font-size:9.4pt; }
.ask b{ color:#fff; }

.nb{ break-inside:avoid; }
section{ break-inside:auto; }
h2,h3{ break-after:avoid; }
"""

HTML_BODY = f"""
<div class="masthead">
  <div class="topbar">
    <div class="kicker">Sales Enablement &amp; Vertriebsentwicklung
      <span class="path">· Pilot Österreich</span></div>
    <div class="conf">Vertraulich · Intern</div>
  </div>
  <h1>Verkauf vom Bauchgefühl<br>zur <span class="accent">messbaren Disziplin</span>.</h1>
  <div class="meta">
    <b>Interner Konzeptvorschlag</b><span class="sep">|</span>
    Vorgelegt von <b>{META['name']}</b><span class="sep">|</span>
    {META['datum']}<span class="sep">|</span>
    Entwicklungspfad: <b>{META['pfad']}</b>
  </div>
  <div class="tickrule"></div>
</div>

<section class="nb">
  <div class="eyebrow"><span class="num">01</span>Ausgangslage / Management Summary</div>
  <p>FitShop verkauft hochpreisige, beratungsintensive Produkte – ein Geschäft, in dem die
  Qualität des einzelnen Verkaufsgesprächs unmittelbar über den Umsatz entscheidet. Heute läuft
  dieser Verkauf weitgehend über Erfahrung, persönliche Verkaufsphilosophie und Tagesform.
  Umsätze sind sichtbar, doch <em>wie</em> sie entstehen – oder warum sie ausbleiben – wird auf
  Gesprächsebene nicht systematisch erfasst.</p>
  <p>In der Filiale Innsbruck wurde bereits begonnen, ein Sales-Tracking-System aufzubauen, das
  genau diese Lücke schließt; der Ansatz wird vom Filialleiter positiv aufgenommen. Dieser
  Vorschlag formalisiert die Initiative zu einer definierten Vollzeitrolle mit eigenem
  Verantwortungsbereich und überführt sie in einen messbaren 90-Tage-Pilot. Ziel ist, Verkauf von
  erfahrungs- und bauchgefühlbasiert zu daten- und prozessbasiert weiterzuentwickeln – mit einem
  Umsatzhebel, der ein Vielfaches der Rollenkosten beträgt und linear über Filialen skaliert.</p>
</section>

<section class="nb">
  <div class="eyebrow"><span class="num">02</span>Problem &amp; Chance</div>
  <h2>Was funktioniert, bleibt im Kopf einzelner Verkäufer.</h2>
  <p>Der Vertrieb im stationären Fitnessfachhandel ist stark von individuellen Stärken geprägt.
  Erfolgreiche Muster bleiben implizit; was nicht funktioniert, wird selten strukturiert sichtbar.
  Konkret wird heute kaum systematisch erfasst:</p>
  <ul>
    <li>Welche Kundengespräche finden statt – und mit welchen Kundentypen?</li>
    <li>Welche Kaufabsichten liegen vor, und wo im Gespräch scheitert der Abschluss?</li>
    <li>Welche Einwände kommen regelmäßig, und welche Argumente wirken tatsächlich?</li>
    <li>Welche Produkte oder Kategorien erzeugen wiederkehrende Probleme?</li>
    <li>Welche Kunden kaufen sofort, welche brauchen ein Follow-up?</li>
    <li>Welche Beratungsfehler und Prozesslücken wiederholen sich?</li>
  </ul>
  <p><strong>Die Chance:</strong> Auf hohen Durchschnittsbons erzeugen schon wenige Prozentpunkte
  mehr Abschlussquote, etwas mehr Upsell und ein systematisches Follow-up spürbar Umsatz pro
  Filiale. Wer diese Muster sichtbar macht und in übertragbare Verkaufsstandards übersetzt, hebt
  nicht einen einzelnen Verkäufer, sondern das Niveau der gesamten Fläche – wiederholbar in jeder
  weiteren Filiale.</p>
</section>

<section class="nb">
  <div class="eyebrow"><span class="num">03</span>Die Lösung</div>
  <div class="onesentence">
    <span class="tag">Die Rolle in einem Satz</span>
    Verkauf im stationären Fitnessfachhandel von erfahrungs- und bauchgefühlbasiert zu daten- und
    prozessbasiert weiterentwickeln – über systematische Gesprächs- und Abschlussanalyse, daraus
    abgeleitetes Training und übertragbare Verkaufsstandards.
  </div>
  <p>Konkret heißt das: Gesprächsdaten strukturiert erfassen, wiederkehrende Muster auswerten,
  daraus konkrete Sales Plays und eine Einwandbibliothek bauen, das Wissen in kurzen, praxisnahen
  Trainingsformaten an das Team zurückspielen und eine verlässliche Follow-up-Struktur etablieren.
  Es geht ausdrücklich <strong>nicht um „ein weiteres Excel-Sheet“</strong>, sondern darum, Verkauf
  systematisch zu entwickeln statt rein reaktiv geschehen zu lassen. Es ist eine definierte
  Vollzeit-Spezialistenrolle mit eigenem Verantwortungsbereich und Review-Gates –
  <strong>keine Zusatzaufgabe zur bestehenden Teilzeit-Verkäuferrolle.</strong></p>
</section>

<section class="nb">
  <div class="eyebrow"><span class="num">04</span>Strategischer Nutzen für FitShop</div>
  <h2>Kleiner Hebel an der Mechanik, große Wirkung auf der Fläche.</h2>
  <p>Der Hebel liegt in der Struktur des Geschäfts: hohe Durchschnittsbons, beratungsintensive
  Entscheidungen, vergleichbare Abläufe über viele Filialen. Eine kleine Verbesserung der
  Abschluss- und Upsell-Mechanik wirkt deshalb absolut spürbar – und ist, einmal als Standard
  dokumentiert, auf weitere Filialen übertragbar. Die Rolle kostet eine Vollzeitkraft; der bewegte
  Umsatzhebel ist ein Vielfaches davon. Exakte Zahlen liefert der Pilot.</p>

  <div class="example nb">
    <span class="flag">Beispielrechnung · illustrativ · mit echten Filialzahlen ersetzen</span>
    <p class="formula">
      <span class="var">X</span> Beratungen/Monat ×
      <span class="var">Y</span> Prozentpunkte Conversion-Lift ×
      <span class="var">Z €</span> Ø-Bonwert =
      <span class="res">+N €/Monat</span>
    </p>
    <p class="num-ex">Beispielhaft eingesetzt: {ROI['beratungen']} Beratungen/Monat ×
      {ROI['lift']} Prozentpunkte × {ROI['bon']} € ≈
      <strong>+{ROI['ergebnis']} €/Monat</strong> je Filiale.</p>
    <p class="note">Diese Zahlen dienen ausschließlich der Veranschaulichung des Hebels. Die
      belastbare Baseline und die tatsächlichen Werte werden im Pilot erhoben.</p>
  </div>

  <div class="proof nb">
    <div class="award">Rising&nbsp;Star</div>
    <div>
      <p><strong>Beweispunkt – Rising-Star-Award.</strong> Als Teil eines 4-köpfigen Teams das
      größte Umsatzwachstum über 60+ europäische FitShop-Filialen. Die Fähigkeit, Verkaufsleistung
      gezielt zu steigern, ist damit nicht Hypothese, sondern bereits belegt.</p>
      <p class="stats">4er-Team &nbsp;·&nbsp; größtes Wachstum von <b>60+</b> Filialen
        &nbsp;·&nbsp; Teamumsatz <b>2 Mio.+&nbsp;€</b></p>
    </div>
  </div>
</section>

<section class="nb">
  <div class="eyebrow"><span class="num">05</span>Aufgaben &amp; Verantwortlichkeiten</div>
  <ul>
    <li><strong>Tracking-System aufbauen und pflegen</strong> – Gespräch → Kaufabsicht → Einwand
      → Abschluss / kein Abschluss → Grund → Follow-up.</li>
    <li><strong>Muster auswerten</strong> – wiederkehrende Einwände, Abschlussbarrieren, wirksame
      Argumente und Produktthemen sichtbar machen.</li>
    <li><strong>In Sales Plays übersetzen</strong> – konkrete Gesprächsbausteine und eine
      Einwandbibliothek aus echten Gesprächen ableiten.</li>
    <li><strong>Training liefern</strong> – kurze, praxisnahe Formate aus realen Fällen, kein
      theoretisches Schulungsprogramm.</li>
    <li><strong>Follow-up-Struktur etablieren</strong> – definierter Prozess für Kunden, die nicht
      sofort kaufen.</li>
    <li><strong>Einarbeitung beschleunigen</strong> – Playbook, das neue Verkäufer schneller auf
      Niveau bringt.</li>
    <li><strong>Reporting</strong> – an Filial- und (ab Phase 2) Regionalleitung.</li>
    <li><strong>Übertragbar dokumentieren</strong> – den Piloten so festhalten, dass er in weiteren
      Filialen wiederholbar ist.</li>
  </ul>
</section>

<section class="nb">
  <div class="eyebrow"><span class="num">06</span>KPIs &amp; Erfolgsmessung</div>
  <div class="cols">
    <div class="card lead-card">
      <h3>Leading</h3>
      <p class="sub">wöchentlich · direkt steuerbar</p>
      <ul>
        <li>Erfassungsquote der Gespräche</li>
        <li>Conversion (Beratung → Abschluss)</li>
        <li>Ø-Bonwert / Upsell-Quote</li>
        <li>Follow-up-Quote &amp; -Conversion</li>
        <li>Einarbeitungszeit neuer Verkäufer</li>
        <li>Anzahl wiederverwendbarer Sales Plays</li>
      </ul>
    </div>
    <div class="card lag-card">
      <h3>Lagging</h3>
      <p class="sub">monatlich · Wirkung</p>
      <ul>
        <li>Umsatz Pilotfiliale vs. Vorperiode</li>
        <li>Umsatz vs. Vergleichsfilialen</li>
        <li>Umsatz pro Verkäufer</li>
        <li>Like-for-like-Vergleich</li>
      </ul>
    </div>
  </div>
  <p class="honest"><strong>Ehrlich eingeordnet:</strong> Der Pilot etabliert zunächst eine
  belastbare Baseline und beweist die Bewegung der Leading-Indikatoren; die Umsatzwirkung wird über
  das Pilotfenster validiert. Die Erfolgskriterien werden vorab gemeinsam definiert, damit am
  Review-Gate objektiv entschieden werden kann.</p>
</section>

<section class="nb">
  <div class="eyebrow"><span class="num">07</span>Pilot- &amp; Einführungsplan</div>
  <div class="phase">
    <div class="pnum">1</div>
    <div class="body">
      <h3>Pilot <span>· ~90 Tage · Innsbruck</span></h3>
      <p>Tracking-System scharf stellen, Baseline erheben, erste Sales Plays und Follow-up-Struktur
      etablieren, Leading-Indikatoren bewegen.</p>
    </div>
  </div>
  <div class="gate">Review-Gate →</div>
  <div class="phase">
    <div class="pnum">2</div>
    <div class="body">
      <h3>Österreich-Rollout <span>· ~3–6 Monate</span></h3>
      <p>Ausweitung auf weitere Filialen mit je einem „Champion“ pro Standort, regionales Reporting,
      Standardisierung der Plays.</p>
    </div>
  </div>
  <div class="gate">Review-Gate →</div>
  <div class="phase">
    <div class="pnum">3</div>
    <div class="body">
      <h3>Etablierte Funktion <span>· Dauerrolle</span></h3>
      <p>Verstetigung der Rolle, ggf. Ausweitung auf DACH, Anbindung an Tooling/CRM. An jedem
      Übergang steht eine bewusste Go/No-go-Entscheidung – kein Automatismus.</p>
    </div>
  </div>
</section>

<section class="nb closing">
  <div class="eyebrow"><span class="num">08</span>Schnittstellen &amp; Change Management</div>
  <ul>
    <li><strong>Filialleiter</strong> – Sponsor der Rolle, wöchentlicher Sync, Mit-Ownership der
      Ziele.</li>
    <li><strong>Regionalleitung</strong> – ab Phase 2 für Budget und Mandat.</li>
    <li><strong>Verkaufsteam</strong> – klar als <em>Enablement</em> geframt: Die Rolle hilft
      Verkäufern, mehr abzuschließen und mehr zu verdienen, und erleichtert die Einarbeitung. Daten
      dienen dem Lernen und dem Team, nicht der Kontrolle Einzelner; die Reps gestalten die Plays
      aktiv mit.</li>
    <li><strong>Zentrale</strong> – optional ab Phase 3.</li>
  </ul>
</section>

<section class="nb">
  <div class="eyebrow"><span class="num">09</span>Rollen-, Entwicklungs- &amp; Vergütungslogik</div>
  <p>Der Weg führt von Teilzeit zu Vollzeit. Das Modell besteht aus einem angemessenen Fixum
  (über dem reinen Verkäufer-Fixum) plus einem variablen Anteil, der an die beeinflussbaren KPIs
  gekoppelt ist. Beides ist entlang der Phasen gestaffelt: An jedem Review-Gate werden Titel, Fixum
  und Verantwortung gemeinsam angepasst. Die konkrete Größenordnung ist <strong>[zu besprechen]</strong>
  und folgt der Verantwortung, nicht umgekehrt.</p>
</section>

<section class="nb">
  <div class="eyebrow"><span class="num">10</span>Bedingungen für beidseitigen Erfolg</div>
  <div class="split">
    <ul>
      <li>Klar definierte Rolle – kein „Verkauf plus Zusatzaufgabe“.</li>
      <li>Mandat und Rückendeckung für die Funktion.</li>
      <li>Zugang zu den nötigen Daten und Tools.</li>
    </ul>
    <ul>
      <li>Vorab definierte Erfolgskriterien und Review-Gates.</li>
      <li>Vergütung folgt Verantwortung.</li>
      <li>Echte Entwicklungsperspektive.</li>
    </ul>
  </div>
</section>

<div class="ask nb">
  <div class="eyebrow"><span class="num">11</span>Nächster Schritt</div>
  <h2>Genehmigung eines 90-Tage-Pilots.</h2>
  <p>Mit vorab definierten Erfolgskriterien und einem festen Review an <b>Tag 90</b>. An diesem
  Review wird über Phase 2 sowie über die Formalisierung von Rolle und Vergütung entschieden.</p>
</div>
"""

def main():
    css = CSS.replace("FONT", FONTS.as_posix())
    html = f"<!doctype html><html lang='de'><head><meta charset='utf-8'>" \
           f"<title>Sales Enablement &amp; Vertriebsentwicklung – Pilot Österreich</title>" \
           f"<meta name='author' content='Daniel Bellwinkel'>" \
           f"<meta name='description' content='Interner Konzeptvorschlag – Vertraulich'>" \
           f"<style>{css}</style></head><body>{HTML_BODY}</body></html>"
    OUT.parent.mkdir(parents=True, exist_ok=True)
    HTML(string=html, base_url=BASE.as_posix()).write_pdf(OUT.as_posix())
    print(f"PDF geschrieben: {OUT}")

if __name__ == "__main__":
    main()
