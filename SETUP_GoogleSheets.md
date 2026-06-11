# Voice-Tracking direkt in dein Google Sheet (lokal auf dem Mac)

Ziel: Du diktierst, und die Zeile landet **automatisch in deinem Google Sheet**.
Das geht nur lokal auf deinem Mac, weil dafür dein eigener Google-Zugang nötig
ist (die Web-Session von Claude kann nicht in dein Sheet schreiben).

## Überblick
```
Du sprichst  ->  Claude Code (lokal)  ->  append_log.py  ->  dein Google Sheet
```

## Einmaliges Setup

### 1. Werkzeuge installieren
- Node.js (für Claude Code) und Python 3 müssen vorhanden sein.
- Claude Code CLI:  `npm install -g @anthropic-ai/claude-code`
- Python-Pakete:    `pip install gspread google-auth`

### 2. Google-Zugang für das Skript (Service Account)
1. https://console.cloud.google.com  ->  neues Projekt anlegen.
2. „APIs & Dienste" -> „Bibliothek" -> **Google Sheets API** aktivieren.
3. „APIs & Dienste" -> „Anmeldedaten" -> **Dienstkonto (Service Account)** erstellen.
4. Beim Dienstkonto -> „Schlüssel" -> **JSON-Key** erstellen und herunterladen
   (z. B. nach `~/training-sa.json`).
5. Öffne dein Google Sheet -> **Freigeben** -> die E-Mail des Dienstkontos
   (steht in der JSON, Feld `client_email`) als **Bearbeiter** hinzufügen.

### 3. Umgebungsvariablen setzen (in ~/.zshrc)
```bash
export GOOGLE_APPLICATION_CREDENTIALS="$HOME/training-sa.json"
export TRAINING_SHEET_ID="<die ID aus der Sheet-URL>"
# Sheet-URL: https://docs.google.com/spreadsheets/d/<DIESE_ID>/edit
```
Danach `source ~/.zshrc`.

### 4. Repo lokal holen (enthält append_log.py + Library)
```bash
git clone <dieses Repo>
cd My-Test-repository
```

## Test (eine Zeile anhängen)
```bash
python append_log.py --json '{"date":"2026-06-11","daytype":"T2","week":1,
  "exercise_id":"chest_row","side":"NA","sets":3,"reps":"10","load":75,
  "rpe":8,"pain":1,"notes":"Seal Row"}'
```
Wenn im Tab `B6_Training_Log` eine neue Zeile auftaucht: läuft.

## Täglicher Ablauf mit Claude Code
1. Im Repo-Ordner `claude` starten.
2. Diktieren, z. B.: „T2, Woche 1, Seal Row 3 mal 10 mit 75 Kilo, RPE 8, Rücken 1."
3. Claude Code mappt den Begriff über `B6_Exercise_Library` auf die ExerciseID
   und ruft `append_log.py` mit den Feldern auf -> Zeile steht in deinem Sheet.

## Wichtig zu den Formeln im Sheet
Die aktuelle Datei ist **Excel-nativ** (Tabellen + strukturierte Referenzen wie
`tblLog[Sets]`). Google Sheets versteht das **nicht** — Dashboard/Weekly_Summary
funktionieren dort nicht. Wenn dein Zuhause Google Sheets sein soll, brauchst du
eine **Sheets-native Version** (Formeln über feste Bereiche statt Tabellen).
Sag Bescheid, dann baue ich die.

## Feld-Referenz für das JSON
`date` (YYYY-MM-DD) · `daytype` (T1–T5/Rest) · `week` (1–4) · `exercise_id`
(aus Library) · `side` (L/R/NA) · `sets` · `reps` · `load` · `rpe` · `pain` ·
`quality` · `besthold` · `cleanreps` · `notes`.
`Category` und `Key` werden automatisch gesetzt — nicht angeben.
