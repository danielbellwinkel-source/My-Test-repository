#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
append_log.py — haengt EINE Trainingszeile direkt an dein Google Sheet an.

Laeuft LOKAL auf deinem Mac (nicht in der Web-Session), weil dafuer dein
eigener Google-Zugang noetig ist. Gedacht zum Aufruf durch Claude Code:
Du diktierst -> Claude Code parst -> ruft dieses Skript mit den Feldern auf.

Setup siehe SETUP_GoogleSheets.md.

Aufruf (Beispiel):
  python append_log.py --json '{"date":"2026-06-11","daytype":"T2","week":1,
     "exercise_id":"chest_row","side":"NA","sets":3,"reps":"10","load":75,
     "rpe":8,"pain":1,"notes":"Seal Row"}'

Nicht gesetzte Felder bleiben leer. Category wird aus dem Tab
B6_Exercise_Library gezogen, Key = DayType_Wweek automatisch gebildet.
"""
import os, sys, json, argparse

import gspread                       # pip install gspread google-auth
from google.oauth2.service_account import Credentials

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
# Reihenfolge MUSS den Spalten in B6_Training_Log entsprechen:
COLUMNS = ["Date","DayType","Week","ExerciseID","Category","Side","Sets",
           "Reps_or_Scheme","Load_kg","TopRPE","Pain_0_10","Quality_0_5",
           "BestHold_s","CleanReps","Notes","Key"]

def get_client():
    cred_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    if not cred_path or not os.path.exists(cred_path):
        sys.exit("FEHLER: GOOGLE_APPLICATION_CREDENTIALS zeigt nicht auf die Service-Account-JSON.")
    creds = Credentials.from_service_account_file(cred_path, scopes=SCOPES)
    return gspread.authorize(creds)

def category_map(sheet):
    """ExerciseID -> Category aus dem Library-Tab (Single Source of Truth)."""
    ws = sheet.worksheet("B6_Exercise_Library")
    rows = ws.get_all_values()
    m = {}
    for r in rows:
        if len(r) >= 4 and r[0] and r[0] != "ExerciseID":
            m[r[0].strip()] = r[3].strip()
    return m

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--json", required=True, help="JSON mit den Feldern (siehe Kopf)")
    args = ap.parse_args()
    f = json.loads(args.json)

    sheet_id = os.environ.get("TRAINING_SHEET_ID")
    if not sheet_id:
        sys.exit("FEHLER: Umgebungsvariable TRAINING_SHEET_ID nicht gesetzt.")

    gc = get_client()
    sheet = gc.open_by_key(sheet_id)
    catmap = category_map(sheet)

    exid = str(f.get("exercise_id", "")).strip()
    category = catmap.get(exid, "")
    daytype = str(f.get("daytype", "")).strip()
    week = f.get("week", "")
    key = f"{daytype}_W{week}" if daytype and week != "" else ""

    row = {
        "Date": f.get("date", ""), "DayType": daytype, "Week": week,
        "ExerciseID": exid, "Category": category, "Side": f.get("side", ""),
        "Sets": f.get("sets", ""), "Reps_or_Scheme": f.get("reps", ""),
        "Load_kg": f.get("load", ""), "TopRPE": f.get("rpe", ""),
        "Pain_0_10": f.get("pain", ""), "Quality_0_5": f.get("quality", ""),
        "BestHold_s": f.get("besthold", ""), "CleanReps": f.get("cleanreps", ""),
        "Notes": f.get("notes", ""), "Key": key,
    }
    values = [row[c] for c in COLUMNS]

    ws = sheet.worksheet("B6_Training_Log")
    ws.append_row(values, value_input_option="USER_ENTERED")
    print("OK angehaengt:", exid, "->", category, "| Key", key)

if __name__ == "__main__":
    main()
