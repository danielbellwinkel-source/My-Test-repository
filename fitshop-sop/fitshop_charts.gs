/**
 * FITSHOP HALL/TIROL – Chart Generator 2026
 *
 * Anleitung:
 *  1. Google Sheet öffnen
 *  2. Erweiterungen → Apps Script
 *  3. Diesen Code einfügen (alten Code ersetzen)
 *  4. Speichern (Strg+S)
 *  5. Funktion "createFitshopCharts" auswählen → ▶ Ausführen
 *  6. Beim ersten Mal: Berechtigungen bestätigen
 *
 * Das Script kann jederzeit erneut ausgeführt werden – es löscht
 * zuerst alle bestehenden Diagramme und erstellt sie neu.
 */

function createFitshopCharts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName('KPI Dashboard');

  if (!ws) {
    SpreadsheetApp.getUi().alert('Sheet "KPI Dashboard" nicht gefunden.\nBitte prüfen ob der Tab korrekt benannt ist.');
    return;
  }

  // Bestehende Diagramme entfernen
  var existing = ws.getCharts();
  for (var i = 0; i < existing.length; i++) {
    ws.removeChart(existing[i]);
  }

  // Fitshop Farben: Daniel=Rot, Jonas=Blau, Angie=Grün, Norbert=Orange
  var COLORS = ['#E31E24', '#1565C0', '#2E7D32', '#E65100'];
  var TITLE_STYLE = {color: '#2D2D2D', fontSize: 13, bold: true};
  var AXIS_STYLE  = {color: '#555555', fontSize: 10};
  var TICK_STYLE  = {color: '#333333', fontSize: 9};
  var GRID_COLOR  = '#E0E0E0';

  // ─── CHART 1: Umsatz-Entwicklung (Gestapelte Säulen) ─────────────────────
  // Daten: Hilfstabelle Zeile 111–123, Spalten A (Monate) + C–F (MA-Umsatz)
  var chart1 = ws.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(ws.getRange('A111:A123'))   // Monatsnamen (X-Achse)
    .addRange(ws.getRange('C111:F123'))   // Daniel / Jonas / Angie / Norbert
    .setOption('title', 'Umsatz-Entwicklung 2026 nach Mitarbeiter (EUR)')
    .setOption('titleTextStyle', TITLE_STYLE)
    .setOption('isStacked', true)
    .setOption('colors', COLORS)
    .setOption('backgroundColor', {fill: '#FFFFFF'})
    .setOption('chartArea', {left: 90, top: 55, width: '80%', height: '68%'})
    .setOption('hAxis', {
      title: 'Monat',
      titleTextStyle: AXIS_STYLE,
      textStyle: TICK_STYLE,
      slantedText: true,
      slantedTextAngle: 30
    })
    .setOption('vAxis', {
      title: 'EUR (monatlich)',
      titleTextStyle: AXIS_STYLE,
      textStyle: TICK_STYLE,
      format: '#,##0',
      gridlines: {color: GRID_COLOR},
      minValue: 0
    })
    .setOption('legend', {
      position: 'bottom',
      textStyle: {color: '#333333', fontSize: 10}
    })
    .setOption('bar', {groupWidth: '80%'})
    .setPosition(23, 1, 0, 0)
    .setNumRows(22)
    .setNumColumns(10)
    .build();
  ws.insertChart(chart1);

  // ─── CHART 2: Kumulierter Umsatz Gesamttrend (Linie) ─────────────────────
  // Daten: Spalte A (Monate) + Spalte M (kumulierter Umsatz)
  var chart2 = ws.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(ws.getRange('A111:A123'))   // Monatsnamen
    .addRange(ws.getRange('M111:M123'))   // Kumulierter Umsatz Gesamt
    .setOption('title', 'Kumulierter Umsatz 2026 – Gesamttrend (EUR)')
    .setOption('titleTextStyle', TITLE_STYLE)
    .setOption('colors', ['#E31E24'])
    .setOption('backgroundColor', {fill: '#FFFFFF'})
    .setOption('chartArea', {left: 90, top: 55, width: '80%', height: '68%'})
    .setOption('hAxis', {
      title: 'Monat',
      titleTextStyle: AXIS_STYLE,
      textStyle: TICK_STYLE,
      slantedText: true,
      slantedTextAngle: 30
    })
    .setOption('vAxis', {
      title: 'Kumuliert (EUR)',
      titleTextStyle: AXIS_STYLE,
      textStyle: TICK_STYLE,
      format: '#,##0',
      gridlines: {color: GRID_COLOR},
      minValue: 0
    })
    .setOption('legend', {position: 'none'})
    .setOption('lineWidth', 3)
    .setOption('pointSize', 7)
    .setOption('curveType', 'function')   // geschwungene Linie
    .setPosition(46, 1, 0, 0)
    .setNumRows(17)
    .setNumColumns(10)
    .build();
  ws.insertChart(chart2);

  // ─── CHART 3: Abschlüsse pro Monat nach Mitarbeiter (Gruppierte Säulen) ───
  // Daten: Spalte A (Monate) + Spalten H–K (MA-Abschlüsse)
  var chart3 = ws.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(ws.getRange('A111:A123'))   // Monatsnamen
    .addRange(ws.getRange('H111:K123'))   // Daniel / Jonas / Angie / Norbert
    .setOption('title', 'Abschlüsse pro Monat nach Mitarbeiter')
    .setOption('titleTextStyle', TITLE_STYLE)
    .setOption('isStacked', false)
    .setOption('colors', COLORS)
    .setOption('backgroundColor', {fill: '#FFFFFF'})
    .setOption('chartArea', {left: 90, top: 55, width: '80%', height: '68%'})
    .setOption('hAxis', {
      title: 'Monat',
      titleTextStyle: AXIS_STYLE,
      textStyle: TICK_STYLE,
      slantedText: true,
      slantedTextAngle: 30
    })
    .setOption('vAxis', {
      title: 'Anzahl Abschlüsse',
      titleTextStyle: AXIS_STYLE,
      textStyle: TICK_STYLE,
      format: '0',
      gridlines: {color: GRID_COLOR},
      minValue: 0
    })
    .setOption('legend', {
      position: 'bottom',
      textStyle: {color: '#333333', fontSize: 10}
    })
    .setOption('bar', {groupWidth: '70%'})
    .setPosition(64, 1, 0, 0)
    .setNumRows(17)
    .setNumColumns(10)
    .build();
  ws.insertChart(chart3);

  SpreadsheetApp.getUi().alert(
    '✓ 3 Diagramme erfolgreich erstellt!\n\n' +
    '• Umsatz-Entwicklung (gestapelte Säulen)\n' +
    '• Kumulierter Umsatz Gesamttrend (Linie)\n' +
    '• Abschlüsse pro Monat (gruppierte Säulen)\n\n' +
    'Die Diagramme aktualisieren sich automatisch wenn neue Leads eingetragen werden.'
  );
}
