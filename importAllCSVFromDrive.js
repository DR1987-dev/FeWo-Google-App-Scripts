function importAllCSVFromDrive() {
  var folder = DriveApp.getFoldersByName("EÜR Exporte").next();
  var files = folder.getFiles();

  var latestFilePerYear = {}; // { "25": {date: "20251231", file: File} }

  Logger.log("🔍 Starte Suche im Ordner: EÜR Exporte");

  // 🔎 1. Alle Dateien scannen und pro Jahr die neueste merken
  while (files.hasNext()) {
    var file = files.next();
    var name = file.getName();
    Logger.log("📄 Gefundene Datei: " + name);

    // ✔ Jahr wird NUR aus dem vorderen Datum ermittelt
    var match = name.match(/^(\d{8})_Buchungen_EÜR_\d{2}\.csv$/);
    if (!match) {
      Logger.log("⏭️  Übersprungen (Namensmuster passt nicht)");
      continue;
    }

    var datePart = match[1];                 // YYYYMMDD
    var year = datePart.substring(2, 4);     // "24", "25", ...

    if (
      !latestFilePerYear[year] ||
      datePart > latestFilePerYear[year].date
    ) {
      latestFilePerYear[year] = {
        date: datePart,
        file: file
      };
    }
  }

  var allData = [];
  var header = null;

  var yearCount = {
    "24": 0,
    "25": 0,
    "26": 0,
    "27": 0
  };

  // 📥 2. Nur die neuesten Dateien je Jahr importieren
  for (var year in latestFilePerYear) {
    var entry = latestFilePerYear[year];
    var file = entry.file;
    var name = file.getName();

    Logger.log("✅ Verwende für Jahr 20" + year + ": " + name);

    var csvContent = file.getBlob().getDataAsString();
    var csvData = Utilities.parseCsv(csvContent, ';');

    if (!header) {
      header = csvData[0];
      allData.push(header);
    }

    var rows = csvData.slice(1);

    rows.forEach(function(row) {
      // 📅 Datumsspalten C & D
      row[2] = convertGermanDate(row[2]);
      row[3] = convertGermanDate(row[3]);

      // 💶 Betragsspalten J & K
      row[9]  = convertEuroNumber(row[9]);
      row[10] = convertEuroNumber(row[10]);

      allData.push(row);
      yearCount[year]++;
    });

    Logger.log("📥 Importierte Zeilen aus " + name + ": " + rows.length);
  }

  if (allData.length === 0) {
    Logger.log("❌ Keine passenden Dateien gefunden.");
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Import") || ss.insertSheet("Import");
  sheet.clearContents();

  sheet
    .getRange(1, 1, allData.length, allData[0].length)
    .setValues(allData);

  // 📊 Abschluss-Log
  Logger.log("📊 Import abgeschlossen");
  Logger.log("➡️  Gesamtzeilen (ohne Header): " + (allData.length - 1));
  Logger.log("➡️  Jahr 2024: " + yearCount["24"] + " Zeilen");
  Logger.log("➡️  Jahr 2025: " + yearCount["25"] + " Zeilen");
  Logger.log("➡️  Jahr 2026: " + yearCount["26"] + " Zeilen");
  Logger.log("➡️  Jahr 2027: " + yearCount["27"] + " Zeilen");
}

/* ===============================
   Hilfsfunktionen
   =============================== */

function convertGermanDate(value) {
  if (!value) return value;
  if (value instanceof Date) return value;

  if (typeof value === "string" && /^\d{2}\.\d{2}\.\d{4}$/.test(value)) {
    var p = value.split(".");
    return new Date(p[2], p[1] - 1, p[0]);
  }
  return value;
}

function convertEuroNumber(value) {
  if (value === "" || value === null) return value;
  if (typeof value === "number") return value;

  return parseFloat(
    value.toString().replace(/\./g, "").replace(",", ".")
  );
}
