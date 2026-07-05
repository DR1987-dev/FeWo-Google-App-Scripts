function generateMonatswerte() {
  const ss = SpreadsheetApp.getActive();
  const sourceSheet = ss.getSheetByName("AlleBuchungenPlan");
  if (!sourceSheet) throw new Error("❌ AlleBuchungenPlan nicht gefunden");

  const targetName = "Monatswerte";
  let targetSheet = ss.getSheetByName(targetName);
  if (!targetSheet) {
    targetSheet = ss.insertSheet(targetName);
  } else {
    targetSheet.clear();
  }

  const data = sourceSheet.getDataRange().getValues();
  data.shift(); // Header entfernen

  Logger.log("🔍 Starte Monatsaggregation, Datensätze: " + data.length);

  const map = {}; // key = konto|YYYY-MM → Array von Buchungen

  // 1️⃣ Buchungen sammeln
  data.forEach((row, i) => {
    const konto = row[1];
    const datum = row[2];
    const betrag = Number(row[3]);
    const kumuliert = Number(row[4]);

    if (!konto || !(datum instanceof Date) || isNaN(betrag) || isNaN(kumuliert)) {
      Logger.log("⚠️ Übersprungen Zeile " + (i + 2));
      return;
    }

    const year = datum.getFullYear();
    const month = String(datum.getMonth() + 1).padStart(2, "0");
    const key = konto + "|" + year + "-" + month;

    if (!map[key]) map[key] = [];

    map[key].push({
      year,
      month,
      konto,
      datum,
      betrag,
      kumuliert
    });
  });

  const output = [
    [
      "Jahr",
      "Monat",
      "Buchungskonto",
      "Monatsstartwert",
      "Einnahmen",
      "Ausgaben",
      "Monatsendwert",
      "Monatsdifferenz"
    ]
  ];

  // 2️⃣ Monatsweise korrekt berechnen
  Object.values(map)
    .sort((a, b) =>
      a[0].year !== b[0].year
        ? a[0].year - b[0].year
        : a[0].month.localeCompare(b[0].month) ||
          a[0].konto.localeCompare(b[0].konto)
    )
    .forEach(entries => {
      // 🔑 INNERHALB DES MONATS SORTIEREN
      entries.sort((a, b) => a.datum - b.datum);

      const first = entries[0];
      const last = entries[entries.length - 1];

      let einnahmen = 0;
      let ausgaben = 0;

      entries.forEach(e => {
        if (e.betrag > 0) einnahmen += e.betrag;
        if (e.betrag < 0) ausgaben += Math.abs(e.betrag);
      });

      const startwert = first.kumuliert - first.betrag;
      const endwert = last.kumuliert;
      const diff = einnahmen - ausgaben;

      output.push([
        first.year,
        first.month,
        first.konto,
        round2(startwert),
        round2(einnahmen),
        round2(ausgaben),
        round2(endwert),
        round2(diff)
      ]);
    });

  targetSheet
    .getRange(1, 1, output.length, output[0].length)
    .setValues(output);

  Logger.log("✅ Monatswerte korrekt geschrieben: " + (output.length - 1));
}
