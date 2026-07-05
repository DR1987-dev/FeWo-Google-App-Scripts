function generateAlleBuchungenPlan() {
 
  // Hilfsfunktionen

  function num(v) {
    const n = Number(v);
    return isNaN(n) ? 0 : Number(n.toFixed(2));
  }

  function istDatumImZeitraum(datum, start, ende) {
    if (!start) return false;
    if (datum < start) return false;
    if (ende && datum > ende) return false;
    return true;
  }

  function passtZumMonatsIntervall(importDatum, startDatum, wertstellungstag, toleranzTage = 3) {
    const diffMonate =
      (importDatum.getFullYear() - startDatum.getFullYear()) * 12 +
      (importDatum.getMonth() - startDatum.getMonth());

    if (diffMonate < 0) return false;

    const erwartetesDatum = new Date(startDatum);
    erwartetesDatum.setMonth(startDatum.getMonth() + diffMonate);

    if (wertstellungstag) {
      erwartetesDatum.setDate(wertstellungstag);
    }

    const diffTage = Math.abs(
      (importDatum - erwartetesDatum) / (1000 * 60 * 60 * 24)
    );

    Logger.log(
      `🧮 Intervallprüfung: erwartet=${Utilities.formatDate(erwartetesDatum, Session.getScriptTimeZone(), "yyyy-MM-dd")}, ` +
      `import=${Utilities.formatDate(importDatum, Session.getScriptTimeZone(), "yyyy-MM-dd")}, diff=${diffTage}`
    );

    return diffTage <= toleranzTage;
  }

  function findePassendeFixkosten(importBuchung, fixkostenListe) {
    Logger.log(
      `🔎 Suche Fixkosten für '${importBuchung.Buchungstext}' am ${Utilities.formatDate(
        importBuchung.Datum,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd"
      )}`
    );

    const kandidaten = fixkostenListe.filter(f => {
      if (f.BuchungstextAbgleich !== importBuchung.Buchungstext) return false;

      if (!istDatumImZeitraum(importBuchung.Datum, f.Startdatum, f.Enddatum)) {
        return false;
      }

      if (f.Intervall && f.Intervall.toLowerCase() === "monat") {
        return passtZumMonatsIntervall(
          importBuchung.Datum,
          f.Startdatum,
          f.Wertstellungstag
        );
      }

      return true;
    });

    Logger.log(`📌 Gefundene Fixkosten-Kandidaten: ${kandidaten.length}`);

    if (kandidaten.length === 0) return null;

    // 👉 Wichtig: den zeitlich neuesten Start nehmen
    kandidaten.sort((a, b) => b.Startdatum - a.Startdatum);

    Logger.log(
      `✅ Verwendeter Fixkosten-Eintrag: ${kandidaten[0].Kostenart} → ${kandidaten[0].Buchungskonto}`
    );

    return kandidaten[0];
  }
 
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Sheets
  const sheetFixkosten = ss.getSheetByName("Fixkosten");
  const sheetImport = ss.getSheetByName("Import");
  const sheetImportZuordnung = ss.getSheetByName("Import_Konto_Zuordnung");
  const sheetManuelle = ss.getSheetByName("Manuelle_Buchungen");
  const sheetKontostart = ss.getSheetByName("Kontostartwerte");
  const sheetUmbuchungen = ss.getSheetByName("Umbuchungen");
  const sheetOutput = ss.getSheetByName("AlleBuchungenPlan");

  sheetOutput.clearContents();
  sheetOutput.appendRow(["Kostenart","Buchungskonto","Datum","Betrag","Kumuliert","Monatsstartwert","Monatsendwert"]);

  // -------------------------------
  // 1️⃣ Fixkosten laden
  const fixkostenData = sheetFixkosten.getDataRange().getValues().slice(1);
  const fixkosten = fixkostenData.map(r => ({
    Kostenart: r[0],
    Kategorie: r[1],
    Betrag: parseFloat(r[2]),
    Startdatum: r[3] ? new Date(r[3]) : null,
    Enddatum: r[4] ? new Date(r[4]) : null,
    BuchungstextAbgleich: r[5],
    Wertstellungstag: r[6] || null,
    Intervall: r[7],
    Buchungskonto: r[8]
  }));
  Logger.log(`📌 Fixkosten geladen: ${fixkosten.length}`);

  // -------------------------------
  // 2️⃣ Import-Zuordnungen
  const importZuordnungData = sheetImportZuordnung.getDataRange().getValues().slice(1);
  const importZuordnung = {};
  importZuordnungData.forEach(r => {
    const key = `${new Date(r[0]).toISOString().slice(0,10)}|${r[1]}`;
    importZuordnung[key] = r[2];
    Logger.log(`🔧 Importzuordnung geladen: ${key} → ${r[2]}`);
  });

  // -------------------------------
  // 3️⃣ Kontostartwerte
  const kontostartData = sheetKontostart.getDataRange().getValues().slice(1);
  const kontenStand = {};
  kontostartData.forEach(r => {
    kontenStand[r[0]] = parseFloat(r[1]);
    Logger.log(`💰 Kontostartwert: ${r[0]} = ${r[1]}`);
  });

  // -------------------------------
  // 4️⃣ Manuelle Buchungen
  const manuelleData = sheetManuelle.getDataRange().getValues().slice(1);
  const manuelleBuchungen = manuelleData.map(r => ({
    Kostenart: r[0],
    Buchungstext: r[1],
    Buchungskonto: r[2],
    Datum: new Date(r[3]),
    Betrag: parseFloat(r[4])
  }));
  Logger.log(`📥 Manuelle Buchungen geladen: ${manuelleBuchungen.length}`);

  // -------------------------------
  // 5️⃣ Import Buchungen (KORRIGIERT & STABIL)
  const importData = sheetImport.getDataRange().getValues().slice(1);

  const importBuchungen = importData.map((r, i) => {
    const buchungstext = r[4];
    const datum = new Date(r[3]);

    const soll = Number(r[9]);   // Soll (Ausgabe)
    const haben = Number(r[10]); // Haben (Einnahme)

    let betrag = 0;

    if (soll && soll !== 0) {
      betrag = -Math.abs(soll);
    } else if (haben && haben !== 0) {
      betrag = Math.abs(haben);
    }

    betrag = Number(betrag.toFixed(2));

    Logger.log(
      `📥 Import ${i + 2}: "${buchungstext}", Soll=${soll}, Haben=${haben}, Betrag=${betrag}`
    );

    return {
      Buchungstext: buchungstext,
      Datum: datum,
      Betrag: betrag
    };
  });

  // -------------------------------
  // 6️⃣ Umbuchungen
  const umbData = sheetUmbuchungen.getDataRange().getValues().slice(1);
  const umbuchungen = umbData.map(r => ({
    Datum: new Date(r[0]),
    Von: r[1],
    Nach: r[2],
    Betrag: parseFloat(r[3]),
    Text: r[4]
  }));
  Logger.log(`🔄 Umbuchungen geladen: ${umbuchungen.length}`);

  // -------------------------------
  // 7️⃣ Alle Buchungen zusammenbauen
  let alleBuchungen = [];

  // 7a️⃣ Import Buchungen verarbeiten
  importBuchungen.forEach(imp => {
    if (imp.Buchungstext.toLowerCase().includes("saldovortrag")) {
      Logger.log(`❌ Fiktiver Saldovortrag ignoriert: ${imp.Buchungstext}`);
      return;
    }

    // Abgleich mit Fixkosten innerhalb Start- und Enddatum ±3 Tage
    let konto = "Mietenkonto"; // Fallback IMMER

    // 1️⃣ Fixkosten-Abgleich (zeitlich + Intervall)
    const fix = findePassendeFixkosten(imp, fixkosten);

    if (fix) {
      konto = fix.Buchungskonto;
      Logger.log(
        `🔁 Import → Fixkosten-Zuordnung: '${imp.Buchungstext}' → ${konto}`
      );
    }

    // 2️⃣ Exakte Import-Zuordnung überschreibt Fixkosten
    const key = `${imp.Datum.toISOString().slice(0,10)}|${imp.Buchungstext}`;
    if (importZuordnung[key]) {
      konto = importZuordnung[key];
      Logger.log(
        `📝 Exakte Import-Zuordnung überschreibt: '${imp.Buchungstext}' → ${konto}`
      );
    }

    alleBuchungen.push({
      Kostenart: imp.Buchungstext,
      Buchungskonto: konto,
      Datum: imp.Datum,
      Betrag: Number(imp.Betrag.toFixed(2)),
      Quelle: "Import"
    });
    Logger.log(`✅ Import-Buchung: ${imp.Buchungstext} ${imp.Betrag} → ${konto}`);
  });

  // 7b️⃣ Manuelle Buchungen hinzufügen
  manuelleBuchungen.forEach(m => {
    alleBuchungen.push({
      Kostenart: m.Kostenart,
      Buchungskonto: m.Buchungskonto,
      Datum: m.Datum,
      Betrag: Number(m.Betrag.toFixed(2)),
      Quelle: "Manuell"
    });
    Logger.log(`➕ Manuelle Buchung: ${m.Kostenart} ${m.Betrag} → ${m.Buchungskonto}`);
  });

  // 7c️⃣ Fixkosten Forecast
  const today = new Date();
  const forecastEnd = new Date(today.getFullYear()+2, today.getMonth(), today.getDate());
  fixkosten.forEach(f => {
    if (!f.Startdatum) return;
    let d = new Date(f.Startdatum);
    while(d <= forecastEnd && (!f.Enddatum || d <= f.Enddatum)) {
      // Prüfen, ob schon ein Import existiert ±3 Tage
      let matchImport = alleBuchungen.find(a =>
        a.Kostenart === f.BuchungstextAbgleich &&
        Math.abs((a.Datum - d)/(1000*60*60*24)) <= 3
      );
      if (!matchImport) {
        let werttag = new Date(d);
        if(f.Wertstellungstag) werttag.setDate(f.Wertstellungstag);
        alleBuchungen.push({
          Kostenart: f.Kostenart,
          Buchungskonto: f.Buchungskonto,
          Datum: werttag,
          Betrag: Number((-Math.abs(f.Betrag)).toFixed(2)),
          Quelle: "FixkostenForecast"
        });
        Logger.log(`📅 Fixkosten hinzugefügt: ${f.Kostenart} ${-Math.abs(f.Betrag)} → ${f.Buchungskonto} am ${Utilities.formatDate(werttag, Session.getScriptTimeZone(), "yyyy-MM-dd")}`);
      }

      // Intervall erhöhen
      if(f.Intervall.toLowerCase() === "monat") d.setMonth(d.getMonth()+1);
      else if(f.Intervall.toLowerCase() === "quartal") d.setMonth(d.getMonth()+3);
      else if(f.Intervall.toLowerCase() === "jahr") d.setFullYear(d.getFullYear()+1);
      else break;
    }
  });

  // 7d️⃣ Umbuchungen einfügen
  umbuchungen.forEach(u => {
    alleBuchungen.push({
      Kostenart: u.Text,
      Buchungskonto: u.Von,
      Datum: u.Datum,
      Betrag: -Number(u.Betrag.toFixed(2)),
      Quelle: "Umbuchung"
    });
    alleBuchungen.push({
      Kostenart: u.Text,
      Buchungskonto: u.Nach,
      Datum: u.Datum,
      Betrag: Number(u.Betrag.toFixed(2)),
      Quelle: "Umbuchung"
    });
    Logger.log(`🔄 Umbuchung: ${u.Betrag} von ${u.Von} → ${u.Nach} am ${Utilities.formatDate(u.Datum, Session.getScriptTimeZone(), "yyyy-MM-dd")}`);
  });

  // -------------------------------
  // 8️⃣ Sortieren nach Datum
  alleBuchungen.sort((a,b)=>a.Datum-b.Datum);

  // -------------------------------
  // 9️⃣ Kumuliert, Monatsstartwert, Monatsendwert berechnen
  const kontenKumuliert = {...kontenStand};
  const kontenMonatStart = {};
  let currentMonth = null;

  alleBuchungen.forEach(b => {
    let m = `${b.Datum.getFullYear()}-${b.Datum.getMonth()}`;
    if(currentMonth !== m) {
      currentMonth = m;
      for(const k in kontenKumuliert) kontenMonatStart[k] = kontenKumuliert[k];
      Logger.log(`📆 Neuer Monat: ${m}, Monatsstartwerte: ${JSON.stringify(kontenMonatStart)}`);
    }

    kontenKumuliert[b.Buchungskonto] = (kontenKumuliert[b.Buchungskonto]||0) + b.Betrag;

    sheetOutput.appendRow([
      b.Kostenart,
      b.Buchungskonto,
      Utilities.formatDate(b.Datum, Session.getScriptTimeZone(), "yyyy-MM-dd"),
      num(b.Betrag),
      num(kontenKumuliert[b.Buchungskonto]),
      num(kontenMonatStart[b.Buchungskonto]),
      num(kontenKumuliert[b.Buchungskonto])
    ]);

    Logger.log(`✅ Buchung geschrieben: ${b.Kostenart}, ${b.Buchungskonto}, ${b.Betrag}, Quelle: ${b.Quelle}`);
  });

  Logger.log("🎉 Alle Buchungen generiert!");
}
