function generateAlleBuchungenPlan() {

  // Hilfsfunktionen

  function requireSheet(ss, name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      throw new Error(`❌ Sheet '${name}' nicht gefunden`);
    }
    return sheet;
  }

  function ensureDate(value, fieldName) {
    const d = new Date(value);
    if (isNaN(d.getTime())) {
      throw new Error(`❌ Ungültiges Datum in Feld '${fieldName}': ${value}`);
    }
    return d;
  }

  function num(v) {
    const n = Number(v);
    return isNaN(n) ? 0 : Number(n.toFixed(2));
  }

  function parseDateOrNull(value, fieldName) {
    if (!value) return null;
    const date = new Date(value);
    if (isNaN(date.getTime())) {
      throw new Error(`❌ Ungültiges Datum in Feld '${fieldName}': ${value}`);
    }
    return date;
  }

  function recurringIntervalMonths(interval) {
    const normalized = String(interval || "").toUpperCase();
    if (normalized === "MONTHLY") return 1;
    if (normalized === "QUARTERLY") return 3;
    if (normalized === "SEMIANNUALLY" || normalized === "HALF_YEARLY") return 6;
    if (normalized === "YEARLY" || normalized === "ANNUALLY") return 12;
    throw new Error(`❌ Nicht unterstütztes Dauerbeleg-Intervall: ${interval}`);
  }

  function resolveRecurringAccount(vendorNumber, category, kontoZuordnung) {
    const vendor = String(vendorNumber || "").trim();
    const categoryKey = String(category || "").trim().toLowerCase();
    if (vendor && categoryKey && kontoZuordnung.composite[vendor + "|" + categoryKey]) {
      return kontoZuordnung.composite[vendor + "|" + categoryKey];
    }
    if (vendor && kontoZuordnung.vendorOnly[vendor]) {
      return kontoZuordnung.vendorOnly[vendor];
    }
    if (categoryKey && kontoZuordnung.category[categoryKey]) {
      return kontoZuordnung.category[categoryKey];
    }
    return "Mietenkonto";
  }

  function findePassendesDauerbeleg(importBuchung, dauerbelege) {
    const kandidaten = dauerbelege.filter(dauerbeleg => {
      if (dauerbeleg.BuchungstextAbgleich !== importBuchung.Buchungstext) return false;
      if (dauerbeleg.Startdatum && importBuchung.Datum < dauerbeleg.Startdatum) return false;
      if (dauerbeleg.Enddatum && importBuchung.Datum > dauerbeleg.Enddatum) return false;

      const diffMonate =
        (importBuchung.Datum.getFullYear() - dauerbeleg.Startdatum.getFullYear()) * 12 +
        (importBuchung.Datum.getMonth() - dauerbeleg.Startdatum.getMonth());
      if (diffMonate < 0 || diffMonate % dauerbeleg.IntervallMonate !== 0) return false;

      const erwartetesDatum = new Date(dauerbeleg.Startdatum);
      erwartetesDatum.setMonth(
        dauerbeleg.Startdatum.getMonth() + diffMonate
      );
      const diffTage = Math.abs(
        (importBuchung.Datum - erwartetesDatum) / (1000 * 60 * 60 * 24)
      );
      return diffTage <= 3;
    });

    kandidaten.sort((a, b) => b.Startdatum - a.Startdatum);
    return kandidaten[0] || null;
  }

  function loadLexwareDauerbelege() {
    const kontoZuordnung = buildKontoZuordnungIndex_();
    const contactIdToVendorNumber = buildContactIdToVendorNumberIndex_();
    const dauerbelege = [];
    const pageSize = 100;
    let page = 0;
    let totalPages = 1;

    do {
      const result = lexwareGetRecurringTemplates_(page, pageSize);
      const body = result.body;
      if (!body || !Array.isArray(body.content)) {
        throw new Error(
          `❌ Unerwartete Antwort von Lexware-Dauerbelegen auf Seite ${page}: ${JSON.stringify(body)}`
        );
      }

      totalPages = body.totalPages !== undefined
        ? body.totalPages
        : (body.page && body.page.totalPages !== undefined ? body.page.totalPages : 1);

      body.content.forEach(templateSummary => {
        const templateId = String(templateSummary.id || "").trim();
        if (!templateId) return;

        const detail = lexwareGetRecurringTemplateDetail_(templateId).body || {};
        if (detail.archived) return;

        const settings = detail.recurringTemplateSettings || templateSummary.recurringTemplateSettings || {};
        const status = String(settings.executionStatus || "").toUpperCase();
        if (status === "PAUSED" || status === "INACTIVE" || status === "ARCHIVED") return;

        const startDate = parseDateOrNull(settings.startDate, `Dauerbeleg ${templateId}.startDate`);
        if (!startDate) {
          throw new Error(`❌ Dauerbeleg '${templateId}' hat kein Startdatum`);
        }
        const endDate = parseDateOrNull(settings.endDate, `Dauerbeleg ${templateId}.endDate`);
        const intervalMonths = recurringIntervalMonths(settings.executionInterval);
        const address = detail.address || templateSummary.address || {};
        const vendorNumber = address.contactId
          ? contactIdToVendorNumber[String(address.contactId).trim()] || ""
          : "";
        const lineItems = Array.isArray(detail.lineItems) && detail.lineItems.length
          ? detail.lineItems
          : [{
              name: detail.title || templateSummary.title || address.name || "Dauerbeleg",
              lineItemAmount: detail.totalPrice && detail.totalPrice.totalGrossAmount
            }];

        lineItems.forEach((item, index) => {
          const category = String(item.name || item.description || detail.title || address.name || "Dauerbeleg").trim();
          const amount = Number(
            item.lineItemAmount !== undefined
              ? item.lineItemAmount
              : (item.unitPrice && item.unitPrice.grossAmount !== undefined
                ? Number(item.unitPrice.grossAmount) * Number(item.quantity || 1)
                : 0)
          );
          if (!isFinite(amount) || amount === 0) {
            throw new Error(
              `❌ Dauerbeleg '${templateId}', Position ${index + 1} hat keinen gültigen Betrag`
            );
          }

          dauerbelege.push({
            Kostenart: category,
            BuchungstextAbgleich: category,
            Betrag: -Math.abs(Number(amount.toFixed(2))),
            Startdatum: startDate,
            Enddatum: endDate,
            IntervallMonate: intervalMonths,
            Buchungskonto: resolveRecurringAccount(vendorNumber, category, kontoZuordnung)
          });
        });
      });

      page++;
    } while (page < totalPages);

    Logger.log(`📌 Lexware-Dauerbelege geladen: ${dauerbelege.length} Position(en)`);
    return dauerbelege;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("❌ Kein aktives Spreadsheet verfügbar");
  }

  // Sheets
  const sheetImport = requireSheet(ss, "Import");
  const sheetImportZuordnung = requireSheet(ss, "Import_Konto_Zuordnung");
  const sheetManuelle = requireSheet(ss, "Manuelle_Buchungen");
  const sheetKontostart = requireSheet(ss, "Kontostartwerte");
  const sheetUmbuchungen = requireSheet(ss, "Umbuchungen");
  const sheetOutput = requireSheet(ss, "AlleBuchungenPlan");

  sheetOutput.clearContents();
  sheetOutput.appendRow(["Kostenart", "Buchungskonto", "Datum", "Betrag", "Kumuliert", "Monatsstartwert", "Monatsendwert"]);

  // -------------------------------
  // 1️⃣ Lexware-Dauerbelege laden
  const dauerbelege = loadLexwareDauerbelege();

  // -------------------------------
  // 2️⃣ Import-Zuordnungen
  const importZuordnungData = sheetImportZuordnung.getDataRange().getValues().slice(1);
  const importZuordnung = {};
  importZuordnungData.forEach(r => {
    if (!r[0] || !r[1] || !r[2]) return;
    const importZuordnungsDatum = new Date(r[0]);
    if (isNaN(importZuordnungsDatum.getTime())) return;
    const key = buildDateTextKey(importZuordnungsDatum, r[1]);
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
    const datum = ensureDate(r[3], `Import.Datum Zeile ${i + 2}`);

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
  // 5b️⃣ Lexware Umsatz Import (ab 2026) – eine Zeile pro Position mit vorberechnetem Konto
  const START_2026 = new Date(2026, 0, 1);
  let lexwareUmsaetze = [];

  const sheetLexwareUmsaetze = ss.getSheetByName("Lexware Umsatz Import");
  if (sheetLexwareUmsaetze) {
    const lexwareUmsaetzeData = sheetLexwareUmsaetze.getDataRange().getValues().slice(1);
    lexwareUmsaetzeData.forEach((r, i) => {
      // Spalten (0-basiert): A=Zeilen_ID, B=Beleg_ID, C=Belegtyp, D=Status,
      //   E=Belegnummer, F=Belegdatum, G=Fälligkeitsdatum, H=Kontakt,
      //   I=Lieferantennummer, J=Gesamtbetrag, K=Währung, L=Bemerkung, M=Position,
      //   N=Pos_Kategorie, O=Pos_Betrag_Brutto, P=Pos_MwSt_Satz,
      //   Q=Pos_MwSt_Betrag, R=Konto
      const belegtyp = String(r[2] || "").toLowerCase().trim();
      if (belegtyp !== "salesinvoice" && belegtyp !== "purchaseinvoice") return;

      const rawDatum = r[5];
      if (!rawDatum) return;
      const belegdatum = new Date(rawDatum);
      if (isNaN(belegdatum.getTime()) || belegdatum < START_2026) return;

      const kontakt   = String(r[7]  || "");
      const kategorie = String(r[13] || "");
      const posBetragStr = String(r[14] || "0").replace(",", ".");
      const posBetrag = parseFloat(posBetragStr) || 0;
      const konto = String(r[17] || "").trim() || "Mietenkonto";

      const kostenart = kategorie || kontakt;
      const betrag = belegtyp === "salesinvoice"
        ? Math.abs(posBetrag)
        : -Math.abs(posBetrag);

      lexwareUmsaetze.push({
        Buchungstext:  kostenart,
        Datum:         belegdatum,
        Betrag:        Number(betrag.toFixed(2)),
        Buchungskonto: konto
      });

      Logger.log(
        `📦 Lexware Umsatz Import ${i + 2}: "${kostenart}", Typ=${belegtyp}, Betrag=${betrag}, Konto=${konto}`
      );
    });
    Logger.log(
      `📌 Lexware Umsatz Import geladen: ${lexwareUmsaetze.length}`
    );
  }

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
    const buchungstextLower = String(imp.Buchungstext || "").toLowerCase();
    if (buchungstextLower.includes("saldovortrag")) {
      Logger.log(`❌ Fiktiver Saldovortrag ignoriert: ${imp.Buchungstext}`);
      return;
    }

    let konto = "Mietenkonto";

    // Dauerbeleg-Zuordnung als Ersatz für die frühere Fixkosten-Zuordnung.
    const dauerbeleg = findePassendesDauerbeleg(imp, dauerbelege);
    if (dauerbeleg) {
      konto = dauerbeleg.Buchungskonto;
      Logger.log(
        `🔁 Import → Dauerbeleg-Zuordnung: '${imp.Buchungstext}' → ${konto}`
      );
    }

    // Exakte Import-Zuordnung überschreibt den Dauerbeleg-Fallback.
    const key = buildDateTextKey(imp.Datum, imp.Buchungstext);
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

  // 7c️⃣ Lexware Umsatz Import Buchungen (ab 2026) – Konto ist bereits vorberechnet
  lexwareUmsaetze.forEach(lx => {
    alleBuchungen.push({
      Kostenart: lx.Buchungstext,
      Buchungskonto: lx.Buchungskonto,
      Datum: lx.Datum,
      Betrag: lx.Betrag,
      Quelle: "LexwareImport"
    });
    Logger.log(`✅ Lexware-Buchung: ${lx.Buchungstext} ${lx.Betrag} → ${lx.Buchungskonto}`);
  });

  // 7d️⃣ Lexware-Dauerbelege forecasten
  const today = new Date();
  const forecastEnd = new Date(today.getFullYear() + 2, today.getMonth(), today.getDate());
  dauerbelege.forEach(dauerbeleg => {
    let d = new Date(dauerbeleg.Startdatum);
    while (d <= forecastEnd && (!dauerbeleg.Enddatum || d <= dauerbeleg.Enddatum)) {
      // Prüfen, ob schon ein Import existiert ±3 Tage
      let matchImport = alleBuchungen.find(a =>
        a.Kostenart === dauerbeleg.BuchungstextAbgleich &&
        Math.abs((a.Datum - d) / (1000 * 60 * 60 * 24)) <= 3
      );
      if (!matchImport) {
        alleBuchungen.push({
          Kostenart: dauerbeleg.Kostenart,
          Buchungskonto: dauerbeleg.Buchungskonto,
          Datum: new Date(d),
          Betrag: dauerbeleg.Betrag,
          Quelle: "LexwareDauerbeleg"
        });
        Logger.log(`📅 Dauerbeleg hinzugefügt: ${dauerbeleg.Kostenart} ${dauerbeleg.Betrag} → ${dauerbeleg.Buchungskonto} am ${Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd")}`);
      }

      // Intervall erhöhen
      d.setMonth(d.getMonth() + dauerbeleg.IntervallMonate);
    }
  });

  // 7e️⃣ Umbuchungen einfügen
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
  alleBuchungen.sort((a, b) => a.Datum - b.Datum);

  // -------------------------------
  // 9️⃣ Kumuliert, Monatsstartwert, Monatsendwert berechnen
  const kontenKumuliert = { ...kontenStand };
  const kontenMonatStart = {};
  let currentMonth = null;
  const outputRows = [];

  alleBuchungen.forEach(b => {
    let m = `${b.Datum.getFullYear()}-${b.Datum.getMonth()}`;
    if (currentMonth !== m) {
      currentMonth = m;
      for (const k in kontenKumuliert) kontenMonatStart[k] = kontenKumuliert[k];
      Logger.log(`📆 Neuer Monat: ${m}, Monatsstartwerte: ${JSON.stringify(kontenMonatStart)}`);
    }

    kontenKumuliert[b.Buchungskonto] = (kontenKumuliert[b.Buchungskonto] || 0) + b.Betrag;

    outputRows.push([
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

  if (outputRows.length > 0) {
    sheetOutput
      .getRange(2, 1, outputRows.length, outputRows[0].length)
      .setValues(outputRows);
  }

  Logger.log("🎉 Alle Buchungen generiert!");
}
