function runAll() {
  try {
    importLodgifyEinnahmenToImport();
    importLexwareAll();
    processLodgifyInvoiceUploadToLexware();
    createLexwareFixkosten();
    var manuelleUmsaetzeResult = createLexwareManuelleUmsaetze();
    if (manuelleUmsaetzeResult && manuelleUmsaetzeResult.ok === false) {
      throw new Error(
        manuelleUmsaetzeResult.error ||
        (manuelleUmsaetzeResult.messages && manuelleUmsaetzeResult.messages.join(" | ")) ||
        "createLexwareManuelleUmsaetze fehlgeschlagen"
      );
    }
    generateAlleBuchungenPlan();
    generateMonatswerte();
    Logger.log("✅ Alle Skripte erfolgreich ausgeführt");
  } catch (e) {
    Logger.log("❌ Abbruch wegen Fehler: " + e.message);
    throw e;
  }
}
