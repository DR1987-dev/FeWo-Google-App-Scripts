function runAll() {
  try {
    importLodgifyEinnahmenToImport();
    importLexwareAll();
    processLodgifyInvoiceUploadToLexware();
    createLexwareFixkosten();
    generateAlleBuchungenPlan();
    generateMonatswerte();
    Logger.log("✅ Alle Skripte erfolgreich ausgeführt");
  } catch (e) {
    Logger.log("❌ Abbruch wegen Fehler: " + e.message);
    throw e;
  }
}
