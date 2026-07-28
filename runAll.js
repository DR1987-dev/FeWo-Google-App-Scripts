function runAll() {
  try {
    importAllCSVFromDrive();
    generateAlleBuchungenPlan();
    generateMonatswerte();
    importLodgifyEinnahmenToImport();
    importLexwareAll();
    Logger.log("✅ Alle Skripte erfolgreich ausgeführt");
  } catch (e) {
    Logger.log("❌ Abbruch wegen Fehler: " + e.message);
    throw e;
  }
}
