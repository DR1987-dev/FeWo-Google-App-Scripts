function runAll() {
  try {
    importLodgifyEinnahmenToImport();
    importLexwareAll();
    processLodgifyInvoiceUploadToLexware();
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
    runDoorKeySafely_();
    Logger.log("✅ Alle Skripte erfolgreich ausgeführt");
  } catch (e) {
    Logger.log("❌ Abbruch wegen Fehler: " + e.message);
    throw e;
  }
}

// Webhook-getriggerter Lauf: wie runAll, aber mit inkrementellem Lexware-Import.
function runDelta() {
  try {
    var lodgifyBookingId = PropertiesService.getScriptProperties().getProperty("LODGIFY_BOOKING_ID");
    Logger.log("ℹ️ LODGIFY_BOOKING_ID=" + JSON.stringify(lodgifyBookingId));
    if (lodgifyBookingId) {
      importLodgifyBookingToImport(lodgifyBookingId);
    } else {
      // Kein Einzelbuchungs-Event (z. B. Lexware-Webhook): nicht alle Lodgify-Buchungen laden.
      Logger.log("ℹ️ Keine Lodgify-Buchungs-ID – Lodgify-Import übersprungen (Voll-Abgleich nachts via runAll)");
    }
    importLexwareDelta();
    processLodgifyInvoiceUploadToLexware();
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
    runDoorKeySafely_();
    Logger.log("✅ Delta-Lauf erfolgreich ausgeführt");
  } catch (e) {
    Logger.log("❌ Abbruch wegen Fehler: " + e.message);
    throw e;
  }
}

// Türcode-Fehler sollen den restlichen Lauf nicht abbrechen.
function runDoorKeySafely_() {
  try {
    var result = processLodgifyDoorKeys();
    if (result && result.ok === false) {
      Logger.log("⚠️ DoorKey-Fehler: " + JSON.stringify(result.errors));
    }
  } catch (e) {
    Logger.log("⚠️ DoorKey-Lauf fehlgeschlagen: " + e.message);
  }
}
