/**
 * Einstiegspunkt für GitHub Actions (Apps Script API: scripts.run).
 * Setzt Konfiguration/Secrets als Script Properties, führt eine erlaubte Funktion aus
 * und liefert Logger-Ausgabe, Dauer und Fehler strukturiert zurück.
 */
var GITHUB_RUNNER_ALLOWED_FUNCTIONS = {
  runAll: runAll,
  importLodgifyEinnahmenToImport: importLodgifyEinnahmenToImport,
  importLexwareAll: importLexwareAll,
  processLodgifyInvoiceUploadToLexware: processLodgifyInvoiceUploadToLexware,
  createLexwareManuelleUmsaetze: createLexwareManuelleUmsaetze,
  generateAlleBuchungenPlan: generateAlleBuchungenPlan,
  generateMonatswerte: generateMonatswerte
};

function githubRunnerExecute(functionName, config, debug) {
  var started = new Date().getTime();
  var result = { ok: true, functionName: functionName, appliedKeys: [], result: null, error: null, logs: "", durationMs: 0 };
  try {
    if (!Object.prototype.hasOwnProperty.call(GITHUB_RUNNER_ALLOWED_FUNCTIONS, functionName)) {
      throw new Error("Funktion nicht erlaubt: " + functionName);
    }
    if (config && typeof config === "object") {
      var keys = Object.keys(config).filter(function (k) {
        return config[k] !== null && config[k] !== undefined && String(config[k]) !== "";
      });
      var props = {};
      keys.forEach(function (k) { props[k] = String(config[k]); });
      if (keys.length) PropertiesService.getScriptProperties().setProperties(props, false);
      result.appliedKeys = keys; // nur Namen, niemals Werte
    }
    if (debug) Logger.log("[DEBUG] Start " + functionName + " @ " + new Date().toISOString());
    var r = GITHUB_RUNNER_ALLOWED_FUNCTIONS[functionName]();
    result.result = (r === undefined) ? null : JSON.parse(JSON.stringify(r));
    if (r && r.ok === false) result.ok = false;
  } catch (e) {
    result.ok = false;
    result.error = { message: String(e && e.message || e), stack: String(e && e.stack || "") };
  }
  result.durationMs = new Date().getTime() - started;
  result.logs = Logger.getLog();
  return result;
}
