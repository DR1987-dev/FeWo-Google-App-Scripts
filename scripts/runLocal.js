// Führt die Apps-Script-Funktionen (Root-*.js) direkt in Node/GitHub Actions aus.
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";
import { startBridge, stopBridge } from "./runtime/bridge.js";
import { globals, flushAll, logLines } from "./runtime/appsScriptShim.js";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const ALLOWED = [
    "runAll",
    "importLodgifyEinnahmenToImport",
    "importLexwareAll",
    "importLexwareDelta",
    "processLodgifyInvoiceUploadToLexware",
    "createLexwareManuelleUmsaetze",
    "generateAlleBuchungenPlan",
    "generateMonatswerte",
    "importAllCSVFromDrive",
    "createLexwareFixkosten",
];

function summary(md) {
    if (process.env.GITHUB_STEP_SUMMARY) fs.appendFileSync(process.env.GITHUB_STEP_SUMMARY, md + "\n");
}

async function main() {
    const fn = String(process.env.FUNCTION_NAME || "runAll").trim();
    if (!ALLOWED.includes(fn)) throw new Error(`Funktion nicht erlaubt: ${fn}`);
    for (const k of ["GOOGLE_SERVICE_ACCOUNT", "SPREADSHEET_ID"]) {
        if (!String(process.env[k] || "").trim()) throw new Error(`Missing required env var: ${k}`);
    }
    for (const k of ["LEXWARE_API_KEY", "LODGIFY_API_KEY"]) {
        if (process.env[k]) console.log(`::add-mask::${process.env[k]}`);
    }

    await startBridge();
    Object.assign(globalThis, globals);
    const files = fs.readdirSync(ROOT).filter((f) => f.endsWith(".js")).sort();
    for (const f of files) {
        new vm.Script(fs.readFileSync(path.join(ROOT, f), "utf8"), { filename: f }).runInThisContext();
    }
    if (typeof globalThis[fn] !== "function") throw new Error(`Funktion ${fn} nicht gefunden`);

    console.log(`Starte ${fn} (${files.length} Quelldateien geladen, Debug=${process.env.DEBUG})`);
    const started = Date.now();
    let error = null, result;
    try {
        result = globalThis[fn]();
        if (result && result.ok === false) error = new Error(result.error || "Funktion meldete ok=false");
    } catch (e) { error = e; }
    try { flushAll(); } catch (e) { error = error || e; console.error("Flush fehlgeschlagen:", e.message); }
    const ms = Date.now() - started;

    if (result !== undefined) console.log("Ergebnis:", JSON.stringify(result, null, 2));
    summary(`## ${fn}: ${error ? "❌ Fehler" : "✅ Erfolg"} (${ms} ms)`);
    if (error) summary("```\n" + (error.stack || error.message) + "\n```");
    summary("<details><summary>Logs</summary>\n\n```\n" + logLines.join("\n").slice(-50000) + "\n```\n</details>");
    stopBridge();
    if (error) { console.error("::error::" + (error.message || error)); console.error(error.stack || ""); process.exit(1); }
    console.log(`✅ ${fn} erfolgreich (${ms} ms)`);
    process.exit(0);
}

main().catch((e) => { console.error("Run failed:", e?.stack || e); process.exit(1); });
