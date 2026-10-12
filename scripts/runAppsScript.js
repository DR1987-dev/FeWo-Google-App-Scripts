import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { google } from "googleapis";

const PROJECT_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

// Env-Namen, die (falls gesetzt) als Script Properties an Apps Script übergeben werden.
const CONFIG_KEYS = [
    "LEXWARE_API_KEY", "LODGIFY_API_KEY", "ALERT_EMAIL",
    "LODGIFY_API_BASE_URL", "LODGIFY_PROPERTY_ID", "LODGIFY_TIMEOUT_MS",
    "LODGIFY_BOOKINGS_PATH", "LODGIFY_RESERVATIONS_PATH", "LODGIFY_GUEST_MESSAGE_PATH",
    "LODGIFY_GUEST_MESSAGE_TYPE", "LODGIFY_GUEST_MESSAGE_SUBJECT_TEMPLATE",
    "INVOICE_EINMALKUNDEN", "INVOICE_UPLOAD_CUTOFF_DATE", "INVOICE_VOUCHER_DAYS_BEFORE_CHECKIN",
    "PAYMENT_REQUEST_ALERT_EMAILS", "PAYMENT_REQUEST_ERROR_NOTIFY_EMAILS",
    "PAYMENT_REQUEST_DAYS_BEFORE_CHECKIN", "PAYMENT_REQUEST_WEEKS_BEFORE_CHECKIN",
    "PAYMENT_REQUEST_MESSAGE_TEMPLATE",
];

const MASKED_KEYS = ["LEXWARE_API_KEY", "LODGIFY_API_KEY"];

function required(name) {
    const v = String(process.env[name] || "").trim();
    if (!v) throw new Error(`Missing required env var: ${name}`);
    return v;
}

function debugLog(...args) {
    if (process.env.DEBUG === "true") console.log("[debug]", ...args);
}

async function main() {
    const scriptId = required("SCRIPT_ID");
    const functionName = String(process.env.FUNCTION_NAME || "runAll").trim();
    const impersonate = required("GOOGLE_IMPERSONATE_USER");
    const debug = process.env.DEBUG === "true";

    const credentialsPath = process.env.GOOGLE_APPLICATION_CREDENTIALS || path.resolve(PROJECT_ROOT, "credentials.json");
    const credentials = JSON.parse(await fs.readFile(credentialsPath, "utf8"));

    const auth = new google.auth.JWT({
        email: credentials.client_email,
        key: credentials.private_key,
        subject: impersonate,
        scopes: [
            "https://www.googleapis.com/auth/script.projects",
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/script.external_request",
            "https://www.googleapis.com/auth/script.send_mail",
        ],
    });
    await auth.authorize();

    const config = {};
    for (const key of CONFIG_KEYS) {
        const v = String(process.env[key] || "").trim();
        if (v) config[key] = v;
    }
    // Secrets im Log maskieren
    for (const key of MASKED_KEYS) if (config[key]) console.log(`::add-mask::${config[key]}`);
    console.log(`Script: ${scriptId} | Funktion: ${functionName} | Debug: ${debug}`);
    console.log(`Übergebene Config-Keys: ${Object.keys(config).join(", ") || "(keine)"}`);

    const script = google.script({ version: "v1", auth });
    const started = Date.now();
    const res = await script.scripts.run({
        scriptId,
        requestBody: {
            function: "githubRunnerExecute",
            parameters: [functionName, config, debug],
            devMode: process.env.DEV_MODE !== "false",
        },
    });
    debugLog(`API-Antwort nach ${Date.now() - started} ms`);

    if (res.data.error) {
        console.error("Apps Script API Fehler:", JSON.stringify(res.data.error, null, 2));
        process.exit(1);
    }

    const out = res.data.response?.result || {};
    console.log("::group::Apps Script Logs");
    console.log(out.logs || "(keine Logs)");
    console.log("::endgroup::");
    if (out.result !== null && out.result !== undefined) {
        console.log("Ergebnis:", JSON.stringify(out.result, null, 2));
    }
    console.log(`Dauer (Apps Script): ${out.durationMs} ms`);

    if (!out.ok) {
        console.error("Fehler:", out.error?.message || "Funktion meldete ok=false");
        if (out.error?.stack) console.error(out.error.stack);
        process.exit(1);
    }
    console.log("✅ Erfolgreich");
}

main().catch((err) => {
    console.error("Run failed:", err?.message || err);
    if (err?.response?.data) console.error(JSON.stringify(err.response.data, null, 2));
    process.exit(1);
});
