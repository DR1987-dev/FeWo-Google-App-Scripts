import { google } from "googleapis";
import https from "node:https";

function requiredEnv(name) {
    const value = process.env[name];
    if (!value || String(value).trim() === "") {
        throw new Error(`Missing required env var: ${name}`);
    }
    return String(value).trim();
}

function optionalEnv(name, fallback = "") {
    const value = process.env[name];
    if (value === undefined || value === null) return fallback;
    return String(value).trim();
}

function asBerlinDateKey(date) {
    const formatter = new Intl.DateTimeFormat("en-CA", {
        timeZone: "Europe/Berlin",
        year: "numeric",
        month: "2-digit",
        day: "2-digit"
    });
    return formatter.format(date);
}

function berlinHour(date) {
    const formatter = new Intl.DateTimeFormat("en-GB", {
        timeZone: "Europe/Berlin",
        hour: "2-digit",
        hour12: false
    });
    return Number(formatter.format(date));
}

function parseSheetDate(value) {
    if (value === null || value === undefined || value === "") return null;

    // Google Sheets serial date number.
    if (typeof value === "number" && Number.isFinite(value)) {
        const millis = Math.round((value - 25569) * 86400 * 1000);
        const date = new Date(millis);
        return Number.isNaN(date.getTime()) ? null : date;
    }

    const text = String(value).trim();
    if (!text) return null;

    // ISO or RFC-like string.
    const isoDate = new Date(text);
    if (!Number.isNaN(isoDate.getTime())) return isoDate;

    // dd.mm.yyyy
    let match = text.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (match) {
        const day = Number(match[1]);
        const month = Number(match[2]);
        const year = Number(match[3]);
        const d = new Date(Date.UTC(year, month - 1, day, 12, 0, 0));
        return Number.isNaN(d.getTime()) ? null : d;
    }

    // yyyy-mm-dd
    match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (match) {
        const year = Number(match[1]);
        const month = Number(match[2]);
        const day = Number(match[3]);
        const d = new Date(Date.UTC(year, month - 1, day, 12, 0, 0));
        return Number.isNaN(d.getTime()) ? null : d;
    }

    return null;
}

function diffDaysCalendar(checkIn, checkOut) {
    const a = Date.UTC(checkIn.getUTCFullYear(), checkIn.getUTCMonth(), checkIn.getUTCDate());
    const b = Date.UTC(checkOut.getUTCFullYear(), checkOut.getUTCMonth(), checkOut.getUTCDate());
    return Math.max(0, Math.round((b - a) / 86400000));
}

function normalizeChannel(raw) {
    if (raw === null || raw === undefined) return null;
    const text = String(raw).trim();
    if (!text) return null;
    const match = text.match(/\d+/);
    if (!match) return null;
    const num = Number(match[0]);
    if (!Number.isInteger(num) || num < 1 || num > 64) return null;
    return num;
}

function findHeaderIndex(headers, candidates) {
    const lowered = headers.map((h) => String(h || "").trim().toLowerCase());
    for (const candidate of candidates) {
        const idx = lowered.indexOf(candidate.toLowerCase());
        if (idx >= 0) return idx;
    }
    return -1;
}

async function ccuGetXml(url, insecureTls) {
    const agent = url.startsWith("https://") && insecureTls
        ? new https.Agent({ rejectUnauthorized: false })
        : undefined;

    const response = await fetch(url, {
        method: "GET",
        agent
    });
    const body = await response.text();
    return { status: response.status, body };
}

function extractChannelIseIds(devicelistXml, keypadSerial) {
    const serialEscaped = keypadSerial.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const channelRegex = new RegExp(
        `<channel[^>]*address=['\"]${serialEscaped}:(\\d+)['\"][^>]*ise_id=['\"](\\d+)['\"][^>]*/?>`,
        "gi"
    );

    const result = new Map();
    let m;
    while ((m = channelRegex.exec(devicelistXml)) !== null) {
        const channel = Number(m[1]);
        const iseId = Number(m[2]);
        if (Number.isInteger(channel) && Number.isInteger(iseId)) {
            result.set(channel, iseId);
        }
    }
    return result;
}

function extractMasterValue(masterXml, name) {
    const nameEscaped = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const rx = new RegExp(`<mastervalue[^>]*name=['\"]${nameEscaped}['\"][^>]*value=['\"]([^'\"]*)['\"][^>]*/?>`, "i");
    const m = masterXml.match(rx);
    return m ? m[1] : null;
}

function buildPin(checkIn, checkOut) {
    const day = String(checkIn.getUTCDate()).padStart(2, "0");
    const year = String(checkIn.getUTCFullYear());
    const stayDays = String(diffDaysCalendar(checkIn, checkOut));
    return `${day}${year}${stayDays}`;
}

function decideActions(rows, checkInIdx, checkOutIdx, channelIdx, todayKey) {
    const actions = new Map();

    for (const row of rows) {
        const channel = normalizeChannel(row[channelIdx]);
        if (!channel) continue;

        const checkInDate = parseSheetDate(row[checkInIdx]);
        const checkOutDate = parseSheetDate(row[checkOutIdx]);
        if (!checkInDate || !checkOutDate) continue;

        const checkInKey = asBerlinDateKey(checkInDate);
        const checkOutKey = asBerlinDateKey(checkOutDate);

        let candidate = null;
        if (checkOutKey === todayKey) {
            candidate = {
                priority: 2,
                reason: "checkout_today",
                pin: "",
                checkInDate,
                checkOutDate
            };
        } else if (checkInKey === todayKey) {
            candidate = {
                priority: 1,
                reason: "checkin_today",
                pin: buildPin(checkInDate, checkOutDate),
                checkInDate,
                checkOutDate
            };
        }

        if (!candidate) continue;

        const existing = actions.get(channel);
        if (!existing || candidate.priority > existing.priority) {
            actions.set(channel, candidate);
        }
    }

    return actions;
}

async function main() {
    const spreadsheetId = requiredEnv("FEWO_SPREADSHEET_ID");
    const sheetName = optionalEnv("FEWO_SHEET_NAME", "AlleBuchungen");
    const ccuIp = requiredEnv("CCU_IP");
    const ccuScheme = optionalEnv("CCU_SCHEME", "https").toLowerCase() === "http" ? "http" : "https";
    const ccuBaseUrl = `${ccuScheme}://${ccuIp}`;
    const ccuSid = requiredEnv("CCU_XMLAPI_SID");
    const keypadSerial = requiredEnv("CCU_KEYPAD_SERIAL");
    const masterName = optionalEnv("CCU_KEYPAD_MASTER_NAME", "NUMERIC_PIN_CODE");
    const insecureTls = optionalEnv("CCU_INSECURE_TLS", "true").toLowerCase() === "true";
    const runOnlyBerlinHourRaw = optionalEnv("RUN_ONLY_BERLIN_HOUR", "");
    const runOnlyBerlinHour = runOnlyBerlinHourRaw ? Number(runOnlyBerlinHourRaw) : null;

    if (runOnlyBerlinHour !== null) {
        const currentBerlinHour = berlinHour(new Date());
        if (!Number.isInteger(runOnlyBerlinHour) || runOnlyBerlinHour < 0 || runOnlyBerlinHour > 23) {
            throw new Error(`Invalid RUN_ONLY_BERLIN_HOUR: ${runOnlyBerlinHourRaw}`);
        }
        if (currentBerlinHour !== runOnlyBerlinHour) {
            console.log(
                `Skip run: current Berlin hour is ${currentBerlinHour}, required ${runOnlyBerlinHour}.`
            );
            return;
        }
    }

    const serviceAccountJson = requiredEnv("GOOGLE_SERVICE_ACCOUNT");
    const serviceAccount = JSON.parse(serviceAccountJson);

    const auth = new google.auth.GoogleAuth({
        credentials: serviceAccount,
        scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    });

    const sheets = google.sheets({ version: "v4", auth });
    const range = `${sheetName}!A:Z`;
    const sheetRes = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range,
        valueRenderOption: "UNFORMATTED_VALUE",
        dateTimeRenderOption: "SERIAL_NUMBER"
    });

    const values = sheetRes.data.values || [];
    if (values.length < 2) {
        console.log("No booking rows found.");
        return;
    }

    const headers = values[0];
    const rows = values.slice(1);

    const checkInIdx = findHeaderIndex(headers, ["CheckIn", "Anreise"]);
    const checkOutIdx = findHeaderIndex(headers, ["CheckOut", "Abreise", "Checkout"]);
    const channelIdx = findHeaderIndex(headers, ["Kanal", "Channel"]);

    if (checkInIdx < 0 || checkOutIdx < 0 || channelIdx < 0) {
        throw new Error(
            `Required headers missing. Found: ${headers.join(", ")}. Need CheckIn/CheckOut/Kanal.`
        );
    }

    const todayKey = asBerlinDateKey(new Date());
    const actions = decideActions(rows, checkInIdx, checkOutIdx, channelIdx, todayKey);

    if (actions.size === 0) {
        console.log(`No keypad action for today (${todayKey}).`);
        return;
    }

    const devicelistUrl = `${ccuBaseUrl}/addons/xmlapi/devicelist.cgi?sid=${encodeURIComponent(ccuSid)}`;
    const devRes = await ccuGetXml(devicelistUrl, insecureTls);
    if (devRes.status !== 200) {
        throw new Error(`devicelist.cgi failed with HTTP ${devRes.status}`);
    }

    const channelIseIds = extractChannelIseIds(devRes.body, keypadSerial);
    if (channelIseIds.size === 0) {
        throw new Error(`No channel ise_ids found for keypad serial ${keypadSerial}`);
    }

    let failures = 0;

    for (const [channel, action] of actions.entries()) {
        const channelIseId = channelIseIds.get(channel);
        if (!channelIseId) {
            console.log(`WARN: no channel ise_id for ${keypadSerial}:${channel}`);
            failures += 1;
            continue;
        }

        const updateUrl = new URL(`${ccuBaseUrl}/addons/xmlapi/mastervaluechange.cgi`);
        updateUrl.searchParams.set("sid", ccuSid);
        updateUrl.searchParams.set("device_id", String(channelIseId));
        updateUrl.searchParams.set("name", masterName);
        updateUrl.searchParams.set("value", action.pin);

        const updRes = await ccuGetXml(updateUrl.toString(), insecureTls);
        if (updRes.status !== 200 || /<not_authenticated\s*\/>/i.test(updRes.body)) {
            console.log(`WARN: update failed for channel ${channel} (HTTP ${updRes.status})`);
            failures += 1;
            continue;
        }

        const verifyUrl = new URL(`${ccuBaseUrl}/addons/xmlapi/mastervalue.cgi`);
        verifyUrl.searchParams.set("sid", ccuSid);
        verifyUrl.searchParams.set("device_id", String(channelIseId));

        const verRes = await ccuGetXml(verifyUrl.toString(), insecureTls);
        const currentValue = extractMasterValue(verRes.body, masterName);

        if (verRes.status !== 200 || /<not_authenticated\s*\/>/i.test(verRes.body)) {
            console.log(`WARN: verify auth failed for channel ${channel} (HTTP ${verRes.status})`);
            failures += 1;
            continue;
        }

        if (currentValue !== action.pin) {
            console.log(
                `WARN: verify mismatch for channel ${channel}. expected='${action.pin}' actual='${currentValue}' reason=${action.reason}`
            );
            failures += 1;
            continue;
        }

        console.log(
            `OK: channel ${channel} (${keypadSerial}:${channel} -> ise_id ${channelIseId}) set ${masterName} to '${action.pin}' reason=${action.reason}`
        );
    }

    if (failures > 0) {
        throw new Error(`PIN sync finished with ${failures} failure(s).`);
    }

    console.log(`PIN sync finished successfully for ${actions.size} channel action(s).`);
}

main().catch((err) => {
    console.error(err && err.stack ? err.stack : err);
    process.exit(1);
});
