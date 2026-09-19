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

function formatRowLabel(row, rowNumber, bookingRefIdx) {
    if (bookingRefIdx < 0) return `row ${rowNumber}`;
    const bookingRef = String(row[bookingRefIdx] ?? "").trim();
    if (!bookingRef) return `row ${rowNumber}`;
    return `row ${rowNumber} booking '${bookingRef}'`;
}

function hasSheetValue(value) {
    if (value === null || value === undefined) return false;
    return String(value).trim() !== "";
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

function decideAction(rows, checkInIdx, checkOutIdx, todayKey, bookingRefIdx = -1) {
    let action = null;
    const todaySkips = [];

    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const rowNumber = i + 2; // header is row 1
        const rowLabel = formatRowLabel(row, rowNumber, bookingRefIdx);

        const rawCheckIn = row[checkInIdx];
        const rawCheckOut = row[checkOutIdx];
        const checkInDate = parseSheetDate(rawCheckIn);
        const checkOutDate = parseSheetDate(rawCheckOut);
        if (!checkInDate || !checkOutDate) {
            const invalidDateParts = [];
            if (hasSheetValue(rawCheckIn) && !checkInDate) {
                invalidDateParts.push(`CheckIn='${String(rawCheckIn)}'`);
            }
            if (hasSheetValue(rawCheckOut) && !checkOutDate) {
                invalidDateParts.push(`CheckOut='${String(rawCheckOut)}'`);
            }
            if (invalidDateParts.length > 0) {
                todaySkips.push(
                    `SKIP: ${rowLabel} has invalid date value(s): ${invalidDateParts.join(", ")}.`
                );
            }
            continue;
        }

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

        if (!action || candidate.priority > action.priority) {
            action = candidate;
        }
    }

    return { action, todaySkips };
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
    const keypadUpdateChannelRaw = optionalEnv("CCU_KEYPAD_UPDATE_CHANNEL", "5");
    const keypadUpdateChannel = normalizeChannel(keypadUpdateChannelRaw);
    const insecureTls = optionalEnv("CCU_INSECURE_TLS", "true").toLowerCase() === "true";
    const runOnlyBerlinHourRaw = optionalEnv("RUN_ONLY_BERLIN_HOUR", "");
    const runOnlyBerlinHour = runOnlyBerlinHourRaw ? Number(runOnlyBerlinHourRaw) : null;

    if (!keypadUpdateChannel) {
        throw new Error(`Invalid CCU_KEYPAD_UPDATE_CHANNEL: ${keypadUpdateChannelRaw}`);
    }

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
    const bookingRefIdx = findHeaderIndex(headers, ["Buchungsnummer", "Booking ID", "BookingID", "ID"]);

    if (checkInIdx < 0 || checkOutIdx < 0) {
        throw new Error(
            `Required headers missing. Found: ${headers.join(", ")}. Need CheckIn/CheckOut.`
        );
    }

    const todayKey = asBerlinDateKey(new Date());
    const { action, todaySkips } = decideAction(
        rows,
        checkInIdx,
        checkOutIdx,
        todayKey,
        bookingRefIdx
    );

    if (!action) {
        console.log(`No keypad action for today (${todayKey}).`);
        for (const line of todaySkips) {
            console.log(line);
        }
        return;
    }

    for (const line of todaySkips) {
        console.log(line);
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

    const channelIseId = channelIseIds.get(keypadUpdateChannel);
    if (!channelIseId) {
        throw new Error(`No channel ise_id found for ${keypadSerial}:${keypadUpdateChannel}`);
    }

    const updateUrl = new URL(`${ccuBaseUrl}/addons/xmlapi/mastervaluechange.cgi`);
    updateUrl.searchParams.set("sid", ccuSid);
    updateUrl.searchParams.set("device_id", String(channelIseId));
    updateUrl.searchParams.set("name", masterName);
    updateUrl.searchParams.set("value", action.pin);

    const updRes = await ccuGetXml(updateUrl.toString(), insecureTls);
    if (updRes.status !== 200 || /<not_authenticated\s*\/>/i.test(updRes.body)) {
        throw new Error(`Update failed for channel ${keypadUpdateChannel} (HTTP ${updRes.status})`);
    }

    const verifyUrl = new URL(`${ccuBaseUrl}/addons/xmlapi/mastervalue.cgi`);
    verifyUrl.searchParams.set("sid", ccuSid);
    verifyUrl.searchParams.set("device_id", String(channelIseId));

    const verRes = await ccuGetXml(verifyUrl.toString(), insecureTls);
    const currentValue = extractMasterValue(verRes.body, masterName);

    if (verRes.status !== 200 || /<not_authenticated\s*\/>/i.test(verRes.body)) {
        throw new Error(`Verify auth failed for channel ${keypadUpdateChannel} (HTTP ${verRes.status})`);
    }

    if (currentValue !== action.pin) {
        throw new Error(
            `Verify mismatch for channel ${keypadUpdateChannel}. expected='${action.pin}' actual='${currentValue}' reason=${action.reason}`
        );
    }

    console.log(
        `OK: channel ${keypadUpdateChannel} (${keypadSerial}:${keypadUpdateChannel} -> ise_id ${channelIseId}) set ${masterName} to '${action.pin}' reason=${action.reason}`
    );

    console.log("PIN sync finished successfully for 1 channel action.");
}

main().catch((err) => {
    console.error(err && err.stack ? err.stack : err);
    process.exit(1);
});
