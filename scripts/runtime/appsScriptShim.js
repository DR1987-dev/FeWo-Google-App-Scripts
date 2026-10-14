// Emuliert die genutzten Apps-Script-Globals (SpreadsheetApp, UrlFetchApp, Utilities, ...) in Node.
import crypto from "node:crypto";
import { call, sleepMs } from "./bridge.js";

export const TZ = "Europe/Berlin";
const DEBUG = String(process.env.DEBUG || "").toLowerCase() === "true";

// ---------- Zeitzonen / Datum ----------
function tzParts(ms, tz) {
    const f = new Intl.DateTimeFormat("en-CA", {
        timeZone: tz, hourCycle: "h23", year: "numeric", month: "2-digit", day: "2-digit",
        hour: "2-digit", minute: "2-digit", second: "2-digit",
    });
    const p = {};
    for (const part of f.formatToParts(new Date(ms))) p[part.type] = part.value;
    return p;
}
function tzOffsetMs(ms, tz) {
    const p = tzParts(ms, tz);
    const wall = Date.UTC(+p.year, +p.month - 1, +p.day, +p.hour, +p.minute, +p.second);
    return wall - Math.floor(ms / 1000) * 1000;
}
const SERIAL_EPOCH_OFFSET = 25569; // Tage zwischen 1899-12-30 und 1970-01-01
export function serialToDate(serial) {
    const wall = Math.round((serial - SERIAL_EPOCH_OFFSET) * 86400000);
    let inst = wall - tzOffsetMs(wall, TZ);
    inst = wall - tzOffsetMs(inst, TZ);
    return new Date(inst);
}
export function dateToSerial(date) {
    const wall = date.getTime() + tzOffsetMs(date.getTime(), TZ);
    return wall / 86400000 + SERIAL_EPOCH_OFFSET;
}

function formatDate(date, tz, pattern) {
    const p = tzParts(date.getTime(), tz || TZ);
    const ms = String(date.getMilliseconds()).padStart(3, "0");
    return String(pattern).replace(/'([^']*)'|yyyy|yy|MM|dd|HH|mm|ss|SSS/g, (m, lit) => {
        if (lit !== undefined) return lit === "" ? "'" : lit;
        switch (m) {
            case "yyyy": return p.year;
            case "yy": return p.year.slice(2);
            case "MM": return p.month;
            case "dd": return p.day;
            case "HH": return p.hour;
            case "mm": return p.minute;
            case "ss": return p.second;
            case "SSS": return ms;
            default: return m;
        }
    });
}

function parseCsv(text, delimiter) {
    const d = delimiter || ",";
    const rows = [];
    let row = [], field = "", inQ = false;
    const s = String(text);
    for (let i = 0; i < s.length; i++) {
        const ch = s[i];
        if (inQ) {
            if (ch === '"') { if (s[i + 1] === '"') { field += '"'; i++; } else inQ = false; }
            else field += ch;
        } else if (ch === '"') inQ = true;
        else if (ch === d) { row.push(field); field = ""; }
        else if (ch === "\n" || ch === "\r") {
            if (ch === "\r" && s[i + 1] === "\n") i++;
            row.push(field); field = ""; rows.push(row); row = [];
        } else field += ch;
    }
    if (field !== "" || row.length) { row.push(field); rows.push(row); }
    return rows;
}

// ---------- Logger ----------
const logLines = [];
const Logger = {
    log(msg) {
        const line = String(msg);
        logLines.push(line);
        console.log(line);
    },
    getLog() { return logLines.join("\n"); },
};

// ---------- Properties ----------
const runtimeProps = {};
const PropertiesService = {
    getScriptProperties() {
        return {
            getProperty(key) {
                if (Object.prototype.hasOwnProperty.call(runtimeProps, key)) return runtimeProps[key];
                const v = process.env[key];
                return v === undefined || v === "" ? null : v;
            },
            setProperty(key, value) { runtimeProps[key] = String(value); return this; },
            setProperties(obj) { Object.keys(obj).forEach((k) => { runtimeProps[k] = String(obj[k]); }); return this; },
            deleteProperty(key) { delete runtimeProps[key]; return this; },
        };
    },
};

// ---------- UrlFetchApp ----------
const UrlFetchApp = {
    fetch(url, options) {
        const o = options || {};
        const headers = Object.assign({}, o.headers || {});
        let body = o.payload;
        if (body !== undefined && body !== null && typeof body !== "string") {
            if (typeof body === "object" && !Buffer.isBuffer(body) && !body.getBytes) {
                body = new URLSearchParams(body).toString();
                if (!o.contentType) headers["Content-Type"] = "application/x-www-form-urlencoded";
            } else throw new Error("UrlFetchApp: Blob-Payload wird im GitHub-Runner nicht unterstützt");
        }
        const hasCT = Object.keys(headers).some((k) => k.toLowerCase() === "content-type");
        if (o.contentType) headers["Content-Type"] = o.contentType;
        else if (body !== undefined && !hasCT) headers["Content-Type"] = "application/x-www-form-urlencoded";
        const method = String(o.method || "get").toLowerCase();
        const started = Date.now();
        const res = call("fetch", { url: String(url), method, headers, body });
        if (DEBUG) console.log(`[http] ${method.toUpperCase()} ${String(url).split("?")[0]} -> ${res.status} (${Date.now() - started} ms)`);
        if (!o.muteHttpExceptions && res.status >= 400) {
            throw new Error(`Request failed for ${url} returned code ${res.status}. Truncated server response: ${res.body.slice(0, 500)}`);
        }
        return {
            getResponseCode: () => res.status,
            getContentText: () => res.body,
            getAllHeaders: () => res.headers,
            getHeaders: () => res.headers,
        };
    },
};

// ---------- Utilities / Session / MailApp / ScriptApp / DriveApp ----------
const Utilities = {
    formatDate, parseCsv,
    sleep: sleepMs,
    getUuid: () => crypto.randomUUID(),
};
const Session = {
    getScriptTimeZone: () => TZ,
    getEffectiveUser: () => ({ getEmail: () => process.env.GOOGLE_IMPERSONATE_USER || "" }),
};
const MailApp = {
    sendEmail(opts) {
        const from = process.env.GOOGLE_IMPERSONATE_USER;
        if (!from) throw new Error("GOOGLE_IMPERSONATE_USER fehlt (Absender für E-Mails)");
        call("mail.send", { from, to: opts.to, subject: opts.subject, body: opts.body });
    },
};
const ScriptApp = {
    getProjectTriggers: () => [],
    deleteTrigger() {},
    newTrigger() { throw new Error("Trigger werden im GitHub-Runner über Workflow-Schedules abgebildet"); },
};
const DriveApp = {
    getFoldersByName(name) {
        const list = call("drive.findFolders", { name });
        let i = 0;
        const mkFolder = (f) => ({
            getName: () => f.name,
            getFiles() {
                const files = call("drive.listFiles", { folderId: f.id });
                let j = 0;
                return {
                    hasNext: () => j < files.length,
                    next() {
                        const file = files[j++];
                        return {
                            getName: () => file.name,
                            getBlob: () => ({ getDataAsString: () => call("drive.download", { fileId: file.id }) }),
                        };
                    },
                };
            },
        });
        return { hasNext: () => i < list.length, next: () => mkFolder(list[i++]) };
    },
};
const ContentService = {
    MimeType: { JSON: "JSON" },
    createTextOutput: () => ({ setMimeType() { return this; } }),
};

// ---------- SpreadsheetApp ----------
const isEmpty = (v) => v === "" || v === null || v === undefined;
const colLetter = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const a1 = (sheet, r, c, nr, nc) => `'${sheet.name.replace(/'/g, "''")}'!${colLetter(c)}${r}:${colLetter(c + nc - 1)}${r + nr - 1}`;
const toApiValue = (v) => (v instanceof Date ? dateToSerial(v) : v === null || v === undefined ? "" : v);

class Spreadsheet {
    constructor(id) {
        this.id = id;
        this.sheets = [];
        this.queue = [];
        const raw = call("sheets.load", { spreadsheetId: id });
        for (const s of raw) this.sheets.push(this._parse(s));
    }
    _parse(s) {
        const gp = s.properties.gridProperties || {};
        const rows = [];
        const rowData = (s.data && s.data[0] && s.data[0].rowData) || [];
        for (const rd of rowData) {
            rows.push((rd.values || []).map((cell) => {
                const ev = cell.effectiveValue;
                if (!ev) return "";
                if (ev.numberValue !== undefined) {
                    const t = cell.effectiveFormat && cell.effectiveFormat.numberFormat && cell.effectiveFormat.numberFormat.type;
                    return t === "DATE" || t === "DATE_TIME" ? serialToDate(ev.numberValue) : ev.numberValue;
                }
                if (ev.boolValue !== undefined) return ev.boolValue;
                if (ev.stringValue !== undefined) return ev.stringValue;
                if (ev.errorValue) return "#ERROR";
                return "";
            }));
        }
        return new Sheet(this, s.properties.sheetId, s.properties.title, rows, gp.rowCount || 1000, gp.columnCount || 26, gp.frozenRowCount || 0);
    }
    getSheetByName(name) { return this.sheets.find((s) => s.name === name) || null; }
    getSheets() { return this.sheets; }
    insertSheet(name) {
        this.flush();
        const reply = call("sheets.batch", { spreadsheetId: this.id, requests: [{ addSheet: { properties: { title: name } } }] });
        const p = reply[0].addSheet.properties;
        const sheet = new Sheet(this, p.sheetId, p.title, [], 1000, 26, 0);
        this.sheets.push(sheet);
        return sheet;
    }
    toast() {}
    enqueue(op) { this.queue.push(op); }
    flush() {
        const q = this.queue;
        this.queue = [];
        let i = 0;
        while (i < q.length) {
            if (q[i].t === "values") {
                const data = [];
                while (i < q.length && q[i].t === "values") { data.push({ range: q[i].range, values: q[i].values }); i++; }
                if (DEBUG) console.log(`[sheets] values.batchUpdate (${data.length} Bereiche)`);
                call("values.batchUpdate", { spreadsheetId: this.id, data });
            } else if (q[i].t === "clearValues") {
                const ranges = [];
                while (i < q.length && q[i].t === "clearValues") { ranges.push(q[i].range); i++; }
                call("values.batchClear", { spreadsheetId: this.id, ranges });
            } else {
                const requests = [];
                while (i < q.length && q[i].t === "request") { requests.push(q[i].request); i++; }
                if (DEBUG) console.log(`[sheets] batchUpdate (${requests.length} Requests)`);
                call("sheets.batch", { spreadsheetId: this.id, requests });
            }
        }
    }
}

class Sheet {
    constructor(ss, id, name, rows, rowCount, columnCount, frozen) {
        Object.assign(this, { ss, id, name, rows, rowCount, columnCount, frozen });
    }
    getName() { return this.name; }
    getLastRow() {
        for (let i = this.rows.length - 1; i >= 0; i--) if (this.rows[i].some((v) => !isEmpty(v))) return i + 1;
        return 0;
    }
    getLastColumn() {
        let max = 0;
        for (const r of this.rows) for (let c = r.length - 1; c >= max; c--) if (!isEmpty(r[c])) { max = c + 1; break; }
        return max;
    }
    getRange(r, c, nr, nc) { return new Range(this, r, c, nr === undefined ? 1 : nr, nc === undefined ? 1 : nc); }
    getDataRange() { return new Range(this, 1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)); }
    appendRow(values) { this.getRange(this.getLastRow() + 1, 1, 1, values.length).setValues([values]); }
    deleteRow(n) {
        this.rows.splice(n - 1, 1);
        this.rowCount = Math.max(1, this.rowCount - 1);
        this.ss.enqueue({ t: "request", request: { deleteDimension: { range: { sheetId: this.id, dimension: "ROWS", startIndex: n - 1, endIndex: n } } } });
    }
    clear() {
        this.rows = [];
        this.ss.enqueue({ t: "request", request: { updateCells: { range: { sheetId: this.id }, fields: "userEnteredValue,userEnteredFormat" } } });
    }
    clearContents() {
        this.rows = [];
        this.ss.enqueue({ t: "clearValues", range: `'${this.name.replace(/'/g, "''")}'` });
        return this;
    }
    setFrozenRows(n) {
        this.frozen = n;
        this.ss.enqueue({ t: "request", request: { updateSheetProperties: { properties: { sheetId: this.id, gridProperties: { frozenRowCount: n } }, fields: "gridProperties.frozenRowCount" } } });
    }
    getFrozenRows() { return this.frozen; }
    _ensureGrid(rows, cols) {
        if (rows > this.rowCount) {
            this.ss.enqueue({ t: "request", request: { appendDimension: { sheetId: this.id, dimension: "ROWS", length: rows - this.rowCount } } });
            this.rowCount = rows;
        }
        if (cols > this.columnCount) {
            this.ss.enqueue({ t: "request", request: { appendDimension: { sheetId: this.id, dimension: "COLUMNS", length: cols - this.columnCount } } });
            this.columnCount = cols;
        }
    }
}

class Range {
    constructor(sheet, r, c, nr, nc) { Object.assign(this, { sheet, r, c, nr, nc }); }
    getValues() {
        const out = [];
        for (let i = 0; i < this.nr; i++) {
            const row = this.sheet.rows[this.r - 1 + i] || [];
            const line = [];
            for (let j = 0; j < this.nc; j++) { const v = row[this.c - 1 + j]; line.push(v === undefined ? "" : v); }
            out.push(line);
        }
        return out;
    }
    getValue() { return this.getValues()[0][0]; }
    setValues(values) {
        const sh = this.sheet;
        const nr = values.length, nc = nr ? values[0].length : 0;
        if (!nr || !nc) return this;
        sh._ensureGrid(this.r + nr - 1, this.c + nc - 1);
        for (let i = 0; i < nr; i++) {
            const idx = this.r - 1 + i;
            while (sh.rows.length <= idx) sh.rows.push([]);
            const row = sh.rows[idx];
            for (let j = 0; j < nc; j++) {
                while (row.length < this.c - 1 + j) row.push("");
                row[this.c - 1 + j] = values[i][j] === undefined || values[i][j] === null ? "" : values[i][j];
            }
        }
        sh.ss.enqueue({ t: "values", range: a1(sh, this.r, this.c, nr, nc), values: values.map((row) => row.map(toApiValue)) });
        return this;
    }
    setValue(v) {
        const vals = [];
        for (let i = 0; i < this.nr; i++) vals.push(new Array(this.nc).fill(v));
        return this.setValues(vals);
    }
    clearContents() {
        const sh = this.sheet;
        for (let i = 0; i < this.nr; i++) {
            const row = sh.rows[this.r - 1 + i];
            if (row) for (let j = 0; j < this.nc; j++) if (this.c - 1 + j < row.length) row[this.c - 1 + j] = "";
        }
        sh.ss.enqueue({ t: "clearValues", range: a1(sh, this.r, this.c, this.nr, this.nc) });
        return this;
    }
    setFontWeight(weight) {
        const sh = this.sheet;
        sh.ss.enqueue({ t: "request", request: { repeatCell: {
            range: { sheetId: sh.id, startRowIndex: this.r - 1, endRowIndex: this.r - 1 + this.nr, startColumnIndex: this.c - 1, endColumnIndex: this.c - 1 + this.nc },
            cell: { userEnteredFormat: { textFormat: { bold: weight === "bold" } } },
            fields: "userEnteredFormat.textFormat.bold",
        } } });
        return this;
    }
}

let activeSpreadsheet = null;
export function getActiveSpreadsheet() { return activeSpreadsheet; }
export function flushAll() { if (activeSpreadsheet) activeSpreadsheet.flush(); }

const SpreadsheetApp = {
    getActiveSpreadsheet() {
        if (!activeSpreadsheet) {
            const id = String(process.env.SPREADSHEET_ID || "").trim();
            if (!id) throw new Error("SPREADSHEET_ID fehlt");
            activeSpreadsheet = new Spreadsheet(id);
        }
        return activeSpreadsheet;
    },
    getActive() { return this.getActiveSpreadsheet(); },
    getUi() { throw new Error("SpreadsheetApp.getUi() ist im GitHub-Runner nicht verfügbar (interaktive Funktion)"); },
};

export const globals = { Logger, PropertiesService, UrlFetchApp, Utilities, Session, MailApp, ScriptApp, DriveApp, ContentService, SpreadsheetApp };
export { logLines };
