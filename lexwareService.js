// ============================================================
// Lexware Office API Integration
// API Docs: https://developers.lexware.io/docs/public-api/
// Auth:     ****** stored in Script Property LEXWARE_API_KEY
// ============================================================

var LEXWARE_BASE_URL = "https://api.lexware.io/v1";
var LEXWARE_DEFAULT_SHEET_NAME = "Lexware";

// ---- Config ------------------------------------------------

function getLexwareConfig() {
    var props = PropertiesService.getScriptProperties();
    var apiKey = (props.getProperty("LEXWARE_API_KEY") || "").trim();
    var sheetName = (props.getProperty("LEXWARE_SHEET_NAME") || LEXWARE_DEFAULT_SHEET_NAME).trim();
    return {
        apiKey: apiKey,
        sheetName: sheetName || LEXWARE_DEFAULT_SHEET_NAME
    };
}

function validateLexwareConfig() {
    var config = getLexwareConfig();
    if (!config.apiKey) {
        throw new Error("Missing Script Property: LEXWARE_API_KEY");
    }
    return config;
}

// ---- HTTP helper -------------------------------------------

function lexwareRequest(path, queryParams) {
    var config = validateLexwareConfig();
    var url = LEXWARE_BASE_URL + path;

    if (queryParams) {
        var parts = [];
        Object.keys(queryParams).forEach(function (key) {
            var val = queryParams[key];
            if (val === undefined || val === null || val === "") return;
            parts.push(encodeURIComponent(key) + "=" + encodeURIComponent(String(val)));
        });
        if (parts.length > 0) url += "?" + parts.join("&");
    }

    var options = {
        method: "get",
        muteHttpExceptions: true,
        headers: {
            "Authorization": "Bearer " + config.apiKey,
            "Accept": "application/json"
        }
    };

    var response = UrlFetchApp.fetch(url, options);
    var status = response.getResponseCode();
    var bodyText = response.getContentText() || "";

    var body;
    try {
        body = bodyText ? JSON.parse(bodyText) : null;
    } catch (e) {
        body = bodyText;
    }

    if (status < 200 || status >= 300) {
        throw new Error("Lexware request failed (" + status + ") for " + url + ": " + bodyText);
    }

    return { status: status, body: body };
}

// ---- API calls ---------------------------------------------

/**
 * Fetches a page of voucher metadata from GET /v1/voucherlist.
 *
 * The /v1/vouchers endpoint is a single-voucher lookup (requires voucherNumber).
 * /v1/voucherlist is the correct endpoint for paginated listing.
 * Both voucherType and voucherStatus are required by the API.
 *
 * @param {string|null} voucherType - e.g. "invoice", "purchaseinvoice", or null/"any" for all types.
 * @param {number} page - 0-based page index.
 * @param {number} pageSize - items per page (max 250).
 */
function lexwareGetVoucherlist_(voucherType, page, pageSize) {
    var params = {
        voucherType: voucherType || "any",
        voucherStatus: "any",
        page: page || 0,
        size: pageSize || 100
    };
    return lexwareRequest("/voucherlist", params);
}

function lexwareHealthCheck() {
    return lexwareRequest("/ping");
}

// ---- Import ------------------------------------------------

/**
 * Fetches all Lexware invoices (outgoing, voucherType=invoice) and writes
 * them to the "Lexware" sheet (or the sheet configured via the
 * LEXWARE_SHEET_NAME script property). Uses the /v1/voucherlist endpoint.
 */
function importLexwareToSheet() {
    var config = getLexwareConfig();
    return lexwareImportVouchersToSheet_("invoice", config.sheetName);
}

// ---- Voucher import (Einnahmen / Ausgaben / Umsätze) -------

var LEXWARE_VOUCHER_HEADERS = [
    "ID",
    "Belegtyp",
    "Status",
    "Belegnummer",
    "Belegdatum",
    "Fälligkeitsdatum",
    "Kontakt",
    "Gesamtbetrag",
    "Währung",
    "Bemerkung"
];

/**
 * Fetches all pages of vouchers (optionally filtered by voucherType) and
 * upserts them into the given sheet. Rows are matched by voucher ID.
 *
 * @param {string|null} voucherType - API filter value ("invoice", "purchaseinvoice", …)
 *                                    or null/undefined to fetch all types.
 * @param {string} sheetName - target sheet name (created if absent).
 * @return {{ok:boolean, sheet:string, total:number, inserted:number, updated:number}}
 */
function lexwareImportVouchersToSheet_(voucherType, sheetName) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("No active spreadsheet");

    var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

    if (sheet.getLastRow() === 0) {
        sheet.appendRow(LEXWARE_VOUCHER_HEADERS);
        sheet.getRange(1, 1, 1, LEXWARE_VOUCHER_HEADERS.length).setFontWeight("bold");
    }

    // Index existing rows by voucher ID (column 1)
    var existingById = {};
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        var existingData = sheet.getRange(2, 1, lastRow - 1, LEXWARE_VOUCHER_HEADERS.length).getValues();
        existingData.forEach(function (row, idx) {
            var id = String(row[0] || "").trim();
            if (id) existingById[id] = { rowIndex: idx + 2, data: row };
        });
    }

    // Fetch all pages
    var allVouchers = [];
    var page = 0;
    var pageSize = 100;
    var totalPages = 1;

    do {
        var result = lexwareGetVoucherlist_(voucherType || null, page, pageSize);
        var body = result.body;

        if (!body || !body.content) {
            Logger.log("Lexware voucherlist: unexpected response on page " + page + ": " + JSON.stringify(body));
            break;
        }

        allVouchers = allVouchers.concat(body.content);

        // The voucherlist endpoint returns pagination at the top level;
        // some older endpoints wrap it in a nested "page" object.
        var pageInfo = body.page || {};
        totalPages = pageInfo.totalPages !== undefined ? pageInfo.totalPages
                   : (body.totalPages !== undefined ? body.totalPages : 1);
        page++;
    } while (page < totalPages);

    Logger.log(
        "Lexware voucherlist (" + (voucherType || "all") + "): fetched " +
        allVouchers.length + " records across " + totalPages + " page(s)"
    );

    var newRows = [];
    var updatedCount = 0;

    allVouchers.forEach(function (v) {
        var id = String(v.id || "").trim();
        if (!id) return;

        var row = [
            id,
            v.voucherType || "",
            v.voucherStatus || "",
            v.voucherNumber || "",
            v.voucherDate ? String(v.voucherDate).slice(0, 10) : "",
            v.dueDate ? String(v.dueDate).slice(0, 10) : "",
            v.contactName || "",
            v.totalAmount !== undefined ? v.totalAmount : "",
            v.currency || "EUR",
            v.remark || ""
        ];

        if (existingById[id]) {
            var existing = existingById[id].data;
            var changed = row.some(function (val, i) { return String(val) !== String(existing[i]); });
            if (changed) {
                sheet.getRange(existingById[id].rowIndex, 1, 1, row.length).setValues([row]);
                updatedCount++;
            }
        } else {
            newRows.push(row);
        }
    });

    if (newRows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    }

    Logger.log(
        "Lexware voucherlist (" + (voucherType || "all") + ") import complete: " +
        "total=" + allVouchers.length +
        ", inserted=" + newRows.length +
        ", updated=" + updatedCount
    );

    return {
        ok: true,
        sheet: sheetName,
        total: allVouchers.length,
        inserted: newRows.length,
        updated: updatedCount
    };
}

/**
 * Imports income vouchers (Einnahmen) into the "Lexware_Einnahmen" sheet.
 * Uses GET /v1/voucherlist?voucherType=invoice&voucherStatus=any.
 * Override the sheet name with the script property LEXWARE_EINNAHMEN_SHEET_NAME.
 */
function importLexwareEinnahmen() {
    var props = PropertiesService.getScriptProperties();
    var sheetName = (props.getProperty("LEXWARE_EINNAHMEN_SHEET_NAME") || "Lexware_Einnahmen").trim();
    return lexwareImportVouchersToSheet_("invoice", sheetName);
}

/**
 * Imports expense vouchers (Ausgaben) into the "Lexware_Ausgaben" sheet.
 * Uses GET /v1/voucherlist?voucherType=purchaseinvoice&voucherStatus=any.
 * Override the sheet name with the script property LEXWARE_AUSGABEN_SHEET_NAME.
 */
function importLexwareAusgaben() {
    var props = PropertiesService.getScriptProperties();
    var sheetName = (props.getProperty("LEXWARE_AUSGABEN_SHEET_NAME") || "Lexware_Ausgaben").trim();
    return lexwareImportVouchersToSheet_("purchaseinvoice", sheetName);
}

/**
 * Imports all vouchers (Umsätze) into the "Lexware_Umsaetze" sheet.
 * Uses GET /v1/voucherlist?voucherType=any&voucherStatus=any.
 * Override the sheet name with the script property LEXWARE_UMSAETZE_SHEET_NAME.
 */
function importLexwareUmsaetze() {
    var props = PropertiesService.getScriptProperties();
    var sheetName = (props.getProperty("LEXWARE_UMSAETZE_SHEET_NAME") || "Lexware_Umsaetze").trim();
    return lexwareImportVouchersToSheet_(null, sheetName);
}

// ---- File upload -------------------------------------------

/**
 * Uploads a Blob as a file to Lexware Office via POST /v1/files-api/files.
 * Returns { ok: true, fileId: string, body: object }.
 *
 * @param {Blob}   blob      The file blob to upload.
 * @param {string} fileName  The file name to use for the upload (e.g. "Rechnung_123.pdf").
 *                           Falls back to the blob's existing name when omitted.
 */
function lexwareUploadFile_(blob, fileName) {
    var config = validateLexwareConfig();
    var url = LEXWARE_BASE_URL + "/files-api/files";

    var namedBlob = blob.setName(fileName || blob.getName() || "upload.pdf");

    var options = {
        method: "post",
        payload: { file: namedBlob },
        muteHttpExceptions: true,
        headers: {
            "Authorization": "Bearer " + config.apiKey,
            "Accept": "application/json"
        }
    };

    var response = UrlFetchApp.fetch(url, options);
    var status = response.getResponseCode();
    var bodyText = response.getContentText() || "";

    var body;
    try {
        body = bodyText ? JSON.parse(bodyText) : null;
    } catch (e) {
        body = bodyText;
    }

    if (status < 200 || status >= 300) {
        throw new Error("Lexware file upload failed (" + status + ") for " + url + ": " + bodyText);
    }

    var fileId = String(
        (body && (body.id || body.fileId || body.file_id || body.uuid)) || ""
    );

    Logger.log("Lexware file upload: status=" + status + ", fileId=" + fileId);

    return { ok: true, status: status, fileId: fileId, body: body };
}

// ---- All imports -------------------------------------------

/**
 * Runs all four Lexware imports:
 *   1. importLexwareToSheet()   – outgoing invoices (Rechnungen)
 *   2. importLexwareEinnahmen() – income vouchers (Einnahmen)
 *   3. importLexwareAusgaben()  – expense vouchers (Ausgaben)
 *   4. importLexwareUmsaetze()  – all vouchers (Umsätze)
 */
function importLexwareAll() {
    importLexwareToSheet();
    importLexwareEinnahmen();
    importLexwareAusgaben();
    importLexwareUmsaetze();
}
