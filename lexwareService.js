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

function getRetryAfterMs_(headers) {
    if (!headers) return 0;
    var retryAfter = headers["Retry-After"];
    if (retryAfter === undefined || retryAfter === null) retryAfter = headers["retry-after"];
    if (Array.isArray(retryAfter)) retryAfter = retryAfter[0];
    var seconds = Number(retryAfter);
    return seconds > 0 ? Math.round(seconds * 1000) : 0;
}

function isLexware404_(error) {
    return !!(error && error.message && /\(404\)/.test(String(error.message)));
}

function lexwareRequest(path, queryParams, baseUrl) {
    var config = validateLexwareConfig();
    var resolvedBaseUrl = baseUrl || LEXWARE_BASE_URL;
    var url = resolvedBaseUrl + path;

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

    var response;
    var status;
    var bodyText = "";
    var maxAttempts = 4;
    for (var attempt = 0; attempt < maxAttempts; attempt++) {
        response = UrlFetchApp.fetch(url, options);
        status = response.getResponseCode();
        bodyText = response.getContentText() || "";
        if (status !== 429 || attempt === maxAttempts - 1) break;
        var retryAfterMs = getRetryAfterMs_(response.getAllHeaders());
        var backoffMs = retryAfterMs || Math.min(8000, Math.pow(2, attempt) * 1000);
        Logger.log(
            "Lexware request hit rate limit for " + url +
            " – retry " + (attempt + 1) + "/" + (maxAttempts - 1) +
            " in " + backoffMs + "ms."
        );
        Utilities.sleep(backoffMs);
    }

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

/**
 * Sends a POST or PUT request to the Lexware API with a JSON body.
 *
 * @param {string} path       API path, e.g. "/vouchers".
 * @param {Object} payload    Object to JSON-encode as request body.
 * @param {string} [method]   HTTP method: "post" (default) or "put".
 * @param {string} [baseUrl]  Override for LEXWARE_BASE_URL.
 * @return {{status:number, body:*}}
 */
function lexwarePostRequest_(path, payload, method, baseUrl) {
    var config = validateLexwareConfig();
    var resolvedBaseUrl = baseUrl || LEXWARE_BASE_URL;
    var url = resolvedBaseUrl + path;

    var options = {
        method: method || "post",
        muteHttpExceptions: true,
        contentType: "application/json",
        payload: JSON.stringify(payload),
        headers: {
            "Authorization": "Bearer " + config.apiKey,
            "Accept": "application/json"
        }
    };

    var response;
    var status;
    var bodyText = "";
    var maxAttempts = 4;
    for (var attempt = 0; attempt < maxAttempts; attempt++) {
        response = UrlFetchApp.fetch(url, options);
        status = response.getResponseCode();
        bodyText = response.getContentText() || "";
        if (status !== 429 || attempt === maxAttempts - 1) break;
        var retryAfterMs = getRetryAfterMs_(response.getAllHeaders());
        var backoffMs = retryAfterMs || Math.min(8000, Math.pow(2, attempt) * 1000);
        Logger.log(
            "Lexware POST request hit rate limit for " + url +
            " – retry " + (attempt + 1) + "/" + (maxAttempts - 1) +
            " in " + backoffMs + "ms."
        );
        Utilities.sleep(backoffMs);
    }

    var body;
    try {
        body = bodyText ? JSON.parse(bodyText) : null;
    } catch (e) {
        body = bodyText;
    }

    if (status < 200 || status >= 300) {
        throw new Error(
            "Lexware POST request failed (" + status + ") for " + url + ": " + bodyText
        );
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
    return lexwareImportVouchersToSheet_("salesinvoice", config.sheetName);
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
 * Uses GET /v1/voucherlist?voucherType=salesinvoice&voucherStatus=any.
 * Override the sheet name with the script property LEXWARE_EINNAHMEN_SHEET_NAME.
 */
function importLexwareEinnahmen() {
    var props = PropertiesService.getScriptProperties();
    var sheetName = (props.getProperty("LEXWARE_EINNAHMEN_SHEET_NAME") || "Lexware_Einnahmen").trim();
    return lexwareImportVouchersToSheet_("salesinvoice", sheetName);
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

// ---- Konto-Zuordnung (account mapping for Umsätze) ---------

var LEXWARE_KONTO_ZUORDNUNG_DEFAULT_SHEET_NAME = "Lexware_Konto_Zuordnung";

var LEXWARE_KONTO_ZUORDNUNG_HEADERS = [
    "Lieferantennummer",  // A  Lieferantennummer des Belegs (optional, wenn nur Kategorie gesetzt)
    "Kategorie",          // B  Buchungskategoriename aus Lexware (optional, wenn nur Lieferantennummer gesetzt)
    "Konto"               // C  Kontonummer oder Kontobezeichnung
];

/**
 * Erstellt das Tabellenblatt "Lexware_Konto_Zuordnung" mit den
 * erforderlichen Spaltenüberschriften, falls es noch nicht existiert.
 * Bestehende Daten (Zuordnungen) bleiben erhalten.
 *
 * Pflege:
 *   - Spalte A (Lieferantennummer): Lieferantennummer des Belegs.
 *   - Spalte B (Kategorie): exakter Buchungskategoriename aus Lexware.
 *   - Spalte C (Konto): gewünschte Kontobezeichnung.
 *
 * Priorität bei der Zuordnung:
 *   1. Kombination Lieferantennummer + Kategorie (beide Spalten A und B gefüllt) – spezifischste Regel
 *   2. Nur Lieferantennummer (Spalte A gesetzt, Spalte B leer)
 *   3. Nur Kategorie (Spalte B gesetzt, Spalte A leer)
 *
 * Override für den Blattnamen: Script Property LEXWARE_KONTO_ZUORDNUNG_SHEET_NAME
 *
 * @return {{ok:boolean, sheet:string}}
 */
function setupLexwareKontoZuordnungSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("No active spreadsheet");
    var props = PropertiesService.getScriptProperties();
    var sheetName = (props.getProperty("LEXWARE_KONTO_ZUORDNUNG_SHEET_NAME") ||
        LEXWARE_KONTO_ZUORDNUNG_DEFAULT_SHEET_NAME).trim();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        Logger.log("Konto-Zuordnung: Blatt '" + sheetName + "' erstellt.");
    }
    if (sheet.getLastRow() === 0) {
        sheet.appendRow(LEXWARE_KONTO_ZUORDNUNG_HEADERS);
        sheet.getRange(1, 1, 1, LEXWARE_KONTO_ZUORDNUNG_HEADERS.length).setFontWeight("bold");
        Logger.log("Konto-Zuordnung: Spaltenüberschriften gesetzt.");
    } else {
        Logger.log("Konto-Zuordnung: Blatt existiert bereits, keine Änderung an der Kopfzeile.");
    }
    return { ok: true, sheet: sheetName };
}

/**
 * Liest das Konto-Zuordnung-Sheet und gibt ein Objekt mit Lookup-Maps zurück.
 *
 * Lookup-Strategien (in Prioritätsreihenfolge):
 *   1. composite["lieferantennummer|kategorie"] → konto  (Spalten A + B beide gefüllt)
 *   2. vendorOnly["lieferantennummer"]          → konto  (nur Spalte A gefüllt, B leer)
 *   3. category["kategorie.toLowerCase()"]      → konto  (nur Spalte B gefüllt, A leer)
 *
 * @return {{composite:Object, vendorOnly:Object, category:Object}}
 */
function buildKontoZuordnungIndex_() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var props = PropertiesService.getScriptProperties();
    var sheetName = (props.getProperty("LEXWARE_KONTO_ZUORDNUNG_SHEET_NAME") ||
        LEXWARE_KONTO_ZUORDNUNG_DEFAULT_SHEET_NAME).trim();
    var sheet = ss.getSheetByName(sheetName);
    var result = { composite: {}, vendorOnly: {}, category: {} };
    if (!sheet || sheet.getLastRow() < 2) return result;
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    data.forEach(function (row) {
        var lieferantennummer = String(row[0] || "").trim();
        var kategorie         = String(row[1] || "").trim();
        var konto             = String(row[2] || "").trim();
        if (!konto) return;
        if (lieferantennummer && kategorie) {
            result.composite[lieferantennummer + "|" + kategorie.toLowerCase()] = konto;
        } else if (lieferantennummer) {
            result.vendorOnly[lieferantennummer] = konto;
        } else if (kategorie) {
            result.category[kategorie.toLowerCase()] = konto;
        }
    });
    return result;
}

function extractLieferantennummer_(voucherSummary, voucherDetail) {
    var candidates = [];

    function pushCandidate(value) {
        if (value === null || value === undefined) return;
        var trimmed = String(value).trim();
        if (trimmed) candidates.push(trimmed);
    }

    function pushVendorRoleNumbers(contact) {
        if (!contact || !contact.roles || !contact.roles.vendor) return;
        var vendorRole = contact.roles.vendor;
        pushCandidate(vendorRole.number);
        pushCandidate(vendorRole.vendorNumber);
        pushCandidate(vendorRole.contactNumber);
    }

    pushVendorRoleNumbers(voucherDetail && voucherDetail.contact);
    pushVendorRoleNumbers(voucherSummary && voucherSummary.contact);

    pushCandidate(voucherDetail && voucherDetail.vendorNumber);
    pushCandidate(voucherDetail && voucherDetail.contactNumber);
    pushCandidate(voucherSummary && voucherSummary.vendorNumber);
    pushCandidate(voucherSummary && voucherSummary.contactNumber);

    return candidates.length ? candidates[0] : "";
}

function resolveLexwareContactId_(voucherSummary, voucherDetail) {
    var candidates = [
        voucherSummary && voucherSummary.contactId,
        voucherSummary && voucherSummary.contact && voucherSummary.contact.id,
        voucherSummary && voucherSummary.contact && voucherSummary.contact.contactId,
        voucherDetail && voucherDetail.contactId,
        voucherDetail && voucherDetail.contact && voucherDetail.contact.id,
        voucherDetail && voucherDetail.contact && voucherDetail.contact.contactId
    ];
    for (var i = 0; i < candidates.length; i++) {
        var id = String(candidates[i] || "").trim();
        if (id) return id;
    }
    return "";
}

// ---- Umsätze with line items --------------------------------

var LEXWARE_UMSAETZE_DETAIL_HEADERS = [
    "Zeilen_ID",          // A  Eindeutiger Schlüssel: {voucherID}_{posNr}
    "Beleg_ID",           // B  Lexware-UUID des Belegs
    "Belegtyp",           // C
    "Status",             // D
    "Belegnummer",        // E
    "Belegdatum",         // F
    "Fälligkeitsdatum",   // G
    "Kontakt",            // H
    "Lieferantennummer",  // I  Lieferantennummer des Belegs (contact.roles.vendor.number)
    "Gesamtbetrag",       // J  Gesamtbetrag des Belegs
    "Währung",            // K
    "Bemerkung",          // L
    "Position",           // M  Positionsnummer (1, 2, 3 …)
    "Pos_Kategorie",      // N  Buchungskategoriename dieser Position
    "Pos_Betrag_Brutto",  // O  Bruttobetrag dieser Position
    "Pos_MwSt_Satz",      // P  Steuersatz % dieser Position
    "Pos_MwSt_Betrag",    // Q  Steuerbetrag dieser Position
    "Konto"               // R  Konto aus dem Blatt Lexware_Konto_Zuordnung
];

/**
 * Liest das Lexware_Kategorien-Sheet und gibt eine Lookup-Map zurück.
 *
 * @return {Object}  { categoryId → categoryName }
 */
function buildKategorienIdToNameIndex_() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var props = PropertiesService.getScriptProperties();
    var sheetName = (props.getProperty("LEXWARE_KATEGORIEN_SHEET_NAME") ||
        LEXWARE_KATEGORIEN_SHEET_NAME).trim();
    var sheet = ss.getSheetByName(sheetName);
    var index = {};
    if (!sheet || sheet.getLastRow() < 2) return index;
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    data.forEach(function (row) {
        var id = String(row[0] || "").trim();
        var name = String(row[1] || "").trim();
        if (id) index[id] = name;
    });
    return index;
}

/**
 * Ruft die vollständigen Belegdetails (inkl. voucherItems) von GET /vouchers/{id} ab.
 *
 * @param {string} voucherId  Lexware-UUID des Belegs.
 * @return {{status:number, body:*}}
 */
function lexwareGetVoucherDetail_(voucherId) {
    return lexwareRequest("/vouchers/" + encodeURIComponent(voucherId));
}

/**
 * Holt alle Belege (optional gefiltert nach voucherType), ruft für jeden Beleg
 * die vollständigen Details inkl. Positionen (voucherItems) ab und schreibt
 * eine Zeile pro Position in das Zielblatt.
 *
 * Zeilenkennung (Spalte A "Zeilen_ID"): "{voucherId}_{posNr}"
 * Die Konto-Spalte wird aus dem Blatt "Lexware_Konto_Zuordnung" befüllt
 * (Buchungskategoriename → Konto).
 *
 * Falls ein Beleg keine Positionen liefert, wird eine zusammenfassende
 * Zeile mit dem Gesamtbetrag eingefügt.
 *
 * Beim ersten Lauf nach einer Formatumstellung (altes Format mit Spalte A = "ID")
 * wird das Sheet automatisch geleert und neu aufgebaut.
 *
 * @param {string|null} voucherType  API-Filterwert oder null für alle Typen.
 * @param {string} sheetName         Name des Zielblatts.
 * @return {{ok:boolean, sheet:string, total:number, inserted:number, updated:number}}
 */
function lexwareImportVouchersWithLineItemsToSheet_(voucherType, sheetName) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("No active spreadsheet");

    var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

    // Migrate from old single-row format (column A header = "ID")
    var firstHeader = sheet.getLastRow() > 0 ? String(sheet.getRange(1, 1).getValue() || "").trim() : "";
    if (sheet.getLastRow() === 0 || firstHeader === "ID") {
        sheet.clearContents();
        sheet.appendRow(LEXWARE_UMSAETZE_DETAIL_HEADERS);
        sheet.getRange(1, 1, 1, LEXWARE_UMSAETZE_DETAIL_HEADERS.length).setFontWeight("bold");
        if (firstHeader === "ID") {
            Logger.log("Lexware Umsätze: altes Format erkannt – Blatt '" + sheetName + "' neu aufgebaut.");
        }
    }

    // Build lookup tables
    var kontoZuordnung = buildKontoZuordnungIndex_();
    var kategorienIndex = buildKategorienIdToNameIndex_();
    // UUID → Lieferantennummer (from Lexware Kunden sheet; voucher API does not include vendor role)
    var contactIdToVendorNumber = buildContactIdToVendorNumberIndex_();

    // Index existing rows by Zeilen_ID (column A)
    var existingById = {};
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        var existingData = sheet.getRange(2, 1, lastRow - 1, LEXWARE_UMSAETZE_DETAIL_HEADERS.length).getValues();
        existingData.forEach(function (row, idx) {
            var id = String(row[0] || "").trim();
            if (id) existingById[id] = { rowIndex: idx + 2, data: row };
        });
    }

    // Fetch voucherlist (all pages)
    var allVouchers = [];
    var page = 0;
    var pageSize = 100;
    var totalPages = 1;

    do {
        var listResult = lexwareGetVoucherlist_(voucherType || null, page, pageSize);
        var listBody = listResult.body;
        if (!listBody || !listBody.content) {
            Logger.log("Lexware voucherlist (details): unerwartete Antwort auf Seite " + page +
                ": " + JSON.stringify(listBody));
            break;
        }
        allVouchers = allVouchers.concat(listBody.content);
        var pageInfo = listBody.page || {};
        totalPages = pageInfo.totalPages !== undefined ? pageInfo.totalPages
                   : (listBody.totalPages !== undefined ? listBody.totalPages : 1);
        page++;
    } while (page < totalPages);

    Logger.log(
        "Lexware Umsätze (details) (" + (voucherType || "all") + "): " +
        allVouchers.length + " Belege auf " + totalPages + " Seite(n) geholt"
    );

    var newRows = [];
    var updatedCount = 0;

    allVouchers.forEach(function (v) {
        var voucherId = String(v.id || "").trim();
        if (!voucherId) return;

        // Summary fields available from voucherlist
        var belegtyp          = v.voucherType || "";
        var status            = v.voucherStatus || "";
        var belegnummer       = v.voucherNumber || "";
        var belegdatum        = v.voucherDate ? String(v.voucherDate).slice(0, 10) : "";
        var faelligkeitsdatum = v.dueDate ? String(v.dueDate).slice(0, 10) : "";
        var kontakt           = v.contactName || "";
        var gesamtbetrag      = v.totalAmount !== undefined ? v.totalAmount : "";
        var waehrung          = v.currency || "EUR";
        var bemerkung         = v.remark || "";

        // Resolve Lieferantennummer from the voucherlist summary using the pre-built
        // Kunden index (UUID → Lieferantennummer). No extra API request is needed here.
        var contactId = resolveLexwareContactId_(v, null);
        var lieferantennummer = contactId && contactIdToVendorNumber[contactId]
            ? contactIdToVendorNumber[contactId]
            : extractLieferantennummer_(v, null);

        // Fetch full voucher detail to get line items only
        var voucherItems = [];
        try {
            var detailResult = lexwareGetVoucherDetail_(voucherId);
            var detail = detailResult.body;
            if (detail && Array.isArray(detail.voucherItems)) {
                voucherItems = detail.voucherItems;
            } else if (detail && Array.isArray(detail.lineItems)) {
                voucherItems = detail.lineItems;
            }
            // If Lieferantennummer not found via summary index, try the detail response as fallback
            if (!lieferantennummer) {
                var detailContactId = resolveLexwareContactId_(null, detail);
                if (detailContactId && contactIdToVendorNumber[detailContactId]) {
                    lieferantennummer = contactIdToVendorNumber[detailContactId];
                } else {
                    lieferantennummer = extractLieferantennummer_(v, detail);
                }
            }
        } catch (e) {
            Logger.log("Lexware Umsätze: Detail-Abruf für Beleg " + voucherId + " fehlgeschlagen: " + e.message);
        }
        // Courtesy pause between individual detail requests to avoid rate limiting
        Utilities.sleep(300);

        // If the detail endpoint returned no items, fall back to a single summary row
        if (voucherItems.length === 0) {
            voucherItems = [{ _summaryFallback: true }];
        }

        voucherItems.forEach(function (item, idx) {
            var posNr  = idx + 1;
            var rowKey = voucherId + "_" + posNr;

            // Resolve category name: prefer direct field, then ID lookup
            var categoryName = "";
            if (!item._summaryFallback) {
                categoryName = String(
                    item.categoryName || item.name || item.description || ""
                ).trim();
                if (!categoryName) {
                    var categoryId = String(item.categoryId || "").trim();
                    if (categoryId) categoryName = kategorienIndex[categoryId] || "";
                }
            }

            var posGross = item._summaryFallback ? gesamtbetrag
                         : (item.amount !== undefined ? item.amount : "");
            var posRate  = item._summaryFallback ? "" : (item.taxRatePercent !== undefined ? item.taxRatePercent : "");
            var posTax   = item._summaryFallback ? "" : (item.taxAmount !== undefined ? item.taxAmount : "");

            // Look up account from Konto-Zuordnung with priority:
            //   1. Lieferantennummer + Kategorie (composite key)
            //   2. Nur Lieferantennummer
            //   3. Nur Kategorie
            //   4. Fallback "Mietenkonto"
            var konto = "Mietenkonto";
            var categoryKey = categoryName ? categoryName.toLowerCase() : "";
            if (lieferantennummer && categoryKey &&
                kontoZuordnung.composite[lieferantennummer + "|" + categoryKey]) {
                konto = kontoZuordnung.composite[lieferantennummer + "|" + categoryKey];
            } else if (lieferantennummer && kontoZuordnung.vendorOnly[lieferantennummer]) {
                konto = kontoZuordnung.vendorOnly[lieferantennummer];
            } else if (categoryKey && kontoZuordnung.category[categoryKey]) {
                konto = kontoZuordnung.category[categoryKey];
            }

            var row = [
                rowKey,
                voucherId,
                belegtyp,
                status,
                belegnummer,
                belegdatum,
                faelligkeitsdatum,
                kontakt,
                lieferantennummer,
                gesamtbetrag,
                waehrung,
                bemerkung,
                posNr,
                categoryName,
                posGross,
                posRate,
                posTax,
                konto
            ];

            if (existingById[rowKey]) {
                var existing = existingById[rowKey].data;
                var changed = row.some(function (val, i) { return String(val) !== String(existing[i]); });
                if (changed) {
                    sheet.getRange(existingById[rowKey].rowIndex, 1, 1, row.length).setValues([row]);
                    updatedCount++;
                }
            } else {
                newRows.push(row);
            }
        });
    });

    if (newRows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length,
            LEXWARE_UMSAETZE_DETAIL_HEADERS.length).setValues(newRows);
    }

    Logger.log(
        "Lexware Umsätze (details) import abgeschlossen: Belege=" + allVouchers.length +
        ", eingefügt=" + newRows.length +
        ", aktualisiert=" + updatedCount
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
 * Imports all vouchers (Umsätze) into the "Lexware Umsatz Import" sheet.
 * Each voucher line item is written as a separate row. The "Konto" column
 * is filled from the "Lexware_Konto_Zuordnung" mapping sheet.
 * Uses GET /v1/voucherlist?voucherType=any&voucherStatus=any.
 * Override the sheet name with the script property LEXWARE_UMSAETZE_SHEET_NAME.
 */
function importLexwareUmsaetze() {
    var props = PropertiesService.getScriptProperties();
    var sheetName = (props.getProperty("LEXWARE_UMSAETZE_SHEET_NAME") || "Lexware Umsatz Import").trim();
    return lexwareImportVouchersWithLineItemsToSheet_(null, sheetName);
}


// ---- Sheet: Buchungskategorien (posting categories) --------

var LEXWARE_KATEGORIEN_SHEET_NAME = "Lexware_Kategorien";

var LEXWARE_KATEGORIEN_HEADERS = [
    "ID",
    "Name",
    "Typ",
    "API-Typ"
];

/**
 * Maps the raw API type string returned by Lexware to a human-readable
 * German label and a top-level bucket (Einnahmen / Ausgaben / Sonstige).
 *
 * Known Lexware posting-category types (case-insensitive):
 *   revenue / revenues / income / sales / einnahmen  → Einnahmen
 *   expense / expenses / costs  / aufwand / ausgaben → Ausgaben
 *   everything else                                  → Sonstige
 *
 * @param  {string} apiType  Raw type value from the API.
 * @return {string}          "Einnahmen", "Ausgaben" or "Sonstige"
 */
function mapKategorienTyp_(apiType) {
    if (!apiType) return "Sonstige";
    var t = String(apiType).toLowerCase();
    if (/^(revenue|revenues|income|sales|einnahmen|ertrag|ertraege)$/.test(t)) return "Einnahmen";
    if (/^(expense|expenses|costs?|cost|aufwand|aufwendungen|ausgaben)$/.test(t)) return "Ausgaben";
    return "Sonstige";
}

/**
 * Fetches all Lexware posting categories from GET /v1/posting-categories and
 * writes them to the "Lexware_Kategorien" sheet, sorted by Typ then Name.
 * Rows are matched by category ID to enable incremental updates.
 *
 * The Typ column contains "Einnahmen", "Ausgaben" or "Sonstige" so that the
 * sheet can easily be filtered or used as a lookup reference.
 *
 * Override the sheet name via script property LEXWARE_KATEGORIEN_SHEET_NAME.
 *
 * @return {{ok:boolean, sheet:string, total:number, inserted:number, updated:number}}
 */
function importLexwareKategorien() {
    var props = PropertiesService.getScriptProperties();
    var sheetName = (props.getProperty("LEXWARE_KATEGORIEN_SHEET_NAME") || LEXWARE_KATEGORIEN_SHEET_NAME).trim();

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("No active spreadsheet");

    var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

    // Fetch categories
    var categories = [];
    try {
        var result = lexwareRequest("/posting-categories");
        var body = result.body;
        if (Array.isArray(body)) {
            categories = body;
        } else if (body && Array.isArray(body.content)) {
            categories = body.content;
        } else if (body && Array.isArray(body.categories)) {
            categories = body.categories;
        } else if (body) {
            categories = [body];
        }
    } catch (e) {
        Logger.log("Lexware: /posting-categories request failed: " + e.message);
        return { ok: false, sheet: sheetName, total: 0, inserted: 0, updated: 0, error: e.message };
    }

    Logger.log("Lexware posting-categories: fetched " + categories.length + " record(s)");

    // Write header row if the sheet is empty
    if (sheet.getLastRow() === 0) {
        sheet.appendRow(LEXWARE_KATEGORIEN_HEADERS);
        sheet.getRange(1, 1, 1, LEXWARE_KATEGORIEN_HEADERS.length).setFontWeight("bold");
    }

    // Index existing rows by category ID (column 1) for upsert
    var existingById = {};
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        var existingData = sheet.getRange(2, 1, lastRow - 1, LEXWARE_KATEGORIEN_HEADERS.length).getValues();
        existingData.forEach(function (row, idx) {
            var id = String(row[0] || "").trim();
            if (id) existingById[id] = { rowIndex: idx + 2, data: row };
        });
    }

    var newRows = [];
    var updatedCount = 0;

    categories.forEach(function (cat) {
        var id = String(cat.id || "").trim();
        if (!id) return;

        var name = String(cat.name || "").trim();
        var apiType = String(cat.type || cat.categoryType || "").trim();
        var typ = mapKategorienTyp_(apiType);

        var row = [id, name, typ, apiType];

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
        sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, LEXWARE_KATEGORIEN_HEADERS.length).setValues(newRows);
    }

    Logger.log(
        "Lexware Kategorien import complete: total=" + categories.length +
        ", inserted=" + newRows.length +
        ", updated=" + updatedCount
    );

    return {
        ok: true,
        sheet: sheetName,
        total: categories.length,
        inserted: newRows.length,
        updated: updatedCount
    };
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

// ---- Sheet: Zahlungen (payments per voucher) ---------------

var LEXWARE_ZAHLUNGEN_SHEET_NAME = "Lexware_Zahlungen";

var LEXWARE_ZAHLUNGEN_HEADERS = [
    "ID",
    "Roheintrag"
];

/**
 * Fetches payment information for every salesinvoice and purchaseinvoice voucher
 * by calling GET /v1/payments/{voucherId} once per voucher, then writes one
 * raw entry row per voucher into the "Lexware_Zahlungen" sheet:
 *   - column A: voucher ID
 *   - column B: raw payment response as string
 *
 * Override the sheet name via script property LEXWARE_ZAHLUNGEN_SHEET_NAME.
 *
 * @return {{ok:boolean, sheet:string, total:number, inserted:number, updated:number}}
 */
function importLexwarePayments() {
    var props = PropertiesService.getScriptProperties();
    var sheetName = (props.getProperty("LEXWARE_ZAHLUNGEN_SHEET_NAME") || LEXWARE_ZAHLUNGEN_SHEET_NAME).trim();

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("No active spreadsheet");

    var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

    if (sheet.getLastRow() === 0) {
        sheet.appendRow(LEXWARE_ZAHLUNGEN_HEADERS);
        sheet.getRange(1, 1, 1, LEXWARE_ZAHLUNGEN_HEADERS.length).setFontWeight("bold");
    } else {
        var currentColumnCount = sheet.getLastColumn();
        var isExpectedHeader = false;
        if (currentColumnCount === LEXWARE_ZAHLUNGEN_HEADERS.length) {
            var currentHeaders = sheet.getRange(1, 1, 1, LEXWARE_ZAHLUNGEN_HEADERS.length).getValues()[0];
            isExpectedHeader = LEXWARE_ZAHLUNGEN_HEADERS.every(function (h, i) {
                return String(currentHeaders[i] || "").trim() === h;
            });
        }
        if (!isExpectedHeader) {
            Logger.log("Lexware Zahlungen: Schema-Migration erkannt – Blattinhalt wird auf 2-Spalten-Rohformat zurückgesetzt.");
            sheet.clearContents();
            sheet.appendRow(LEXWARE_ZAHLUNGEN_HEADERS);
            sheet.getRange(1, 1, 1, LEXWARE_ZAHLUNGEN_HEADERS.length).setFontWeight("bold");
        }
    }

    var toRawString_ = function (value) {
        if (value === undefined) return "";
        if (typeof value === "string") return value;
        try {
            return JSON.stringify(value);
        } catch (e) {
            return String(value);
        }
    };

    // Index existing rows by ID in column A.
    var existingById = {};
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        var existingData = sheet.getRange(2, 1, lastRow - 1, LEXWARE_ZAHLUNGEN_HEADERS.length).getValues();
        existingData.forEach(function (row, idx) {
            var key = String(row[0] || "").trim();
            if (key) existingById[key] = { rowIndex: idx + 2, data: row };
        });
    }

    // Fetch all vouchers (Einnahmen = salesinvoice, Ausgaben = purchaseinvoice)
    var allVouchers = [];
    ["salesinvoice", "purchaseinvoice"].forEach(function (voucherType) {
        var page = 0;
        var pageSize = 100;
        var totalPages = 1;
        do {
            try {
                var listResult = lexwareGetVoucherlist_(voucherType, page, pageSize);
                var listBody = listResult.body;
                if (!listBody || !listBody.content) break;
                allVouchers = allVouchers.concat(listBody.content);
                var pageInfo = listBody.page || {};
                totalPages = pageInfo.totalPages !== undefined ? pageInfo.totalPages
                           : (listBody.totalPages !== undefined ? listBody.totalPages : 1);
                page++;
                Utilities.sleep(300);
            } catch (e) {
                Logger.log("Lexware Zahlungen: voucherlist (" + voucherType + ") fehlgeschlagen: " + e.message);
                break;
            }
        } while (page < totalPages);
    });

    Logger.log("Lexware Zahlungen: " + allVouchers.length + " Belege zum Abrufen der Zahlungen gefunden");

    var newRows = [];
    var updatedCount = 0;

    allVouchers.forEach(function (v) {
        var voucherId = String(v.id || "").trim();
        if (!voucherId) return;

        var rawEntry = "";

        try {
            var payResult = lexwareRequest("/payments/" + voucherId);
            rawEntry = toRawString_(payResult ? payResult.body : null);
        } catch (e) {
            if (!isLexware404_(e)) {
                Logger.log("Lexware Zahlungen: Zahlung für Beleg " + voucherId + " fehlgeschlagen: " + e.message);
                rawEntry = toRawString_({
                    error: e.message,
                    voucherId: voucherId
                });
            } else {
                rawEntry = "";
            }
        }

        var row = [voucherId, rawEntry];
        if (existingById[voucherId]) {
            var existing = existingById[voucherId].data;
            var changed = row.some(function (val, i) { return String(val) !== String(existing[i]); });
            if (changed) {
                sheet.getRange(existingById[voucherId].rowIndex, 1, 1, row.length).setValues([row]);
                updatedCount++;
            }
        } else {
            newRows.push(row);
        }

        // Courtesy pause after each payment request to respect the 2 req/s rate limit.
        Utilities.sleep(300);
    });

    if (newRows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, LEXWARE_ZAHLUNGEN_HEADERS.length).setValues(newRows);
    }

    Logger.log(
        "Lexware Zahlungen import abgeschlossen: Belege=" + allVouchers.length +
        ", eingefügt=" + newRows.length +
        ", aktualisiert=" + updatedCount
    );

    return {
        ok: true,
        sheet: sheetName,
        total: allVouchers.length,
        inserted: newRows.length,
        updated: updatedCount
    };
}

// ---- All imports -------------------------------------------

/**
 * Runs all Lexware imports:
 *   1. importLexwareToSheet()             – outgoing invoices (Rechnungen)
 *   2. importLexwareEinnahmen()           – income vouchers (Einnahmen)
 *   3. importLexwareAusgaben()            – expense vouchers (Ausgaben)
 *   4. importLexwareKategorien()          – posting categories (Buchungskategorien)
 *   5. setupLexwareKontoZuordnungSheet()  – ensures Konto-Zuordnung sheet exists
 *   6. syncLexwareKundenSheet()           – contacts cache (Kunden & Lieferanten)
 *   7. importLexwareUmsaetze()            – all vouchers with line items (Umsätze)
 *   8. importLexwarePayments()            – payment status per voucher (Zahlungen)
 */
function importLexwareAll() {
    importLexwareToSheet();
    importLexwareEinnahmen();
    importLexwareAusgaben();
    try {
        importLexwareKategorien();
    } catch (e) {
        Logger.log("importLexwareKategorien skipped: " + e.message);
    }
    try {
        setupLexwareKontoZuordnungSheet();
    } catch (e) {
        Logger.log("setupLexwareKontoZuordnungSheet skipped: " + e.message);
    }
    try {
        syncLexwareKundenSheet();
    } catch (e) {
        Logger.log("syncLexwareKundenSheet skipped: " + e.message);
    }
    importLexwareUmsaetze();
    try {
        importLexwarePayments();
    } catch (e) {
        Logger.log("importLexwarePayments skipped: " + e.message);
    }
}
