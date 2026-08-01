// ============================================================
// Lexware Office API Integration
// API Docs: https://developers.lexware.io/docs/public-api/
// Auth:     ****** stored in Script Property LEXWARE_API_KEY
// ============================================================

var LEXWARE_BASE_URL = "https://api.lexware.io/v1";
var LEXWARE_BANKTRANSACTIONS_BASE_URL = "https://api.lexware.io/banktransactions/v1";
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

// ---- Finance / Bank accounts & transactions ----------------

/**
 * Fetches all bank accounts from GET /v1/bankaccounts.
 * Returns the parsed response body.
 */
function lexwareGetBankAccounts_() {
    return lexwareRequest("/bankaccounts");
}

/**
 * Fetches one page of bank transactions from Lexware's banktransactions API.
 *
 * @param {string|null} bankAccountId - Optional: limit results to one account.
 * @param {number} page               - 0-based page index.
 * @param {number} pageSize           - Items per page (max 250).
 */
function lexwareGetBankTransactions_(bankAccountId, page, pageSize, requestMode) {
    var safePage = page || 0;
    var safePageSize = pageSize || 100;
    var allRequests = [
        {
            name: "banktransactions-api",
            baseUrl: LEXWARE_BANKTRANSACTIONS_BASE_URL,
            // The dedicated banktransactions API uses 1-based page numbering.
            params: buildBankTransactionParams_(safePage + 1, safePageSize, bankAccountId, "limit")
        },
        {
            name: "legacy-public-api",
            baseUrl: LEXWARE_BASE_URL,
            params: buildBankTransactionParams_(safePage, safePageSize, bankAccountId, "size")
        }
    ];
    var requests = requestMode
        ? allRequests.filter(function (request) { return request.name === requestMode; })
        : allRequests;
    var lastError = null;
    var attemptedVariants = [];
    for (var i = 0; i < requests.length; i++) {
        try {
            return {
                result: lexwareRequest("/banktransactions", requests[i].params, requests[i].baseUrl),
                requestMode: requests[i].name
            };
        } catch (e) {
            lastError = e;
            attemptedVariants.push(requests[i].name);
            Logger.log(
                "Lexware banktransactions request variant failed (" +
                requests[i].name +
                ", page=" + safePage +
                ", baseUrl=" + requests[i].baseUrl +
                "): " + e.message
            );
        }
    }
    Logger.log("Lexware banktransactions: all request variants failed");
    throw new Error(
        "Lexware banktransactions request failed for variants [" +
        attemptedVariants.join(", ") +
        "] on page " + safePage + ": " +
        (lastError ? lastError.message : "unknown error")
    );
}

function buildBankTransactionParams_(page, pageSize, bankAccountId, sizeKey) {
    var params = { page: page };
    params[sizeKey] = pageSize;
    if (bankAccountId) {
        params.bankAccountId = bankAccountId;
    }
    return params;
}

// ---- Sheet: Kontostand (bank account balances) -------------

var LEXWARE_KONTOSTAND_SHEET_NAME = "Lexware_Kontostand";

var LEXWARE_KONTOSTAND_HEADERS = [
    "ID",
    "Kontoname",
    "IBAN",
    "Kontonummer",
    "BIC",
    "Kontostand",
    "Währung",
    "Zuletzt aktualisiert"
];

/**
 * Fetches all bank accounts from Lexware and writes them to the
 * "Lexware_Kontostand" sheet (or the sheet configured via the
 * LEXWARE_KONTOSTAND_SHEET_NAME script property).
 *
 * @return {{ok:boolean, sheet:string, total:number}}
 */
function importLexwareKontostand() {
    var props = PropertiesService.getScriptProperties();
    var sheetName = (props.getProperty("LEXWARE_KONTOSTAND_SHEET_NAME") || LEXWARE_KONTOSTAND_SHEET_NAME).trim();

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("No active spreadsheet");

    var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

    // API may return an array directly or wrap it in a "content" or "bankAccounts" property
    var accounts = [];
    try {
        var result = lexwareGetBankAccounts_();
        var body = result.body;
        if (Array.isArray(body)) {
            accounts = body;
        } else if (body && Array.isArray(body.content)) {
            accounts = body.content;
        } else if (body && Array.isArray(body.bankAccounts)) {
            accounts = body.bankAccounts;
        } else if (body) {
            // Single account object
            accounts = [body];
        }
    } catch (e) {
        Logger.log("Lexware: /bankaccounts endpoint not available – skipping Kontostand: " + e.message);
        return { ok: false, sheet: sheetName, total: 0, error: e.message };
    }

    Logger.log("Lexware bankaccounts: fetched " + accounts.length + " account(s)");

    // Rebuild the sheet from scratch (small dataset – overwrite is fine)
    sheet.clearContents();
    sheet.appendRow(LEXWARE_KONTOSTAND_HEADERS);
    sheet.getRange(1, 1, 1, LEXWARE_KONTOSTAND_HEADERS.length).setFontWeight("bold");

    var rows = accounts.map(function (acc) {
        var balance = "";
        if (acc.balance !== undefined && acc.balance !== null) {
            balance = typeof acc.balance === "object"
                ? (acc.balance.value !== undefined ? acc.balance.value : JSON.stringify(acc.balance))
                : acc.balance;
        }
        var currency = "";
        if (acc.currency) {
            currency = acc.currency;
        } else if (acc.balance && acc.balance.currency) {
            currency = acc.balance.currency;
        } else {
            currency = "EUR";
        }
        return [
            String(acc.id || ""),
            acc.name || acc.accountName || "",
            acc.iban || "",
            acc.accountNumber || "",
            acc.bic || "",
            balance,
            currency,
            acc.lastUpdated ? String(acc.lastUpdated).slice(0, 19).replace("T", " ") : ""
        ];
    });

    if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, LEXWARE_KONTOSTAND_HEADERS.length).setValues(rows);
    }

    Logger.log("Lexware Kontostand import complete: total=" + accounts.length);

    return { ok: true, sheet: sheetName, total: accounts.length };
}

// ---- Sheet: Finanzen (bank transactions) -------------------

var LEXWARE_FINANZEN_SHEET_NAME = "Lexware_Finanzen";

var LEXWARE_FINANZEN_HEADERS = [
    "ID",
    "Konto-ID",
    "Kontoname",
    "Datum",
    "Betrag",
    "Währung",
    "Typ",
    "Zahlungsreferenz",
    "Status",
    "Buchungstext"
];

/**
 * Fetches all bank transactions from Lexware and upserts them into the
 * "Lexware_Finanzen" sheet (or the sheet configured via the
 * LEXWARE_FINANZEN_SHEET_NAME script property).
 *
 * Transactions are matched by their ID (column 1) to avoid duplicates.
 *
 * @return {{ok:boolean, sheet:string, total:number, inserted:number, updated:number}}
 */
function importLexwareFinanzen() {
    var props = PropertiesService.getScriptProperties();
    var sheetName = (props.getProperty("LEXWARE_FINANZEN_SHEET_NAME") || LEXWARE_FINANZEN_SHEET_NAME).trim();

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("No active spreadsheet");

    var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

    if (sheet.getLastRow() === 0) {
        sheet.appendRow(LEXWARE_FINANZEN_HEADERS);
        sheet.getRange(1, 1, 1, LEXWARE_FINANZEN_HEADERS.length).setFontWeight("bold");
    }

    // Index existing rows by transaction ID (column 1)
    var existingById = {};
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        var existingData = sheet.getRange(2, 1, lastRow - 1, LEXWARE_FINANZEN_HEADERS.length).getValues();
        existingData.forEach(function (row, idx) {
            var id = String(row[0] || "").trim();
            if (id) existingById[id] = { rowIndex: idx + 2, data: row };
        });
    }

    // Optionally fetch account names once for display
    var accountNames = {};
    try {
        var accResult = lexwareGetBankAccounts_();
        var accBody = accResult.body;
        var accList = Array.isArray(accBody) ? accBody
                    : (accBody && Array.isArray(accBody.content) ? accBody.content
                    : (accBody && Array.isArray(accBody.bankAccounts) ? accBody.bankAccounts : []));
        accList.forEach(function (a) {
            if (a.id) accountNames[String(a.id)] = a.name || a.accountName || "";
        });
    } catch (e) {
        Logger.log("Lexware: could not fetch bank account names: " + e.message);
    }

    // Fetch all transaction pages
    var allTransactions = [];
    var page = 0;
    var pageSize = 100;
    var explicitTotalPages = null;
    var bankTransactionsRequestMode = null;

    while (true) {
        var bankTransactionsResult;
        var body;
        try {
            bankTransactionsResult = lexwareGetBankTransactions_(null, page, pageSize, bankTransactionsRequestMode);
            bankTransactionsRequestMode = bankTransactionsResult.requestMode;
            body = bankTransactionsResult.result.body;
        } catch (e) {
            Logger.log("Lexware: banktransactions endpoint not available – skipping Finanzen: " + e.message);
            return { ok: false, status: "skipped", sheet: sheetName, total: 0, inserted: 0, updated: 0, error: e.message };
        }

        if (!body) {
            Logger.log("Lexware banktransactions: empty response on page " + page);
            break;
        }

        var content = Array.isArray(body) ? body
                    : (Array.isArray(body.content) ? body.content
                    : (Array.isArray(body.bankTransactions) ? body.bankTransactions : []));

        allTransactions = allTransactions.concat(content);

        var pageInfo = body.page || {};
        if (pageInfo.totalPages !== undefined) {
            explicitTotalPages = pageInfo.totalPages;
        } else if (body.totalPages !== undefined) {
            explicitTotalPages = body.totalPages;
        }
        page++;
        // Stop either when the API tells us we've reached the last page,
        // or when an undelimited result set returns a short final page.
        if (
            (explicitTotalPages !== null && page >= explicitTotalPages) ||
            (explicitTotalPages === null && content.length < pageSize)
        ) break;
    }

    Logger.log(
        "Lexware banktransactions: fetched " +
        allTransactions.length + " record(s) across " + page + " page(s)"
    );

    var newRows = [];
    var updatedCount = 0;

    allTransactions.forEach(function (t) {
        var id = String(
            t.id || t.bankTransactionReference || t.transactionId || ""
        ).trim();
        if (!id) return;

        var accountId = String(
            (t.bankAccount && t.bankAccount.id) || t.bankAccountId || ""
        );
        var accountName = accountNames[accountId] || "";

        var date = t.date || t.bookingDate || t.valueDate || "";
        if (date) date = String(date).slice(0, 10);

        var amount = "";
        if (t.amount !== undefined && t.amount !== null) {
            amount = typeof t.amount === "object"
                ? (t.amount.value !== undefined ? t.amount.value : "")
                : t.amount;
        }

        var currency = "";
        if (t.currency) {
            currency = t.currency;
        } else if (t.amount && t.amount.currency) {
            currency = t.amount.currency;
        } else {
            currency = "EUR";
        }

        var type = t.type || t.transactionType || "";
        var paymentRef = "";
        if (t.paymentReference) {
            paymentRef = typeof t.paymentReference === "object"
                ? (t.paymentReference.value || JSON.stringify(t.paymentReference))
                : String(t.paymentReference);
        }
        var status = t.status || t.bookingStatus || "";
        var purpose = t.purpose || t.description || t.bookingText || t.note || "";

        var row = [
            id,
            accountId,
            accountName,
            date,
            amount,
            currency,
            type,
            paymentRef,
            status,
            purpose
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
        sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, LEXWARE_FINANZEN_HEADERS.length).setValues(newRows);
    }

    Logger.log(
        "Lexware Finanzen import complete: total=" + allTransactions.length +
        ", inserted=" + newRows.length +
        ", updated=" + updatedCount
    );

    return {
        ok: true,
        sheet: sheetName,
        total: allTransactions.length,
        inserted: newRows.length,
        updated: updatedCount
    };
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

// ---- All imports -------------------------------------------

/**
 * Runs all Lexware imports:
 *   1. importLexwareToSheet()    – outgoing invoices (Rechnungen)
 *   2. importLexwareEinnahmen()  – income vouchers (Einnahmen)
 *   3. importLexwareAusgaben()   – expense vouchers (Ausgaben)
 *   4. importLexwareUmsaetze()   – all vouchers (Umsätze)
 *   5. importLexwareKontostand() – bank account balances (Kontostand)
 *   6. importLexwareFinanzen()   – all bank transactions (Finanzen)
 *   7. importLexwareKategorien() – posting categories (Buchungskategorien)
 */
function importLexwareAll() {
    importLexwareToSheet();
    importLexwareEinnahmen();
    importLexwareAusgaben();
    importLexwareUmsaetze();
    try {
        importLexwareKontostand();
    } catch (e) {
        Logger.log("importLexwareKontostand skipped: " + e.message);
    }
    importLexwareFinanzen();
    try {
        importLexwareKategorien();
    } catch (e) {
        Logger.log("importLexwareKategorien skipped: " + e.message);
    }
}
