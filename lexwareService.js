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

function lexwareGetInvoices(page, pageSize) {
    return lexwareRequest("/invoices", {
        page: page || 0,
        size: pageSize || 100
    });
}

function lexwareHealthCheck() {
    return lexwareRequest("/ping");
}

// ---- Invoice field extractors ------------------------------

function lexwareInvoiceContactName_(invoice) {
    var addr = invoice.address || {};
    if (addr.name) return addr.name;
    if (addr.contactId) return addr.contactId;
    return "";
}

function lexwareInvoiceTotalNet_(invoice) {
    var tp = invoice.totalPrice || {};
    return tp.totalNetAmount !== undefined ? tp.totalNetAmount : "";
}

function lexwareInvoiceTotalGross_(invoice) {
    var tp = invoice.totalPrice || {};
    return tp.totalGrossAmount !== undefined ? tp.totalGrossAmount : "";
}

function lexwareInvoiceTaxAmount_(invoice) {
    var tp = invoice.totalPrice || {};
    return tp.totalTaxAmount !== undefined ? tp.totalTaxAmount : "";
}

// ---- Import ------------------------------------------------

/**
 * Fetches all Lexware invoices and writes them to the "Lexware" sheet
 * (or the sheet configured via the LEXWARE_SHEET_NAME script property).
 * Existing rows are matched by invoice ID and updated in place; new
 * invoices are appended.
 */
function importLexwareToSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("No active spreadsheet");

    var config = getLexwareConfig();
    var sheetName = config.sheetName;
    var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

    var headers = [
        "ID",
        "Rechnungsnummer",
        "Status",
        "Rechnungsdatum",
        "Fälligkeitsdatum",
        "Kontakt",
        "Nettobetrag",
        "Steuerbetrag",
        "Bruttobetrag",
        "Währung"
    ];

    // Write header row if the sheet is empty
    if (sheet.getLastRow() === 0) {
        sheet.appendRow(headers);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    }

    // Read existing rows indexed by invoice ID (column 1)
    var existingById = {};
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        var existingData = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
        existingData.forEach(function (row, idx) {
            var id = String(row[0] || "").trim();
            if (id) existingById[id] = { rowIndex: idx + 2, data: row };
        });
    }

    // Fetch all pages
    var allInvoices = [];
    var page = 0;
    var pageSize = 100;
    var totalPages = 1;

    do {
        var result = lexwareGetInvoices(page, pageSize);
        var body = result.body;

        if (!body || !body.content) {
            Logger.log("Lexware: unexpected response on page " + page + ": " + JSON.stringify(body));
            break;
        }

        allInvoices = allInvoices.concat(body.content);
        totalPages = body.totalPages !== undefined ? body.totalPages : 1;
        page++;
    } while (page < totalPages);

    Logger.log("Lexware: fetched " + allInvoices.length + " invoices across " + totalPages + " page(s)");

    var newRows = [];
    var updatedCount = 0;

    allInvoices.forEach(function (invoice) {
        var id = String(invoice.id || "").trim();
        if (!id) return;

        var tp = invoice.totalPrice || {};
        var row = [
            id,
            invoice.voucherNumber || "",
            invoice.voucherStatus || "",
            invoice.voucherDate || "",
            invoice.dueDate || "",
            lexwareInvoiceContactName_(invoice),
            lexwareInvoiceTotalNet_(invoice),
            lexwareInvoiceTaxAmount_(invoice),
            lexwareInvoiceTotalGross_(invoice),
            tp.currency || "EUR"
        ];

        if (existingById[id]) {
            // Update existing row if anything changed
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
        "Lexware import complete: total=" + allInvoices.length +
        ", inserted=" + newRows.length +
        ", updated=" + updatedCount
    );

    return {
        ok: true,
        sheet: sheetName,
        total: allInvoices.length,
        inserted: newRows.length,
        updated: updatedCount
    };
}
