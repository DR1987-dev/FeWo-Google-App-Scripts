// ============================================================
// Lodgify → Lexware Voucher Creation
//
// Two days BEFORE the check-in date (configurable), a Lexware
// sales-invoice (Einnahmebeleg) is created for the booking with
// two line items:
//   1. Einnahmen  – 7 % VAT, amount = total booking amount − 50 €
//   2. Dienstleistung – 7 % VAT, amount = 50 €
//
// The Lexware contact is looked up by the guest name; if no
// matching contact exists it is created automatically.
//
// Tracking columns appended to the AlleBuchungen sheet:
//   LexwareVoucherCreated  – ISO timestamp of the voucher creation
//   LexwareVoucherId       – Lexware voucher UUID
//   LexwareVoucherError    – Error message of the last failed attempt
//                            (booking is permanently skipped once this is set)
//
// Script Properties (alle optional):
//   INVOICE_VOUCHER_DAYS_BEFORE_CHECKIN – Tage vor CheckIn (Standard: 2)
//   INVOICE_UPLOAD_SHEET_NAME           – Sheet-Name (Standard: AlleBuchungen)
//   INVOICE_UPLOAD_CUTOFF_DATE          – Buchungen mit CheckIn vor diesem Datum
//                                          werden ignoriert (ISO: YYYY-MM-DD)
//                                          Standard: 2026-01-01
//   INVOICE_EINMALKUNDEN                – "true": Einmalkunden-Modus – Gastnamen werden
//                                          direkt als Adresse in den Beleg geschrieben,
//                                          kein Lexware-Kontakt wird angelegt/gesucht.
//                                          Standard: "false" (bisheriges Verhalten)
//
// ---- PDF-Download (Lodgify public API nicht verfügbar – für später auskommentiert) ----
//
// INVOICE_UPLOAD_DAYS_AFTER_CHECKOUT  – Tage nach CheckOut (Standard: 2)
// LODGIFY_INVOICE_PATH_TEMPLATE       – Legacy-Pfad-Template für den Invoice-Download.
//                                        Standard: /v1/reservation/booking/{booking_id}/invoice
//                                        {booking_id} wird durch die URL-kodierte Buchungs-ID ersetzt.
// ============================================================

var INVOICE_UPLOAD_COL_UPLOADED_ = "LexwareVoucherCreated";
var INVOICE_UPLOAD_COL_FILE_ID_  = "LexwareVoucherId";
var INVOICE_UPLOAD_COL_ERROR_    = "LexwareVoucherError";

// ---- Config ------------------------------------------------

function getInvoiceUploadConfig_() {
    var props = PropertiesService.getScriptProperties();

    var rawDays = props.getProperty("INVOICE_VOUCHER_DAYS_BEFORE_CHECKIN");
    var daysBeforeCheckin = (rawDays !== null && rawDays !== "" && !isNaN(Number(rawDays)) && Number(rawDays) >= 0)
        ? Number(rawDays)
        : 2;

    var sheetName = (props.getProperty("INVOICE_UPLOAD_SHEET_NAME") || "AlleBuchungen").trim();

    var cutoffDateStr = (props.getProperty("INVOICE_UPLOAD_CUTOFF_DATE") || "2026-01-01").trim();
    var cutoffDate = cutoffDateStr ? new Date(cutoffDateStr) : new Date("2026-01-01");
    if (cutoffDate && isNaN(cutoffDate.getTime())) cutoffDate = new Date("2026-01-01");

    var einmalkunden = (props.getProperty("INVOICE_EINMALKUNDEN") || "").trim().toLowerCase() === "true";

    return {
        daysBeforeCheckin: daysBeforeCheckin,
        sheetName: sheetName,
        cutoffDate: cutoffDate,
        einmalkunden: einmalkunden
    };
}

// ---- Lodgify invoice download (auskommentiert – PDF-Download über Lodgify public API nicht verfügbar) ------

//
// /**
//  * Downloads the invoice document for a reservation from Lodgify.
//  *
//  * Uses the reservations invoices endpoints to fetch invoice metadata and then
//  * downloads the PDF from publicInvoiceLink/pdfLink. Falls back to the legacy
//  * direct invoice path for older tenants.
//  *
//  * @param  {string} bookingId  Lodgify reservation/booking ID.
//  * @param  {string} pathTemplate  Path template with {booking_id} placeholder.
//  * @return {Blob}  The invoice blob (PDF or other format).
//  */
// function getLodgifyInvoicePdf_(bookingId, pathTemplate) {
//     var lodgifyConfig = validateLodgifyConfig();
//     var encodedId = encodeURIComponent(String(bookingId));
//     var invoiceListPath = "/api/v1/reservations/" + encodedId + "/invoices";
//     var invoiceListResponse;
//     var invoiceCandidates = [];
//     var lastError = "no path attempted";
//
//     try {
//         invoiceListResponse = lodgifyRequest(invoiceListPath, { method: "get" });
//         invoiceCandidates = normalizeLodgifyList(invoiceListResponse.body);
//     } catch (e) {
//         lastError = String(e && e.message ? e.message : e) + " (path: " + invoiceListPath + ")";
//     }
//
//     for (var i = 0; i < invoiceCandidates.length; i++) {
//         var invoice = invoiceCandidates[i] || {};
//         var invoiceId = invoice.id || invoice.invoiceId || invoice.invoice_id || "";
//         var detail = invoice;
//
//         if (invoiceId) {
//             try {
//                 detail = lodgifyRequest(
//                     invoiceListPath + "/" + encodeURIComponent(String(invoiceId)),
//                     { method: "get" }
//                 ).body || invoice;
//             } catch (e) {
//                 lastError = String(e && e.message ? e.message : e) + " (invoiceId: " + invoiceId + ")";
//             }
//         }
//
//         var docUrl = extractLodgifyInvoiceDocumentUrl_(detail) || extractLodgifyInvoiceDocumentUrl_(invoice);
//         if (!docUrl) {
//             lastError = "Invoice metadata contained no publicInvoiceLink/pdfLink for booking " + bookingId;
//             continue;
//         }
//
//         try {
//             return fetchLodgifyInvoiceBlobFromUrl_(String(docUrl), lodgifyConfig.apiKey);
//         } catch (e) {
//             lastError = String(e && e.message ? e.message : e) + " (docUrl: " + docUrl + ")";
//         }
//     }
//
//     var template = pathTemplate || "/v1/reservation/booking/{booking_id}/invoice";
//     var paths = [
//         template.replace("{booking_id}", encodedId),
//         "/v2/reservations/bookings/" + encodedId + "/invoice"
//     ];
//     var uniquePaths = [];
//     var seen = {};
//     paths.forEach(function (p) {
//         if (!seen[p]) { seen[p] = true; uniquePaths.push(p); }
//     });
//
//     for (var j = 0; j < uniquePaths.length; j++) {
//         try {
//             return fetchLegacyLodgifyInvoiceBlob_(uniquePaths[j], lodgifyConfig.apiKey);
//         } catch (e) {
//             lastError = String(e && e.message ? e.message : e) + " (path: " + uniquePaths[j] + ")";
//         }
//     }
//
//     throw new Error("Lodgify invoice download failed for booking " + bookingId + ". Last error: " + lastError);
// }
//
// function extractLodgifyInvoiceDocumentUrl_(invoice) {
//     if (!invoice || typeof invoice !== "object") return "";
//     return String(
//         invoice.publicInvoiceLink ||
//         invoice.pdfLink ||
//         invoice.pdf_link ||
//         invoice.public_invoice_link ||
//         ""
//     ).trim();
// }
//
// function fetchLodgifyInvoiceBlobFromUrl_(docUrl, apiKey) {
//     var response = UrlFetchApp.fetch(String(docUrl), {
//         method: "get",
//         muteHttpExceptions: true,
//         headers: {
//             "X-ApiKey": apiKey,
//             "Accept": "application/pdf,application/octet-stream,*/*"
//         }
//     });
//     var status = response.getResponseCode();
//     if (status < 200 || status >= 300) {
//         throw new Error("Document URL fetch failed (" + status + "): " + docUrl);
//     }
//     return response.getBlob();
// }
//
// function fetchLegacyLodgifyInvoiceBlob_(path, apiKey) {
//     var url = lodgifyBuildUrl(path);
//     var response = UrlFetchApp.fetch(url, {
//         method: "get",
//         muteHttpExceptions: true,
//         headers: {
//             "X-ApiKey": apiKey,
//             "Accept": "application/pdf,application/octet-stream,application/json,*/*"
//         }
//     });
//
//     var status = response.getResponseCode();
//     if (status < 200 || status >= 300) {
//         throw new Error("HTTP " + status + " at " + path + ": " + response.getContentText());
//     }
//
//     var contentType = String(
//         (response.getAllHeaders()["Content-Type"] || response.getAllHeaders()["content-type"]) || ""
//     ).toLowerCase();
//
//     if (contentType.indexOf("application/json") !== -1 || contentType.indexOf("text/") !== -1) {
//         var bodyText = response.getContentText() || "";
//         var parsed;
//         try { parsed = JSON.parse(bodyText); } catch (e) { parsed = null; }
//
//         var docUrl = parsed
//             ? (parsed.url || parsed.invoice_url || parsed.document_url ||
//                parsed.download_url || parsed.pdf_url || parsed.pdfUrl || "")
//             : "";
//
//         if (!docUrl) {
//             throw new Error("JSON response at " + path + " contained no downloadable URL. Body: " + bodyText.slice(0, 200));
//         }
//
//         return fetchLodgifyInvoiceBlobFromUrl_(String(docUrl), apiKey);
//     }
//
//     return response.getBlob();
// }
//


// ---- Lexware contact helper --------------------------------

/**
 * Looks up a Lexware contact by display name. If none is found, creates a new
 * customer contact with the given name and returns its UUID.
 *
 * @param  {string} guestName  Full name of the guest.
 * @return {string}  Lexware contact UUID.
 */
function findOrCreateLexwareContactByGuestName_(guestName) {
    var name = String(guestName || "").trim();
    if (!name) throw new Error("findOrCreateLexwareContactByGuestName_: guest name is empty");

    // Search for existing contact by name
    var searchResult;
    try {
        searchResult = lexwareRequest("/contacts", { name: name });
    } catch (e) {
        Logger.log("VoucherCreate: Kontaktsuche fehlgeschlagen für '" + name + "': " + e.message);
        searchResult = null;
    }

    if (searchResult && searchResult.body) {
        var body = searchResult.body;
        var contacts = [];
        if (Array.isArray(body)) {
            contacts = body;
        } else if (body && Array.isArray(body.content)) {
            contacts = body.content;
        } else if (body && Array.isArray(body.contacts)) {
            contacts = body.contacts;
        } else if (body && body.id) {
            contacts = [body];
        }

        for (var i = 0; i < contacts.length; i++) {
            var c = contacts[i];
            var displayName = "";
            if (c.company && c.company.name) {
                displayName = String(c.company.name).trim();
            } else if (c.person) {
                var parts = [];
                if (c.person.firstName) parts.push(String(c.person.firstName).trim());
                if (c.person.lastName)  parts.push(String(c.person.lastName).trim());
                displayName = parts.join(" ");
            } else {
                displayName = String(c.displayName || c.name || "").trim();
            }

            if (displayName.toLowerCase() === name.toLowerCase() && c.id) {
                Logger.log("VoucherCreate: Kontakt '" + name + "' gefunden (ID=" + c.id + ").");
                return String(c.id);
            }
        }
    }

    // Not found – create a new customer contact
    var nameParts = name.split(" ");
    var firstName, lastName;
    if (nameParts.length === 1) {
        // Single-word name: store entirely as lastName (common for surnames only)
        firstName = "";
        lastName  = nameParts[0];
    } else {
        firstName = nameParts.slice(0, -1).join(" ");
        lastName  = nameParts[nameParts.length - 1];
    }

    var newContactPayload = { roles: { customer: {} }, person: {} };
    if (firstName) newContactPayload.person.firstName = firstName;
    if (lastName)  newContactPayload.person.lastName  = lastName;

    var createResult = lexwarePostRequest_("/contacts", newContactPayload);
    var created = createResult.body || {};
    var newId = String(created.id || "");
    if (!newId) {
        throw new Error(
            "VoucherCreate: Kontakt erstellt, aber keine ID zurückgegeben. Body: " +
            JSON.stringify(created)
        );
    }

    Logger.log("VoucherCreate: Kontakt '" + name + "' erstellt (ID=" + newId + ").");
    return newId;
}

// ---- Helpers -----------------------------------------------

/**
 * Ensures the three tracking columns exist in the sheet header row and
 * returns their 0-based column indices.
 *
 * @param  {Sheet}  sheet    Target sheet.
 * @param  {Array}  headers  Current header values (mutated in place when columns are added).
 * @return {{ uploadedIdx: number, fileIdIdx: number, errorIdx: number }}
 */
function ensureInvoiceUploadColumns_(sheet, headers) {
    var uploadedIdx = headers.indexOf(INVOICE_UPLOAD_COL_UPLOADED_);
    var fileIdIdx   = headers.indexOf(INVOICE_UPLOAD_COL_FILE_ID_);
    var errorIdx    = headers.indexOf(INVOICE_UPLOAD_COL_ERROR_);

    if (uploadedIdx === -1) {
        uploadedIdx = headers.length;
        sheet.getRange(1, uploadedIdx + 1).setValue(INVOICE_UPLOAD_COL_UPLOADED_);
        headers.push(INVOICE_UPLOAD_COL_UPLOADED_);
    }

    if (fileIdIdx === -1) {
        fileIdIdx = headers.length;
        sheet.getRange(1, fileIdIdx + 1).setValue(INVOICE_UPLOAD_COL_FILE_ID_);
        headers.push(INVOICE_UPLOAD_COL_FILE_ID_);
    }

    if (errorIdx === -1) {
        errorIdx = headers.length;
        sheet.getRange(1, errorIdx + 1).setValue(INVOICE_UPLOAD_COL_ERROR_);
        headers.push(INVOICE_UPLOAD_COL_ERROR_);
    }

    return { uploadedIdx: uploadedIdx, fileIdIdx: fileIdIdx, errorIdx: errorIdx };
}

function toMidnight_(date) {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

// ---- Main --------------------------------------------------

/**
 * Scans the AlleBuchungen sheet for bookings whose check-in date is
 * `daysBeforeCheckin` days from today (or earlier, if not yet processed) and
 * creates a Lexware sales invoice with two line items:
 *   1. Einnahmen     – 7 % VAT, amount = total booking amount − 50 €
 *   2. Dienstleistung – 7 % VAT, amount = 50 €
 *
 * The Lexware contact is looked up by the guest name; if none exists it is
 * created automatically.
 *
 * Already-processed rows (non-empty LexwareVoucherCreated) are skipped.
 *
 * @return {{ ok: boolean, processed: number, skipped: number, failed: number, errors: Array }}
 */
function processLodgifyInvoiceUploadToLexware() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("No active spreadsheet");

    var config = getInvoiceUploadConfig_();
    var sheet = ss.getSheetByName(config.sheetName);

    if (!sheet) {
        Logger.log(
            "VoucherCreate: sheet '" + config.sheetName + "' not found – skipping."
        );
        return { ok: true, processed: 0, skipped: 0, failed: 0, errors: [] };
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
        return { ok: true, processed: 0, skipped: 0, failed: 0, errors: [] };
    }

    var lastCol = sheet.getLastColumn();
    var allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var headers = allData[0].map(function (h) { return String(h || "").trim(); });

    var bookingIdColIdx = headers.indexOf("LodgifyBookingId");
    var checkinColIdx   = headers.indexOf("CheckIn");
    var statusColIdx    = headers.indexOf("Status");
    var gastNameColIdx  = headers.indexOf("GastName");
    var betragColIdx    = headers.indexOf("Betrag");

    if (bookingIdColIdx === -1 || checkinColIdx === -1) {
        Logger.log(
            "VoucherCreate: required columns LodgifyBookingId / CheckIn not found in sheet '" +
            config.sheetName + "' – skipping."
        );
        return { ok: true, processed: 0, skipped: 0, failed: 0, errors: [] };
    }

    var cols = ensureInvoiceUploadColumns_(sheet, headers);

    var today = toMidnight_(new Date());
    var msPerDay = 24 * 60 * 60 * 1000;

    var processed = 0;
    var skipped   = 0;
    var failed    = 0;
    var errors    = [];

    // Cache for guest-name → Lexware contact UUID
    var contactCache = {};

    for (var i = 1; i < allData.length; i++) {
        var row = allData[i];

        var bookingId = String(row[bookingIdColIdx] || "").trim();
        if (!bookingId) { skipped++; continue; }

        // Skip already created
        var uploadedValue = cols.uploadedIdx < row.length
            ? String(row[cols.uploadedIdx] || "").trim()
            : "";
        if (uploadedValue) { skipped++; continue; }

        // Skip bookings with a previously recorded error (prevents repeated API calls)
        var errorValue = cols.errorIdx < row.length
            ? String(row[cols.errorIdx] || "").trim()
            : "";
        if (errorValue) { skipped++; continue; }

        // Parse check-in date
        var checkinRaw = row[checkinColIdx];
        var checkinDate = checkinRaw instanceof Date
            ? checkinRaw
            : (checkinRaw ? new Date(checkinRaw) : null);
        if (!checkinDate || isNaN(checkinDate.getTime())) { skipped++; continue; }

        var checkinDay = toMidnight_(checkinDate);

        // Respect optional cutoff
        if (config.cutoffDate && checkinDay < config.cutoffDate) { skipped++; continue; }

        // Trigger date = checkin − daysBeforeCheckin
        var triggerDate = new Date(checkinDay.getTime() - config.daysBeforeCheckin * msPerDay);

        // Only process when today has reached the trigger date
        if (today < triggerDate) { skipped++; continue; }

        // Skip declined / cancelled bookings
        if (statusColIdx >= 0) {
            var status = String(row[statusColIdx] || "").trim();
            if (status && isDeclinedOrCancelledStatusText_(status)) { skipped++; continue; }
        }

        // Read guest name and total amount
        var guestName = gastNameColIdx >= 0
            ? String(row[gastNameColIdx] || "").trim()
            : "";
        var totalAmount = betragColIdx >= 0
            ? (Number(row[betragColIdx]) || 0)
            : 0;

        if (!guestName) {
            var errMsg = "GastName ist leer für Buchung " + bookingId;
            Logger.log("VoucherCreate: " + errMsg);
            sheet.getRange(i + 1, cols.errorIdx + 1).setValue(errMsg);
            errors.push({ bookingId: bookingId, error: errMsg });
            failed++;
            continue;
        }

        // Compute line item amounts (minimum 0.01 € for item 1)
        var serviceAmount   = 50;
        var einnahmenAmount = round2(totalAmount - serviceAmount);
        if (einnahmenAmount < 0) {
            Logger.log(
                "VoucherCreate: Buchung " + bookingId +
                " – Gesamtbetrag (" + totalAmount + " €) kleiner als Dienstleistungsanteil (" +
                serviceAmount + " €). Einnahmen-Position wird mit 0 € erstellt."
            );
            einnahmenAmount = 0;
        }

        // Belegdatum = trigger date (= checkin − daysBeforeCheckin)
        var voucherDateStr = Utilities.formatDate(
            triggerDate, Session.getScriptTimeZone(), "yyyy-MM-dd"
        );

        try {
            // Resolve Lexware contact (skipped in Einmalkunden-Modus)
            var contactId = null;
            if (!config.einmalkunden) {
                contactId = contactCache[guestName];
                if (!contactId) {
                    contactId = findOrCreateLexwareContactByGuestName_(guestName);
                    contactCache[guestName] = contactId;
                }
            }

            // Build voucher reference from booking ID
            var belegRef = "Lodgify-" + bookingId;

            // Create voucher
            var voucherParams = {
                typ:          "salesinvoice",
                kontaktnummer: "",
                belegRef:     belegRef,
                voucherDate:  voucherDateStr,
                notiz:        "Lodgify Buchung " + bookingId + " – CheckIn " +
                               Utilities.formatDate(checkinDay, Session.getScriptTimeZone(), "yyyy-MM-dd"),
                lineItems: [
                    {
                        kategorieName: "Einnahmen",
                        categoryId:    null,
                        betragBrutto:  einnahmenAmount,
                        mwstSatz:      7
                    },
                    {
                        kategorieName: "Dienstleistung",
                        categoryId:    null,
                        betragBrutto:  serviceAmount,
                        mwstSatz:      7
                    }
                ]
            };
            if (contactId) {
                voucherParams.contactId = contactId;
            } else {
                voucherParams.address = { name: guestName };
            }
            var voucherId = createLexwareManuellerUmsatz_(voucherParams);

            var timestamp = Utilities.formatDate(
                new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"
            );

            var sheetRow = i + 1;
            sheet.getRange(sheetRow, cols.uploadedIdx + 1).setValue(timestamp);
            sheet.getRange(sheetRow, cols.fileIdIdx + 1).setValue(voucherId || "");

            Logger.log(
                "VoucherCreate: Buchung " + bookingId +
                " → Beleg erstellt, voucherId=" + (voucherId || "n/a") +
                ", gast=" + guestName +
                (config.einmalkunden ? " [Einmalkunde]" : ", kontaktId=" + (contactId || "n/a")) +
                ", betrag=" + totalAmount +
                ", checkin=" + Utilities.formatDate(checkinDay, Session.getScriptTimeZone(), "yyyy-MM-dd")
            );
            processed++;

        } catch (e) {
            var errMsg2 = String(e && e.message ? e.message : e);
            Logger.log("VoucherCreate: Buchung " + bookingId + " fehlgeschlagen: " + errMsg2);
            sheet.getRange(i + 1, cols.errorIdx + 1).setValue(errMsg2);
            errors.push({ bookingId: bookingId, error: errMsg2 });
            failed++;
        }
    }

    Logger.log(
        "VoucherCreate complete: processed=" + processed +
        ", skipped=" + skipped +
        ", failed=" + failed
    );

    return {
        ok: true,
        processed: processed,
        skipped: skipped,
        failed: failed,
        errors: errors
    };
}

/**
 * Creates a daily time-based trigger for processLodgifyInvoiceUploadToLexware
 * at approximately 08:00 (script timezone).
 * Safe to call multiple times – deletes any existing trigger for the same
 * function before creating a new one.
 */
function setupInvoiceUploadTrigger() {
    var fnName = "processLodgifyInvoiceUploadToLexware";

    ScriptApp.getProjectTriggers().forEach(function (trigger) {
        if (trigger.getHandlerFunction() === fnName) {
            ScriptApp.deleteTrigger(trigger);
        }
    });

    ScriptApp.newTrigger(fnName)
        .timeBased()
        .everyDays(1)
        .atHour(8)
        .create();

    Logger.log("VoucherCreate trigger created: daily at 08:00 for " + fnName);
}
