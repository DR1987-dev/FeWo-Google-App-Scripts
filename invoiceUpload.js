// ============================================================
// Lodgify → Lexware Invoice Upload
//
// Two days after the checkout date (configurable), the invoice
// PDF for a booking is downloaded from the Lodgify API and
// uploaded to Lexware Office via the files API.
//
// Tracking columns appended to the AlleBuchungen sheet:
//   LexwareInvoiceUploaded  – ISO timestamp of the upload
//   LexwareFileId           – Lexware file UUID returned by the upload
//
// Script Properties (alle optional):
//   INVOICE_UPLOAD_DAYS_AFTER_CHECKOUT  – Tage nach CheckOut (Standard: 2)
//   INVOICE_UPLOAD_SHEET_NAME           – Sheet-Name (Standard: AlleBuchungen)
//   INVOICE_UPLOAD_CUTOFF_DATE          – Buchungen mit CheckOut vor diesem Datum
//                                          werden ignoriert (ISO: YYYY-MM-DD)
//   LODGIFY_INVOICE_PATH_TEMPLATE       – Pfad-Template für den Invoice-Download.
//                                          Standard: /v1/reservation/booking/{booking_id}/invoice
//                                          {booking_id} wird durch die URL-kodierte Buchungs-ID ersetzt.
// ============================================================

var INVOICE_UPLOAD_COL_UPLOADED_ = "LexwareInvoiceUploaded";
var INVOICE_UPLOAD_COL_FILE_ID_  = "LexwareFileId";

// ---- Config ------------------------------------------------

function getInvoiceUploadConfig_() {
    var props = PropertiesService.getScriptProperties();

    var rawDays = props.getProperty("INVOICE_UPLOAD_DAYS_AFTER_CHECKOUT");
    var daysAfterCheckout = (rawDays !== null && rawDays !== "" && !isNaN(Number(rawDays)) && Number(rawDays) >= 0)
        ? Number(rawDays)
        : 2;

    var sheetName = (props.getProperty("INVOICE_UPLOAD_SHEET_NAME") || "AlleBuchungen").trim();

    var cutoffDateStr = (props.getProperty("INVOICE_UPLOAD_CUTOFF_DATE") || "").trim();
    var cutoffDate = cutoffDateStr ? new Date(cutoffDateStr) : null;
    if (cutoffDate && isNaN(cutoffDate.getTime())) cutoffDate = null;

    var invoicePathTemplate = (
        props.getProperty("LODGIFY_INVOICE_PATH_TEMPLATE") ||
        "/v1/reservation/booking/{booking_id}/invoice"
    ).trim();

    return {
        daysAfterCheckout: daysAfterCheckout,
        sheetName: sheetName,
        cutoffDate: cutoffDate,
        invoicePathTemplate: invoicePathTemplate
    };
}

// ---- Lodgify invoice download ------------------------------

/**
 * Downloads the invoice document for a booking from Lodgify.
 *
 * Tries the configured path template first; if the response contains JSON
 * with a URL field, the URL is fetched to obtain the actual binary blob.
 *
 * @param  {string} bookingId  Lodgify booking ID.
 * @param  {string} pathTemplate  Path template with {booking_id} placeholder.
 * @return {Blob}  The invoice blob (PDF or other format).
 */
function getLodgifyInvoicePdf_(bookingId, pathTemplate) {
    var lodgifyConfig = validateLodgifyConfig();
    var encodedId = encodeURIComponent(String(bookingId));
    var template = pathTemplate || "/v1/reservation/booking/{booking_id}/invoice";

    var paths = [
        template.replace("{booking_id}", encodedId),
        "/v2/reservations/bookings/" + encodedId + "/invoice"
    ];

    // De-duplicate (in case the template already matches the fallback)
    var uniquePaths = [];
    var seen = {};
    paths.forEach(function (p) {
        if (!seen[p]) { seen[p] = true; uniquePaths.push(p); }
    });

    var lastError = "no path attempted";

    for (var i = 0; i < uniquePaths.length; i++) {
        try {
            var url = lodgifyBuildUrl(uniquePaths[i]);
            var response = UrlFetchApp.fetch(url, {
                method: "get",
                muteHttpExceptions: true,
                headers: {
                    "X-ApiKey": lodgifyConfig.apiKey,
                    "Accept": "application/pdf,application/octet-stream,application/json,*/*"
                }
            });

            var status = response.getResponseCode();
            if (status < 200 || status >= 300) {
                lastError = "HTTP " + status + " at " + uniquePaths[i] + ": " + response.getContentText();
                continue;
            }

            var contentType = String(
                (response.getAllHeaders()["Content-Type"] || response.getAllHeaders()["content-type"]) || ""
            ).toLowerCase();

            // If the API returns JSON, look for a URL pointing to the actual document
            if (contentType.indexOf("application/json") !== -1 || contentType.indexOf("text/") !== -1) {
                var bodyText = response.getContentText() || "";
                var parsed;
                try { parsed = JSON.parse(bodyText); } catch (e) { parsed = null; }

                var docUrl = parsed
                    ? (parsed.url || parsed.invoice_url || parsed.document_url ||
                       parsed.download_url || parsed.pdf_url || parsed.pdfUrl || "")
                    : "";

                if (docUrl) {
                    var docResponse = UrlFetchApp.fetch(String(docUrl), {
                        method: "get",
                        muteHttpExceptions: true,
                        headers: { "X-ApiKey": lodgifyConfig.apiKey }
                    });
                    var docStatus = docResponse.getResponseCode();
                    if (docStatus >= 200 && docStatus < 300) {
                        return docResponse.getBlob();
                    }
                    lastError = "Document URL fetch failed (" + docStatus + "): " + docUrl;
                    continue;
                }

                lastError = "JSON response at " + uniquePaths[i] + " contained no downloadable URL. Body: " + bodyText.slice(0, 200);
                continue;
            }

            // Binary response – return directly
            return response.getBlob();

        } catch (e) {
            lastError = String(e && e.message ? e.message : e) + " (path: " + uniquePaths[i] + ")";
        }
    }

    throw new Error(
        "Lodgify invoice download failed for booking " + bookingId + ". Last error: " + lastError
    );
}

// ---- Helpers -----------------------------------------------

/**
 * Ensures the two tracking columns exist in the sheet header row and
 * returns their 0-based column indices.
 *
 * @param  {Sheet}  sheet    Target sheet.
 * @param  {Array}  headers  Current header values (mutated in place when columns are added).
 * @return {{ uploadedIdx: number, fileIdIdx: number }}
 */
function ensureInvoiceUploadColumns_(sheet, headers) {
    var uploadedIdx = headers.indexOf(INVOICE_UPLOAD_COL_UPLOADED_);
    var fileIdIdx   = headers.indexOf(INVOICE_UPLOAD_COL_FILE_ID_);

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

    return { uploadedIdx: uploadedIdx, fileIdIdx: fileIdIdx };
}

function toMidnight_(date) {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

// ---- Main --------------------------------------------------

/**
 * Scans the AlleBuchungen sheet for bookings whose checkout date was
 * `daysAfterCheckout` days ago (or earlier, if not yet processed) and:
 *   1. Downloads the invoice PDF from Lodgify.
 *   2. Uploads the PDF to Lexware Office via the files API.
 *   3. Writes the upload timestamp and Lexware file ID back to the sheet.
 *
 * Already-processed rows (non-empty LexwareInvoiceUploaded) are skipped.
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
            "InvoiceUpload: sheet '" + config.sheetName + "' not found – skipping."
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
    var checkoutColIdx  = headers.indexOf("CheckOut");
    var statusColIdx    = headers.indexOf("Status");

    if (bookingIdColIdx === -1 || checkoutColIdx === -1) {
        Logger.log(
            "InvoiceUpload: required columns LodgifyBookingId / CheckOut not found in sheet '" +
            config.sheetName + "' – skipping."
        );
        return { ok: true, processed: 0, skipped: 0, failed: 0, errors: [] };
    }

    var cols = ensureInvoiceUploadColumns_(sheet, headers);

    var today = toMidnight_(new Date());

    var processed = 0;
    var skipped   = 0;
    var failed    = 0;
    var errors    = [];

    for (var i = 1; i < allData.length; i++) {
        var row = allData[i];

        var bookingId = String(row[bookingIdColIdx] || "").trim();
        if (!bookingId) { skipped++; continue; }

        // Skip already uploaded
        var uploadedValue = cols.uploadedIdx < row.length
            ? String(row[cols.uploadedIdx] || "").trim()
            : "";
        if (uploadedValue) { skipped++; continue; }

        // Parse checkout date
        var checkoutRaw = row[checkoutColIdx];
        var checkoutDate = checkoutRaw instanceof Date
            ? checkoutRaw
            : (checkoutRaw ? new Date(checkoutRaw) : null);
        if (!checkoutDate || isNaN(checkoutDate.getTime())) { skipped++; continue; }

        var checkoutDay = toMidnight_(checkoutDate);

        // Respect optional cutoff
        if (config.cutoffDate && checkoutDay < config.cutoffDate) { skipped++; continue; }

        // Trigger date = checkout + daysAfterCheckout
        var msPerDay = 24 * 60 * 60 * 1000;
        var triggerDate = new Date(checkoutDay.getTime() + config.daysAfterCheckout * msPerDay);

        // Only process if the trigger date has been reached
        if (triggerDate > today) { skipped++; continue; }

        // Skip declined / cancelled bookings
        if (statusColIdx >= 0) {
            var status = String(row[statusColIdx] || "").trim();
            if (status && isDeclinedOrCancelledStatusText_(status)) { skipped++; continue; }
        }

        // Download & upload
        try {
            var blob = getLodgifyInvoicePdf_(bookingId, config.invoicePathTemplate);
            var fileName = "Lodgify_Invoice_" + bookingId + ".pdf";
            blob.setName(fileName);
            if (!blob.getContentType() || blob.getContentType() === "application/octet-stream") {
                blob.setContentType("application/pdf");
            }

            var uploadResult = lexwareUploadFile_(blob, fileName);
            var timestamp = Utilities.formatDate(
                new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"
            );

            var sheetRow = i + 1; // 1-based, row 1 is the header
            sheet.getRange(sheetRow, cols.uploadedIdx + 1).setValue(timestamp);
            sheet.getRange(sheetRow, cols.fileIdIdx + 1).setValue(uploadResult.fileId || "");

            Logger.log(
                "InvoiceUpload: booking " + bookingId +
                " → uploaded, fileId=" + (uploadResult.fileId || "n/a") +
                ", checkout=" + Utilities.formatDate(checkoutDay, Session.getScriptTimeZone(), "yyyy-MM-dd")
            );
            processed++;

        } catch (e) {
            var errMsg = String(e && e.message ? e.message : e);
            Logger.log("InvoiceUpload: booking " + bookingId + " failed: " + errMsg);
            errors.push({ bookingId: bookingId, error: errMsg });
            failed++;
        }
    }

    Logger.log(
        "InvoiceUpload complete: processed=" + processed +
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

    Logger.log("InvoiceUpload trigger created: daily at 08:00 for " + fnName);
}
