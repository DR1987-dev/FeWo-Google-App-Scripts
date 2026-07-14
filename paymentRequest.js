// ========== ZAHLUNGSAUFFORDERUNG / PAYMENT REQUEST ==========
//
// Diese Datei implementiert die automatische Zahlungsaufforderung:
//
// 1. Alle Lodgify-Buchungen werden in das Sheet "AlleBuchungen" (konfigurierbar)
//    eingepflegt/aktualisiert.
// 2. Beim nächsten Lodgify-Import werden für externe Buchungen mit
//    "RequestFullPayment" automatische Zahlungsanforderungen in Lodgify ausgelöst,
//    sobald das konfigurierte Zeitfenster vor CheckIn erreicht ist.
// 3. Bereits ausgelöste Anforderungen werden über "ZahlungsUpdateDurchgefuehrt"
//    erkannt und nicht erneut angefordert.
//
// Script Properties (konfigurierbar):
//   PAYMENT_REQUEST_SHEET_NAME          – Sheet-Name (Standard: "AlleBuchungen")
//   PAYMENT_REQUEST_DAYS_BEFORE_CHECKIN – Tage vor CheckIn (Standard: 14)
//   PAYMENT_REQUEST_WEEKS_BEFORE_CHECKIN – alternativ in Wochen
// =========================================================================

var ALLE_BUCHUNGEN_HEADERS_ = [
    "LodgifyBookingId",
    "GastName",
    "CheckIn",
    "CheckOut",
    "Betrag",
    "Status",
    "Kanal",
    "IsExternal",
    "ZahlungsAufforderungAktiv",
    "PaymentOption",
    "RequestFullPayment",
    "FullPaymentDaysBeforeCheckin",
    "FullPaymentWeeksBeforeCheckin",
    "ZahlungsUpdateDurchgefuehrt",
    "BookingType",
    "BookingDate",
    "Note",
    "Account",
    "FeesTotal",
    "NetAmount",
    "PayoutAmount"
];

// 0-basierter Index der "ZahlungsAufforderungAktiv"-Spalte in ALLE_BUCHUNGEN_HEADERS_
var ALLE_BUCHUNGEN_MARKER_COL_IDX_ = 8;

// 0-basierter Index der "ZahlungsUpdateDurchgefuehrt"-Spalte
var ALLE_BUCHUNGEN_TIMESTAMP_COL_IDX_ = 13;
var ALLE_BUCHUNGEN_GUEST_NAME_COL_IDX_ = 1;
var ALLE_BUCHUNGEN_CHECKIN_COL_IDX_ = 2;
var ALLE_BUCHUNGEN_CHECKOUT_COL_IDX_ = 3;

/**
 * Wandelt einen Wert in einen Boolean um.
 * Akzeptiert: true/false, "TRUE"/"FALSE", "1"/"0", "ja"/"nein", "yes"/"no".
 */
function toBoolean(value) {
    if (typeof value === "boolean") return value;
    const text = String(value === null || value === undefined ? "" : value).trim().toLowerCase();
    return text === "true" || text === "1" || text === "yes" || text === "ja";
}

/**
 * Liest die Konfiguration für die Zahlungsaufforderung aus den Script Properties.
 */
function getPaymentRequestConfig_() {
    const props = PropertiesService.getScriptProperties();
    const sheetName = (props.getProperty("PAYMENT_REQUEST_SHEET_NAME") || "AlleBuchungen").trim();

    const rawDays = props.getProperty("PAYMENT_REQUEST_DAYS_BEFORE_CHECKIN");
    const rawWeeks = props.getProperty("PAYMENT_REQUEST_WEEKS_BEFORE_CHECKIN");

    let daysBeforeCheckin = 14; // Standard: 2 Wochen
    if (rawDays !== null && rawDays !== "" && !isNaN(Number(rawDays)) && Number(rawDays) > 0) {
        daysBeforeCheckin = Number(rawDays);
    } else if (rawWeeks !== null && rawWeeks !== "" && !isNaN(Number(rawWeeks)) && Number(rawWeeks) > 0) {
        daysBeforeCheckin = Number(rawWeeks) * 7;
    }

    return { sheetName, daysBeforeCheckin };
}

/**
 * Prüft, ob das aktuelle Datum im Zahlungsaufforderungs-Fenster liegt.
 * Das Fenster beginnt `daysBeforeCheckin` Tage vor dem CheckIn.
 */
function isWithinPaymentRequestWindow_(checkinDate, daysBeforeCheckin) {
    if (!checkinDate) return false;
    const checkin = checkinDate instanceof Date ? checkinDate : new Date(checkinDate);
    if (isNaN(checkin.getTime())) return false;

    const now = new Date();
    if (checkin <= now) return false;
    const windowStart = new Date(checkin.getTime() - daysBeforeCheckin * 24 * 60 * 60 * 1000);

    // Auslösen wenn: jetzt >= Fensteranfang (d.h. wir sind im Fenster oder danach)
    return now >= windowStart;
}

function parsePositiveNumber_(value) {
    if (value === null || value === undefined || value === "") return null;
    const normalized = Number(value);
    if (isNaN(normalized) || normalized <= 0) return null;
    return normalized;
}

function hasPaymentRequestTimestamp_(value) {
    return String(value === null || value === undefined ? "" : value).trim() !== "";
}

function buildPaymentRequestTimestamp_() {
    return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}

function getBookingPaymentRequestLeadDays_(booking, config) {
    const bookingDays = parsePositiveNumber_(firstDefined(booking || {}, [
        "full_payment_days_before_checkin",
        "fullPaymentDaysBeforeCheckin"
    ]));
    if (bookingDays !== null) return bookingDays;

    const bookingWeeks = parsePositiveNumber_(firstDefined(booking || {}, [
        "full_payment_weeks_before_checkin",
        "fullPaymentWeeksBeforeCheckin"
    ]));
    if (bookingWeeks !== null) return bookingWeeks * 7;

    const fallbackDays = parsePositiveNumber_(config && config.daysBeforeCheckin);
    return fallbackDays !== null ? fallbackDays : null;
}

function evaluateAutomaticPaymentRequest_(booking, paymentUpdateCompleted, config) {
    if (!booking || typeof booking !== "object") {
        return { shouldRequest: false, reason: "missingBooking" };
    }

    if (hasPaymentRequestTimestamp_(paymentUpdateCompleted)) {
        return { shouldRequest: false, reason: "alreadyRequested" };
    }

    const isExternal = toBoolean(firstDefined(booking, ["is_external", "isExternal", "external"]));
    if (!isExternal) {
        return { shouldRequest: false, reason: "notExternal" };
    }

    const requestFullPayment = toBoolean(firstDefined(booking, ["request_full_payment", "requestFullPayment"]));
    if (!requestFullPayment) {
        return { shouldRequest: false, reason: "fullPaymentDisabled" };
    }

    const checkinDate = extractLodgifyCheckinDate_(booking);
    if (!checkinDate) {
        return { shouldRequest: false, reason: "missingCheckin" };
    }

    const daysBeforeCheckin = getBookingPaymentRequestLeadDays_(booking, config);
    if (daysBeforeCheckin === null) {
        return { shouldRequest: false, reason: "missingLeadTime" };
    }

    if (!isWithinPaymentRequestWindow_(checkinDate, daysBeforeCheckin)) {
        return { shouldRequest: false, reason: "outsideWindow", daysBeforeCheckin: daysBeforeCheckin };
    }

    return {
        shouldRequest: true,
        reason: "eligible",
        daysBeforeCheckin: daysBeforeCheckin
    };
}

/**
 * Gibt eine Map von Header-Name -> 1-basierter Spaltenindex für das Sheet zurück.
 */
function getSheetHeaderMap_(sheet) {
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return {};
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const map = {};
    headers.forEach(function (h, idx) {
        const key = String(h || "").trim();
        if (key) map[key] = idx + 1;
    });
    return map;
}

/**
 * Aktualisiert einzelne Zellen einer Sheet-Zeile anhand von Header-Namen.
 * Für jeden Eintrag in `updates` wird der erste gefundene Spalten-Header verwendet.
 *
 * @param {string} sheetName  Name des Sheets
 * @param {number} row        1-basierte Zeilennummer
 * @param {Array}  updates    Array von { headers: [string, ...], value: any }
 * @returns {{ updated: number, skipped: number, skippedHeaders: string[] }}
 */
function updateSheetRowByHeaders(sheetName, row, updates) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { updated: 0, skipped: updates.length, skippedHeaders: [] };

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { updated: 0, skipped: updates.length, skippedHeaders: [] };

    const headerMap = getSheetHeaderMap_(sheet);
    let updated = 0;
    const skippedHeaders = [];

    updates.forEach(function (update) {
        const headerAliases = update.headers || [];
        let colIndex = -1;
        for (let i = 0; i < headerAliases.length; i++) {
            if (headerMap[headerAliases[i]] !== undefined) {
                colIndex = headerMap[headerAliases[i]];
                break;
            }
        }
        if (colIndex === -1) {
            skippedHeaders.push(headerAliases[0] || "?");
            return;
        }
        sheet.getRange(row, colIndex).setValue(update.value);
        updated++;
    });

    return { updated, skipped: skippedHeaders.length, skippedHeaders };
}

/**
 * Schreibt die Zahlungsfelder einer Lodgify-Buchung in das Sheet.
 * Sucht die Spalten anhand der konfigurierten Header-Namen (inkl. Aliase).
 *
 * @param {string} sheetName  Name des Sheets
 * @param {number} row        1-basierte Zeilennummer
 * @param {Object} booking    Lodgify-Buchungsobjekt
 */
function updateLodgifyBookingInSheet(sheetName, row, booking) {
    const updates = [
        {
            headers: ["PaymentOption", "Zahlungsoption"],
            value: String(booking.payment_option || ""),
        },
        {
            headers: ["RequestFullPayment", "VollbetragAnfordern", "ZahlungVollAnfordern"],
            value: toBoolean(booking.request_full_payment) ? "TRUE" : "FALSE",
        },
        {
            headers: ["FullPaymentDaysBeforeCheckin", "VollbetragTageVorAnreise", "PaymentRequestDays"],
            value: String(booking.full_payment_days_before_checkin || ""),
        },
        {
            headers: ["FullPaymentWeeksBeforeCheckin", "VollbetragWochenVorAnreise", "PaymentRequestWeeks"],
            value: String(booking.full_payment_weeks_before_checkin || ""),
        },
        {
            headers: ["IsExternal", "Extern", "ExternBuchung", "NichtLodgify"],
            value: toBoolean(booking.is_external) ? "TRUE" : "FALSE",
        },
        {
            headers: ["LodgifyBookingId", "BookingId", "BuchungsID"],
            value: String(booking.lodgify_booking_id || booking.id || ""),
        },
    ];

    return updateSheetRowByHeaders(sheetName, row, updates);
}

/**
 * Stellt sicher, dass das AlleBuchungen-Sheet mit den korrekten Headern existiert.
 */
function ensureAlleBuchungenSheet_(sheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("No active spreadsheet");

    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        sheet.getRange(1, 1, 1, ALLE_BUCHUNGEN_HEADERS_.length).setValues([ALLE_BUCHUNGEN_HEADERS_]);
        sheet.setFrozenRows(1);
    }

    const headerMap = getSheetHeaderMap_(sheet);
    const missingHeaders = ALLE_BUCHUNGEN_HEADERS_.filter(function (header) {
        return headerMap[header] === undefined;
    });

    if (missingHeaders.length > 0) {
        const startCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, startCol, 1, missingHeaders.length).setValues([missingHeaders]);
    }

    if (sheet.getFrozenRows() < 1) {
        sheet.setFrozenRows(1);
    }

    return sheet;
}

function preserveExistingAlleBuchungenCellValue_(targetRow, existingRow, columnIndex) {
    if (!targetRow[columnIndex] && existingRow[columnIndex]) {
        targetRow[columnIndex] = existingRow[columnIndex];
    }
}

function extractLodgifyBookingId_(item) {
    return String(
        firstDefined(item || {}, ["id", "bookingId", "booking_id", "reservationId", "reservation_id"]) || ""
    ).trim();
}

function getRowValueByHeaders_(row, headerMap, headers, fallback) {
    const aliases = headers || [];
    for (let i = 0; i < aliases.length; i++) {
        const colIndex = headerMap[aliases[i]];
        if (colIndex === undefined) continue;
        const value = row[colIndex - 1];
        if (value !== undefined && value !== null && value !== "") {
            return value;
        }
    }
    return fallback;
}

function getOwnValueForKeys_(obj, keys) {
    if (!obj || typeof obj !== "object") return undefined;
    const aliases = keys || [];
    for (let i = 0; i < aliases.length; i++) {
        if (Object.prototype.hasOwnProperty.call(obj, aliases[i])) {
            return obj[aliases[i]];
        }
    }
    return undefined;
}

function getBookingFieldOverride_(booking, keys, rawKeys) {
    const directValue = getOwnValueForKeys_(booking, keys || []);
    if (directValue !== undefined) return directValue;

    const raw = booking && booking.raw && typeof booking.raw === "object" ? booking.raw : null;
    if (!raw) return undefined;

    return getOwnValueForKeys_(raw, rawKeys || keys || []);
}

function isValuePresent_(value) {
    return value !== null && value !== undefined && value !== "";
}

function defaultLodgifyBookingType_(isExternal) {
    return isExternal ? "lodgify_external" : "lodgify";
}

function computeLodgifyNetAmount_(grossAmount, feesTotal) {
    const gross = toNumberOrZero_(grossAmount);
    const fees = toNumberOrZero_(feesTotal);
    if (gross >= 0 && fees > gross) {
        return 0;
    }
    return Number((gross - fees).toFixed(2));
}

function extractHttpStatusFromErrorSafe_(msg) {
    if (typeof extractHttpStatusFromError_ === "function") {
        return extractHttpStatusFromError_(msg);
    }
    const match = String(msg || "").match(/\((\d{3})\)/);
    return match ? Number(match[1]) : null;
}

function extractAmountFromPaths_(item, directKeys, deepPaths) {
    const direct = firstDefined(item, directKeys || []);
    if (direct !== null && direct !== undefined && direct !== "") {
        return resolveAmountObject_(direct);
    }

    const nested = firstDefinedDeep(item, deepPaths || []);
    return resolveAmountObject_(nested);
}

function extractLodgifyFeesTotal_(item) {
    return extractAmountFromPaths_(item || {}, [
        "fees", "fee", "totalFees", "feesTotal", "serviceFee", "serviceFees",
        "commission", "commissionAmount", "channelFee", "hostFee"
    ], [
        "quote.fees", "quote.totalFees", "quote.feesTotal",
        "reservation.fees", "reservation.totalFees", "reservation.feesTotal",
        "financials.fees", "financials.totalFees", "financials.commission",
        "charges.fees", "charges.totalFees", "invoice.fees"
    ]);
}

function buildLodgifyEditableBookingFromSheetRow_(sheetName, rowNo, row, headerMap) {
    const bookingId = String(getRowValueByHeaders_(row, headerMap, ["LodgifyBookingId", "BookingId", "BuchungsID"], "") || "").trim();
    if (!bookingId) return null;

    const guestName = String(getRowValueByHeaders_(row, headerMap, ["GastName", "GuestName"], "") || "").trim();
    const checkin = formatDateCell_(getRowValueByHeaders_(row, headerMap, ["CheckIn", "Anreise"], ""));
    const checkout = formatDateCell_(getRowValueByHeaders_(row, headerMap, ["CheckOut", "Abreise"], ""));
    const bookingDate = formatDateCell_(getRowValueByHeaders_(row, headerMap, ["BookingDate", "Buchungstag", "Datum"], checkin || checkout || ""));
    const grossAmount = toNumberOrZero_(getRowValueByHeaders_(row, headerMap, ["Betrag", "GrossAmount"], 0));
    const feesTotal = toNumberOrZero_(getRowValueByHeaders_(row, headerMap, ["FeesTotal", "Gebuehren", "Fees"], 0));
    const netAmountRaw = getRowValueByHeaders_(row, headerMap, ["NetAmount", "Netto"], undefined);
    const payoutAmountRaw = getRowValueByHeaders_(row, headerMap, ["PayoutAmount", "Auszahlung"], undefined);
    const status = String(getRowValueByHeaders_(row, headerMap, ["Status"], "") || "").trim();
    const channel = String(getRowValueByHeaders_(row, headerMap, ["Kanal", "Channel"], "") || "").trim();
    const note = String(getRowValueByHeaders_(row, headerMap, ["Note", "Notiz"], "") || "").trim();
    const account = String(getRowValueByHeaders_(row, headerMap, ["Account", "Buchungskonto"], "Lodgify")).trim();
    const paymentOption = String(getRowValueByHeaders_(row, headerMap, ["PaymentOption", "Zahlungsoption"], "") || "").trim();
    const fullPaymentDays = String(getRowValueByHeaders_(row, headerMap, ["FullPaymentDaysBeforeCheckin", "VollbetragTageVorAnreise", "PaymentRequestDays"], "") || "").trim();
    const fullPaymentWeeks = String(getRowValueByHeaders_(row, headerMap, ["FullPaymentWeeksBeforeCheckin", "VollbetragWochenVorAnreise", "PaymentRequestWeeks"], "") || "").trim();
    const requestFullPayment = toBoolean(getRowValueByHeaders_(row, headerMap, ["RequestFullPayment", "VollbetragAnfordern", "ZahlungVollAnfordern"], false));
    const paymentRequestActive = toBoolean(getRowValueByHeaders_(row, headerMap, ["ZahlungsAufforderungAktiv", "PaymentRequestActive"], false));
    const paymentUpdateCompleted = String(getRowValueByHeaders_(row, headerMap, ["ZahlungsUpdateDurchgefuehrt"], "") || "").trim();
    const isExternal = toBoolean(getRowValueByHeaders_(row, headerMap, ["IsExternal", "Extern", "ExternBuchung", "NichtLodgify"], false));
    const defaultBookingType = defaultLodgifyBookingType_(isExternal);
    const bookingType = String(getRowValueByHeaders_(row, headerMap, ["BookingType"], defaultBookingType) || defaultBookingType).trim();
    const computedNetAmount = computeLodgifyNetAmount_(grossAmount, feesTotal);
    const netAmount = netAmountRaw === undefined ? computedNetAmount : toNumberOrZero_(netAmountRaw);
    const payoutAmount = payoutAmountRaw === undefined ? netAmount : toNumberOrZero_(payoutAmountRaw);

    return {
        id: `lodgify:${encodeURIComponent(sheetName)}:${rowNo}`,
        booking_type: bookingType || defaultBookingType,
        source_sheet: sheetName,
        source_row: rowNo,
        guest_name: guestName || bookingId,
        checkin: checkin,
        checkout: checkout || checkin,
        gross_amount: grossAmount,
        fees_total: feesTotal,
        net_amount: netAmount,
        payout_amount: payoutAmount,
        date: bookingDate || checkin || checkout || "",
        note: note || status,
        account: account || "Lodgify",
        lodgify_booking_id: bookingId,
        status: status,
        channel: channel,
        is_external: isExternal,
        payment_option: paymentOption,
        request_full_payment: requestFullPayment,
        full_payment_days_before_checkin: fullPaymentDays,
        full_payment_weeks_before_checkin: fullPaymentWeeks,
        payment_request_active: paymentRequestActive,
        payment_update_completed: paymentUpdateCompleted,
        raw: {
            lodgify_booking_id: bookingId,
            guest_name: guestName,
            checkin: checkin,
            checkout: checkout || checkin,
            booking_date: bookingDate || checkin || checkout || "",
            gross_amount: grossAmount,
            fees_total: feesTotal,
            net_amount: netAmount,
            payout_amount: payoutAmount,
            status: status,
            channel: channel,
            note: note,
            account: account || "Lodgify",
            is_external: isExternal,
            payment_request_active: paymentRequestActive,
            payment_option: paymentOption,
            request_full_payment: requestFullPayment,
            full_payment_days_before_checkin: fullPaymentDays,
            full_payment_weeks_before_checkin: fullPaymentWeeks,
            payment_update_completed: paymentUpdateCompleted,
            booking_type: bookingType || defaultBookingType
        }
    };
}

function mergeLodgifyEditableBooking_(currentBooking, booking) {
    const current = currentBooking || {};
    const currentRaw = current.raw || {};

    const guestName = getBookingFieldOverride_(booking, ["guest_name", "guestName"], ["guest_name", "guestName"]);
    const checkin = getBookingFieldOverride_(booking, ["checkin", "checkIn"], ["checkin", "checkIn"]);
    const checkout = getBookingFieldOverride_(booking, ["checkout", "checkOut"], ["checkout", "checkOut"]);
    const bookingDate = getBookingFieldOverride_(booking, ["date", "booking_date", "bookingDate"], ["date", "booking_date", "bookingDate"]);
    const grossAmount = getBookingFieldOverride_(booking, ["gross_amount", "grossAmount", "betrag", "amount"], ["gross_amount", "grossAmount", "betrag", "amount"]);
    const feesTotal = getBookingFieldOverride_(booking, ["fees_total", "feesTotal"], ["fees_total", "feesTotal"]);
    const netAmount = getBookingFieldOverride_(booking, ["net_amount", "netAmount"], ["net_amount", "netAmount"]);
    const payoutAmount = getBookingFieldOverride_(booking, ["payout_amount", "payoutAmount"], ["payout_amount", "payoutAmount"]);
    const status = getBookingFieldOverride_(booking, ["status"], ["status"]);
    const channel = getBookingFieldOverride_(booking, ["channel", "source"], ["channel", "source"]);
    const note = getBookingFieldOverride_(booking, ["note"], ["note"]);
    const account = getBookingFieldOverride_(booking, ["account"], ["account"]);
    const paymentOption = getBookingFieldOverride_(booking, ["payment_option", "paymentOption"], ["payment_option", "paymentOption"]);
    const requestFullPayment = getBookingFieldOverride_(booking, ["request_full_payment", "requestFullPayment"], ["request_full_payment", "requestFullPayment"]);
    const fullPaymentDays = getBookingFieldOverride_(booking, ["full_payment_days_before_checkin", "fullPaymentDaysBeforeCheckin"], ["full_payment_days_before_checkin", "fullPaymentDaysBeforeCheckin"]);
    const fullPaymentWeeks = getBookingFieldOverride_(booking, ["full_payment_weeks_before_checkin", "fullPaymentWeeksBeforeCheckin"], ["full_payment_weeks_before_checkin", "fullPaymentWeeksBeforeCheckin"]);
    const paymentRequestActive = getBookingFieldOverride_(booking, ["payment_request_active", "paymentRequestActive", "zahlungsaufforderung_aktiv"], ["payment_request_active", "paymentRequestActive", "zahlungsaufforderung_aktiv"]);
    const isExternalOverride = getBookingFieldOverride_(booking, ["is_external", "isExternal", "external"], ["is_external", "isExternal", "external"]);
    const bookingTypeOverride = getBookingFieldOverride_(booking, ["booking_type", "bookingType"], ["booking_type", "bookingType"]);
    const lodgifyBookingId = getBookingFieldOverride_(booking, ["lodgify_booking_id", "lodgifyBookingId", "lodgify_buchungs_id", "booking_id", "bookingId"], ["lodgify_booking_id", "lodgifyBookingId", "lodgify_buchungs_id", "booking_id", "bookingId"]);

    const merged = {
        id: current.id,
        booking_type: bookingTypeOverride !== undefined
            ? String(bookingTypeOverride || "")
            : (isExternalOverride !== undefined
                ? (toBoolean(isExternalOverride) ? "lodgify_external" : "lodgify")
                : (current.booking_type || currentRaw.booking_type || "lodgify")),
        source_sheet: current.source_sheet,
        source_row: current.source_row,
        guest_name: guestName !== undefined ? String(guestName || "") : String(current.guest_name || currentRaw.guest_name || ""),
        checkin: checkin !== undefined ? String(checkin || "") : String(current.checkin || currentRaw.checkin || ""),
        checkout: checkout !== undefined ? String(checkout || "") : String(current.checkout || currentRaw.checkout || ""),
        gross_amount: grossAmount !== undefined ? toNumberOrZero_(grossAmount) : toNumberOrZero_(current.gross_amount || currentRaw.gross_amount),
        fees_total: feesTotal !== undefined ? toNumberOrZero_(feesTotal) : toNumberOrZero_(current.fees_total || currentRaw.fees_total),
        net_amount: netAmount !== undefined ? toNumberOrZero_(netAmount) : toNumberOrZero_(current.net_amount || currentRaw.net_amount),
        payout_amount: payoutAmount !== undefined ? toNumberOrZero_(payoutAmount) : toNumberOrZero_(current.payout_amount || currentRaw.payout_amount),
        date: bookingDate !== undefined ? String(bookingDate || "") : String(current.date || currentRaw.booking_date || ""),
        note: note !== undefined ? String(note || "") : String(current.note || currentRaw.note || ""),
        account: account !== undefined ? String(account || "") : String(current.account || currentRaw.account || "Lodgify"),
        lodgify_booking_id: lodgifyBookingId !== undefined ? String(lodgifyBookingId || "") : String(current.lodgify_booking_id || currentRaw.lodgify_booking_id || ""),
        status: status !== undefined ? String(status || "") : String(current.status || currentRaw.status || ""),
        channel: channel !== undefined ? String(channel || "") : String(current.channel || currentRaw.channel || ""),
        is_external: isExternalOverride !== undefined ? toBoolean(isExternalOverride) : toBoolean(current.is_external || currentRaw.is_external),
        payment_option: paymentOption !== undefined ? String(paymentOption || "") : String(current.payment_option || currentRaw.payment_option || ""),
        request_full_payment: requestFullPayment !== undefined ? toBoolean(requestFullPayment) : toBoolean(current.request_full_payment || currentRaw.request_full_payment),
        full_payment_days_before_checkin: fullPaymentDays !== undefined ? String(fullPaymentDays || "") : String(current.full_payment_days_before_checkin || currentRaw.full_payment_days_before_checkin || ""),
        full_payment_weeks_before_checkin: fullPaymentWeeks !== undefined ? String(fullPaymentWeeks || "") : String(current.full_payment_weeks_before_checkin || currentRaw.full_payment_weeks_before_checkin || ""),
        payment_request_active: paymentRequestActive !== undefined ? toBoolean(paymentRequestActive) : toBoolean(current.payment_request_active || currentRaw.payment_request_active),
        payment_update_completed: String(current.payment_update_completed || currentRaw.payment_update_completed || "")
    };

    if (!merged.booking_type) {
        merged.booking_type = merged.is_external ? "lodgify_external" : "lodgify";
    }

    if (!merged.checkout && merged.checkin) {
        merged.checkout = merged.checkin;
    }
    if (!merged.date) {
        merged.date = merged.checkin || merged.checkout || "";
    }
    if (!merged.account) {
        merged.account = "Lodgify";
    }

    merged.raw = {
        lodgify_booking_id: merged.lodgify_booking_id,
        guest_name: merged.guest_name,
        checkin: merged.checkin,
        checkout: merged.checkout,
        booking_date: merged.date,
        gross_amount: merged.gross_amount,
        fees_total: merged.fees_total,
        net_amount: merged.net_amount,
        payout_amount: merged.payout_amount,
        status: merged.status,
        channel: merged.channel,
        note: merged.note,
        account: merged.account,
        is_external: merged.is_external,
        payment_request_active: merged.payment_request_active,
        payment_option: merged.payment_option,
        request_full_payment: merged.request_full_payment,
        full_payment_days_before_checkin: merged.full_payment_days_before_checkin,
        full_payment_weeks_before_checkin: merged.full_payment_weeks_before_checkin,
        payment_update_completed: merged.payment_update_completed,
        booking_type: merged.booking_type
    };

    return merged;
}

function shouldTriggerLodgifyPaymentUpdate_(booking) {
    const trigger = getBookingFieldOverride_(booking, [
        "trigger_lodgify_payment",
        "triggerLodgifyPayment",
        "trigger_payment_request",
        "triggerPaymentRequest",
        "apply_payment_request_now",
        "applyPaymentRequestNow"
    ], [
        "trigger_lodgify_payment",
        "triggerLodgifyPayment",
        "trigger_payment_request",
        "triggerPaymentRequest",
        "apply_payment_request_now",
        "applyPaymentRequestNow"
    ]);

    return trigger !== undefined ? toBoolean(trigger) : false;
}

function buildLodgifyPaymentPatchPayloadCandidates_(booking) {
    const paymentOption = String(booking.payment_option || "").trim();
    const requestFullPayment = toBoolean(booking.request_full_payment);
    const fullPaymentDays = String(booking.full_payment_days_before_checkin || "").trim();
    const fullPaymentWeeks = String(booking.full_payment_weeks_before_checkin || "").trim();

    const snakePayload = { payment_option: {} };
    const camelPayload = { paymentOption: {} };

    if (paymentOption) {
        snakePayload.payment_option.payment_option = paymentOption;
        camelPayload.paymentOption.paymentOption = paymentOption;
    }

    snakePayload.payment_option.request_full_payment = requestFullPayment;
    camelPayload.paymentOption.requestFullPayment = requestFullPayment;

    if (fullPaymentDays !== "") {
        snakePayload.payment_option.full_payment_days_before_checkin = Number(fullPaymentDays);
        camelPayload.paymentOption.fullPaymentDaysBeforeCheckin = Number(fullPaymentDays);
    }

    if (fullPaymentWeeks !== "") {
        snakePayload.payment_option.full_payment_weeks_before_checkin = Number(fullPaymentWeeks);
        camelPayload.paymentOption.fullPaymentWeeksBeforeCheckin = Number(fullPaymentWeeks);
    }

    return [snakePayload, camelPayload];
}

function buildLodgifyPaymentPatchPathCandidates_(bookingId) {
    const props = PropertiesService.getScriptProperties();
    const configured = String(props.getProperty("LODGIFY_PAYMENT_PATCH_PATH_TEMPLATE") || "").trim();
    const rawCandidates = [
        configured,
        "/api/reservations/{reservationId}",
        "/api/reservations/{id}",
        "/v2/reservations/{reservationId}",
        "/v2/reservations/{id}",
        "/v2/reservations/bookings/{reservationId}",
        "/v2/reservations/bookings/{id}",
        "/v1/reservation/booking/{reservationId}",
        "/v1/reservation/booking/{id}"
    ];

    const resolved = [];
    rawCandidates.forEach(function (template) {
        if (!template) return;
        const path = String(template)
            .replace(/\{reservationId\}/g, encodeURIComponent(bookingId))
            .replace(/\{id\}/g, encodeURIComponent(bookingId))
            .replace(/:reservationId/g, encodeURIComponent(bookingId))
            .replace(/:id/g, encodeURIComponent(bookingId));
        if (resolved.indexOf(path) === -1) {
            resolved.push(path);
        }
    });

    return resolved;
}

function triggerLodgifyPaymentUpdate_(booking) {
    const bookingId = String(booking.lodgify_booking_id || booking.id || "").trim();
    if (!bookingId) {
        throw new Error("LodgifyBookingId fehlt für die angeforderte Zahlungsaktualisierung.");
    }

    const pathCandidates = buildLodgifyPaymentPatchPathCandidates_(bookingId);
    const payloadCandidates = buildLodgifyPaymentPatchPayloadCandidates_(booking);
    const attempts = [];
    let lastError = null;

    for (let i = 0; i < pathCandidates.length; i++) {
        const path = pathCandidates[i];
        for (let j = 0; j < payloadCandidates.length; j++) {
            const payload = payloadCandidates[j];
            try {
                const response = lodgifyRequest(path, {
                    method: "patch",
                    payload: payload
                });
                return {
                    ok: true,
                    path: path,
                    status: response.status,
                    payload: payload,
                    response: response.body
                };
            } catch (err) {
                const msg = String(err && err.message ? err.message : err);
                attempts.push(`PATCH ${path}: ${msg}`);
                lastError = msg;
                const status = extractHttpStatusFromErrorSafe_(msg);
                if (status === 404 || status === 405) {
                    continue;
                }
            }
        }
    }

    const attemptPreview = attempts.slice(-4).join(" | ");
    const attemptedPaths = pathCandidates.join(", ");
    Logger.log(`Lodgify payment patch paths tried (${bookingId}): ${attemptedPaths}`);
    const attemptedPathsPreview = attemptedPaths.length > 200 ? attemptedPaths.slice(0, 200) + "..." : attemptedPaths;
    throw new Error(`Lodgify Payment-Update fehlgeschlagen (${bookingId}, Versuche=${attempts.length}, Pfade=${attemptedPathsPreview}): ${attemptPreview || lastError || "unbekannter Fehler"}`);
}

function updateLodgifyEditableBookingRow_(sheetName, rowNo, booking) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("No active spreadsheet");

    const sheet = ensureAlleBuchungenSheet_(sheetName);
    if (!sheet) throw new Error("Sheet not found: " + sheetName);

    const headerMap = getSheetHeaderMap_(sheet);
    const currentRow = sheet.getRange(rowNo, 1, 1, sheet.getLastColumn()).getValues()[0];
    const currentBooking = buildLodgifyEditableBookingFromSheetRow_(sheetName, rowNo, currentRow, headerMap);
    if (!currentBooking) {
        throw new Error("Lodgify-Buchung in Tabelle " + sheetName + " in Zeile " + rowNo + " nicht gefunden.");
    }

    const merged = mergeLodgifyEditableBooking_(currentBooking, booking || {});
    const updates = [
        { headers: ["LodgifyBookingId", "BookingId", "BuchungsID"], value: String(merged.lodgify_booking_id || "") },
        { headers: ["GastName", "GuestName"], value: String(merged.guest_name || "") },
        { headers: ["CheckIn", "Anreise"], value: parseDateOrBlank_(merged.checkin) },
        { headers: ["CheckOut", "Abreise"], value: parseDateOrBlank_(merged.checkout) },
        { headers: ["Betrag", "GrossAmount"], value: toNumberOrZero_(merged.gross_amount) },
        { headers: ["Status"], value: String(merged.status || "") },
        { headers: ["Kanal", "Channel"], value: String(merged.channel || "") },
        { headers: ["IsExternal", "Extern", "ExternBuchung", "NichtLodgify"], value: merged.is_external ? "TRUE" : "FALSE" },
        { headers: ["ZahlungsAufforderungAktiv", "PaymentRequestActive"], value: merged.payment_request_active ? "TRUE" : "FALSE" },
        { headers: ["PaymentOption", "Zahlungsoption"], value: String(merged.payment_option || "") },
        { headers: ["RequestFullPayment", "VollbetragAnfordern", "ZahlungVollAnfordern"], value: merged.request_full_payment ? "TRUE" : "FALSE" },
        { headers: ["FullPaymentDaysBeforeCheckin", "VollbetragTageVorAnreise", "PaymentRequestDays"], value: String(merged.full_payment_days_before_checkin || "") },
        { headers: ["FullPaymentWeeksBeforeCheckin", "VollbetragWochenVorAnreise", "PaymentRequestWeeks"], value: String(merged.full_payment_weeks_before_checkin || "") },
        { headers: ["BookingType"], value: String(merged.booking_type || (merged.is_external ? "lodgify_external" : "lodgify")) },
        { headers: ["BookingDate", "Buchungstag", "Datum"], value: parseDateOrBlank_(merged.date) },
        { headers: ["Note", "Notiz"], value: String(merged.note || "") },
        { headers: ["Account", "Buchungskonto"], value: String(merged.account || "Lodgify") },
        { headers: ["FeesTotal", "Gebuehren", "Fees"], value: toNumberOrZero_(merged.fees_total) },
        { headers: ["NetAmount", "Netto"], value: toNumberOrZero_(merged.net_amount) },
        { headers: ["PayoutAmount", "Auszahlung"], value: toNumberOrZero_(merged.payout_amount) }
    ];

    const updateResult = updateSheetRowByHeaders(sheetName, rowNo, updates);
    let paymentTriggerResult = null;

    if (shouldTriggerLodgifyPaymentUpdate_(booking || {})) {
        if (hasPaymentRequestTimestamp_(merged.payment_update_completed)) {
            paymentTriggerResult = {
                ok: true,
                skipped: true,
                reason: "alreadyRequested",
                requestedAt: merged.payment_update_completed
            };
        } else {
            paymentTriggerResult = triggerLodgifyPaymentUpdate_(merged);
            updateSheetRowByHeaders(sheetName, rowNo, [{
                headers: ["ZahlungsUpdateDurchgefuehrt"],
                value: buildPaymentRequestTimestamp_()
            }]);
        }
    }

    return {
        ok: true,
        id: currentBooking.id,
        kind: "lodgify",
        sheet: sheetName,
        row: rowNo,
        updated: updateResult,
        paymentTriggered: !!(paymentTriggerResult && !paymentTriggerResult.skipped),
        paymentTriggerResult: paymentTriggerResult
    };
}

/**
 * Mappt ein Lodgify-Buchungsobjekt auf eine AlleBuchungen-Zeile.
 * Gibt null zurück wenn keine BookingId vorhanden ist.
 */
function mapLodgifyItemToAlleBuchungenRow_(item) {
    if (!item || typeof item !== "object") return null;

    const id = extractLodgifyBookingId_(item);
    if (!id) return null;

    const guestName = extractLodgifyGuestName_(item);
    const checkinDate = extractLodgifyCheckinDate_(item);
    const checkoutDate = extractLodgifyCheckoutDate_(item);

    const amount = extractAmountFromPaths_(item, [
        "total", "grandTotal", "grand_total", "totalAmount", "total_amount",
        "price", "bookingAmount", "booking_amount", "amountToPay", "amount_to_pay",
        "amount", "amountDue", "amount_due"
    ], [
        "quote.total", "quote.totalAmount", "quote.total_amount",
        "reservation.total", "reservation.totalAmount", "reservation.total_amount",
        "financials.total", "financials.totalAmount", "financials.total_amount",
        "charges.total", "invoice.total"
    ]);
    const feesTotal = extractLodgifyFeesTotal_(item);
    const payoutAmountDirect = firstDefined(item, [
        "payout", "payoutAmount", "ownerPayout", "hostPayout", "netAmount", "net"
    ]);
    const payoutAmountNested = firstDefinedDeep(item, [
        "quote.payout", "quote.net", "quote.netAmount",
        "reservation.payout", "reservation.net", "reservation.netAmount",
        "financials.payout", "financials.net", "financials.netAmount",
        "charges.payout", "invoice.net"
    ]);
    const hasPayoutAmount = isValuePresent_(payoutAmountDirect) || isValuePresent_(payoutAmountNested);
    const payoutAmountRaw = hasPayoutAmount
        ? resolveAmountObject_(isValuePresent_(payoutAmountDirect)
            ? payoutAmountDirect
            : payoutAmountNested)
        : 0;
    const computedNetAmount = computeLodgifyNetAmount_(amount, feesTotal);
    const netAmount = hasPayoutAmount ? payoutAmountRaw : computedNetAmount;
    const payoutAmount = hasPayoutAmount ? payoutAmountRaw : netAmount;

    const status = String(
        firstDefined(item, [
            "status", "bookingStatus", "booking_status",
            "reservationStatus", "reservation_status", "state"
        ]) || ""
    ).trim();

    const channel = String(
        firstDefined(item, ["source", "channel", "origin", "source_text"]) || ""
    ).trim();

    const isExternal = toBoolean(
        firstDefined(item, ["is_external", "isExternal", "external"])
    );

    const paymentOption = String(
        firstDefined(item, ["payment_option", "paymentOption"]) || ""
    ).trim();

    const requestFullPayment = toBoolean(
        firstDefined(item, ["request_full_payment", "requestFullPayment"])
    );

    const fullPaymentDays = String(
        firstDefined(item, ["full_payment_days_before_checkin", "fullPaymentDaysBeforeCheckin"]) || ""
    ).trim();

    const fullPaymentWeeks = String(
        firstDefined(item, ["full_payment_weeks_before_checkin", "fullPaymentWeeksBeforeCheckin"]) || ""
    ).trim();

    const bookingDate = extractBuchungstagDate_(item) || checkinDate || checkoutDate;
    const bookingType = defaultLodgifyBookingType_(isExternal);
    const note = status && channel ? `${status} | ${channel}` : (status || channel || "");

    return {
        id,
        row: [
            id,                                         // LodgifyBookingId
            guestName,                                  // GastName
            checkinDate || "",                         // CheckIn
            checkoutDate || "",                        // CheckOut
            Number(amount.toFixed(2)),                  // Betrag
            status,                                     // Status
            channel,                                    // Kanal
            isExternal ? "TRUE" : "FALSE",           // IsExternal
            "",                                         // ZahlungsAufforderungAktiv (wird NICHT überschrieben)
            paymentOption,                              // PaymentOption
            requestFullPayment ? "TRUE" : "FALSE",    // RequestFullPayment
            fullPaymentDays,                            // FullPaymentDaysBeforeCheckin
            fullPaymentWeeks,                           // FullPaymentWeeksBeforeCheckin
            "",                                         // ZahlungsUpdateDurchgefuehrt
            bookingType,                                // BookingType
            bookingDate || "",                         // BookingDate
            note,                                       // Note
            "Lodgify",                                 // Account
            Number(feesTotal.toFixed(2)),               // FeesTotal
            Number(netAmount.toFixed(2)),               // NetAmount
            Number(payoutAmount.toFixed(2))             // PayoutAmount
        ],
        checkinDate
    };
}

/**
 * Befüllt/aktualisiert das AlleBuchungen-Sheet aus einer Liste von Lodgify-Items.
 * Die Spalte "ZahlungsAufforderungAktiv" wird dabei NICHT überschrieben (manuelle Eingabe).
 *
 * @param {string} sheetName  Name des Ziel-Sheets
 * @param {Array}  items      Lodgify-Buchungsobjekte
 * @returns {{ inserted: number, updated: number }}
 */
function upsertAlleBuchungenFromItems_(sheetName, items) {
    if (!items || items.length === 0) return { inserted: 0, updated: 0 };

    const sheet = ensureAlleBuchungenSheet_(sheetName);
    const lastRow = sheet.getLastRow();
    const numCols = ALLE_BUCHUNGEN_HEADERS_.length;

    // Vorhandene Daten lesen
    const existingData = lastRow > 1
        ? sheet.getRange(2, 1, lastRow - 1, numCols).getValues()
        : [];

    // Index aufbauen: bookingId -> 0-basierter Datenzeilen-Index
    const existingById = {};
    existingData.forEach(function (row, idx) {
        const id = String(row[0] || "").trim();
        if (id) existingById[id] = idx;
    });

    let inserted = 0;
    let updated = 0;
    const newRows = [];
    const rowUpdates = []; // { dataIdx, values }

    items.forEach(function (item) {
        const mapped = mapLodgifyItemToAlleBuchungenRow_(item);
        if (!mapped) return;

        if (Object.prototype.hasOwnProperty.call(existingById, mapped.id)) {
            const dataIdx = existingById[mapped.id];
            const existingRow = existingData[dataIdx];

            // Neue Werte übernehmen, aber ZahlungsAufforderungAktiv und
            // ZahlungsUpdateDurchgefuehrt aus dem bestehenden Eintrag beibehalten
            const newRowValues = mapped.row.slice();
            preserveExistingAlleBuchungenCellValue_(newRowValues, existingRow, ALLE_BUCHUNGEN_GUEST_NAME_COL_IDX_);
            preserveExistingAlleBuchungenCellValue_(newRowValues, existingRow, ALLE_BUCHUNGEN_CHECKIN_COL_IDX_);
            preserveExistingAlleBuchungenCellValue_(newRowValues, existingRow, ALLE_BUCHUNGEN_CHECKOUT_COL_IDX_);
            newRowValues[ALLE_BUCHUNGEN_MARKER_COL_IDX_] = existingRow[ALLE_BUCHUNGEN_MARKER_COL_IDX_];
            if (existingRow[ALLE_BUCHUNGEN_TIMESTAMP_COL_IDX_]) {
                newRowValues[ALLE_BUCHUNGEN_TIMESTAMP_COL_IDX_] = existingRow[ALLE_BUCHUNGEN_TIMESTAMP_COL_IDX_];
            }

            rowUpdates.push({ dataIdx, values: newRowValues });
            existingData[dataIdx] = newRowValues;
            updated++;
        } else {
            newRows.push(mapped.row);
            existingById[mapped.id] = existingData.length + newRows.length - 1;
            inserted++;
        }
    });

    // Bestehende Zeilen aktualisieren
    rowUpdates.forEach(function (update) {
        const sheetRow = update.dataIdx + 2; // +1 für Header-Zeile, +1 für 1-basiert
        sheet.getRange(sheetRow, 1, 1, numCols).setValues([update.values]);
    });

    // Neue Zeilen anhängen
    if (newRows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, numCols).setValues(newRows);
    }

    Logger.log(
        `AlleBuchungen upsert (${sheetName}): inserted=${inserted}, updated=${updated}`
    );
    return { inserted, updated };
}

/**
 * Scannt das AlleBuchungen-Sheet nach externen Lodgify-Buchungen mit aktivem
 * "RequestFullPayment" und löst die Zahlungsanforderung im passenden Zeitfenster aus.
 *
 * @param {string} sheetName   Name des Sheets
 * @param {Object} itemsById   Map: bookingId -> Lodgify-Buchungsobjekt
 * @param {Object} config      { daysBeforeCheckin: number }
 * @returns {{ applied: number, skippedNoData: number, skippedWindow: number, skippedAlreadyRequested: number, skippedIneligible: number }}
 */
function applyPaymentRequestUpdates_(sheetName, itemsById, config) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { applied: 0, skippedNoData: 0, skippedWindow: 0, skippedAlreadyRequested: 0, skippedIneligible: 0 };

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { applied: 0, skippedNoData: 0, skippedWindow: 0, skippedAlreadyRequested: 0, skippedIneligible: 0 };

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { applied: 0, skippedNoData: 0, skippedWindow: 0, skippedAlreadyRequested: 0, skippedIneligible: 0 };

    const numCols = sheet.getLastColumn();
    const allData = sheet.getRange(1, 1, lastRow, numCols).getValues();
    const headers = allData[0].map(function (h) { return String(h || "").trim(); });

    const bookingIdColIdx = headers.indexOf("LodgifyBookingId");
    const timestampColIdx = headers.indexOf("ZahlungsUpdateDurchgefuehrt");

    if (bookingIdColIdx === -1) {
        Logger.log(
            "⚠️ AlleBuchungen: Benötigte Spalte LodgifyBookingId nicht gefunden."
        );
        return { applied: 0, skippedNoData: 0, skippedWindow: 0, skippedAlreadyRequested: 0, skippedIneligible: 0 };
    }

    let applied = 0;
    let skippedNoData = 0;
    let skippedWindow = 0;
    let skippedAlreadyRequested = 0;
    let skippedIneligible = 0;

    for (let i = 1; i < allData.length; i++) {
        const row = allData[i];
        const bookingId = String(row[bookingIdColIdx] || "").trim();
        if (!bookingId) continue;

        const booking = itemsById[bookingId];
        if (!booking) {
            Logger.log(
                `⚠️ AlleBuchungen Zahlungsupdate: Buchung ${bookingId} nicht in aktuellen Lodgify-Daten gefunden.`
            );
            skippedNoData++;
            continue;
        }

        const paymentUpdateCompleted = timestampColIdx === -1 ? "" : row[timestampColIdx];
        const evaluation = evaluateAutomaticPaymentRequest_(booking, paymentUpdateCompleted, config);
        if (!evaluation.shouldRequest) {
            if (evaluation.reason === "alreadyRequested") {
                skippedAlreadyRequested++;
            } else if (evaluation.reason === "outsideWindow") {
                skippedWindow++;
            } else {
                skippedIneligible++;
            }
            continue;
        }

        const sheetRow = i + 1; // 1-basiert
        const paymentTriggerResult = triggerLodgifyPaymentUpdate_(booking);
        if (!paymentTriggerResult || paymentTriggerResult.ok !== true) {
            throw new Error(`Lodgify Zahlungsanforderung für Buchung ${bookingId} wurde nicht bestätigt.`);
        }

        if (timestampColIdx !== -1) {
            sheet.getRange(sheetRow, timestampColIdx + 1).setValue(
                buildPaymentRequestTimestamp_()
            );
        }

        Logger.log(
            `✅ AlleBuchungen Zahlungsupdate: Buchung ${bookingId} (Zeile ${sheetRow}) in Lodgify angefordert.`
        );
        applied++;
    }

    Logger.log(
        `AlleBuchungen Zahlungsupdate: applied=${applied}, skippedWindow=${skippedWindow}, skippedNoData=${skippedNoData}, skippedAlreadyRequested=${skippedAlreadyRequested}, skippedIneligible=${skippedIneligible}`
    );
    return { applied, skippedNoData, skippedWindow, skippedAlreadyRequested, skippedIneligible };
}

/**
 * Standalone-Funktion: Holt aktuelle Lodgify-Daten, befüllt AlleBuchungen und
 * führt automatische Zahlungsupdates für Zeilen mit gesetztem Kenner durch.
 *
 * Kann als Apps-Script-Trigger oder per API-Aufruf (action: "processPaymentRequests") ausgeführt werden.
 *
 * @param {Object} [params]  Optionale Query-Parameter (z.B. { page, size, from, to })
 */
function processLodgifyPaymentRequestUpdates(params) {
    const config = getPaymentRequestConfig_();
    const queryParams = params || {};

    const bookingsResult = fetchBookingsWithCloudFallback_(queryParams);
    const reservationsResult = fetchReservationsWithFallback_(queryParams);

    const combinedItems = bookingsResult.items.concat(reservationsResult.items);
    const deduped = dedupeBookingsById_(combinedItems);
    const allItems = deduped.items;

    if (allItems.length === 0 && (bookingsResult.error || reservationsResult.error)) {
        throw new Error(
            "Keine Lodgify-Daten geladen. " +
            [bookingsResult.error, reservationsResult.error].filter(Boolean).join(" | ")
        );
    }

    // AlleBuchungen-Sheet befüllen/aktualisieren
    const upsertResult = upsertAlleBuchungenFromItems_(config.sheetName, allItems);

    // Items-by-ID-Map für schnellen Zugriff
    const itemsById = {};
    allItems.forEach(function (item) {
        const id = String(
            firstDefined(item, ["id", "bookingId", "booking_id", "reservationId", "reservation_id"]) || ""
        ).trim();
        if (id) itemsById[id] = item;
    });

    // Zahlungsaufforderungs-Updates anwenden
    const applyResult = applyPaymentRequestUpdates_(config.sheetName, itemsById, config);

    Logger.log(
        `processLodgifyPaymentRequestUpdates: fetched=${allItems.length}, ` +
        `upsertInserted=${upsertResult.inserted}, upsertUpdated=${upsertResult.updated}, ` +
        `paymentUpdatesApplied=${applyResult.applied}`
    );

    return {
        ok: true,
        sheet: config.sheetName,
        config: { daysBeforeCheckin: config.daysBeforeCheckin },
        fetched: allItems.length,
        upsert: upsertResult,
        paymentUpdates: applyResult,
        sources: {
            bookingsFetched: bookingsResult.items.length,
            reservationsFetched: reservationsResult.items.length,
            bookingsWarning: bookingsResult.warning || null,
            reservationsWarning: reservationsResult.warning || null
        }
    };
}
