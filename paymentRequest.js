// ========== ZAHLUNGSAUFFORDERUNG / PAYMENT REQUEST ==========
//
// Diese Datei implementiert die automatische Zahlungsaufforderung:
//
// 1. Alle Lodgify-Buchungen werden in das Sheet "AlleBuchungen" (konfigurierbar)
//    eingepflegt/aktualisiert.
// 2. Die Spalte "ZahlungsAufforderungAktiv" kann manuell auf TRUE gesetzt werden.
// 3. Beim nächsten Lodgify-Import: Wenn der Kenner gesetzt ist UND wir uns im
//    konfigurierten Zeitfenster vor CheckIn befinden, werden die Zahlungsfelder
//    (PaymentOption, RequestFullPayment, …) aus den Lodgify-Daten übernommen.
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
    "ZahlungsUpdateDurchgefuehrt"
];

// 0-basierter Index der "ZahlungsAufforderungAktiv"-Spalte in ALLE_BUCHUNGEN_HEADERS_
var ALLE_BUCHUNGEN_MARKER_COL_IDX_ = 8;

// 0-basierter Index der "ZahlungsUpdateDurchgefuehrt"-Spalte
var ALLE_BUCHUNGEN_TIMESTAMP_COL_IDX_ = 13;

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
    const windowStart = new Date(checkin.getTime() - daysBeforeCheckin * 24 * 60 * 60 * 1000);

    // Auslösen wenn: jetzt >= Fensteranfang (d.h. wir sind im Fenster oder danach)
    return now >= windowStart;
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
    return sheet;
}

/**
 * Mappt ein Lodgify-Buchungsobjekt auf eine AlleBuchungen-Zeile.
 * Gibt null zurück wenn keine BookingId vorhanden ist.
 */
function mapLodgifyItemToAlleBuchungenRow_(item) {
    if (!item || typeof item !== "object") return null;

    const id = String(
        firstDefined(item, ["id", "bookingId", "booking_id", "reservationId", "reservation_id"]) || ""
    ).trim();
    if (!id) return null;

    const guestName = String(
        firstDefined(item, [
            "guestName", "guest_name", "customerName", "customer_name",
            "tenantName", "tenant_name", "name"
        ]) || ""
    ).trim();

    const checkinDate = parseDateOrNull(
        firstDefined(item, [
            "checkIn", "check_in", "checkInDate", "check_in_date",
            "arrival", "arrivalDate", "arrival_date",
            "startDate", "start_date", "from", "dateFrom", "date_from"
        ])
    );

    const checkoutDate = parseDateOrNull(
        firstDefined(item, [
            "checkOut", "check_out", "checkOutDate", "check_out_date",
            "departure", "departureDate", "departure_date",
            "endDate", "end_date", "to", "dateTo", "date_to"
        ])
    );

    const amountRaw = firstDefined(item, [
        "total", "grandTotal", "grand_total", "totalAmount", "total_amount",
        "price", "bookingAmount", "booking_amount", "amountToPay", "amount_to_pay",
        "amount", "amountDue", "amount_due"
    ]);
    const amountNested = firstDefinedDeep(item, [
        "quote.total", "quote.totalAmount", "quote.total_amount",
        "reservation.total", "reservation.totalAmount",
        "financials.total", "financials.totalAmount",
        "charges.total", "invoice.total"
    ]);
    const resolvedAmount = (amountRaw !== null && amountRaw !== undefined && amountRaw !== "")
        ? amountRaw : amountNested;
    const amount = resolveAmountObject_(resolvedAmount);

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

    return {
        id,
        row: [
            id,                                         // LodgifyBookingId
            guestName,                                  // GastName
            checkinDate || "",                          // CheckIn
            checkoutDate || "",                         // CheckOut
            Number(amount.toFixed(2)),                  // Betrag
            status,                                     // Status
            channel,                                    // Kanal
            isExternal ? "TRUE" : "FALSE",              // IsExternal
            "",                                         // ZahlungsAufforderungAktiv (wird NICHT überschrieben)
            paymentOption,                              // PaymentOption
            requestFullPayment ? "TRUE" : "FALSE",      // RequestFullPayment
            fullPaymentDays,                            // FullPaymentDaysBeforeCheckin
            fullPaymentWeeks,                           // FullPaymentWeeksBeforeCheckin
            ""                                          // ZahlungsUpdateDurchgefuehrt
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
 * Scannt das AlleBuchungen-Sheet nach Zeilen mit gesetztem "ZahlungsAufforderungAktiv"-Kenner
 * und schreibt bei passendem Zeitfenster die Zahlungsfelder aus den Lodgify-Daten.
 *
 * @param {string} sheetName   Name des Sheets
 * @param {Object} itemsById   Map: bookingId -> Lodgify-Buchungsobjekt
 * @param {Object} config      { daysBeforeCheckin: number }
 * @returns {{ applied: number, skippedNoData: number, skippedWindow: number }}
 */
function applyPaymentRequestUpdates_(sheetName, itemsById, config) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { applied: 0, skippedNoData: 0, skippedWindow: 0 };

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { applied: 0, skippedNoData: 0, skippedWindow: 0 };

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { applied: 0, skippedNoData: 0, skippedWindow: 0 };

    const numCols = sheet.getLastColumn();
    const allData = sheet.getRange(1, 1, lastRow, numCols).getValues();
    const headers = allData[0].map(function (h) { return String(h || "").trim(); });

    const markerColIdx = headers.indexOf("ZahlungsAufforderungAktiv");
    const checkinColIdx = headers.indexOf("CheckIn");
    const bookingIdColIdx = headers.indexOf("LodgifyBookingId");
    const timestampColIdx = headers.indexOf("ZahlungsUpdateDurchgefuehrt");

    if (markerColIdx === -1 || checkinColIdx === -1 || bookingIdColIdx === -1) {
        Logger.log(
            "⚠️ AlleBuchungen: Benötigte Spalten (ZahlungsAufforderungAktiv, CheckIn, LodgifyBookingId) nicht gefunden."
        );
        return { applied: 0, skippedNoData: 0, skippedWindow: 0 };
    }

    let applied = 0;
    let skippedNoData = 0;
    let skippedWindow = 0;

    for (let i = 1; i < allData.length; i++) {
        const row = allData[i];

        if (!toBoolean(row[markerColIdx])) continue;

        const checkinValue = row[checkinColIdx];
        if (!isWithinPaymentRequestWindow_(checkinValue, config.daysBeforeCheckin)) {
            skippedWindow++;
            continue;
        }

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

        const sheetRow = i + 1; // 1-basiert
        updateLodgifyBookingInSheet(sheetName, sheetRow, booking);

        // Zeitstempel setzen
        if (timestampColIdx !== -1) {
            sheet.getRange(sheetRow, timestampColIdx + 1).setValue(
                Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss")
            );
        }

        Logger.log(
            `✅ AlleBuchungen Zahlungsupdate: Buchung ${bookingId} (Zeile ${sheetRow}) aktualisiert.`
        );
        applied++;
    }

    Logger.log(
        `AlleBuchungen Zahlungsupdate: applied=${applied}, skippedWindow=${skippedWindow}, skippedNoData=${skippedNoData}`
    );
    return { applied, skippedNoData, skippedWindow };
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
