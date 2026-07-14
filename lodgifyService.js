function getLodgifyConfig() {
    const props = PropertiesService.getScriptProperties();
    const apiKey = (props.getProperty("LODGIFY_API_KEY") || "").trim();
    const baseUrl = (props.getProperty("LODGIFY_API_BASE_URL") || "https://api.lodgify.com").trim();
    const timeoutMs = Number(props.getProperty("LODGIFY_TIMEOUT_MS") || 30000);
    const propertyId = (props.getProperty("LODGIFY_PROPERTY_ID") || "").trim();
    const sheetName = (props.getProperty("LODGIFY_SHEET_NAME") || "Lodgify").trim();
    const auditSheetName = (props.getProperty("LODGIFY_AUDIT_SHEET_NAME") || "Lodgify_Audit").trim();
    const bookingsPath = (props.getProperty("LODGIFY_BOOKINGS_PATH") || "/v2/reservations/bookings").trim();
    const reservationsPath = (props.getProperty("LODGIFY_RESERVATIONS_PATH") || "/v1/reservation").trim();

    return {
        apiKey,
        baseUrl: baseUrl.replace(/\/+$/, ""),
        timeoutMs: Number.isFinite(timeoutMs) && timeoutMs > 0 ? timeoutMs : 30000,
        propertyId,
        sheetName: sheetName || "Lodgify",
        auditSheetName: auditSheetName || "Lodgify_Audit",
        bookingsPath: bookingsPath || "/v2/reservations/bookings",
        reservationsPath: reservationsPath || "/v1/reservation"
    };
}

function validateLodgifyConfig() {
    const config = getLodgifyConfig();
    if (!config.apiKey) {
        throw new Error("Missing Script Property: LODGIFY_API_KEY");
    }
    if (!config.baseUrl) {
        throw new Error("Missing Script Property: LODGIFY_API_BASE_URL");
    }
    return config;
}

function setLodgifyConfig(apiKey, baseUrl, timeoutMs) {
    if (!apiKey || !String(apiKey).trim()) {
        throw new Error("apiKey must not be empty");
    }

    const props = PropertiesService.getScriptProperties();
    props.setProperty("LODGIFY_API_KEY", String(apiKey).trim());

    if (baseUrl && String(baseUrl).trim()) {
        props.setProperty("LODGIFY_API_BASE_URL", String(baseUrl).trim());
    }

    if (timeoutMs !== undefined && timeoutMs !== null) {
        props.setProperty("LODGIFY_TIMEOUT_MS", String(timeoutMs));
    }
}

function setLodgifyEndpointPaths(bookingsPath, reservationsPath) {
    const props = PropertiesService.getScriptProperties();

    if (bookingsPath && String(bookingsPath).trim()) {
        props.setProperty("LODGIFY_BOOKINGS_PATH", String(bookingsPath).trim());
    }

    if (reservationsPath && String(reservationsPath).trim()) {
        props.setProperty("LODGIFY_RESERVATIONS_PATH", String(reservationsPath).trim());
    }
}

function lodgifyBuildUrl(path, queryParams) {
    const config = validateLodgifyConfig();
    const normalizedPath = path.startsWith("/") ? path : `/${path}`;
    let baseUrl = config.baseUrl;

    // Vermeidet doppelte Versionssegmente wie /v1/v2/... wenn die Base-URL bereits versioniert ist.
    if (/\/v\d+$/i.test(baseUrl) && /^\/v\d+\//i.test(normalizedPath)) {
        baseUrl = baseUrl.replace(/\/v\d+$/i, "");
    }

    const base = `${baseUrl}${normalizedPath}`;

    if (!queryParams) return base;

    const parts = [];
    Object.keys(queryParams).forEach(key => {
        const value = queryParams[key];
        if (value === undefined || value === null || value === "") return;
        parts.push(`${encodeURIComponent(key)}=${encodeURIComponent(String(value))}`);
    });

    if (parts.length === 0) return base;
    return `${base}?${parts.join("&")}`;
}

function lodgifyRequest(path, options) {
    const config = validateLodgifyConfig();
    const requestOptions = options || {};
    const method = (requestOptions.method || "get").toLowerCase();

    const fetchOptions = {
        method,
        muteHttpExceptions: true,
        headers: {
            "X-ApiKey": config.apiKey,
            Accept: "application/json"
        }
    };

    if (requestOptions.payload !== undefined && requestOptions.payload !== null) {
        fetchOptions.contentType = "application/json";
        fetchOptions.payload = JSON.stringify(requestOptions.payload);
    }

    const url = lodgifyBuildUrl(path, requestOptions.queryParams);
    const response = UrlFetchApp.fetch(url, fetchOptions);

// ===== DEBUG START =====
Logger.log("===== LODGIFY REQUEST =====");
Logger.log("URL: " + url);
Logger.log("Status: " + response.getResponseCode());

Logger.log("===== HEADERS =====");
Logger.log(JSON.stringify(response.getAllHeaders(), null, 2));

Logger.log("===== BODY =====");
Logger.log(response.getContentText());
// ===== DEBUG END =====

const status = response.getResponseCode();
const bodyText = response.getContentText() || "";

let body;
try {
    body = bodyText ? JSON.parse(bodyText) : null;
} catch (e) {
    body = bodyText;
}

    if (status < 200 || status >= 300) {
        throw new Error(`Lodgify request failed (${status}) for ${url}: ${bodyText}`);
    }

    return {
        status,
        body,
        headers: response.getAllHeaders()
    };
}

function lodgifyGetBookings(queryParams) {

  const response = lodgifyRequest("/v2/reservations/bookings", {
    method: "get",
    queryParams: Object.assign({
      page: 1,
      size: 50,
      includeCount: true,
      stayFilter: "All",
      includeTransactions: false,
      includeExternal: true,
      includeQuoteDetails: true
    }, queryParams || {})
  });

  Logger.log("========== REQUEST ==========");
  Logger.log(JSON.stringify(queryParams || {}, null, 2));

  Logger.log("========== RESPONSE ==========");
  Logger.log(JSON.stringify(response.body, null, 2));

  return response;
}

function lodgifyGetReservations(queryParams) {
    const config = getLodgifyConfig();
    return lodgifyRequest(config.reservationsPath, {
        method: "get",
        queryParams: queryParams || {}
    });
}

function lodgifyHealthCheck() {
    return lodgifyGetBookings({ page: 1, size: 1 });
}

function importLodgifyEinnahmenToImport(queryParams) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
        throw new Error("No active spreadsheet");
    }

    const config = getLodgifyConfig();
    const targetSheetName = config.sheetName;
    const sheet = ss.getSheetByName(targetSheetName) || ss.insertSheet(targetSheetName);
    const params = queryParams || {};
    const excludeDeclinedCancelled = toBooleanWithDefault_(params.excludeDeclinedCancelled, true);

    const bookingsResult = fetchBookingsWithCloudFallback_(params);
    const reservationsResult = fetchReservationsWithFallback_(params);

    const combinedItems = bookingsResult.items.concat(reservationsResult.items);
    const statusSummaryBefore = summarizeBookingStatuses_(combinedItems);
    const confirmedRawItems = combinedItems.filter(item => isConfirmedBooking_(item, excludeDeclinedCancelled));
    const deduped = dedupeBookingsById_(confirmedRawItems);
    const allItems = deduped.items;
    const statusSummaryAfter = summarizeBookingStatuses_(allItems);
    const yearSummaryRaw = summarizeYearsFromItems_(combinedItems);
    const yearSummaryConfirmed = summarizeYearsFromItems_(allItems);

    if (allItems.length === 0 && (bookingsResult.error || reservationsResult.error)) {
        throw new Error(
            "No Lodgify data fetched. " +
            [bookingsResult.error, reservationsResult.error].filter(Boolean).join(" | ")
        );
    }

    if (sheet.getLastRow() === 0) {
        sheet.appendRow([
            "Konto",
            "Gegenkonto",
            "Buchungstag",
            "Wertstellung",
            "Buchungstext",
            "Verwendungszweck",
            "Beguenstigter/Zahlungspflichtiger",
            "Kontonummer",
            "BLZ",
            "Soll",
            "Haben"
        ]);
    }

    const existingRows = sheet.getDataRange().getValues().slice(1);
    const existingKeys = {};
    const existingByBookingId = {};
    existingRows.forEach((r, idx) => {
        const key = buildImportDedupKey(r[3], r[4], r[10]);
        if (key) existingKeys[key] = true;

        const bookingId = parseBookingIdFromText_(r[4]);
        if (bookingId) {
            // Sheet ist 1-basiert, Datenzeilen starten ab Zeile 2.
            const rowNumber = idx + 2;
            if (!existingByBookingId[bookingId]) {
                existingByBookingId[bookingId] = rowNumber;
            }
        }
    });

    const newRows = [];
    const rowUpdates = [];
    let skippedInvalid = 0;
    let skippedMissingDate = 0;
    let skippedDuplicate = 0;
    let updatedExisting = 0;
    let unchangedExisting = 0;
    let mappedAmountZero = 0;
    let mappedAmountPositive = 0;

    allItems.forEach(item => {
        const mappedDetails = mapLodgifyItemToImportRowDetailed_(item);
        const mapped = mappedDetails.row;
        if (!mapped) {
            if (mappedDetails.reason === "missingDate") skippedMissingDate++;
            else skippedInvalid++;
            return;
        }

        if (mappedDetails.amount > 0) mappedAmountPositive++;
        else mappedAmountZero++;

        const bookingId = mappedDetails.bookingId;
        if (bookingId && existingByBookingId[bookingId]) {
            const existingRowNumber = existingByBookingId[bookingId];
            const existingRow = existingRows[existingRowNumber - 2];
            if (isSameLodgifyRow_(existingRow, mapped)) {
                unchangedExisting++;
                skippedDuplicate++;
                return;
            }

            rowUpdates.push({ rowNumber: existingRowNumber, values: mapped });
            existingRows[existingRowNumber - 2] = mapped.slice();
            updatedExisting++;
            return;
        }

        const key = buildImportDedupKey(mapped[3], mapped[4], mapped[10]);
        if (!key || existingKeys[key]) {
            skippedDuplicate++;
            return;
        }

        existingKeys[key] = true;
        newRows.push(mapped);

        if (bookingId) {
            const newRowNumber = existingRows.length + newRows.length + 1;
            existingByBookingId[bookingId] = newRowNumber;
        }
    });

    if (rowUpdates.length > 0) {
        rowUpdates.forEach(update => {
            sheet.getRange(update.rowNumber, 1, 1, update.values.length).setValues([update.values]);
        });
    }

    if (newRows.length > 0) {
        sheet
            .getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length)
            .setValues(newRows);
    }

    Logger.log(`Lodgify import summary: raw=${combinedItems.length}, excludeDeclinedCancelled=${excludeDeclinedCancelled}, confirmedRaw=${confirmedRawItems.length}, confirmedUnique=${allItems.length}, sourceDuplicates=${deduped.duplicateCount}, inserted=${newRows.length}, updated=${updatedExisting}, unchanged=${unchangedExisting}, duplicateSkipped=${skippedDuplicate}, missingDate=${skippedMissingDate}, invalid=${skippedInvalid}, amountPositive=${mappedAmountPositive}, amountZero=${mappedAmountZero}`);
    Logger.log(`Lodgify status summary before filter: ${JSON.stringify(statusSummaryBefore)}`);
    Logger.log(`Lodgify status summary after filter: ${JSON.stringify(statusSummaryAfter)}`);
    Logger.log(`Lodgify year summary raw: ${JSON.stringify(yearSummaryRaw)}`);
    Logger.log(`Lodgify year summary confirmedUnique: ${JSON.stringify(yearSummaryConfirmed)}`);

    // Zahlungsaufforderung: AlleBuchungen-Sheet befüllen und Zahlungsupdates anwenden.
    // Fehler werden abgefangen, damit der Finance-Import nicht abbricht.
    let paymentRequestResult = null;
    try {
        const prConfig = getPaymentRequestConfig_();
        upsertAlleBuchungenFromItems_(prConfig.sheetName, allItems);

        const itemsById = {};
        allItems.forEach(item => {
            const id = String(
                firstDefined(item, ["id", "bookingId", "booking_id", "reservationId", "reservation_id"]) || ""
            ).trim();
            if (id) itemsById[id] = item;
        });

        paymentRequestResult = applyPaymentRequestUpdates_(prConfig.sheetName, itemsById, prConfig);
    } catch (prErr) {
        Logger.log(`⚠️ Zahlungsaufforderungs-Update fehlgeschlagen: ${String(prErr && prErr.message ? prErr.message : prErr)}`);
    }

    return {
        ok: true,
        sheet: targetSheetName,
        fetched: allItems.length,
        inserted: newRows.length,
        skipped: allItems.length - newRows.length,
        diagnostics: {
            rawFetched: combinedItems.length,
            excludeDeclinedCancelled,
            confirmedAfterFilter: confirmedRawItems.length,
            confirmedUnique: allItems.length,
            sourceDuplicateConfirmed: deduped.duplicateCount,
            updatedExisting,
            unchangedExisting,
            skippedDuplicate,
            skippedMissingDate,
            skippedInvalid,
            mappedAmountPositive,
            mappedAmountZero,
            statusSummaryBefore,
            statusSummaryAfter,
            yearSummaryRaw,
            yearSummaryConfirmed
        },
        sources: {
            bookingsFetched: bookingsResult.items.length,
            reservationsFetched: reservationsResult.items.length,
            bookingsWarning: bookingsResult.warning || null,
            reservationsWarning: reservationsResult.warning || null
        },
        paymentRequests: paymentRequestResult
    };
}

function auditLodgifyBookingsToSheet(queryParams) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
        throw new Error("No active spreadsheet");
    }

    const config = getLodgifyConfig();
    const sheet = ss.getSheetByName(config.auditSheetName) || ss.insertSheet(config.auditSheetName);
    const params = queryParams || {};
    const excludeDeclinedCancelled = toBooleanWithDefault_(params.excludeDeclinedCancelled, false);

    const bookingsResult = fetchBookingsWithCloudFallback_(params);
    const reservationsResult = fetchReservationsWithFallback_(params);

    const tagged = [];
    (bookingsResult.items || []).forEach(item => tagged.push({ sourceEndpoint: "bookings", item }));
    (reservationsResult.items || []).forEach(item => tagged.push({ sourceEndpoint: "reservations", item }));

    const deduped = dedupeTaggedBookingsById_(tagged);
    const rows = deduped.items.map(entry => mapAuditRow_(entry.item, entry.sourceEndpoint));

    sheet.clearContents();
    const header = [
        "sourceEndpoint",
        "bookingId",
        "status",
        "channel",
        "type",
        "bookingDate",
        "wertstellungDate",
        "yearBooking",
        "yearWertstellung",
        "amount",
        "includeWhenExcludeDeclinedCancelled",
        "rawSnippet"
    ];
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    }

    const statusSummary = summarizeBookingStatuses_(deduped.items.map(entry => entry.item));
    const yearSummary = summarizeYearsFromItems_(deduped.items.map(entry => entry.item));

    Logger.log(`Lodgify audit summary: rawTagged=${tagged.length}, unique=${deduped.items.length}, sourceDuplicates=${deduped.duplicateCount}, excludedByFilter=${countExcludedByFilter_(deduped.items.map(entry => entry.item), excludeDeclinedCancelled)}`);
    Logger.log(`Lodgify audit status summary: ${JSON.stringify(statusSummary)}`);
    Logger.log(`Lodgify audit year summary: ${JSON.stringify(yearSummary)}`);

    return {
        ok: true,
        sheet: config.auditSheetName,
        rawTagged: tagged.length,
        unique: deduped.items.length,
        sourceDuplicates: deduped.duplicateCount,
        excludeDeclinedCancelled,
        excludedByFilter: countExcludedByFilter_(deduped.items.map(entry => entry.item), excludeDeclinedCancelled),
        statusSummary,
        yearSummary,
        sources: {
            bookingsFetched: bookingsResult.items.length,
            reservationsFetched: reservationsResult.items.length,
            bookingsWarning: bookingsResult.warning || null,
            reservationsWarning: reservationsResult.warning || null
        }
    };
}

function fetchBookingsWithCloudFallback_(queryParams) {
    const config = getLodgifyConfig();
    const endpointCandidates = uniquePaths_([
        "/v2/reservations/bookings",
        "/v2/bookings",
        config.bookingsPath
    ]);

    const baseParams = {};
    Object.keys(queryParams || {}).forEach(key => {
        const value = queryParams[key];
        if (value !== undefined && value !== null && value !== "") {
            baseParams[key] = value;
        }
    });
    baseParams.includeCanceled = "false";
    if (baseParams.includeCount === undefined) baseParams.includeCount = true;
    if (baseParams.stayFilter === undefined) baseParams.stayFilter = "All";
    if (baseParams.includeTransactions === undefined) baseParams.includeTransactions = false;
    if (baseParams.includeExternal === undefined) baseParams.includeExternal = true;
    if (baseParams.includeQuoteDetails === undefined) baseParams.includeQuoteDetails = true;
    if (config.propertyId) {
        baseParams.propertyId = config.propertyId;
    }

    const warnings = [];
    let lastError = null;

    for (let i = 0; i < endpointCandidates.length; i++) {
        const endpoint = endpointCandidates[i];
        try {
            const items = fetchAllPagesFromEndpoint_(endpoint, baseParams);
            return {
                items,
                warning: endpoint === config.bookingsPath ? null : `bookings fallback endpoint used: GET ${endpoint}`,
                error: null
            };
        } catch (err) {
            const msg = String(err && err.message ? err.message : err);
            lastError = msg;
            const status = extractHttpStatusFromError_(msg);

            if (status === 404 || status === 405) {
                warnings.push(`GET ${endpoint}: ${status}`);
                continue;
            }

            if (status === 401 || status === 403) {
                warnings.push(`GET ${endpoint}: ${status}`);
                continue;
            }

            return {
                items: [],
                warning: null,
                error: `bookings: ${msg}`
            };
        }
    }

    return {
        items: [],
        warning: warnings.length
            ? `bookings unavailable on known paths (${warnings.join(", ")})`
            : "bookings endpoint not found on known paths",
        error: lastError && extractHttpStatusFromError_(lastError) ? null : lastError
    };
}

function safeLodgifyListFetch_(fetchFn, params, sourceName) {
    try {
        const response = fetchFn(params || {});
        return {
            items: normalizeLodgifyList(response.body),
            warning: null,
            error: null
        };
    } catch (err) {
        const msg = String(err && err.message ? err.message : err);

        const status = extractHttpStatusFromError_(msg);

        // Einige Lodgify-Tenants haben nicht beide Endpunkte aktiv.
        if (status === 404) {
            Logger.log(`⚠️ Lodgify ${sourceName} endpoint not found (404), skipping source.`);
            return {
                items: [],
                warning: `${sourceName} endpoint not found (404)`,
                error: null
            };
        }

        // Endpoint existiert, aber Key-Rechte/Scope verhindern Zugriff.
        if (status === 401 || status === 403) {
            Logger.log(`⚠️ Lodgify ${sourceName} endpoint reachable but access denied (${status}), skipping source.`);
            return {
                items: [],
                warning: `${sourceName} endpoint reachable but access denied (${status})`,
                error: null
            };
        }

        return {
            items: [],
            warning: null,
            error: `${sourceName}: ${msg}`
        };
    }
}

function fetchReservationsWithFallback_(params) {
    const config = getLodgifyConfig();
    const candidates = uniqueEndpointCandidates_([
        { path: config.reservationsPath, method: "get" },
        { path: config.reservationsPath, method: "post" },
        { path: "/v2/reservations/bookings", method: "get" },
        { path: "/v2/reservations/bookings", method: "post" },
        { path: "/v2/reservations", method: "get" },
        { path: "/v1/reservation/booking", method: "get" },
        { path: "/v1/reservation/booking", method: "post" },
        { path: "/v1/reservation/bookings", method: "get" },
        { path: "/v1/reservation", method: "get" }
    ]);

    const warnings = [];
    let lastError = null;
    for (let i = 0; i < candidates.length; i++) {
        const candidate = candidates[i];
        const path = candidate.path;
        const method = candidate.method;
        try {
            const items = method === "get"
                ? fetchAllPagesFromEndpoint_(path, params || {})
                : normalizeLodgifyList(lodgifyRequest(path, {
                    method,
                    queryParams: null,
                    payload: params || {}
                }).body);
            return {
                items,
                warning:
                    path === config.reservationsPath && method === "get"
                        ? null
                        : `reservations fallback endpoint used: ${method.toUpperCase()} ${path}`,
                error: null
            };
        } catch (err) {
            const msg = String(err && err.message ? err.message : err);
            lastError = msg;
            const status = extractHttpStatusFromError_(msg);
            if (status === 404) {
                warnings.push(`${method.toUpperCase()} ${path}: 404`);
                continue;
            }
            if (status === 401 || status === 403 || status === 405) {
                warnings.push(`${method.toUpperCase()} ${path}: ${status}`);
                continue;
            }

            return {
                items: [],
                warning: null,
                error: `reservations: ${msg}`
            };
        }
    }

    Logger.log("⚠️ Lodgify reservations endpoint not found for all known candidates.");
    return {
        items: [],
        warning: warnings.length
            ? `reservations unavailable on known paths (${warnings.join(", ")})`
            : "reservations endpoint not found (404) on all known paths",
        error: lastError && extractHttpStatusFromError_(lastError) ? null : lastError
    };
}

function fetchAllPagesFromEndpoint_(endpoint, baseParams) {
    const params = Object.assign({}, baseParams || {});
    const requestedPage = toPositiveInt_(params.page);
    const pageSize = toPositiveInt_(params.size) || 100;
    const maxPages = 200;
    const isV1Endpoint = /^\/v1\//i.test(String(endpoint || ""));

    const startPage = requestedPage || 1;
    const endPage = requestedPage || maxPages;
    const collected = [];
    const seenKeys = {};

    for (let page = startPage; page <= endPage; page++) {
        const pageParams = buildPagingParams_(endpoint, params, page, pageSize);
        const response = lodgifyRequest(endpoint, {
            method: "get",
            queryParams: pageParams
        });

        const body = response.body || {};
        const items = normalizeLodgifyList(body);
        let newOnPage = 0;
        for (let i = 0; i < items.length; i++) {
            const item = items[i];
            const key = buildLodgifyItemKey_(item);
            if (!key || seenKeys[key]) continue;
            seenKeys[key] = true;
            collected.push(item);
            newOnPage++;
        }

        Logger.log(`Lodgify page fetch: endpoint=${endpoint}, page=${page}, size=${pageSize}, items=${items.length}, newItems=${newOnPage}, totalUnique=${collected.length}, count=${body.count !== undefined ? body.count : "n/a"}`);

        if (!items.length) {
            break;
        }

        // Wenn die API dieselbe Seite erneut liefert, stoppen wir sicher.
        if (newOnPage === 0) {
            Logger.log(`Lodgify pagination stop: repeated page detected for endpoint=${endpoint} page=${page}`);
            break;
        }

        if (requestedPage) {
            break;
        }

        const totalCount = toPositiveInt_(body.count);
        if (totalCount && collected.length >= totalCount) {
            break;
        }

        // Nicht-v1: klassisches Ende bei kurzer Seite.
        // v1: einige Endpunkte cappen serverseitig auf 50, daher weiter paginieren bis leere/duplizierte Seite.
        if (!isV1Endpoint && items.length < pageSize) {
            break;
        }
    }

    return collected;
}

function buildLodgifyItemKey_(item) {
    if (!item || typeof item !== "object") return "";

    const id = firstDefined(item, ["id", "bookingId", "booking_id", "reservationId", "reservation_id"]);
    if (id !== null && id !== undefined && id !== "") {
        return `id:${String(id).trim()}`;
    }

    const arrivalDate = extractLodgifyCheckinDate_(item);
    const departureDate = extractLodgifyCheckoutDate_(item);
    const arrival = arrivalDate ? arrivalDate.toISOString() : "";
    const departure = departureDate ? departureDate.toISOString() : "";
    const guest = extractLodgifyGuestName_(item) || "";
    const amount = firstDefined(item, ["total", "totalAmount", "total_amount", "amount", "amountPaid", "amount_paid"]) || "";
    return `fallback:${String(arrival)}|${String(departure)}|${String(guest)}|${String(amount)}`;
}

function hasMeaningfulLodgifyValue_(value) {
    if (value === null || value === undefined) return false;
    if (typeof value === "string") return value.trim() !== "";
    if (Array.isArray(value)) return value.length > 0;
    if (value instanceof Date) return isFinite(value.getTime());
    if (typeof value === "object") return Object.keys(value).length > 0;
    return true;
}

function isMergeableLodgifyObject_(value) {
    return !!value && typeof value === "object" && !Array.isArray(value) && !(value instanceof Date);
}

function cloneLodgifyValue_(value) {
    if (Array.isArray(value)) {
        return value.map(cloneLodgifyValue_);
    }
    if (isMergeableLodgifyObject_(value)) {
        const cloned = {};
        Object.keys(value).forEach(function (key) {
            cloned[key] = cloneLodgifyValue_(value[key]);
        });
        return cloned;
    }
    return value;
}

function mergeLodgifyItemData_(preferred, fallback) {
    if (!isMergeableLodgifyObject_(preferred)) {
        return cloneLodgifyValue_(preferred);
    }
    if (!isMergeableLodgifyObject_(fallback)) {
        return cloneLodgifyValue_(preferred);
    }

    const merged = cloneLodgifyValue_(preferred);
    Object.keys(fallback).forEach(function (key) {
        const preferredValue = merged[key];
        const fallbackValue = fallback[key];

        if (!hasMeaningfulLodgifyValue_(preferredValue)) {
            merged[key] = cloneLodgifyValue_(fallbackValue);
            return;
        }

        if (isMergeableLodgifyObject_(preferredValue) && isMergeableLodgifyObject_(fallbackValue)) {
            merged[key] = mergeLodgifyItemData_(preferredValue, fallbackValue);
            return;
        }

        if (Array.isArray(preferredValue) && Array.isArray(fallbackValue) && preferredValue.length === 0 && fallbackValue.length > 0) {
            merged[key] = cloneLodgifyValue_(fallbackValue);
        }
    });

    return merged;
}

var LODGIFY_ITEM_COMPLETENESS_WEIGHTS_ = {
    bookingId: 1,
    guestName: 5,
    checkin: 5,
    checkout: 4,
    bookingDate: 2,
    amount: 3,
    fees: 1,
    status: 1,
    channel: 1,
    paymentOption: 1
};

function scoreLodgifyItemCompleteness_(item) {
    if (!item || typeof item !== "object") return 0;

    let score = 0;
    if (extractLodgifyBookingId_(item)) score += LODGIFY_ITEM_COMPLETENESS_WEIGHTS_.bookingId;
    if (extractLodgifyGuestName_(item)) score += LODGIFY_ITEM_COMPLETENESS_WEIGHTS_.guestName;
    if (extractLodgifyCheckinDate_(item)) score += LODGIFY_ITEM_COMPLETENESS_WEIGHTS_.checkin;
    if (extractLodgifyCheckoutDate_(item)) score += LODGIFY_ITEM_COMPLETENESS_WEIGHTS_.checkout;
    if (extractBuchungstagDate_(item)) score += LODGIFY_ITEM_COMPLETENESS_WEIGHTS_.bookingDate;
    if (extractAmountFromPaths_(item, [
        "total", "grandTotal", "grand_total", "totalAmount", "total_amount",
        "price", "bookingAmount", "booking_amount", "amountToPay", "amount_to_pay",
        "amount", "amountDue", "amount_due"
    ], [
        "quote.total", "quote.totalAmount", "quote.total_amount",
        "reservation.total", "reservation.totalAmount", "reservation.total_amount",
        "financials.total", "financials.totalAmount", "financials.total_amount",
        "charges.total", "invoice.total"
    ]) > 0) score += LODGIFY_ITEM_COMPLETENESS_WEIGHTS_.amount;

    if (extractLodgifyFeesTotal_(item) > 0) score += LODGIFY_ITEM_COMPLETENESS_WEIGHTS_.fees;

    const status = firstDefined(item, [
        "status", "bookingStatus", "booking_status", "reservationStatus", "reservation_status", "state"
    ]);
    if (hasMeaningfulLodgifyValue_(status)) score += LODGIFY_ITEM_COMPLETENESS_WEIGHTS_.status;

    const channel = firstDefined(item, ["source", "channel", "origin", "source_text"]);
    if (hasMeaningfulLodgifyValue_(channel)) score += LODGIFY_ITEM_COMPLETENESS_WEIGHTS_.channel;

    const paymentOption = firstDefined(item, ["payment_option", "paymentOption"]);
    if (hasMeaningfulLodgifyValue_(paymentOption)) score += LODGIFY_ITEM_COMPLETENESS_WEIGHTS_.paymentOption;

    return score;
}

function choosePreferredLodgifyItem_(existingItem, incomingItem) {
    const existingScore = scoreLodgifyItemCompleteness_(existingItem);
    const incomingScore = scoreLodgifyItemCompleteness_(incomingItem);
    const preferIncoming = incomingScore > existingScore;
    const primary = preferIncoming ? incomingItem : existingItem;
    const secondary = preferIncoming ? existingItem : incomingItem;

    return {
        item: mergeLodgifyItemData_(primary, secondary),
        preferred: preferIncoming ? "incoming" : "existing"
    };
}

function dedupeBookingsById_(items) {
    const unique = [];
    const seen = {};
    let duplicateCount = 0;

    (items || []).forEach(item => {
        const key = buildLodgifyItemKey_(item);
        if (!key) {
            unique.push(item);
            return;
        }

        if (seen[key] !== undefined) {
            duplicateCount++;
            const preferred = choosePreferredLodgifyItem_(unique[seen[key]], item);
            // Für den produktiven Import zählen nur die zusammengeführten Buchungsdaten.
            // Wenn die Herkunft relevant ist, nutzt der Audit-Pfad `dedupeTaggedBookingsById_`
            // und behält die bevorzugte Quelle dort separat bei.
            unique[seen[key]] = preferred.item;
            return;
        }

        seen[key] = unique.length;
        unique.push(item);
    });

    return {
        items: unique,
        duplicateCount
    };
}

function dedupeTaggedBookingsById_(taggedItems) {
    const unique = [];
    const seen = {};
    let duplicateCount = 0;

    (taggedItems || []).forEach(entry => {
        const item = entry && entry.item ? entry.item : null;
        const key = buildLodgifyItemKey_(item);

        if (!key) {
            unique.push({
                sourceEndpoint: entry && entry.sourceEndpoint ? entry.sourceEndpoint : "unknown",
                item
            });
            return;
        }

        if (seen[key] !== undefined) {
            duplicateCount++;
            const seenIndex = seen[key];
            const currentEntry = unique[seenIndex];
            const preferred = choosePreferredLodgifyItem_(currentEntry.item, item);
            const incomingSource = entry && entry.sourceEndpoint ? entry.sourceEndpoint : "unknown";
            const nextSource = currentEntry.sourceEndpoint === incomingSource
                ? currentEntry.sourceEndpoint
                : `${currentEntry.sourceEndpoint}+${incomingSource}`;
            unique[seenIndex] = {
                sourceEndpoint: nextSource,
                item: preferred.item
            };
            return;
        }

        seen[key] = unique.length;
        unique.push({
            sourceEndpoint: entry && entry.sourceEndpoint ? entry.sourceEndpoint : "unknown",
            item
        });
    });

    return {
        items: unique,
        duplicateCount
    };
}

function mapAuditRow_(item, sourceEndpoint) {
    const status = extractBookingStatusText_(item);
    const channel = String(firstDefined(item, ["source", "channel", "origin", "source_text"]) || "").trim();
    const type = String(firstDefined(item, ["type", "reservationType", "reservation_type"]) || "").trim();
    const id = String(firstDefined(item, ["id", "bookingId", "booking_id", "reservationId", "reservation_id"]) || "").trim();
    const bookingDate = extractBuchungstagDate_(item);
    const wertstellungDate = extractWertstellungDate_(item);
    const amount = extractAmountForAudit_(item);
    const include = isConfirmedBooking_(item, true);

    return [
        sourceEndpoint || "unknown",
        id,
        status,
        channel,
        type,
        bookingDate || "",
        wertstellungDate || "",
        yearOrUnknown_(bookingDate),
        yearOrUnknown_(wertstellungDate),
        Number(amount.toFixed(2)),
        include,
        safeJsonSnippet_(item, 600)
    ];
}

function extractAmountForAudit_(item) {
    const amountValue = firstDefined(item, [
        "total", "grandTotal", "grand_total", "totalAmount", "total_amount",
        "price", "bookingAmount", "booking_amount", "amountToPay", "amount_to_pay",
        "quoteTotal", "quote_total", "invoiceTotal", "invoice_total",
        "amount", "amountDue", "amount_due", "balance", "balance_due",
        "amountPaid", "amount_paid", "paidAmount", "paid_amount", "totalPaid", "total_paid"
    ]);

    const nestedAmount = firstDefinedDeep(item, [
        "quote.total", "quote.totalAmount", "quote.total_amount",
        "reservation.total", "reservation.totalAmount", "reservation.total_amount",
        "financials.total", "financials.totalAmount", "financials.total_amount",
        "charges.total", "invoice.total"
    ]);

    const resolvedAmountValue = amountValue !== null && amountValue !== undefined && amountValue !== ""
        ? amountValue
        : nestedAmount;
    const amount = resolveAmountObject_(resolvedAmountValue);
    return amount > 0 ? amount : 0;
}

function extractBookingStatusText_(item) {
    return String(firstDefined(item, [
        "status", "bookingStatus", "booking_status", "reservationStatus", "reservation_status", "state"
    ]) || "unknown")
        .trim()
        .toLowerCase();
}

function safeJsonSnippet_(value, maxLen) {
    let text = "";
    try {
        text = JSON.stringify(value);
    } catch (e) {
        text = String(value || "");
    }
    const limit = Number.isFinite(maxLen) && maxLen > 0 ? maxLen : 600;
    if (text.length <= limit) return text;
    return `${text.slice(0, limit)}...`;
}

function countExcludedByFilter_(items, excludeDeclinedCancelled) {
    if (!excludeDeclinedCancelled) return 0;
    let excluded = 0;
    (items || []).forEach(item => {
        if (!isConfirmedBooking_(item, true)) excluded++;
    });
    return excluded;
}

function buildPagingParams_(endpoint, baseParams, page, pageSize) {
    const params = Object.assign({}, baseParams || {});
    params.page = page;
    params.size = pageSize;

    // Einige v1-Endpunkte erwarten alternative Parameternamen.
    if (/^\/v1\//i.test(String(endpoint || ""))) {
        if (params.Page === undefined) params.Page = page;
        if (params.Size === undefined) params.Size = pageSize;
        if (params.pageSize === undefined) params.pageSize = pageSize;
        if (params.limit === undefined) params.limit = pageSize;
        if (params.offset === undefined) params.offset = (page - 1) * pageSize;
        if (params.skip === undefined) params.skip = (page - 1) * pageSize;
        if (params.includeCount === undefined) params.includeCount = true;
    }

    return params;
}

function toPositiveInt_(value) {
    const n = Number(value);
    if (!Number.isFinite(n)) return null;
    const i = Math.floor(n);
    return i > 0 ? i : null;
}

function uniquePaths_(paths) {
    const unique = [];
    const seen = {};
    (paths || []).forEach(p => {
        const path = String(p || "").trim();
        if (!path || seen[path]) return;
        seen[path] = true;
        unique.push(path);
    });
    return unique;
}

function uniqueEndpointCandidates_(candidates) {
    const unique = [];
    const seen = {};
    (candidates || []).forEach(c => {
        const path = String(c && c.path ? c.path : "").trim();
        const method = String(c && c.method ? c.method : "get").toLowerCase();
        if (!path) return;
        const key = `${method}|${path}`;
        if (seen[key]) return;
        seen[key] = true;
        unique.push({ path, method });
    });
    return unique;
}

function extractHttpStatusFromError_(msg) {
    const text = String(msg || "");
    const match = text.match(/\((\d{3})\)/);
    return match ? Number(match[1]) : null;
}

function isConfirmedBooking_(item, excludeDeclinedCancelled) {
    if (!item || typeof item !== "object") return false;

    const type = String(firstDefined(item, ["type", "reservationType", "reservation_type"]) || "")
        .trim()
        .toLowerCase();
    if (type && type.indexOf("enquiry") !== -1) {
        return false;
    }

    const status = String(firstDefined(item, [
        "status", "bookingStatus", "booking_status", "reservationStatus", "reservation_status", "state"
    ]) || "")
        .trim()
        .toLowerCase();

    if (!status) {
        return true;
    }

    if (excludeDeclinedCancelled) {
        // Gewuenscht: nur abgelehnte/stornierte Buchungen filtern.
        const blockedTokens = [
            "cancel", "canceled", "cancelled", "declin", "reject", "denied"
        ];
        for (let i = 0; i < blockedTokens.length; i++) {
            if (status.indexOf(blockedTokens[i]) !== -1) return false;
        }
    }

    const allowedTokens = ["book", "confirm", "paid", "active", "completed", "checked"];
    for (let i = 0; i < allowedTokens.length; i++) {
        if (status.indexOf(allowedTokens[i]) !== -1) return true;
    }

    return true;
}

function toBooleanWithDefault_(value, fallback) {
    if (value === undefined || value === null || value === "") return fallback;
    if (typeof value === "boolean") return value;
    const text = String(value).trim().toLowerCase();
    if (text === "true" || text === "1" || text === "yes" || text === "ja") return true;
    if (text === "false" || text === "0" || text === "no" || text === "nein") return false;
    return fallback;
}

function summarizeBookingStatuses_(items) {
    const summary = {};
    (items || []).forEach(item => {
        const rawStatus = firstDefined(item, [
            "status", "bookingStatus", "booking_status", "reservationStatus", "reservation_status", "state"
        ]);
        const status = String(rawStatus || "(empty)").trim().toLowerCase() || "(empty)";
        if (!summary[status]) summary[status] = 0;
        summary[status]++;
    });
    return summary;
}

function normalizeLodgifyList(body) {
    if (Array.isArray(body)) return body;
    if (!body || typeof body !== "object") return [];

    const keys = ["items", "results", "data", "bookings", "reservations"];
    for (let i = 0; i < keys.length; i++) {
        const value = body[keys[i]];
        if (Array.isArray(value)) return value;
    }

    return [];
}

function mapLodgifyItemToImportRow(item) {
    const details = mapLodgifyItemToImportRowDetailed_(item);
    return details.row;
}

function mapLodgifyItemToImportRowDetailed_(item) {
    if (!item || typeof item !== "object") {
        return {
            row: null,
            reason: "invalidItem",
            amount: 0
        };
    }

    const buchungstagDate = extractBuchungstagDate_(item);

    // Wertstellung soll strikt nach Check-in/Arrival laufen.
    // Nur wenn kein Anreisedatum vorhanden ist, faellt sie auf Buchungstag zurueck.
    const wertstellungDate = extractWertstellungDate_(item) || buchungstagDate;
    if (!wertstellungDate) {
        return {
            row: null,
            reason: "missingDate",
            amount: 0
        };
    }
    const safeBuchungstagDate = buchungstagDate || wertstellungDate;

    // Wichtig: nicht-bezahlte, aber bestaetigte Buchungen sollen mit Gesamtwert erscheinen.
    const amountValue = firstDefined(item, [
        "total", "grandTotal", "grand_total", "totalAmount", "total_amount",
        "price", "bookingAmount", "booking_amount", "amountToPay", "amount_to_pay",
        "quoteTotal", "quote_total", "invoiceTotal", "invoice_total",
        "amount", "amountDue", "amount_due", "balance", "balance_due",
        "amountPaid", "amount_paid", "paidAmount", "paid_amount", "totalPaid", "total_paid"
    ]);

    const nestedAmount = firstDefinedDeep(item, [
        "quote.total", "quote.totalAmount", "quote.total_amount",
        "reservation.total", "reservation.totalAmount", "reservation.total_amount",
        "financials.total", "financials.totalAmount", "financials.total_amount",
        "charges.total", "invoice.total"
    ]);

    const resolvedAmountValue = amountValue !== null && amountValue !== undefined && amountValue !== ""
        ? amountValue
        : nestedAmount;
    const amount = resolveAmountObject_(resolvedAmountValue);
    const finalAmount = amount > 0 ? amount : 0;

    const id = firstDefined(item, ["id", "bookingId", "reservationId"]) || "Lodgify";
    const bookingId = String(id || "").trim();
    const channel = firstDefined(item, ["source", "channel", "origin"]) || "Lodgify";
    const status = String(firstDefined(item, ["status", "bookingStatus", "booking_status", "reservationStatus", "reservation_status", "state"]) || "unknown")
        .trim()
        .toLowerCase();
    const text = `Lodgify ${channel} ${id}`;

    const row = ["", "", "", "", "", "", "", "", "", 0, 0];
    row[2] = safeBuchungstagDate;
    row[3] = wertstellungDate;
    row[4] = text;
    row[5] = `status:${status}`;
    row[9] = 0;
    row[10] = Number(finalAmount.toFixed(2));

    return {
        row,
        reason: finalAmount > 0 ? "ok" : "amountZero",
        amount: finalAmount,
        bookingId
    };
}

function extractWertstellungDate_(item) {
    const directValue = firstDefined(item, [
        "checkIn", "check_in", "checkInDate", "check_in_date",
        "arrival", "arrivalDate", "arrival_date", "dateArrival", "date_arrival",
        "startDate", "start_date", "from", "dateFrom", "date_from"
    ]);
    const directDate = parseDateOrNull(directValue);
    if (directDate) return directDate;

    const nestedValue = firstDefinedDeep(item, [
        "reservation.arrival", "reservation.arrivalDate", "reservation.arrival_date",
        "reservation.checkIn", "reservation.check_in", "reservation.startDate", "reservation.start_date",
        "reservation.from", "reservation.dateFrom", "reservation.date_from",
        "booking.arrival", "booking.arrivalDate", "booking.arrival_date",
        "booking.checkIn", "booking.check_in", "booking.startDate", "booking.start_date",
        "booking.from", "booking.dateFrom", "booking.date_from",
        "period.from", "period.start", "stay.from", "stay.start", "dates.arrival", "dates.checkIn"
    ]);
    const nestedDate = parseDateOrNull(nestedValue);
    if (nestedDate) return nestedDate;

    return null;
}

function extractLodgifyCheckinDate_(item) {
    const directValue = firstDefined(item || {}, [
        "checkIn", "check_in", "checkInDate", "check_in_date",
        "arrival", "arrivalDate", "arrival_date", "dateArrival", "date_arrival",
        "startDate", "start_date", "from", "dateFrom", "date_from"
    ]);
    const directDate = parseDateOrNull(directValue);
    if (directDate) return directDate;

    const nestedValue = firstDefinedDeep(item || {}, [
        "reservation.arrival", "reservation.arrivalDate", "reservation.arrival_date",
        "reservation.checkIn", "reservation.check_in", "reservation.startDate", "reservation.start_date",
        "reservation.from", "reservation.dateFrom", "reservation.date_from",
        "booking.arrival", "booking.arrivalDate", "booking.arrival_date",
        "booking.checkIn", "booking.check_in", "booking.startDate", "booking.start_date",
        "booking.from", "booking.dateFrom", "booking.date_from",
        "period.from", "period.start", "stay.from", "stay.start", "dates.arrival", "dates.checkIn"
    ]);
    const nestedDate = parseDateOrNull(nestedValue);
    if (nestedDate) return nestedDate;

    return findFirstParsableDateByKeyHint_(item || {}, ["arrival", "checkin", "start", "from"]);
}

function extractLodgifyCheckoutDate_(item) {
    const directValue = firstDefined(item || {}, [
        "checkOut", "check_out", "checkOutDate", "check_out_date",
        "departure", "departureDate", "departure_date", "dateDeparture", "date_departure",
        "endDate", "end_date", "to", "dateTo", "date_to"
    ]);
    const directDate = parseDateOrNull(directValue);
    if (directDate) return directDate;

    const nestedValue = firstDefinedDeep(item || {}, [
        "reservation.departure", "reservation.departureDate", "reservation.departure_date",
        "reservation.checkOut", "reservation.check_out", "reservation.endDate", "reservation.end_date",
        "reservation.to", "reservation.dateTo", "reservation.date_to",
        "booking.departure", "booking.departureDate", "booking.departure_date",
        "booking.checkOut", "booking.check_out", "booking.endDate", "booking.end_date",
        "booking.to", "booking.dateTo", "booking.date_to",
        "period.to", "period.end", "stay.to", "stay.end", "dates.departure", "dates.checkOut"
    ]);
    const nestedDate = parseDateOrNull(nestedValue);
    if (nestedDate) return nestedDate;

    return findFirstParsableDateByKeyHint_(item || {}, ["departure", "checkout", "end", "to"]);
}

function extractBuchungstagDate_(item) {
    const directValue = firstDefined(item, [
        "bookingDate", "booking_date", "createdAt", "created_at", "updatedAt", "updated_at",
        "paymentDate", "paidAt", "paid_at", "reservationDate", "reservation_date",
        "checkIn", "check_in", "arrival", "arrivalDate", "arrival_date"
    ]);
    const directDate = parseDateOrNull(directValue);
    if (directDate) return directDate;

    const nestedValue = firstDefinedDeep(item, [
        "reservation.createdAt", "reservation.created_at", "reservation.bookingDate", "reservation.booking_date",
        "booking.createdAt", "booking.created_at", "booking.bookingDate", "booking.booking_date",
        "meta.createdAt", "meta.created_at", "audit.createdAt", "audit.created_at"
    ]);
    const nestedDate = parseDateOrNull(nestedValue);
    if (nestedDate) return nestedDate;

    return findFirstParsableDateByKeyHint_(item, ["booking", "created", "payment", "paid", "date"]);
}

function extractLodgifyGuestName_(item) {
    if (!item || typeof item !== "object") return "";

    const directName = firstDefined(item, [
        "guestName", "guest_name", "customerName", "customer_name",
        "tenantName", "tenant_name", "name"
    ]);
    const normalizedDirectName = normalizeGuestNameValue_(directName);
    if (normalizedDirectName) return normalizedDirectName;

    const directGuestObject = firstDefined(item, ["guest", "customer", "tenant", "contact", "leadGuest", "booker"]);
    const normalizedDirectGuestObjectName = normalizeGuestNameValue_(directGuestObject);
    if (normalizedDirectGuestObjectName) return normalizedDirectGuestObjectName;

    const nestedName = firstDefinedDeep(item, [
        "guest.guest_name.full_name", "guest.guest_name.fullName", "guest.guest_name.name",
        "guest.guest_name.first_name", "guest.guest_name.firstName",
        "guest.name", "guest.fullName", "guest.full_name",
        "customer.name", "customer.fullName", "customer.full_name",
        "tenant.name", "tenant.fullName", "tenant.full_name",
        "contact.name", "contact.fullName", "contact.full_name",
        "leadGuest.name", "leadGuest.fullName", "leadGuest.full_name",
        "booker.name", "booker.fullName", "booker.full_name",
        "reservation.guestName", "reservation.guest_name",
        "reservation.customerName", "reservation.customer_name",
        "booking.guestName", "booking.guest_name",
        "booking.customerName", "booking.customer_name"
    ]);
    const normalizedNestedName = normalizeGuestNameValue_(nestedName);
    if (normalizedNestedName) return normalizedNestedName;

    const namePathPairs = [
        ["guest.firstName", "guest.lastName"],
        ["guest.first_name", "guest.last_name"],
        ["customer.firstName", "customer.lastName"],
        ["customer.first_name", "customer.last_name"],
        ["tenant.firstName", "tenant.lastName"],
        ["tenant.first_name", "tenant.last_name"],
        ["contact.firstName", "contact.lastName"],
        ["contact.first_name", "contact.last_name"],
        ["leadGuest.firstName", "leadGuest.lastName"],
        ["leadGuest.first_name", "leadGuest.last_name"],
        ["booker.firstName", "booker.lastName"],
        ["booker.first_name", "booker.last_name"],
        ["firstName", "lastName"],
        ["first_name", "last_name"]
    ];

    for (let i = 0; i < namePathPairs.length; i++) {
        const pair = namePathPairs[i];
        const first = String(getByPath_(item, pair[0]) || "").trim();
        const last = String(getByPath_(item, pair[1]) || "").trim();
        const combined = `${first} ${last}`.trim();
        if (combined) return combined;
    }

    return "";
}

function normalizeGuestNameValue_(value) {
    if (value === null || value === undefined || value === "") return "";

    if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
        return String(value).trim();
    }

    if (typeof value === "object") {
        const directName = firstDefined(value, ["name", "fullName", "full_name", "displayName", "display_name"]);
        if (directName !== null && directName !== undefined && directName !== "") {
            return String(directName).trim();
        }

        const nestedName = firstDefinedDeep(value, [
            "guest_name.full_name", "guest_name.fullName", "guest_name.name",
            "profile.full_name", "profile.fullName", "profile.name"
        ]);
        if (nestedName !== null && nestedName !== undefined && nestedName !== "") {
            return String(nestedName).trim();
        }

        const firstName = firstDefined(value, ["firstName", "first_name"]);
        const lastName = firstDefined(value, ["lastName", "last_name"]);
        const combined = `${String(firstName || "").trim()} ${String(lastName || "").trim()}`.trim();
        if (combined) return combined;
    }

    return "";
}

function findFirstParsableDateByKeyHint_(obj, keyHints) {
    const maxDepth = 4;

    function visit(node, depth) {
        if (depth > maxDepth || node === null || node === undefined) return null;

        if (typeof node === "string" || typeof node === "number") {
            const parsed = parseDateOrNull(node);
            return parsed || null;
        }

        if (Array.isArray(node)) {
            for (let i = 0; i < node.length; i++) {
                const found = visit(node[i], depth + 1);
                if (found) return found;
            }
            return null;
        }

        if (typeof node === "object") {
            const keys = Object.keys(node);

            for (let i = 0; i < keys.length; i++) {
                const key = keys[i];
                const normalized = key.toLowerCase();
                let hinted = false;
                for (let h = 0; h < keyHints.length; h++) {
                    if (normalized.indexOf(keyHints[h]) !== -1) {
                        hinted = true;
                        break;
                    }
                }
                if (!hinted) continue;

                const parsed = parseDateOrNull(node[key]);
                if (parsed) return parsed;
            }

            for (let i = 0; i < keys.length; i++) {
                const found = visit(node[keys[i]], depth + 1);
                if (found) return found;
            }
        }

        return null;
    }

    return visit(obj, 0);
}

function summarizeYearsFromItems_(items) {
    const byWertstellung = {};
    const byBuchungstag = {};

    (items || []).forEach(item => {
        const wertstellungDate = extractWertstellungDate_(item);
        const buchungstagDate = extractBuchungstagDate_(item);

        const wertstellungYear = yearOrUnknown_(wertstellungDate);
        const buchungstagYear = yearOrUnknown_(buchungstagDate);

        byWertstellung[wertstellungYear] = (byWertstellung[wertstellungYear] || 0) + 1;
        byBuchungstag[buchungstagYear] = (byBuchungstag[buchungstagYear] || 0) + 1;
    });

    return {
        byWertstellung,
        byBuchungstag
    };
}

function yearOrUnknown_(dateObj) {
    if (!dateObj || isNaN(dateObj.getTime())) return "unknown";
    return String(dateObj.getFullYear());
}

function parseBookingIdFromText_(textValue) {
    const text = String(textValue || "").trim();
    if (!text) return "";
    const match = text.match(/(\d+)\s*$/);
    return match ? match[1] : "";
}

function isSameLodgifyRow_(existingRow, mappedRow) {
    if (!existingRow || !mappedRow) return false;

    const existingDate = parseDateOrNull(existingRow[3]);
    const mappedDate = parseDateOrNull(mappedRow[3]);
    const existingDateText = existingDate ? Utilities.formatDate(existingDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : "";
    const mappedDateText = mappedDate ? Utilities.formatDate(mappedDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : "";

    const existingText = String(existingRow[4] || "").trim();
    const mappedText = String(mappedRow[4] || "").trim();
    const existingStatus = String(existingRow[5] || "").trim();
    const mappedStatus = String(mappedRow[5] || "").trim();
    const existingHaben = toNumberOrZero(existingRow[10]);
    const mappedHaben = toNumberOrZero(mappedRow[10]);

    return existingDateText === mappedDateText &&
        existingText === mappedText &&
        existingStatus === mappedStatus &&
        Number(existingHaben.toFixed(2)) === Number(mappedHaben.toFixed(2));
}

function firstDefined(obj, keys) {
    for (let i = 0; i < keys.length; i++) {
        const value = obj[keys[i]];
        if (value !== undefined && value !== null && value !== "") {
            return value;
        }
    }
    return null;
}

function firstDefinedDeep(obj, paths) {
    for (let i = 0; i < paths.length; i++) {
        const value = getByPath_(obj, paths[i]);
        if (value !== undefined && value !== null && value !== "") {
            return value;
        }
    }
    return null;
}

function getByPath_(obj, path) {
    if (!obj || typeof obj !== "object") return undefined;
    const parts = String(path || "").split(".").filter(Boolean);
    let current = obj;
    for (let i = 0; i < parts.length; i++) {
        if (current === null || current === undefined) return undefined;
        current = current[parts[i]];
    }
    return current;
}

function parseDateOrNull(value) {
    if (!value) return null;
    const d = new Date(value);
    return isNaN(d.getTime()) ? null : d;
}

function toNumberOrZero(value) {
    if (value === undefined || value === null || value === "") return 0;
    if (typeof value === "number") return Number.isFinite(value) ? value : 0;

    // Erlaubt auch Formate wie "1.234,56 EUR" oder "EUR 1234.56".
    const normalized = String(value)
        .replace(/[^\d,.-]/g, "")
        .replace(/\.(?=\d{3}(\D|$))/g, "")
        .replace(/,/g, ".");

    const n = Number(normalized);
    return isNaN(n) ? 0 : n;
}

function resolveAmountObject_(value) {
    if (value === null || value === undefined || value === "") return 0;
    if (typeof value === "number") return Number.isFinite(value) ? value : 0;
    if (typeof value === "string") return toNumberOrZero(value);
    // v2 Lodgify amount objects can use various shapes, e.g.
    // { "amount": 123.45, "currency": "EUR" } or { "value": 123.45 }
    // Additional field names are included defensively for non-standard variants.
    if (typeof value === "object" && !Array.isArray(value)) {
        const inner = firstDefined(value, ["amount", "value", "gross", "net", "total"]);
        if (inner !== null && inner !== undefined) {
            return resolveAmountObject_(inner);
        }
    }
    return 0;
}

function buildImportDedupKey(dateValue, textValue, habenValue) {
    const dateObj = parseDateOrNull(dateValue);
    const text = String(textValue || "").trim();
    const haben = toNumberOrZero(habenValue);
    if (!dateObj || !text) return null;

    const dateText = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
    return `${dateText}|${text}|${haben.toFixed(2)}`;
}

function diagnoseLodgifyEndpoints() {
    const config = validateLodgifyConfig();
    const candidates = uniqueEndpointCandidates_([
        { path: "/v1/properties", method: "get" },
        { path: "/bookings", method: "get" },
        { path: "/reservations", method: "get" },
        { path: "/v1/reservation/booking", method: "get" },
        { path: "/v1/reservation/booking", method: "post" },
        { path: "/v1/reservation/bookings", method: "get" },
        { path: "/v1/reservation", method: "get" },
        { path: "/v1/bookings", method: "get" },
        { path: "/v1/reservations", method: "get" },
        { path: "/v2/bookings", method: "get" },
        { path: "/v2/reservations", method: "get" },
        { path: "/v2/reservations/bookings", method: "get" },
        { path: "/v2/reservations/bookings", method: "post" }
    ]);

    const payload = {
        page: 1,
        size: 1
    };

    const queryParams = {
        page: 1,
        size: 1
    };

    const results = [];
    candidates.forEach(candidate => {
        const path = candidate.path;
        const method = candidate.method;
        const url = lodgifyBuildUrl(path, method === "get" ? queryParams : null);
        const response = UrlFetchApp.fetch(url, {
            method,
            muteHttpExceptions: true,
            contentType: method === "post" ? "application/json" : undefined,
            payload: method === "post" ? JSON.stringify(payload) : undefined,
            headers: {
                "X-ApiKey": config.apiKey,
                Accept: "application/json"
            }
        });

        const status = response.getResponseCode();
        const preview = (response.getContentText() || "").slice(0, 180);
        results.push({ method: method.toUpperCase(), path, status, preview });
    });

    Logger.log(JSON.stringify(results));
    return results;
}

function diagnoseLodgifyAuthModes(path, queryParams) {
    const config = validateLodgifyConfig();
    const testPath = String(path || "/v2/reservations/bookings").trim() || "/v2/reservations/bookings";
    const params = queryParams || { page: 1, size: 1, includeCanceled: "false" };

    const modes = [
        {
            mode: "X-ApiKey",
            headers: {
                Accept: "application/json",
                "X-ApiKey": config.apiKey
            }
        },
        {
            mode: "Authorization Bearer",
            headers: {
                Accept: "application/json",
                Authorization: `Bearer ${config.apiKey}`
            }
        },
        {
            mode: "Both",
            headers: {
                Accept: "application/json",
                "X-ApiKey": config.apiKey,
                Authorization: `Bearer ${config.apiKey}`
            }
        }
    ];

    const url = lodgifyBuildUrl(testPath, params);
    const results = modes.map(entry => {
        const response = UrlFetchApp.fetch(url, {
            method: "get",
            muteHttpExceptions: true,
            headers: entry.headers
        });

        const status = response.getResponseCode();
        const preview = (response.getContentText() || "").slice(0, 220);
        return {
            mode: entry.mode,
            status,
            preview
        };
    });

    Logger.log(JSON.stringify({ path: testPath, url, results }));
    return {
        path: testPath,
        url,
        results
    };
}
