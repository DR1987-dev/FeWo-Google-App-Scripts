function initApiConfig(key) {
    setLodgifyApiKey(key);
}

function initEnv(key) {
    return initApiConfig(key);
}

function setLodgifyApiKey(apiKey) {
    var key = String(apiKey || "").trim();
    if (!key) {
        throw new Error("LODGIFY_API_KEY must not be empty");
    }
    PropertiesService.getScriptProperties().setProperty("LODGIFY_API_KEY", key);
}

function doGet() {
    try {
        return handleGetAction_(arguments[0]);
    } catch (err) {
        return jsonResponse_(500, {
            ok: false,
            error: String(err && err.message ? err.message : err)
        });
    }
}

function doPost(e) {
    try {
        var payload = parsePostBody_(e);
        var action = String(payload.action || "").trim();

        if (action === "setLodgifyApiKey" || action === "setKey") {
            setLodgifyApiKey(payload.key);
            return jsonResponse_(200, {
                ok: true,
                message: "LODGIFY_API_KEY updated",
                action: "setLodgifyApiKey"
            });
        }

        if (action === "runAll") {
            runAll();
            return jsonResponse_(200, {
                ok: true,
                message: "Pipeline runAll executed"
            });
        }

        if (action === "importAllCSV") {
            importAllCSVFromDrive();
            return jsonResponse_(200, {
                ok: true,
                message: "Import executed"
            });
        }

        if (action === "generateAlleBuchungenPlan") {
            generateAlleBuchungenPlan();
            return jsonResponse_(200, {
                ok: true,
                message: "AlleBuchungenPlan generated"
            });
        }

        if (action === "generateMonatswerte") {
            generateMonatswerte();
            return jsonResponse_(200, {
                ok: true,
                message: "Monatswerte generated"
            });
        }

        if (action === "lodgifyHealth") {
            var health = lodgifyHealthCheck();
            return jsonResponse_(200, {
                ok: true,
                result: health
            });
        }

        if (action === "lodgifyBookings") {
            var bookings = lodgifyGetBookings(payload.queryParams || {});
            return jsonResponse_(200, {
                ok: true,
                result: bookings
            });
        }

        if (action === "lodgifyReservations") {
            var reservations = lodgifyGetReservations(payload.queryParams || {});
            return jsonResponse_(200, {
                ok: true,
                result: reservations
            });
        }

        if (action === "lodgifyAuthDiag") {
            var diag = diagnoseLodgifyAuthModes(payload.path, payload.queryParams || {});
            return jsonResponse_(200, {
                ok: true,
                result: diag
            });
        }

        if (action === "importLodgifyEinnahmen") {
            var importResult = importLodgifyEinnahmenToImport(payload.queryParams || {});
            return jsonResponse_(200, {
                ok: true,
                result: importResult
            });
        }

        if (action === "lodgifyAudit") {
            var auditResult = auditLodgifyBookingsToSheet(payload.queryParams || {});
            return jsonResponse_(200, {
                ok: true,
                result: auditResult
            });
        }

        if (action === "processPaymentRequests") {
            var prResult = processLodgifyPaymentRequestUpdates(payload.queryParams || {});
            return jsonResponse_(200, {
                ok: true,
                result: prResult
            });
        }

        if (action === "createExpense") {
            var createdExpense = createExpenseRow_(payload.expense || payload.data || {});
            return jsonResponse_(200, {
                ok: true,
                result: createdExpense
            });
        }

        if (action === "updateEditableBooking") {
            var updatedBooking = updateEditableBooking_(payload.booking || payload.data || {});
            return jsonResponse_(200, {
                ok: true,
                result: updatedBooking
            });
        }

        return jsonResponse_(400, {
            ok: false,
            error: "Unknown action"
        });
    } catch (err) {
        return jsonResponse_(500, {
            ok: false,
            error: String(err && err.message ? err.message : err)
        });
    }
}

function handleGetAction_(e) {
    var action = "";
    if (e && e.parameter && e.parameter.action) {
        action = String(e.parameter.action).trim();
    }

    if (!action || action === "ping") {
        return jsonResponse_(200, {
            ok: true,
            message: "Web app is reachable"
        });
    }

    if (action === "health") {
        return jsonResponse_(200, {
            ok: true,
            result: {
                reachable: true,
                timestamp: new Date().toISOString()
            }
        });
    }

    if (action === "getAlleBuchungenPlan") {
        return jsonResponse_(200, {
            ok: true,
            result: readSheetObjects_("AlleBuchungenPlan")
        });
    }

    if (action === "getMonatswerte") {
        return jsonResponse_(200, {
            ok: true,
            result: readSheetObjects_("Monatswerte")
        });
    }

    if (action === "getSummary") {
        return jsonResponse_(200, {
            ok: true,
            result: buildSummaryFromSheets_()
        });
    }

    if (action === "getEuer") {
        return jsonResponse_(200, {
            ok: true,
            result: buildEuerFromMonatswerte_()
        });
    }

    if (action === "getPayments") {
        return jsonResponse_(200, {
            ok: true,
            result: []
        });
    }

    if (action === "getEditableBookings") {
        return jsonResponse_(200, {
            ok: true,
            result: getEditableBookings_()
        });
    }

    if (action === "lodgifyHealth") {
        var health = lodgifyHealthCheck();
        return jsonResponse_(200, {
            ok: true,
            result: health
        });
    }

    if (action === "lodgifyBookings") {
        var bookingsParams = {
            page: toNumberOrUndefined_(e && e.parameter ? e.parameter.page : undefined),
            size: toNumberOrUndefined_(e && e.parameter ? e.parameter.size : undefined),
            from: e && e.parameter ? e.parameter.from : undefined,
            to: e && e.parameter ? e.parameter.to : undefined
        };

        var bookings = lodgifyGetBookings(bookingsParams);
        return jsonResponse_(200, {
            ok: true,
            result: bookings
        });
    }

    if (action === "lodgifyReservations") {
        var reservationParams = {
            page: toNumberOrUndefined_(e && e.parameter ? e.parameter.page : undefined),
            size: toNumberOrUndefined_(e && e.parameter ? e.parameter.size : undefined),
            from: e && e.parameter ? e.parameter.from : undefined,
            to: e && e.parameter ? e.parameter.to : undefined
        };

        var reservations = lodgifyGetReservations(reservationParams);
        return jsonResponse_(200, {
            ok: true,
            result: reservations
        });
    }

    if (action === "lodgifyAudit") {
        var auditParams = {
            page: toNumberOrUndefined_(e && e.parameter ? e.parameter.page : undefined),
            size: toNumberOrUndefined_(e && e.parameter ? e.parameter.size : undefined),
            excludeDeclinedCancelled: e && e.parameter ? e.parameter.excludeDeclinedCancelled : undefined
        };

        var auditResultGet = auditLodgifyBookingsToSheet(auditParams);
        return jsonResponse_(200, {
            ok: true,
            result: auditResultGet
        });
    }

    if (action === "getAlleBuchungen") {
        return jsonResponse_(200, {
            ok: true,
            result: readSheetObjects_("AlleBuchungen")
        });
    }

    if (action === "lodgifyAuthDiag") {
        var path = e && e.parameter ? e.parameter.path : undefined;
        var queryParams = {
            page: toNumberOrUndefined_(e && e.parameter ? e.parameter.page : undefined),
            size: toNumberOrUndefined_(e && e.parameter ? e.parameter.size : undefined),
            includeCanceled: e && e.parameter ? e.parameter.includeCanceled : undefined
        };

        var diagResult = diagnoseLodgifyAuthModes(path, queryParams);
        return jsonResponse_(200, {
            ok: true,
            result: diagResult
        });
    }

    return jsonResponse_(400, {
        ok: false,
        error: "Unknown action"
    });
}

function toNumberOrUndefined_(value) {
    if (value === undefined || value === null || value === "") {
        return undefined;
    }

    var n = Number(value);
    return isNaN(n) ? undefined : n;
}

function readSheetObjects_(sheetName) {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        throw new Error("Sheet not found: " + sheetName);
    }

    var values = sheet.getDataRange().getValues();
    if (!values || values.length < 2) {
        return [];
    }

    var header = values[0].map(function (col, idx) {
        var label = String(col || "").trim();
        return label || "col_" + idx;
    });

    var rows = [];
    for (var i = 1; i < values.length; i++) {
        var row = values[i];
        var obj = {};
        for (var j = 0; j < header.length; j++) {
            var cell = row[j];
            if (cell instanceof Date) {
                obj[header[j]] = Utilities.formatDate(cell, Session.getScriptTimeZone(), "yyyy-MM-dd");
            } else {
                obj[header[j]] = cell;
            }
        }
        rows.push(obj);
    }

    return rows;
}

function buildSummaryFromSheets_() {
    var monatswerte = readSheetObjects_("Monatswerte");
    var alle = readSheetObjects_("AlleBuchungenPlan");
    var now = new Date();
    var y = now.getFullYear();
    var m = now.getMonth() + 1;

    var monthRows = monatswerte.filter(function (r) {
        return Number(r.Jahr) === y && Number(r.Monat) === m;
    });

    var monthIncome = monthRows.reduce(function (sum, r) {
        return sum + toNumberOrZero_(r.Einnahmen);
    }, 0);
    var monthExpenses = monthRows.reduce(function (sum, r) {
        return sum + toNumberOrZero_(r.Ausgaben);
    }, 0);
    var monthProfit = monthRows.reduce(function (sum, r) {
        return sum + toNumberOrZero_(r.Monatsdifferenz);
    }, 0);

    var expenses = alle
        .filter(function (r) { return toNumberOrZero_(r.Betrag) < 0; })
        .map(function (r, idx) {
            return {
                id: String(r.Datum || "") + "-" + String(r.Kostenart || "expense") + "-" + idx,
                category: String(r.Kostenart || ""),
                date: r.Datum || "",
                amount: Math.abs(toNumberOrZero_(r.Betrag)),
                note: String(r.Buchungskonto || "")
            };
        });

    return {
        month_income: round2(monthIncome),
        month_expenses: round2(monthExpenses),
        month_profit: round2(monthProfit),
        open_payments: 0,
        expenses: expenses
    };
}

function buildEuerFromMonatswerte_() {
    var monatswerte = readSheetObjects_("Monatswerte");
    var now = new Date();
    var y = now.getFullYear();
    var m = now.getMonth() + 1;

    var monthRows = monatswerte.filter(function (r) {
        return Number(r.Jahr) === y && Number(r.Monat) === m;
    });

    var yearRows = monatswerte.filter(function (r) {
        return Number(r.Jahr) === y;
    });

    return {
        month_income: round2(monthRows.reduce(function (sum, r) { return sum + toNumberOrZero_(r.Einnahmen); }, 0)),
        month_expenses: round2(monthRows.reduce(function (sum, r) { return sum + toNumberOrZero_(r.Ausgaben); }, 0)),
        month_profit: round2(monthRows.reduce(function (sum, r) { return sum + toNumberOrZero_(r.Monatsdifferenz); }, 0)),
        year_income: round2(yearRows.reduce(function (sum, r) { return sum + toNumberOrZero_(r.Einnahmen); }, 0)),
        year_expenses: round2(yearRows.reduce(function (sum, r) { return sum + toNumberOrZero_(r.Ausgaben); }, 0)),
        year_profit: round2(yearRows.reduce(function (sum, r) { return sum + toNumberOrZero_(r.Monatsdifferenz); }, 0))
    };
}

function createExpenseRow_(expense) {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName("Manuelle_Buchungen");
    if (!sheet) {
        throw new Error("Sheet not found: Manuelle_Buchungen");
    }

    var dateText = String(expense.date || "").trim();
    var parsedDate = dateText ? new Date(dateText) : new Date();
    if (isNaN(parsedDate.getTime())) {
        parsedDate = new Date();
    }

    var category = String(expense.category || "Sonstiges");
    var note = String(expense.note || "Ausgabe");
    var amount = Number(expense.amount || 0);
    var signedAmount = -Math.abs(amount);

    var row = [
        category,
        note,
        "Mietenkonto",
        parsedDate,
        round2(signedAmount)
    ];

    sheet.appendRow(row);

    return {
        category: category,
        note: note,
        amount: Math.abs(round2(signedAmount)),
        date: Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "yyyy-MM-dd")
    };
}

function toNumberOrZero_(value) {
    var n = Number(value);
    return isNaN(n) ? 0 : n;
}

function getEditableBookings_() {
    var ss = SpreadsheetApp.getActive();
    var rows = [];

    rows = rows
        .concat(readFixkostenEditable_(ss))
        .concat(readManualEditable_(ss))
        .concat(readUmbuchungenEditable_(ss))
        .concat(readLodgifyEditable_(ss));

    rows.sort(function (a, b) {
        return String(b.date || "").localeCompare(String(a.date || ""));
    });

    return rows;
}

function readFixkostenEditable_(ss) {
    var sheet = ss.getSheetByName("Fixkosten");
    if (!sheet) return [];

    var values = sheet.getDataRange().getValues();
    if (!values || values.length < 2) return [];

    var out = [];
    for (var i = 1; i < values.length; i++) {
        var row = values[i];
        var rowNo = i + 1;
        var amount = toNumberOrZero_(row[2]);
        out.push({
            id: "fixkosten:" + rowNo,
            booking_type: "fixkosten",
            source_sheet: "Fixkosten",
            source_row: rowNo,
            guest_name: String(row[0] || "Fixkosten"),
            checkin: formatDateCell_(row[3]),
            checkout: formatDateCell_(row[4]),
            gross_amount: Math.abs(amount),
            fees_total: 0,
            net_amount: -Math.abs(amount),
            payout_amount: -Math.abs(amount),
            date: formatDateCell_(row[3]),
            note: String(row[5] || ""),
            account: String(row[8] || ""),
            raw: {
                kostenart: String(row[0] || ""),
                kategorie: String(row[1] || ""),
                betrag: Math.abs(amount),
                startdatum: formatDateCell_(row[3]),
                enddatum: formatDateCell_(row[4]),
                buchungstextabgleich: String(row[5] || ""),
                wertstellungstag: row[6] === null || row[6] === undefined ? "" : String(row[6]),
                intervall: String(row[7] || ""),
                buchungskonto: String(row[8] || "")
            }
        });
    }

    return out;
}

function readManualEditable_(ss) {
    var sheet = ss.getSheetByName("Manuelle_Buchungen");
    if (!sheet) return [];

    var values = sheet.getDataRange().getValues();
    if (!values || values.length < 2) return [];

    var out = [];
    for (var i = 1; i < values.length; i++) {
        var row = values[i];
        var rowNo = i + 1;
        var amount = toNumberOrZero_(row[4]);
        out.push({
            id: "manual:" + rowNo,
            booking_type: "manual",
            source_sheet: "Manuelle_Buchungen",
            source_row: rowNo,
            guest_name: String(row[0] || "Manuelle Buchung"),
            checkin: formatDateCell_(row[3]),
            checkout: formatDateCell_(row[3]),
            gross_amount: Math.abs(amount),
            fees_total: 0,
            net_amount: amount,
            payout_amount: amount,
            date: formatDateCell_(row[3]),
            note: String(row[1] || ""),
            account: String(row[2] || ""),
            raw: {
                kostenart: String(row[0] || ""),
                buchungstext: String(row[1] || ""),
                buchungskonto: String(row[2] || ""),
                datum: formatDateCell_(row[3]),
                betrag: amount
            }
        });
    }

    return out;
}

function readUmbuchungenEditable_(ss) {
    var sheet = ss.getSheetByName("Umbuchungen");
    if (!sheet) return [];

    var values = sheet.getDataRange().getValues();
    if (!values || values.length < 2) return [];

    var out = [];
    for (var i = 1; i < values.length; i++) {
        var row = values[i];
        var rowNo = i + 1;
        var amount = Math.abs(toNumberOrZero_(row[3]));
        out.push({
            id: "transfer:" + rowNo,
            booking_type: "transfer",
            source_sheet: "Umbuchungen",
            source_row: rowNo,
            guest_name: String(row[4] || "Umbuchung"),
            checkin: formatDateCell_(row[0]),
            checkout: formatDateCell_(row[0]),
            gross_amount: amount,
            fees_total: 0,
            net_amount: 0,
            payout_amount: 0,
            date: formatDateCell_(row[0]),
            note: String(row[4] || ""),
            account: String(row[1] || "") + " -> " + String(row[2] || ""),
            raw: {
                datum: formatDateCell_(row[0]),
                von: String(row[1] || ""),
                nach: String(row[2] || ""),
                betrag: amount,
                text: String(row[4] || "")
            }
        });
    }

    return out;
}

function readLodgifyEditable_(ss) {
    if (!ss || typeof getPaymentRequestConfig_ !== "function") return [];

    var config = getPaymentRequestConfig_();
    var sheetName = String(config && config.sheetName ? config.sheetName : "AlleBuchungen").trim() || "AlleBuchungen";
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];

    var values = sheet.getDataRange().getValues();
    if (!values || values.length < 2) return [];

    var headerMap = typeof getSheetHeaderMap_ === "function"
        ? getSheetHeaderMap_(sheet)
        : {};
    var out = [];

    for (var i = 1; i < values.length; i++) {
        var row = values[i];
        var rowNo = i + 1;
        var booking = typeof buildLodgifyEditableBookingFromSheetRow_ === "function"
            ? buildLodgifyEditableBookingFromSheetRow_(sheetName, rowNo, row, headerMap)
            : null;
        if (booking) {
            out.push(booking);
        }
    }

    return out;
}

function updateEditableBooking_(booking) {
    var id = String(booking.id || "").trim();
    if (!id || id.indexOf(":") < 0) {
        throw new Error("booking.id fehlt oder ist ungueltig");
    }

    var parts = id.split(":");
    var kind = parts[0];
    var rowNo = 0;
    var encodedSheetName = "";

    if (kind === "lodgify" || kind === "lodgify_external") {
        if (parts.length < 3) {
            throw new Error("Lodgify booking.id ist ungueltig");
        }
        rowNo = Number(parts[parts.length - 1]);
        encodedSheetName = parts.slice(1, parts.length - 1).join(":");
    } else {
        rowNo = Number(parts[1]);
    }

    if (!rowNo || isNaN(rowNo) || rowNo < 2) {
        throw new Error("Zeilennummer ungueltig");
    }

    var ss = SpreadsheetApp.getActive();

    if (kind === "fixkosten") {
        var fixSheet = ss.getSheetByName("Fixkosten");
        if (!fixSheet) throw new Error("Sheet not found: Fixkosten");

        fixSheet.getRange(rowNo, 1, 1, 9).setValues([[
            String(booking.kostenart || ""),
            String(booking.kategorie || ""),
            Math.abs(toNumberOrZero_(booking.betrag)),
            parseDateOrToday_(booking.startdatum),
            parseDateOrBlank_(booking.enddatum),
            String(booking.buchungstextabgleich || ""),
            toNumberOrBlank_(booking.wertstellungstag),
            String(booking.intervall || ""),
            String(booking.buchungskonto || "")
        ]]);

        regenerateDerivedSheets_();
        return { ok: true, id: id, kind: kind };
    }

    if (kind === "manual") {
        var manualSheet = ss.getSheetByName("Manuelle_Buchungen");
        if (!manualSheet) throw new Error("Sheet not found: Manuelle_Buchungen");

        manualSheet.getRange(rowNo, 1, 1, 5).setValues([[
            String(booking.kostenart || ""),
            String(booking.buchungstext || ""),
            String(booking.buchungskonto || ""),
            parseDateOrToday_(booking.datum),
            toNumberOrZero_(booking.betrag)
        ]]);

        regenerateDerivedSheets_();
        return { ok: true, id: id, kind: kind };
    }

    if (kind === "transfer") {
        var transferSheet = ss.getSheetByName("Umbuchungen");
        if (!transferSheet) throw new Error("Sheet not found: Umbuchungen");

        transferSheet.getRange(rowNo, 1, 1, 5).setValues([[
            parseDateOrToday_(booking.datum),
            String(booking.von || ""),
            String(booking.nach || ""),
            Math.abs(toNumberOrZero_(booking.betrag)),
            String(booking.text || "")
        ]]);

        regenerateDerivedSheets_();
        return { ok: true, id: id, kind: kind };
    }

    if (kind === "lodgify" || kind === "lodgify_external") {
        if (typeof updateLodgifyEditableBookingRow_ !== "function") {
            throw new Error("Lodgify-Buchungen koennen in diesem Deployment nicht bearbeitet werden.");
        }

        var sheetName = encodedSheetName
            ? decodeURIComponent(encodedSheetName)
            : String((getPaymentRequestConfig_() || {}).sheetName || "AlleBuchungen");

        return updateLodgifyEditableBookingRow_(sheetName, rowNo, booking || {});
    }

    throw new Error("Nicht unterstuetzte Buchungsart: " + kind);
}

function regenerateDerivedSheets_() {
    generateAlleBuchungenPlan();
    generateMonatswerte();
}

function formatDateCell_(value) {
    if (!value) return "";
    var d = value instanceof Date ? value : new Date(value);
    if (isNaN(d.getTime())) return "";
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function parseDateOrToday_(value) {
    var d = value ? new Date(value) : new Date();
    if (isNaN(d.getTime())) return new Date();
    return d;
}

function parseDateOrBlank_(value) {
    if (!value) return "";
    var d = new Date(value);
    if (isNaN(d.getTime())) return "";
    return d;
}

function toNumberOrBlank_(value) {
    if (value === "" || value === null || value === undefined) return "";
    var n = Number(value);
    return isNaN(n) ? "" : n;
}

function parsePostBody_(e) {
    if (!e || !e.postData || !e.postData.contents) {
        throw new Error("Missing POST body");
    }

    try {
        return JSON.parse(e.postData.contents);
    } catch (parseError) {
        throw new Error("Invalid JSON body");
    }
}

function jsonResponse_(status, data) {
    var body = {
        status: status,
        data: data
    };

    return ContentService
        .createTextOutput(JSON.stringify(body))
        .setMimeType(ContentService.MimeType.JSON);
}