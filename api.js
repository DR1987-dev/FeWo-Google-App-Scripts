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