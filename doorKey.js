// =========================================================================
// Türcode (Doorkey) – wird DOORKEY_DAYS_BEFORE_CHECKIN Tage (Standard: 3) vor
// Anreise in Lodgify gesetzt und im Sheet (Spalte "DoorKey") gespeichert.
// Format: TT + MM + Personen + Nächte (z. B. Anreise 07.03., 4 Personen,
// 5 Nächte => "070345").
// Der Daily-Keypad-PIN-Sync liest den Code aus der Spalte "DoorKey".
//
// Script Properties (optional):
//   DOORKEY_SHEET_NAME            – Standard: "AlleBuchungen"
//   DOORKEY_DAYS_BEFORE_CHECKIN   – Standard: 3
// =========================================================================

var DOORKEY_COL_GUESTS_ = "Personen";
var DOORKEY_COL_KEY_ = "DoorKey";

function buildDoorKey_(checkIn, guests, nights) {
    var day = String(checkIn.getDate()).padStart(2, "0");
    var month = String(checkIn.getMonth() + 1).padStart(2, "0");
    return day + month + String(guests) + String(nights);
}

function extractLodgifyGuestCount_(booking) {
    var rooms = booking && Array.isArray(booking.rooms) ? booking.rooms : [];
    var sum = 0;
    rooms.forEach(function (room) {
        var people = Number(firstDefined(room || {}, ["people", "guests", "guest_count"]));
        if (isFinite(people) && people > 0) sum += people;
    });
    if (sum > 0) return sum;

    var direct = Number(firstDefined(booking || {}, ["people", "guests", "guest_count", "number_of_guests", "numberOfGuests"]));
    if (isFinite(direct) && direct > 0) return direct;

    var adults = Number(firstDefined(booking || {}, ["adults"])) || 0;
    var children = Number(firstDefined(booking || {}, ["children"])) || 0;
    return adults + children;
}

function ensureDoorKeyColumns_(sheet, headers) {
    [DOORKEY_COL_GUESTS_, DOORKEY_COL_KEY_].forEach(function (name) {
        if (headers.indexOf(name) === -1) {
            sheet.getRange(1, headers.length + 1).setValue(name);
            headers.push(name);
        }
    });
}

function setLodgifyDoorKey_(booking, bookingId, key) {
    var rooms = booking && Array.isArray(booking.rooms) ? booking.rooms : [];
    var payloadRooms = rooms
        .map(function (room) {
            var roomTypeId = firstDefined(room || {}, ["room_type_id", "roomTypeId"]);
            return roomTypeId === null ? null : { room_type_id: Number(roomTypeId), key_code: key };
        })
        .filter(function (room) { return room !== null; });
    if (payloadRooms.length === 0) {
        throw new Error("Keine room_type_id in Lodgify-Buchung " + bookingId + " gefunden.");
    }

    lodgifyRequest("/v2/reservations/bookings/" + encodeURIComponent(bookingId) + "/keyCodes", {
        method: "put",
        contentType: "application/json-patch+json",
        payload: { rooms: payloadRooms }
    });
}

function processLodgifyDoorKeys() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("No active spreadsheet");

    var props = PropertiesService.getScriptProperties();
    var sheetName = (props.getProperty("DOORKEY_SHEET_NAME") || "AlleBuchungen").trim();
    var rawDays = props.getProperty("DOORKEY_DAYS_BEFORE_CHECKIN");
    var daysBefore = (rawDays !== null && rawDays !== "" && !isNaN(Number(rawDays)) && Number(rawDays) >= 0)
        ? Number(rawDays)
        : 3;

    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) {
        return { ok: true, set: 0, skipped: 0, failed: 0, errors: [] };
    }

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
        .map(function (h) { return String(h || "").trim(); });
    var idIdx = headers.indexOf("LodgifyBookingId");
    var checkInIdx = headers.indexOf("CheckIn");
    var checkOutIdx = headers.indexOf("CheckOut");
    var statusIdx = headers.indexOf("Status");
    if (idIdx === -1 || checkInIdx === -1 || checkOutIdx === -1) {
        Logger.log("DoorKey: Spalten LodgifyBookingId/CheckIn/CheckOut fehlen – übersprungen.");
        return { ok: true, set: 0, skipped: 0, failed: 0, errors: [] };
    }

    ensureDoorKeyColumns_(sheet, headers);
    var guestsIdx = headers.indexOf(DOORKEY_COL_GUESTS_);
    var keyIdx = headers.indexOf(DOORKEY_COL_KEY_);

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
    var msPerDay = 24 * 60 * 60 * 1000;
    var today = toMidnight_(new Date());
    var result = { ok: true, set: 0, skipped: 0, failed: 0, errors: [] };

    data.forEach(function (row, i) {
        var bookingId = String(row[idIdx] || "").trim();
        var checkIn = row[checkInIdx] instanceof Date ? row[checkInIdx] : new Date(row[checkInIdx]);
        var checkOut = row[checkOutIdx] instanceof Date ? row[checkOutIdx] : new Date(row[checkOutIdx]);
        if (!bookingId || isNaN(checkIn.getTime()) || isNaN(checkOut.getTime())) { result.skipped++; return; }

        var status = statusIdx >= 0 ? String(row[statusIdx] || "").trim() : "";
        if (status && isDeclinedOrCancelledStatusText_(status)) { result.skipped++; return; }

        var checkInDay = toMidnight_(checkIn);
        var triggerDate = new Date(checkInDay.getTime() - daysBefore * msPerDay);
        if (today < triggerDate || today > checkInDay) { result.skipped++; return; }

        try {
            var booking = fetchLodgifyBookingById_(bookingId);
            var guests = extractLodgifyGuestCount_(booking);
            if (!guests) throw new Error("Personenanzahl nicht ermittelbar.");
            var nights = Math.round((toMidnight_(checkOut).getTime() - checkInDay.getTime()) / msPerDay);
            if (nights < 1) throw new Error("Ungültige Anzahl Nächte: " + nights);

            var key = buildDoorKey_(checkIn, guests, nights);
            var rowNo = i + 2;
            sheet.getRange(rowNo, guestsIdx + 1).setValue(guests);
            if (String(row[keyIdx] || "").trim() === key) { result.skipped++; return; }

            setLodgifyDoorKey_(booking, bookingId, key);
            sheet.getRange(rowNo, keyIdx + 1).setValue("'" + key);
            Logger.log("DoorKey gesetzt: Buchung " + bookingId + " -> " + key);
            result.set++;
        } catch (e) {
            result.failed++;
            result.errors.push({ bookingId: bookingId, error: e.message });
            Logger.log("DoorKey Fehler Buchung " + bookingId + ": " + e.message);
        }
    });

    if (result.failed > 0) result.ok = false;
    return result;
}
