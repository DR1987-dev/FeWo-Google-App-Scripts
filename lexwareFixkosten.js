// ============================================================
// Lexware Fixkosten – periodische Eingangsrechnungen erzeugen
//
// Liest das Tabellenblatt "Lexware Fixkosten" und erzeugt für
// jeden fälligen Eintrag automatisch eine Eingangsrechnung
// (purchaseinvoice) in Lexware Office.
//
// Rhythmus: monatlich | quartalsweise | jährlich
//
// Erforderliche Spalten im Blatt "Lexware Fixkosten"
// (Reihenfolge muss exakt eingehalten werden):
//
//  Spalte A  Bezeichnung        – Freitext, z. B. "Strom Q1"
//  Spalte B  Lieferant          – Anzeigename (nur zur Übersicht, kein API-Lookup)
//  Spalte C  Lieferantennummer  – Kundennummer des Lieferanten in Lexware (contactNumber)
//  Spalte D  Betrag_Netto       – Nettobetrag in EUR (Zahl, z. B. 120.00)
//  Spalte E  MwSt_Satz          – Steuersatz in % (0, 7 oder 19)
//  Spalte F  Rhythmus           – monatlich | quartalsweise | jährlich
//  Spalte G  Fälligkeitstag     – Tag im Monat (1–28), an dem die Rechnung fällig ist
//  Spalte H  Konto_IBAN         – IBAN des abbuchenden Kontos (fremdes Konto, optional)
//  Spalte I  Aktiv              – TRUE/FALSE – Zeile wird nur verarbeitet wenn TRUE
//  Spalte J  Notiz              – Freitext, wird als Remark in Lexware übernommen
//  Spalte K  Zuletzt_Gebucht    – wird vom Skript zurückgeschrieben (JJJJ-MM-TT)
//  Spalte L  Lexware_Beleg_ID   – wird vom Skript zurückgeschrieben (Lexware-UUID)
//
// Script Properties (optional):
//   FIXKOSTEN_SHEET_NAME  – Name des Tabellenblatts (Standard: "Lexware Fixkosten")
// ============================================================

var FIXKOSTEN_DEFAULT_SHEET_NAME = "Lexware Fixkosten";

var FIXKOSTEN_HEADERS = [
    "Bezeichnung",        // A  1
    "Lieferant",          // B  2  (Anzeigename, kein API-Lookup)
    "Lieferantennummer",  // C  3  (contactNumber in Lexware – eindeutig)
    "Betrag_Netto",       // D  4
    "MwSt_Satz",          // E  5
    "Rhythmus",           // F  6
    "Fälligkeitstag",     // G  7
    "Konto_IBAN",         // H  8
    "Aktiv",              // I  9
    "Notiz",              // J  10
    "Zuletzt_Gebucht",    // K  11
    "Lexware_Beleg_ID"    // L  12
];

// Column indices (1-based for sheet operations)
var FK_COL = {
    BEZEICHNUNG:       1,
    LIEFERANT:         2,
    LIEFERANTENNUMMER: 3,
    BETRAG_NETTO:      4,
    MWST_SATZ:         5,
    RHYTHMUS:          6,
    FAELLIGKEITSTAG:   7,
    KONTO_IBAN:        8,
    AKTIV:             9,
    NOTIZ:             10,
    ZULETZT_GEBUCHT:   11,
    LEXWARE_BELEG_ID:  12
};

// ---- Config ------------------------------------------------

function getFixkostenSheetName_() {
    var props = PropertiesService.getScriptProperties();
    return (props.getProperty("FIXKOSTEN_SHEET_NAME") || FIXKOSTEN_DEFAULT_SHEET_NAME).trim();
}

// ---- Sheet setup -------------------------------------------

/**
 * Erstellt das Tabellenblatt "Lexware Fixkosten" mit den erforderlichen
 * Spaltenüberschriften, falls es noch nicht existiert.
 * Kann manuell ausgeführt werden, um das Blatt vorzubereiten.
 */
function setupFixkostenSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("No active spreadsheet");

    var sheetName = getFixkostenSheetName_();
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        Logger.log("Fixkosten: Blatt '" + sheetName + "' erstellt.");
    }

    if (sheet.getLastRow() === 0) {
        sheet.appendRow(FIXKOSTEN_HEADERS);
        sheet.getRange(1, 1, 1, FIXKOSTEN_HEADERS.length).setFontWeight("bold");
        Logger.log("Fixkosten: Spaltenüberschriften gesetzt.");
    } else {
        Logger.log("Fixkosten: Blatt existiert bereits, keine Änderung an der Kopfzeile.");
    }

    return { ok: true, sheet: sheetName };
}

// ---- Due-date logic ----------------------------------------

/**
 * Prüft, ob eine Fixkosten-Zeile in der aktuellen Periode fällig ist und
 * noch nicht gebucht wurde.
 *
 * Fälligkeitstag-Logik:
 *   monatlich    – Buchung am Tag X des laufenden Monats.
 *   quartalsweise – Buchung am Tag X des ersten Monats im Quartal
 *                   (z. B. Tag 15 → 15. Jan / 15. Apr / 15. Jul / 15. Okt).
 *   jährlich     – Buchung am Tag X des Januars des laufenden Jahres.
 *
 * @param {string}   rhythmus        "monatlich" | "quartalsweise" | "jährlich"
 * @param {number}   faelligkeitstag Tag im Monat (1–28)
 * @param {string}   zuletztGebucht  Datum der letzten Buchung ("JJJJ-MM-TT") oder ""
 * @param {Date}     now             Heutiges Datum
 * @return {{isDue:boolean, voucherDate:string, dueDate:string}|null}
 *         null wenn noch nicht fällig; andernfalls Objekt mit berechneten Daten.
 */
function isFixkostenDue_(rhythmus, faelligkeitstag, zuletztGebucht, now) {
    var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    var day = Math.max(1, Math.min(28, parseInt(faelligkeitstag, 10) || 1));

    // Determine the start and end of the current billing period
    var periodStart;
    var periodEnd;

    var r = String(rhythmus || "").toLowerCase().trim();

    if (r === "monatlich") {
        // Current period: 1st to last day of current month
        periodStart = new Date(today.getFullYear(), today.getMonth(), 1);
        periodEnd   = new Date(today.getFullYear(), today.getMonth() + 1, 0);
    } else if (r === "quartalsweise") {
        // Quarter starts in month 0, 3, 6 or 9; due date falls on day X of that first month
        var qMonth = Math.floor(today.getMonth() / 3) * 3; // 0, 3, 6, 9
        periodStart = new Date(today.getFullYear(), qMonth, 1);
        periodEnd   = new Date(today.getFullYear(), qMonth + 3, 0);
    } else if (r === "jährlich") {
        // Annual: due on day X of January; period spans the whole calendar year
        periodStart = new Date(today.getFullYear(), 0, 1);
        periodEnd   = new Date(today.getFullYear(), 11, 31);
    } else {
        Logger.log("Fixkosten: Unbekannter Rhythmus '" + rhythmus + "' – übersprungen.");
        return null;
    }

    // The voucher date is day X of the first month of the current period
    var dueDateInPeriod = new Date(periodStart.getFullYear(), periodStart.getMonth(), day);

    // If we haven't reached the due day yet this period, not due
    if (today < dueDateInPeriod) {
        return null;
    }

    // Already booked this period?
    if (zuletztGebucht) {
        var lastBooked = new Date(String(zuletztGebucht).trim());
        if (!isNaN(lastBooked.getTime()) && lastBooked >= periodStart) {
            return null; // already handled this period
        }
    }

    return {
        isDue: true,
        voucherDate: formatDate_(dueDateInPeriod),
        dueDate: formatDate_(periodEnd)
    };
}

function formatDate_(d) {
    var y = d.getFullYear();
    var m = String(d.getMonth() + 1).padStart(2, "0");
    var day = String(d.getDate()).padStart(2, "0");
    return y + "-" + m + "-" + day;
}

// ---- Lexware contact lookup --------------------------------

/**
 * Sucht einen Kontakt in Lexware anhand seiner Lieferantennummer (contactNumber).
 * Diese Nummer ist in Lexware eindeutig und vermeidet Verwechslungen bei
 * gleichlautenden Lieferantennamen.
 *
 * @param  {string} contactNumber  Lieferantennummer aus Lexware (z. B. "L-1042")
 * @return {string|null}           Lexware-UUID des Kontakts oder null
 */
function findLexwareContactIdByNumber_(contactNumber) {
    if (!contactNumber) return null;
    var numberTrimmed = String(contactNumber).trim();

    var result;
    try {
        result = lexwareRequest("/contacts", { contactNumber: numberTrimmed });
    } catch (e) {
        Logger.log("Fixkosten: Kontaktsuche fehlgeschlagen für Nummer '" + numberTrimmed + "': " + e.message);
        return null;
    }

    var body = result.body;
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

    // Look for exact contactNumber match
    for (var i = 0; i < contacts.length; i++) {
        var c = contacts[i];
        var cNum = String(c.contactNumber || c.number || c.customerNumber || "").trim();
        if (cNum === numberTrimmed) {
            return String(c.id);
        }
    }

    // If the API returned exactly one result, accept it (some APIs filter server-side)
    if (contacts.length === 1) {
        Logger.log(
            "Fixkosten: Kontakt für Nummer '" + numberTrimmed +
            "' via Einzeltreffer gefunden (ID=" + contacts[0].id + ")."
        );
        return String(contacts[0].id);
    }

    Logger.log("Fixkosten: Kein eindeutiger Kontakt für Lieferantennummer '" + numberTrimmed + "' gefunden.");
    return null;
}

// ---- Voucher creation --------------------------------------

/**
 * Erzeugt eine Eingangsrechnung (purchaseinvoice) in Lexware Office.
 *
 * @param {Object} params
 * @param {string} params.contactId    Lexware-UUID des Lieferanten
 * @param {string} params.voucherDate  Belegdatum (JJJJ-MM-TT)
 * @param {string} params.dueDate      Fälligkeitsdatum (JJJJ-MM-TT)
 * @param {string} params.bezeichnung  Positionstext
 * @param {number} params.betragNetto  Nettobetrag
 * @param {number} params.mwstSatz     Mehrwertsteuersatz in % (0, 7 oder 19)
 * @param {string} params.konto_iban   IBAN des abbuchenden Kontos (optional, in Notiz)
 * @param {string} params.notiz        Freitext-Notiz (optional)
 * @return {string}  Lexware-UUID des erstellten Belegs
 */
function createLexwarePurchaseInvoice_(params) {
    var taxRatePercentage = Number(params.mwstSatz) || 0;
    var netAmount = round2(Number(params.betragNetto) || 0);
    var grossAmount = round2(netAmount * (1 + taxRatePercentage / 100));

    // Build remark: include IBAN of debit account if provided
    var remark = params.notiz ? String(params.notiz).trim() : "";
    if (params.konto_iban) {
        var ibanNote = "Abbuchung von IBAN: " + String(params.konto_iban).trim();
        remark = remark ? remark + " | " + ibanNote : ibanNote;
    }

    var payload = {
        type: "purchaseinvoice",
        voucherDate: params.voucherDate,
        dueDate: params.dueDate,
        voucherStatus: "open",
        // Lexware Office API expects the contact nested under a "contact" object.
        // The flat "contactId" field is kept for older/alternative API versions.
        contact: { contactId: params.contactId },
        contactId: params.contactId,
        lineItems: [
            {
                type: "custom",
                name: String(params.bezeichnung).trim(),
                quantity: 1,
                unitName: "Pauschal",
                unitPrice: {
                    currency: "EUR",
                    netAmount: netAmount,
                    grossAmount: grossAmount,
                    taxRatePercentage: taxRatePercentage
                },
                lineItemAmount: grossAmount
            }
        ]
    };

    if (remark) {
        payload.remark = remark;
    }

    var result = lexwarePostRequest_("/vouchers", payload);
    var body = result.body;

    var voucherId = String(
        (body && (body.id || body.voucherId || body.uuid)) || ""
    );

    Logger.log(
        "Fixkosten: Beleg erstellt – " + params.bezeichnung +
        ", Netto=" + netAmount +
        ", MwSt=" + taxRatePercentage + "%" +
        ", ID=" + voucherId
    );

    return voucherId;
}

// ---- Main entry point --------------------------------------

/**
 * Liest das Tabellenblatt "Lexware Fixkosten" und erzeugt für jeden
 * fälligen, aktiven Eintrag eine Eingangsrechnung in Lexware.
 *
 * Zurückgeschrieben werden:
 *   - Spalte J (Zuletzt_Gebucht)  – heutiges Datum
 *   - Spalte K (Lexware_Beleg_ID) – UUID des erstellten Belegs
 *
 * @return {{ok:boolean, created:number, skipped:number, errors:number}}
 */
function createLexwareFixkosten() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("No active spreadsheet");

    var sheetName = getFixkostenSheetName_();
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
        Logger.log("Fixkosten: Tabellenblatt '" + sheetName + "' nicht gefunden – setupFixkostenSheet() ausführen.");
        return { ok: false, created: 0, skipped: 0, errors: 0, error: "Sheet not found: " + sheetName };
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
        Logger.log("Fixkosten: Keine Datenzeilen im Blatt '" + sheetName + "'.");
        return { ok: true, created: 0, skipped: 0, errors: 0 };
    }

    var numCols = FIXKOSTEN_HEADERS.length;
    var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    var now = new Date();

    var created = 0;
    var skipped = 0;
    var errors  = 0;

    // Cache contact IDs to avoid repeated API calls for same supplier
    var contactCache = {};

    for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var rowNum = i + 2; // 1-based sheet row

        var bezeichnung       = String(row[FK_COL.BEZEICHNUNG - 1]       || "").trim();
        var lieferant         = String(row[FK_COL.LIEFERANT - 1]         || "").trim();
        var lieferantennummer = String(row[FK_COL.LIEFERANTENNUMMER - 1] || "").trim();
        var betragNetto       = Number(row[FK_COL.BETRAG_NETTO - 1])     || 0;
        var mwstSatz          = Number(row[FK_COL.MWST_SATZ - 1])        || 0;
        var rhythmus          = String(row[FK_COL.RHYTHMUS - 1]          || "").trim();
        var faelligkeitstag   = row[FK_COL.FAELLIGKEITSTAG - 1];
        var kontoIban         = String(row[FK_COL.KONTO_IBAN - 1]        || "").trim();
        var aktiv             = row[FK_COL.AKTIV - 1];
        var notiz             = String(row[FK_COL.NOTIZ - 1]             || "").trim();
        var zuletztGebucht    = String(row[FK_COL.ZULETZT_GEBUCHT - 1]   || "").trim();

        // Skip empty or inactive rows
        if (!bezeichnung && !lieferantennummer) { skipped++; continue; }
        if (aktiv === false || String(aktiv).toUpperCase() === "FALSE" || aktiv === 0) {
            Logger.log("Fixkosten: Zeile " + rowNum + " (" + bezeichnung + ") – inaktiv, übersprungen.");
            skipped++;
            continue;
        }

        // Validate required fields
        if (!lieferantennummer) {
            Logger.log(
                "Fixkosten: Zeile " + rowNum + " (" + bezeichnung + ")" +
                " – Lieferantennummer fehlt, übersprungen."
            );
            skipped++;
            continue;
        }
        if (!betragNetto || betragNetto <= 0) {
            Logger.log("Fixkosten: Zeile " + rowNum + " (" + bezeichnung + ") – Betrag fehlt oder 0, übersprungen.");
            skipped++;
            continue;
        }

        // Check due date
        var dueInfo = isFixkostenDue_(rhythmus, faelligkeitstag, zuletztGebucht, now);
        if (!dueInfo) {
            Logger.log("Fixkosten: Zeile " + rowNum + " (" + bezeichnung + ") – noch nicht fällig.");
            skipped++;
            continue;
        }

        // Look up contact by Lieferantennummer (cached per number)
        var contactId = contactCache[lieferantennummer];
        if (!contactId) {
            contactId = findLexwareContactIdByNumber_(lieferantennummer);
            if (contactId) contactCache[lieferantennummer] = contactId;
        }
        if (!contactId) {
            Logger.log(
                "Fixkosten: Zeile " + rowNum + " (" + bezeichnung + ")" +
                " – Lieferantennummer '" + lieferantennummer +
                "' (" + (lieferant || "?") + ") nicht in Lexware gefunden, übersprungen."
            );
            errors++;
            continue;
        }

        // Create purchase invoice
        try {
            var voucherId = createLexwarePurchaseInvoice_({
                contactId:   contactId,
                voucherDate: dueInfo.voucherDate,
                dueDate:     dueInfo.dueDate,
                bezeichnung: bezeichnung,
                betragNetto: betragNetto,
                mwstSatz:    mwstSatz,
                konto_iban:  kontoIban,
                notiz:       notiz
            });

            // Write back booking date and voucher ID
            sheet.getRange(rowNum, FK_COL.ZULETZT_GEBUCHT).setValue(formatDate_(now));
            sheet.getRange(rowNum, FK_COL.LEXWARE_BELEG_ID).setValue(voucherId);

            Logger.log(
                "Fixkosten: ✅ Zeile " + rowNum + " (" + bezeichnung + ")" +
                " – Beleg erstellt: " + voucherId
            );
            created++;
        } catch (e) {
            Logger.log(
                "Fixkosten: ❌ Zeile " + rowNum + " (" + bezeichnung + ")" +
                " – Fehler beim Erstellen: " + e.message
            );
            errors++;
        }
    }

    Logger.log(
        "Fixkosten abgeschlossen: erstellt=" + created +
        ", übersprungen=" + skipped +
        ", Fehler=" + errors
    );

    return { ok: errors === 0, created: created, skipped: skipped, errors: errors };
}
