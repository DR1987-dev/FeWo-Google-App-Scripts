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
//  Spalte A  Kategorie_Nr       – Buchungskategorie-Nummer aus Lexware (bookingKey, z. B. 4300)
//  Spalte B  Lieferant          – Anzeigename (nur zur Übersicht, kein API-Lookup)
//  Spalte C  Lieferantennummer  – Kundennummer des Lieferanten in Lexware (contactNumber)
//  Spalte D  Betrag_Brutto      – Bruttobetrag in EUR (Zahl, z. B. 142.80)
//  Spalte E  MwSt_Satz          – Steuersatz in % (0, 7 oder 19)
//  Spalte F  Rhythmus           – monatlich | quartalsweise | jährlich
//  Spalte G  Fälligkeitstag     – Tag im Monat (1–31); wird auf den letzten Tag des
//                                  Abrechnungsmonats begrenzt (z. B. 31 im Februar → 28/29)
//  Spalte H  Fälligkeitsmonat   – Monat innerhalb der Periode (optional):
//                                  monatlich:    ignoriert
//                                  quartalsweise: 1=erster Monat (Standard), 2=zweiter, 3=dritter
//                                  jährlich:      1=Januar (Standard) … 12=Dezember
//  Spalte I  Konto_IBAN         – IBAN des abbuchenden Kontos (fremdes Konto, optional)
//  Spalte J  Aktiv              – TRUE/FALSE – Zeile wird nur verarbeitet wenn TRUE
//  Spalte K  Notiz              – Freitext, wird als Remark in Lexware übernommen
//  Spalte L  Zuletzt_Gebucht    – wird vom Skript zurückgeschrieben (JJJJ-MM-TT)
//  Spalte M  Lexware_Beleg_ID   – wird vom Skript zurückgeschrieben (Lexware-UUID)
//
// Script Properties (optional):
//   FIXKOSTEN_SHEET_NAME  – Name des Tabellenblatts (Standard: "Lexware Fixkosten")
// ============================================================

var FIXKOSTEN_DEFAULT_SHEET_NAME = "Lexware Fixkosten";

var FIXKOSTEN_HEADERS = [
    "Kategorie_Nr",       // A  1  (bookingKey in Lexware, z. B. 4300)
    "Lieferant",          // B  2  (Anzeigename, kein API-Lookup)
    "Lieferantennummer",  // C  3  (contactNumber in Lexware – eindeutig)
    "Betrag_Brutto",      // D  4
    "MwSt_Satz",          // E  5
    "Rhythmus",           // F  6
    "Fälligkeitstag",     // G  7  (1–31; wird auf letzten Tag des Abrechnungsmonats begrenzt)
    "Fälligkeitsmonat",   // H  8  (1–3 für quartalsweise; 1–12 für jährlich; leer=Standard)
    "Konto_IBAN",         // I  9
    "Aktiv",              // J  10
    "Notiz",              // K  11
    "Zuletzt_Gebucht",    // L  12
    "Lexware_Beleg_ID"    // M  13
];

// Column indices (1-based for sheet operations)
var FK_COL = {
    KATEGORIE_NR:       1,
    LIEFERANT:          2,
    LIEFERANTENNUMMER:  3,
    BETRAG_BRUTTO:      4,
    MWST_SATZ:          5,
    RHYTHMUS:           6,
    FAELLIGKEITSTAG:    7,
    FAELLIGKEITSMONAT:  8,
    KONTO_IBAN:         9,
    AKTIV:              10,
    NOTIZ:              11,
    ZULETZT_GEBUCHT:    12,
    LEXWARE_BELEG_ID:   13
};

// ---- Dynamic column helpers --------------------------------

/**
 * Liest die erste Zeile des Blattes und baut eine Map
 * { "Spaltenname": 0-basierter-Index }.
 * Spalten ohne Beschriftung werden ignoriert.
 *
 * @param  {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {Object}  Mapping Spaltenname → 0-basierter Index
 */
function buildColMap_(sheet) {
    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var map = {};
    for (var i = 0; i < headers.length; i++) {
        var name = String(headers[i] || "").trim();
        if (name) map[name] = i;
    }
    return map;
}

/**
 * Liest einen Wert aus einer Datenzeile anhand des Header-Namens.
 * Versucht zuerst `primary`, dann `fallback`.
 * Gibt undefined zurück, wenn keine passende Spalte gefunden wurde.
 *
 * @param  {Array}  row
 * @param  {Object} colMap   Ergebnis von buildColMap_
 * @param  {string} primary  Bevorzugter Spaltenname
 * @param  {string} [fallback] Alternativer Spaltenname
 * @return {*}
 */
function readCell_(row, colMap, primary, fallback) {
    if (Object.prototype.hasOwnProperty.call(colMap, primary)) return row[colMap[primary]];
    if (fallback && Object.prototype.hasOwnProperty.call(colMap, fallback)) return row[colMap[fallback]];
    return undefined;
}

/**
 * Liefert den 1-basierten Spaltenindex für sheet.getRange(row, col).
 * Versucht zuerst `primary`, dann `fallback`; gibt -1 zurück wenn nicht gefunden.
 *
 * @param  {Object} colMap
 * @param  {string} primary
 * @param  {string} [fallback]
 * @return {number}  1-basierter Index oder -1
 */
function writeCol_(colMap, primary, fallback) {
    if (Object.prototype.hasOwnProperty.call(colMap, primary)) return colMap[primary] + 1;
    if (fallback && Object.prototype.hasOwnProperty.call(colMap, fallback)) return colMap[fallback] + 1;
    return -1;
}

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
 *   quartalsweise – Buchung am Tag X des Monats Y im Quartal
 *                   (Y=1: erster Monat = Jan/Apr/Jul/Okt, Y=2: zweiter, Y=3: dritter).
 *                   Standardwert für Y ist 1.
 *   jährlich     – Buchung am Tag X des Monats Y (1=Januar … 12=Dezember).
 *                   Standardwert für Y ist 1 (Januar).
 *
 * @param {string}   rhythmus          "monatlich" | "quartalsweise" | "jährlich"
 * @param {number}   faelligkeitstag   Tag im Monat (1–31); wird auf den letzten Tag des
 *                                      Abrechnungsmonats begrenzt (z. B. 31 → 28/29 im Februar)
 * @param {number}   faelligkeitsmonat Monat innerhalb der Periode (optional):
 *                                       quartalsweise: 1–3 (Standard 1)
 *                                       jährlich:      1–12 (Standard 1)
 * @param {string}   zuletztGebucht    Datum der letzten Buchung ("JJJJ-MM-TT") oder ""
 * @param {Date}     now               Heutiges Datum
 * @return {{isDue:boolean, voucherDate:string, dueDate:string}|null}
 *         null wenn noch nicht fällig; andernfalls Objekt mit berechneten Daten.
 */
function isFixkostenDue_(rhythmus, faelligkeitstag, faelligkeitsmonat, zuletztGebucht, now) {
    var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    // Accept 1–31; the actual day will be clamped to the last day of the billing month below.
    var day = Math.max(1, Math.min(31, parseInt(faelligkeitstag, 10) || 1));

    // Determine the start and end of the current billing period
    var periodStart;
    var periodEnd;
    var billingMonth; // 0-based month index for the voucher date

    var r = String(rhythmus || "").toLowerCase().trim();

    // Normalize common short / alternate spellings
    if (r === "monat")                                          r = "monatlich";
    if (r === "quartal" || r === "vierteljährlich"
            || r === "vierteljaehrlich")                        r = "quartalsweise";
    if (r === "jahr" || r === "jährlich" || r === "jaehrlich") r = "jährlich";

    if (r === "monatlich") {
        // Current period: 1st to last day of current month
        periodStart  = new Date(today.getFullYear(), today.getMonth(), 1);
        periodEnd    = new Date(today.getFullYear(), today.getMonth() + 1, 0);
        billingMonth = today.getMonth();
    } else if (r === "quartalsweise") {
        // Quarter starts in month 0, 3, 6 or 9
        var qMonth = Math.floor(today.getMonth() / 3) * 3; // 0, 3, 6, 9
        periodStart = new Date(today.getFullYear(), qMonth, 1);
        periodEnd   = new Date(today.getFullYear(), qMonth + 3, 0);
        // faelligkeitsmonat: 1=first month of quarter (default), 2=second, 3=third
        var monthOffset = Math.max(1, Math.min(3, parseInt(faelligkeitsmonat, 10) || 1)) - 1;
        billingMonth = qMonth + monthOffset;
    } else if (r === "jährlich") {
        // Annual: period spans the whole calendar year
        periodStart = new Date(today.getFullYear(), 0, 1);
        periodEnd   = new Date(today.getFullYear(), 12, 0); // last day of December
        // faelligkeitsmonat: 1=January (default), 2=February, ..., 12=December
        // Note: if billingMonth is before today's month, the voucher date will be in the
        // past within this calendar year — this is intentional (the period is the full year).
        billingMonth = Math.max(1, Math.min(12, parseInt(faelligkeitsmonat, 10) || 1)) - 1;
    } else {
        Logger.log("Fixkosten: Unbekannter Rhythmus '" + rhythmus + "' – übersprungen.");
        return null;
    }

    // Clamp day to the actual number of days in the billing month.
    // E.g. day=31 in February → 28 (or 29 in a leap year); day=31 in April → 30.
    var lastDayOfBillingMonth = new Date(today.getFullYear(), billingMonth + 1, 0).getDate();
    var effectiveDay = Math.min(day, lastDayOfBillingMonth);

    // The voucher date is day X of the billing month of the current period
    var dueDateInPeriod = new Date(today.getFullYear(), billingMonth, effectiveDay);

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
 * @param {string} params.kategorieNr  Lexware-Buchungskategorie-Nummer (bookingKey, z. B. "4300")
 * @param {number} params.betragBrutto  Bruttobetrag
 * @param {number} params.mwstSatz     Mehrwertsteuersatz in % (0, 7 oder 19)
 * @param {string} params.konto_iban   IBAN des abbuchenden Kontos (optional, in Notiz)
 * @param {string} params.notiz        Freitext-Notiz (optional)
 * @return {string}  Lexware-UUID des erstellten Belegs
 */
function createLexwarePurchaseInvoice_(params) {
    var taxRatePercentage = Number(params.mwstSatz) || 0;
    var grossAmount = round2(Number(params.betragBrutto) || 0);
    var netAmount = round2(grossAmount / (1 + taxRatePercentage / 100));

    // Build remark: include IBAN of debit account if provided
    var remark = params.notiz ? String(params.notiz).trim() : "";
    if (params.konto_iban) {
        var ibanNote = "Abbuchung von IBAN: " + String(params.konto_iban).trim();
        remark = remark ? remark + " | " + ibanNote : ibanNote;
    }

    // Line item name: use notiz as display text, otherwise fall back to category number or generic label
    var lineItemName = (params.notiz && String(params.notiz).trim())
        || (params.kategorieNr ? "Kategorie " + params.kategorieNr : "Fixkosten");

    var lineItem = {
                type: "custom",
                name: lineItemName,
                quantity: 1,
                unitName: "Pauschal",
                unitPrice: {
                    currency: "EUR",
                    netAmount: netAmount,
                    grossAmount: grossAmount,
                    taxRatePercentage: taxRatePercentage
                },
                lineItemAmount: grossAmount
            };

    var bookingKeyNum = parseInt(params.kategorieNr, 10);
    if (!isNaN(bookingKeyNum) && bookingKeyNum > 0) {
        lineItem.bookingKey = bookingKeyNum;
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
        lineItems: [ lineItem ]
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
        "Fixkosten: Beleg erstellt – Kategorie " + params.kategorieNr +
        ", Brutto=" + grossAmount +
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
 *   - Spalte L (Zuletzt_Gebucht)  – heutiges Datum
 *   - Spalte M (Lexware_Beleg_ID) – UUID des erstellten Belegs
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

    // Build dynamic column map from actual sheet headers (row 1).
    // This makes the script robust against different column orders and
    // against sheets that were created without the "Lieferantennummer" column.
    var colMap = buildColMap_(sheet);

    var numCols = sheet.getLastColumn();
    var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    var now = new Date();

    var created = 0;
    var skipped = 0;
    var errors  = 0;

    // Pre-compute write-back column indices (1-based)
    var zuletztGebuchtCol = writeCol_(colMap, "Zuletzt_Gebucht");
    var lexwareBelegIdCol = writeCol_(colMap, "Lexware_Beleg_ID", "Lexware_Beleg_I");

    // Cache contact IDs to avoid repeated API calls for same supplier
    var contactCache = {};

    for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var rowNum = i + 2; // 1-based sheet row

        var kategorieNr        = String(readCell_(row, colMap, "Kategorie_Nr")                          || "").trim();
        var lieferant          = String(readCell_(row, colMap, "Lieferant")                             || "").trim();
        // Fall back to "Lieferant" when a dedicated "Lieferantennummer" column is absent
        var lieferantennummer  = String(readCell_(row, colMap, "Lieferantennummer", "Lieferant")        || "").trim();
        var betragBrutto       = Number(readCell_(row, colMap, "Betrag_Brutto"))                        || 0;
        var mwstSatz           = Number(readCell_(row, colMap, "MwSt_Satz"))                            || 0;
        var rhythmus           = String(readCell_(row, colMap, "Rhythmus")                              || "").trim();
        var faelligkeitstag    = readCell_(row, colMap, "Fälligkeitstag");
        var faelligkeitsmonat  = readCell_(row, colMap, "Fälligkeitsmonat");
        var kontoIban          = String(readCell_(row, colMap, "Konto_IBAN")                            || "").trim();
        var aktiv              = readCell_(row, colMap, "Aktiv");
        var notiz              = String(readCell_(row, colMap, "Notiz")                                 || "").trim();
        var zuletztGebucht     = String(readCell_(row, colMap, "Zuletzt_Gebucht")                       || "").trim();

        // Skip empty or inactive rows
        if (!kategorieNr && !lieferantennummer) { skipped++; continue; }
        if (aktiv === false || String(aktiv).toUpperCase() === "FALSE" || aktiv === 0) {
            Logger.log("Fixkosten: Zeile " + rowNum + " (Kat. " + kategorieNr + ") – inaktiv, übersprungen.");
            skipped++;
            continue;
        }

        // Validate required fields
        if (!lieferantennummer) {
            Logger.log(
                "Fixkosten: Zeile " + rowNum + " (Kat. " + kategorieNr + ")" +
                " – Lieferantennummer fehlt, übersprungen."
            );
            skipped++;
            continue;
        }
        if (!betragBrutto || betragBrutto <= 0) {
            Logger.log("Fixkosten: Zeile " + rowNum + " (Kat. " + kategorieNr + ") – Betrag fehlt oder 0, übersprungen.");
            skipped++;
            continue;
        }

        // Check due date
        var dueInfo = isFixkostenDue_(rhythmus, faelligkeitstag, faelligkeitsmonat, zuletztGebucht, now);
        if (!dueInfo) {
            Logger.log("Fixkosten: Zeile " + rowNum + " (Kat. " + kategorieNr + ") – noch nicht fällig.");
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
                "Fixkosten: Zeile " + rowNum + " (Kat. " + kategorieNr + ")" +
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
                kategorieNr: kategorieNr,
                betragBrutto: betragBrutto,
                mwstSatz:    mwstSatz,
                konto_iban:  kontoIban,
                notiz:       notiz
            });

            // Write back booking date and voucher ID
            if (zuletztGebuchtCol > 0) {
                sheet.getRange(rowNum, zuletztGebuchtCol).setValue(formatDate_(now));
            }
            if (lexwareBelegIdCol > 0) {
                sheet.getRange(rowNum, lexwareBelegIdCol).setValue(voucherId);
            }

            Logger.log(
                "Fixkosten: ✅ Zeile " + rowNum + " (Kat. " + kategorieNr + ")" +
                " – Beleg erstellt: " + voucherId
            );
            created++;
        } catch (e) {
            Logger.log(
                "Fixkosten: ❌ Zeile " + rowNum + " (Kat. " + kategorieNr + ")" +
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
