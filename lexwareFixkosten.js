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
//  Spalte A  Kategorie          – Buchungskategoriename aus Lexware (name, z. B. "Reise MA")
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
//  Spalte N  Beleggruppe        – Gruppenkennung für mehrzeilige Belege (optional):
//                                  Mehrere Zeilen mit derselben Beleggruppe werden zu
//                                  einer einzigen Eingangsrechnung mit mehreren Positionen
//                                  zusammengefasst. Leer = eigenständige Rechnung.
//                                  Für alle Zeilen einer Gruppe gelten: Lieferantennummer,
//                                  Rhythmus, Fälligkeitstag/-monat, Konto_IBAN, Aktiv und
//                                  Notiz der ersten Zeile der Gruppe.
//                                  Jede Zeile liefert eine eigene Position (Kategorie +
//                                  Betrag_Brutto + MwSt_Satz).
//
// Script Properties (optional):
//   FIXKOSTEN_SHEET_NAME  – Name des Tabellenblatts (Standard: "Lexware Fixkosten")
// ============================================================

var FIXKOSTEN_DEFAULT_SHEET_NAME = "Lexware Fixkosten";

var FIXKOSTEN_HEADERS = [
    "Kategorie",          // A  1  (name in Lexware, z. B. "Reise MA")
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
    "Lexware_Beleg_ID",   // M  13
    "Beleggruppe"         // N  14 (optional; gleiche Kennung → eine Rechnung mit mehreren Positionen)
];

// Column indices (1-based for sheet operations)
var FK_COL = {
    KATEGORIE:          1,
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
    LEXWARE_BELEG_ID:   13,
    BELEGGRUPPE:        14
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

function buildFixkostenVoucherNumber_(params) {
    var datePart = String(params.voucherDate || "").replace(/[^0-9]/g, "");
    var supplierPart = String(params.lieferantennummer || "")
        .replace(/[^A-Za-z0-9]/g, "")
        .toUpperCase();
    var categoryPart = String(params.kategorieName || "")
        .replace(/[^A-Za-z0-9]/g, "")
        .toUpperCase();
    var amountPart = String(Math.round((Number(params.betragBrutto) || 0) * 100));

    var voucherNumber = [
        "FK",
        datePart || "DATE",
        supplierPart || categoryPart || "NA",
        amountPart || "0"
    ].join("-");

    return voucherNumber.slice(0, 60);
}

// ---- Lexware contact lookup --------------------------------

/**
 * Sucht einen Kontakt in Lexware anhand seiner Lieferantennummer.
 * Lexware speichert Lieferantennummern im Vendor-Role-Objekt des Kontakts
 * (roles.vendor.number). Der Endpunkt unterstützt die Filterung per
 * number-Abfrageparameter (GET /v1/contacts?number=...).
 *
 * @param  {string} vendorNumber  Lieferantennummer aus Lexware (z. B. "70009")
 * @return {string|null}          Lexware-UUID des Kontakts oder null
 */
function findLexwareContactIdByNumber_(vendorNumber) {
    if (!vendorNumber) return null;
    var rawValue = vendorNumber; // keep original for type logging
    var numberTrimmed = String(vendorNumber).trim();

    // Log the raw cell value type so we can detect unexpected formatting
    Logger.log(
        "Fixkosten: Kontaktsuche – Rohwert: '" + rawValue +
        "' (Typ=" + typeof rawValue + "), bereinigt: '" + numberTrimmed + "'"
    );

    // Reconstruct the URL for diagnostic logging (mirrors lexwareRequest logic)
    var debugUrl = LEXWARE_BASE_URL + "/contacts?vendor=true&number=" + encodeURIComponent(numberTrimmed);
    Logger.log("Fixkosten: Kontaktsuche – GET " + debugUrl);

    var result;
    try {
        result = lexwareRequest("/contacts", { vendor: true, number: numberTrimmed });
    } catch (e) {
        Logger.log("Fixkosten: Kontaktsuche fehlgeschlagen für Nummer '" + numberTrimmed + "': " + e.message);
        return null;
    }

    // Log HTTP status and key pagination fields
    var body = result.body;
    var totalElements = (body && body.totalElements !== undefined) ? body.totalElements : "n/a";
    var totalPages    = (body && body.totalPages    !== undefined) ? body.totalPages    : "n/a";
    Logger.log(
        "Fixkosten: Kontaktsuche HTTP " + result.status +
        " für Nummer '" + numberTrimmed +
        "' – totalElements=" + totalElements +
        ", totalPages=" + totalPages
    );

    // Log truncated raw body so the exact API response is visible
    var rawBodyStr = body ? JSON.stringify(body) : "null";
    Logger.log(
        "Fixkosten: Kontaktsuche – Raw-Body (max 800 Z.): " +
        rawBodyStr.slice(0, 800) + (rawBodyStr.length > 800 ? "…" : "")
    );

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

    Logger.log(
        "Fixkosten: Kontaktsuche für Nummer '" + numberTrimmed +
        "' – " + contacts.length + " Kontakt(e) im content-Array." +
        (contacts.length > 0
            ? " | Erster Kontakt id=" + (contacts[0].id || "?") +
              " roles=" + JSON.stringify((contacts[0] && contacts[0].roles) || null)
            : "")
    );

    // 1. Exact match on roles.vendor.number (preferred – vendor-specific number)
    for (var i = 0; i < contacts.length; i++) {
        var c = contacts[i];
        if (c.roles && c.roles.vendor) {
            var vendorNum = String(c.roles.vendor.number || "").trim();
            if (vendorNum === numberTrimmed) {
                Logger.log("Fixkosten: Kontakt für Nummer '" + numberTrimmed + "' via roles.vendor.number gefunden (ID=" + c.id + ").");
                return String(c.id);
            }
        }
    }

    // 2. Fallback: match on roles.customer.number
    for (var j = 0; j < contacts.length; j++) {
        var cc = contacts[j];
        if (cc.roles && cc.roles.customer) {
            var customerNum = String(cc.roles.customer.number || "").trim();
            if (customerNum === numberTrimmed) {
                Logger.log("Fixkosten: Kontakt für Nummer '" + numberTrimmed + "' via roles.customer.number gefunden (ID=" + cc.id + ").");
                return String(cc.id);
            }
        }
    }

    // 3. Fallback: match on top-level contact number field
    for (var k = 0; k < contacts.length; k++) {
        var ck = contacts[k];
        var topNum = String(ck.number || "").trim();
        if (topNum && topNum === numberTrimmed) {
            Logger.log("Fixkosten: Kontakt für Nummer '" + numberTrimmed + "' via top-level number gefunden (ID=" + ck.id + ").");
            return String(ck.id);
        }
    }

    // 4. Last resort: if the API (which was already queried with ?number=...) returned
    //    exactly one contact, trust the API filter and accept that result.
    if (contacts.length === 1) {
        var single = contacts[0];
        Logger.log(
            "Fixkosten: Kontakt für Nummer '" + numberTrimmed +
            "' via Einzeltreffer akzeptiert (ID=" + single.id +
            ", roles=" + JSON.stringify(single.roles || null) + ")."
        );
        return String(single.id);
    }

    Logger.log("Fixkosten: Kein eindeutiger Kontakt für Lieferantennummer '" + numberTrimmed + "' gefunden.");
    return null;
}

function listLexwareContacts_() {
    var contacts = [];
    var page = 0;
    var pageSize = 100;
    var totalPages = 1;

    do {
        var result = lexwareRequest("/contacts", { page: page, size: pageSize, vendor: true });
        var body = result.body;
        var pageContacts = [];

        if (Array.isArray(body)) {
            pageContacts = body;
        } else if (body && Array.isArray(body.content)) {
            pageContacts = body.content;
        } else if (body && Array.isArray(body.contacts)) {
            pageContacts = body.contacts;
        } else if (body && body.id) {
            pageContacts = [body];
        }

        contacts = contacts.concat(pageContacts);
        totalPages = (body && body.page && body.page.totalPages !== undefined) ? body.page.totalPages
                   : (body && body.totalPages !== undefined ? body.totalPages : (pageContacts.length === pageSize ? page + 2 : page + 1));
        page++;
    } while (page < totalPages);

    Logger.log(
        "Fixkosten: Kontaktindex – " + contacts.length +
        " Kontakt(e) aus " + totalPages + " Seite(n) geladen."
    );

    return contacts;
}

function buildLexwareContactNumberIndex_() {
    // 1. Versuche, den Index aus dem lokalen Kunden-Sheet zu erstellen.
    //    syncLexwareKundenSheet() muss vorher aufgerufen worden sein (z. B.
    //    durch importLexwareAll()), damit das Sheet aktuell ist.
    var sheetIndex = buildContactIndexFromKundenSheet_();
    if (sheetIndex && Object.keys(sheetIndex).length > 0) {
        Logger.log(
            "Fixkosten: Kontaktindex – " + Object.keys(sheetIndex).length +
            " Einträge aus Kunden-Sheet (kein API-Aufruf erforderlich)."
        );
        return sheetIndex;
    }

    // 2. Fallback: Kontakte per API abrufen (wenn das Sheet leer / nicht vorhanden ist).
    Logger.log("Fixkosten: Kunden-Sheet leer oder nicht vorhanden – lade Kontakte per API.");
    var contacts = listLexwareContacts_();
    var index = {};

    function addNumber(numberValue, contactId, source) {
        var normalized = String(numberValue || "").trim();
        if (!normalized || !contactId) return;
        if (!Object.prototype.hasOwnProperty.call(index, normalized)) {
            index[normalized] = { id: String(contactId), ambiguous: false, source: source };
            return;
        }
        if (index[normalized].id !== String(contactId)) {
            index[normalized].ambiguous = true;
        }
    }

    for (var i = 0; i < contacts.length; i++) {
        var contact = contacts[i] || {};
        var contactId = contact.id;
        addNumber(contact.number, contactId, "number");
        if (contact.roles && contact.roles.vendor) addNumber(contact.roles.vendor.number, contactId, "roles.vendor.number");
        if (contact.roles && contact.roles.customer) addNumber(contact.roles.customer.number, contactId, "roles.customer.number");
    }

    return index;
}

function findLexwareContactIdInIndex_(vendorNumber, contactIndex) {
    if (!vendorNumber || !contactIndex) return null;
    var numberTrimmed = String(vendorNumber).trim();
    var match = contactIndex[numberTrimmed];
    if (!match) return null;
    if (match.ambiguous) {
        Logger.log("Fixkosten: Kontaktindex – Lieferantennummer '" + numberTrimmed + "' ist nicht eindeutig.");
        return null;
    }
    Logger.log(
        "Fixkosten: Kontakt für Nummer '" + numberTrimmed +
        "' via Kontaktindex gefunden (ID=" + match.id +
        ", Quelle=" + match.source + ")."
    );
    return match.id;
}

/**
 * Sucht eine Buchungskategorie anhand ihres Namens.
 *
 * Sucht zuerst im lokalen Sheet "Lexware_Kategorien" (Spalte "Name" → Spalte "ID"),
 * das von importLexwareKategorien() befüllt wird. Nur wenn das Sheet nicht
 * vorhanden ist oder der Name dort nicht gefunden wird, wird als Fallback
 * GET /v1/posting-categories aufgerufen.
 *
 * @param  {string} categoryName  Name der Buchungskategorie (z. B. "Reise MA")
 * @return {string|null}          Lexware-UUID der Kategorie oder null
 */
function findLexwarePostingCategoryId_(categoryName) {
    if (!categoryName) return null;
    var nameStr = String(categoryName).trim();

    // ---- 1. Sheet-Lookup -----------------------------------------------
    var props = PropertiesService.getScriptProperties();
    var kategorienSheetName = (props.getProperty("LEXWARE_KATEGORIEN_SHEET_NAME") || "Lexware_Kategorien").trim();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
        var kategorienSheet = ss.getSheetByName(kategorienSheetName);
        if (kategorienSheet && kategorienSheet.getLastRow() > 1) {
            // Header row: ID(1) | Name(2) | Typ(3) | API-Typ(4)
            var sheetData = kategorienSheet.getRange(2, 1, kategorienSheet.getLastRow() - 1, 2).getValues();
            for (var s = 0; s < sheetData.length; s++) {
                if (String(sheetData[s][1] || "").trim() === nameStr) {
                    var sheetId = String(sheetData[s][0] || "").trim();
                    if (sheetId) {
                        Logger.log(
                            "Fixkosten: Buchungskategorie '" + nameStr +
                            "' aus Sheet '" + kategorienSheetName + "' gefunden (ID=" + sheetId + ")."
                        );
                        return sheetId;
                    }
                }
            }
            Logger.log(
                "Fixkosten: Buchungskategorie '" + nameStr +
                "' nicht im Sheet '" + kategorienSheetName + "' gefunden – versuche API."
            );
        }
    }

    // ---- 2. API-Fallback -----------------------------------------------
    var result;
    try {
        result = lexwareRequest("/posting-categories");
    } catch (e) {
        Logger.log("Fixkosten: Buchungskategorie-Suche fehlgeschlagen für '" + nameStr + "': " + e.message);
        return null;
    }

    var body = result.body;
    var categories = [];

    if (Array.isArray(body)) {
        categories = body;
    } else if (body && Array.isArray(body.content)) {
        categories = body.content;
    } else if (body && Array.isArray(body.categories)) {
        categories = body.categories;
    }

    for (var i = 0; i < categories.length; i++) {
        var cat = categories[i];
        if (String(cat.name || "").trim() === nameStr) {
            return String(cat.id);
        }
    }

    Logger.log("Fixkosten: Keine Buchungskategorie für Name '" + nameStr + "' gefunden.");
    return null;
}

// ---- Voucher creation --------------------------------------

/**
 * Erzeugt eine Eingangsrechnung (purchaseinvoice) in Lexware Office.
 *
 * Unterstützt sowohl Einzelpositionen (params.betragBrutto / params.mwstSatz /
 * params.categoryId) als auch mehrere Positionen über params.lineItems.
 *
 * @param {Object} params
 * @param {string} params.contactId         Lexware-UUID des Lieferanten
 * @param {string} params.lieferantennummer Lieferantennummer (für Belegnummer-Generierung)
 * @param {string} params.voucherDate       Belegdatum (JJJJ-MM-TT)
 * @param {string} params.dueDate           Fälligkeitsdatum (JJJJ-MM-TT)
 * @param {string} [params.konto_iban]      IBAN des abbuchenden Kontos (optional, in Notiz)
 * @param {string} [params.notiz]           Freitext-Notiz (optional)
 *
 * Einzelposition (rückwärtskompatibel):
 * @param {string} [params.kategorieName]   Buchungskategoriename (nur für Logging)
 * @param {string} [params.categoryId]      Lexware-UUID der Buchungskategorie
 * @param {number} [params.betragBrutto]    Bruttobetrag
 * @param {number} [params.mwstSatz]        Mehrwertsteuersatz in % (0, 7 oder 19)
 *
 * Mehrere Positionen (bevorzugt):
 * @param {Array}  [params.lineItems]       Array von Positionen; jede hat:
 *   @param {number} lineItems[].betragBrutto
 *   @param {number} lineItems[].mwstSatz
 *   @param {string} [lineItems[].categoryId]
 *   @param {string} [lineItems[].kategorieName]  (nur für Logging)
 *
 * @return {string}  Lexware-UUID des erstellten Belegs
 */
function createLexwarePurchaseInvoice_(params) {
    // Normalize to lineItems array (supports both single-item legacy and multi-item call)
    var lineItems = params.lineItems;
    if (!lineItems || !lineItems.length) {
        lineItems = [{
            betragBrutto: params.betragBrutto,
            mwstSatz:     params.mwstSatz,
            categoryId:   params.categoryId,
            kategorieName: params.kategorieName
        }];
    }

    // Build voucherItems and accumulate totals
    var totalGross = 0;
    var totalTax   = 0;
    var voucherItems = [];

    lineItems.forEach(function (item) {
        var gross   = round2(Number(item.betragBrutto) || 0);
        var taxRate = Number(item.mwstSatz) || 0;
        var tax     = round2(gross - gross / (1 + taxRate / 100));
        totalGross  = round2(totalGross + gross);
        totalTax    = round2(totalTax + tax);

        var vi = { amount: gross, taxAmount: tax, taxRatePercent: taxRate };
        if (item.categoryId) vi.categoryId = item.categoryId;
        voucherItems.push(vi);
    });

    // For the voucher number use the category of the first (or only) item
    var voucherNumber = buildFixkostenVoucherNumber_({
        voucherDate:      params.voucherDate,
        lieferantennummer: params.lieferantennummer,
        kategorieName:    lineItems.length === 1 ? (lineItems[0].kategorieName || "") : "",
        betragBrutto:     totalGross
    });

    // Build remark: include IBAN of debit account if provided
    var remark = params.notiz ? String(params.notiz).trim() : "";
    if (params.konto_iban) {
        var ibanNote = "Abbuchung von IBAN: " + String(params.konto_iban).trim();
        remark = remark ? remark + " | " + ibanNote : ibanNote;
    }

    var payload = {
        type: "purchaseinvoice",
        voucherNumber: voucherNumber,
        voucherDate: params.voucherDate,
        dueDate: params.dueDate,
        totalGrossAmount: totalGross,
        totalTaxAmount: totalTax,
        taxType: "gross",
        contactId: params.contactId,
        voucherItems: voucherItems
    };

    if (remark) {
        payload.remark = remark;
    }

    var result = lexwarePostRequest_("/vouchers", payload);
    var body = result.body;

    var voucherId = String(
        (body && (body.id || body.voucherId || body.uuid)) || ""
    );

    var categoryLog = lineItems.length === 1
        ? (lineItems[0].kategorieName || "–")
        : lineItems.map(function (li) { return li.kategorieName || "–"; }).join(", ");

    Logger.log(
        "Fixkosten: Beleg erstellt – " + lineItems.length + " Position(en)" +
        " [" + categoryLog + "]" +
        ", Brutto=" + totalGross +
        ", ID=" + voucherId
    );

    return voucherId;
}

// ---- Main entry point --------------------------------------

/**
 * Liest das Tabellenblatt "Lexware Fixkosten" und erzeugt für jeden
 * fälligen, aktiven Eintrag eine Eingangsrechnung in Lexware.
 *
 * Zeilen mit derselben nichtleeren Beleggruppe (Spalte N) werden zu
 * einer einzigen Eingangsrechnung mit mehreren Positionen zusammengefasst.
 * Dabei bestimmt die erste Zeile der Gruppe: Lieferantennummer, Rhythmus,
 * Fälligkeitstag/-monat, Konto_IBAN, Aktiv und Notiz.
 * Jede Zeile der Gruppe liefert eine Position: Kategorie, Betrag_Brutto,
 * MwSt_Satz.
 *
 * Zurückgeschrieben werden (alle Zeilen einer Gruppe):
 *   - Spalte L (Zuletzt_Gebucht)  – heutiges Datum
 *   - Spalte M (Lexware_Beleg_ID) – UUID des erstellten Belegs
 *
 * @return {{ok:boolean, created:number, skipped:number, errors:number}}
 */
function createLexwareFixkosten() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("No active spreadsheet");

    // Kategorien-Sheet aktualisieren, damit findLexwarePostingCategoryId_()
    // die aktuellen UUIDs aus dem Sheet lesen kann.
    try {
        importLexwareKategorien();
    } catch (e) {
        Logger.log("Fixkosten: Kategorien-Import fehlgeschlagen (wird fortgesetzt): " + e.message);
    }

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
    var colMap = buildColMap_(sheet);
    Logger.log("Fixkosten: Spalten-Map: " + JSON.stringify(colMap));

    var numCols = sheet.getLastColumn();
    var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    var now = new Date();

    var created = 0;
    var skipped = 0;
    var errors  = 0;

    // Pre-compute write-back column indices (1-based)
    var zuletztGebuchtCol = writeCol_(colMap, "Zuletzt_Gebucht");
    var lexwareBelegIdCol = writeCol_(colMap, "Lexware_Beleg_ID", "Lexware_Beleg_I");

    // Cache contact IDs and category UUIDs to avoid repeated API calls
    var contactCache = {};
    var contactIndex = null;
    var categoryCache = {};

    // ---- Pass 1: collect candidate rows ---------------------------------
    // Each entry: { rowNum, beleggruppe, kategorieName, lieferant,
    //               lieferantennummer, betragBrutto, mwstSatz, rhythmus,
    //               faelligkeitstag, faelligkeitsmonat, kontoIban, notiz,
    //               zuletztGebucht }
    var candidates = [];

    for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var rowNum = i + 2; // 1-based sheet row

        var kategorieName     = String(readCell_(row, colMap, "Kategorie", "Kategorie_Nr")       || "").trim();
        var lieferant         = String(readCell_(row, colMap, "Lieferant")                       || "").trim();
        var lieferantennummer = String(readCell_(row, colMap, "Lieferantennummer", "Lieferant")  || "").trim();
        var betragBrutto      = Number(readCell_(row, colMap, "Betrag_Brutto"))                  || 0;
        var mwstSatz          = Number(readCell_(row, colMap, "MwSt_Satz"))                      || 0;
        var rhythmus          = String(readCell_(row, colMap, "Rhythmus")                        || "").trim();
        var faelligkeitstag   = readCell_(row, colMap, "Fälligkeitstag");
        var faelligkeitsmonat = readCell_(row, colMap, "Fälligkeitsmonat");
        var kontoIban         = String(readCell_(row, colMap, "Konto_IBAN")                      || "").trim();
        var aktiv             = readCell_(row, colMap, "Aktiv");
        var notiz             = String(readCell_(row, colMap, "Notiz")                           || "").trim();
        var zuletztGebucht    = String(readCell_(row, colMap, "Zuletzt_Gebucht")                 || "").trim();
        var beleggruppe       = String(readCell_(row, colMap, "Beleggruppe")                     || "").trim();

        // Skip empty rows
        if (!kategorieName && !lieferantennummer) { skipped++; continue; }

        // Skip inactive rows
        if (aktiv === false || String(aktiv).toUpperCase() === "FALSE" || aktiv === 0) {
            Logger.log("Fixkosten: Zeile " + rowNum + " (Kat. " + kategorieName + ") – inaktiv, übersprungen.");
            skipped++;
            continue;
        }

        // Validate required fields for standalone rows.
        // For grouped rows the lieferantennummer of the first row is used; log here for traceability.
        if (!lieferantennummer) {
            Logger.log(
                "Fixkosten: Zeile " + rowNum + " (Kat. " + kategorieName + ")" +
                " – Lieferantennummer fehlt, übersprungen."
            );
            skipped++;
            continue;
        }
        if (!betragBrutto || betragBrutto <= 0) {
            Logger.log("Fixkosten: Zeile " + rowNum + " (Kat. " + kategorieName + ") – Betrag fehlt oder 0, übersprungen.");
            skipped++;
            continue;
        }

        candidates.push({
            rowNum:            rowNum,
            beleggruppe:       beleggruppe,
            kategorieName:     kategorieName,
            lieferant:         lieferant,
            lieferantennummer: lieferantennummer,
            betragBrutto:      betragBrutto,
            mwstSatz:          mwstSatz,
            rhythmus:          rhythmus,
            faelligkeitstag:   faelligkeitstag,
            faelligkeitsmonat: faelligkeitsmonat,
            kontoIban:         kontoIban,
            notiz:             notiz,
            zuletztGebucht:    zuletztGebucht
        });
    }

    // ---- Pass 2: group candidates ----------------------------------------
    // Rows without a Beleggruppe are each their own group (standalone).
    // Rows sharing the same non-empty Beleggruppe form a single invoice.
    var groupMap = {}; // beleggruppe string → [candidate, ...]
    var groupOrder = []; // preserves insertion order for deterministic processing

    candidates.forEach(function (c) {
        if (c.beleggruppe) {
            if (!Object.prototype.hasOwnProperty.call(groupMap, c.beleggruppe)) {
                groupMap[c.beleggruppe] = [];
                groupOrder.push(c.beleggruppe);
            }
            groupMap[c.beleggruppe].push(c);
        } else {
            // Standalone: use a unique key so it does not collide with named groups
            var standaloneKey = "__standalone_row_" + c.rowNum;
            groupMap[standaloneKey] = [c];
            groupOrder.push(standaloneKey);
        }
    });

    // ---- Pass 3: process each group --------------------------------------
    groupOrder.forEach(function (key) {
        var group = groupMap[key];
        var first = group[0]; // invoice header comes from the first row of the group

        // For due-date check: use the latest Zuletzt_Gebucht across all rows
        // to prevent re-booking a group where only some rows were already written back.
        var latestZuletztGebucht = group.reduce(function (max, c) {
            if (!c.zuletztGebucht) return max;
            return (!max || c.zuletztGebucht > max) ? c.zuletztGebucht : max;
        }, "");

        // Check due date (use first row's schedule)
        var dueInfo = isFixkostenDue_(
            first.rhythmus,
            first.faelligkeitstag,
            first.faelligkeitsmonat,
            latestZuletztGebucht,
            now
        );
        if (!dueInfo) {
            var groupLabel = first.beleggruppe || ("Zeile " + first.rowNum);
            Logger.log("Fixkosten: Gruppe '" + groupLabel + "' – noch nicht fällig.");
            skipped += group.length;
            return;
        }

        // Look up contact by Lieferantennummer (cached per number)
        var rawLieferantennummerCell = readCell_(data[first.rowNum - 2], colMap, "Lieferantennummer", "Lieferant");
        Logger.log(
            "Fixkosten: Zeile " + first.rowNum + " – Rohwert Lieferantennummer: " +
            JSON.stringify(rawLieferantennummerCell) +
            " (Typ=" + typeof rawLieferantennummerCell + ")" +
            ", bereinigt: '" + first.lieferantennummer + "'"
        );
        var contactId = contactCache[first.lieferantennummer];
        if (!contactId) {
            if (!contactIndex) contactIndex = buildLexwareContactNumberIndex_();
            contactId = findLexwareContactIdInIndex_(first.lieferantennummer, contactIndex);
        }
        if (!contactId) {
            contactId = findLexwareContactIdByNumber_(first.lieferantennummer);
            if (contactId) contactCache[first.lieferantennummer] = contactId;
        }
        if (!contactId) {
            Logger.log(
                "Fixkosten: Zeile " + first.rowNum + " (Kat. " + first.kategorieName + ")" +
                " – Lieferantennummer '" + first.lieferantennummer +
                "' (" + (first.lieferant || "?") + ") nicht in Lexware gefunden, übersprungen."
            );
            errors += group.length;
            return;
        }

        // Resolve category UUID for each line item (cached per name)
        var lineItems = group.map(function (c) {
            var categoryId = null;
            if (c.kategorieName) {
                if (Object.prototype.hasOwnProperty.call(categoryCache, c.kategorieName)) {
                    categoryId = categoryCache[c.kategorieName];
                } else {
                    categoryId = findLexwarePostingCategoryId_(c.kategorieName);
                    categoryCache[c.kategorieName] = categoryId;
                }
                if (!categoryId) {
                    Logger.log(
                        "Fixkosten: Zeile " + c.rowNum + " (Kat. " + c.kategorieName + ")" +
                        " – Buchungskategorie nicht gefunden, Position wird ohne Kategorie erstellt."
                    );
                }
            }
            return {
                betragBrutto:  c.betragBrutto,
                mwstSatz:      c.mwstSatz,
                categoryId:    categoryId,
                kategorieName: c.kategorieName
            };
        });

        // Create the purchase invoice (one invoice, potentially many line items)
        try {
            var voucherId = createLexwarePurchaseInvoice_({
                contactId:         contactId,
                voucherDate:       dueInfo.voucherDate,
                dueDate:           dueInfo.dueDate,
                lieferantennummer: first.lieferantennummer,
                konto_iban:        first.kontoIban,
                notiz:             first.notiz,
                lineItems:         lineItems
            });

            // Write back booking date and voucher ID to ALL rows of this group
            group.forEach(function (c) {
                if (zuletztGebuchtCol > 0) {
                    sheet.getRange(c.rowNum, zuletztGebuchtCol).setValue(formatDate_(now));
                }
                if (lexwareBelegIdCol > 0) {
                    sheet.getRange(c.rowNum, lexwareBelegIdCol).setValue(voucherId);
                }
            });

            var groupLabel = first.beleggruppe || ("Zeile " + first.rowNum);
            Logger.log(
                "Fixkosten: ✅ Gruppe '" + groupLabel + "'" +
                " (" + group.length + " Position(en))" +
                " – Beleg erstellt: " + voucherId
            );
            created++;
        } catch (e) {
            var groupLabelErr = first.beleggruppe || ("Zeile " + first.rowNum);
            Logger.log(
                "Fixkosten: ❌ Gruppe '" + groupLabelErr + "'" +
                " – Fehler beim Erstellen: " + e.message
            );
            errors += group.length;
        }
    });

    Logger.log(
        "Fixkosten abgeschlossen: erstellt=" + created +
        ", übersprungen=" + skipped +
        ", Fehler=" + errors
    );

    return { ok: errors === 0, created: created, skipped: skipped, errors: errors };
}
