// ============================================================
// Lexware Kunden – Kontakte in separatem Tabellenblatt speichern
//
// Speichert alle Lexware-Kontakte (Kunden und Lieferanten) im
// Tabellenblatt "Lexware Kunden" und stellt sie zur Wiederverwendung
// in anderen Skripten (Fixkosten, Manuelle Umsätze) bereit.
//
// Hauptfunktionen:
//   syncLexwareKundenSheet()  – Synchronisiert das Blatt mit den
//                               aktuellen Kontakten aus Lexware.
//   setupLexwareKundenSheet() – Erstellt das Blatt mit Überschriften
//                               (wird von syncLexwareKundenSheet()
//                               automatisch aufgerufen).
//
// Spalten im Blatt "Lexware Kunden":
//
//  Spalte A  UUID              – Lexware-interne ID (Primärschlüssel)
//  Spalte B  Nummer            – Allgemeine Kontaktnummer (falls vorhanden)
//  Spalte C  Kundennummer      – Nummer in der Kunden-Rolle (falls vorhanden)
//  Spalte D  Lieferantennummer – Nummer in der Lieferanten-Rolle (falls vorhanden)
//  Spalte E  Name              – Anzeigename (Firmenname oder Personenname)
//  Spalte F  Typ               – Kunde | Lieferant | Kunde/Lieferant | Kontakt
//  Spalte G  Notiz             – Freitextnotiz am Kontakt (contact.note)
//
// Integration:
//   buildLexwareContactNumberIndex_() (in lexwareFixkosten.js) verwendet
//   dieses Blatt als primären Cache. Nur wenn das Blatt leer ist oder der
//   gesuchte Kontakt dort nicht gefunden wird, wird die Lexware-API abgefragt.
//
// Script Properties (optional):
//   LEXWARE_KUNDEN_SHEET_NAME  – Name des Tabellenblatts
//                                (Standard: "Lexware Kunden")
// ============================================================

var LEXWARE_KUNDEN_DEFAULT_SHEET_NAME = "Lexware Kunden";

var LEXWARE_KUNDEN_HEADERS = [
    "UUID",               // A  1  (Lexware-UUID, Primärschlüssel für Upsert)
    "Nummer",             // B  2  (allgemeine Kontaktnummer, falls vorhanden)
    "Kundennummer",       // C  3  (roles.customer.number)
    "Lieferantennummer",  // D  4  (roles.vendor.number)
    "Name",               // E  5  (Anzeigename: Firmenname oder Vor-/Nachname)
    "Typ",                // F  6  (Kunde | Lieferant | Kunde/Lieferant | Kontakt)
    "Notiz"               // G  7  (Freitextnotiz am Kontakt, contact.note)
];

// ---- Config ------------------------------------------------

function getKundenSheetName_() {
    var props = PropertiesService.getScriptProperties();
    return (
        props.getProperty("LEXWARE_KUNDEN_SHEET_NAME") ||
        LEXWARE_KUNDEN_DEFAULT_SHEET_NAME
    ).trim();
}

// ---- Sheet setup -------------------------------------------

/**
 * Erstellt das Tabellenblatt "Lexware Kunden" mit den erforderlichen
 * Spaltenüberschriften, falls es noch nicht existiert.
 * Kann manuell ausgeführt werden, um das Blatt vorzubereiten.
 *
 * @return {{ok:boolean, sheet:string}}
 */
function setupLexwareKundenSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("No active spreadsheet");

    var sheetName = getKundenSheetName_();
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        Logger.log("Kunden: Blatt '" + sheetName + "' erstellt.");
    }

    if (sheet.getLastRow() === 0) {
        sheet.appendRow(LEXWARE_KUNDEN_HEADERS);
        sheet.getRange(1, 1, 1, LEXWARE_KUNDEN_HEADERS.length).setFontWeight("bold");
        Logger.log("Kunden: Spaltenüberschriften gesetzt.");
    } else {
        Logger.log("Kunden: Blatt existiert bereits, keine Änderung an der Kopfzeile.");
    }

    return { ok: true, sheet: sheetName };
}

// ---- Contact name helper -----------------------------------

/**
 * Leitet den Anzeigenamen aus einem Lexware-Kontaktobjekt ab.
 * Bevorzugt den Firmennamen, dann den Personennamen (Vor- und Nachname),
 * zuletzt displayName/name als Rückfall.
 *
 * @param  {Object} contact  Roher Kontakteintrag aus der Lexware-API.
 * @return {string}
 */
function getContactDisplayName_(contact) {
    if (contact.company && contact.company.name) {
        return String(contact.company.name).trim();
    }
    if (contact.person) {
        var parts = [];
        if (contact.person.firstName) parts.push(String(contact.person.firstName).trim());
        if (contact.person.lastName)  parts.push(String(contact.person.lastName).trim());
        if (parts.length) return parts.join(" ");
    }
    return String(contact.displayName || contact.name || "").trim();
}

// ---- Sync --------------------------------------------------

/**
 * Ruft alle Kontakte (Kunden und Lieferanten) aus Lexware Office ab
 * und schreibt sie in das Tabellenblatt "Lexware Kunden".
 *
 * Vorhandene Zeilen werden anhand der UUID aktualisiert (Upsert);
 * neue Kontakte werden am Ende eingefügt.
 *
 * Nach der Synchronisation kann das Blatt in anderen Skripten als
 * Kontakt-Cache genutzt werden, sodass keine wiederholten API-Abfragen
 * erforderlich sind.
 *
 * @return {{ok:boolean, sheet:string, total:number, inserted:number, updated:number}}
 */
function syncLexwareKundenSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("No active spreadsheet");

    var sheetName = getKundenSheetName_();
    var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

    // Kopfzeile setzen, falls das Blatt leer ist
    if (sheet.getLastRow() === 0) {
        sheet.appendRow(LEXWARE_KUNDEN_HEADERS);
        sheet.getRange(1, 1, 1, LEXWARE_KUNDEN_HEADERS.length).setFontWeight("bold");
    }

    // Alle Kontakte (Kunden + Lieferanten + sonstige) seitenweise abrufen.
    // Einmalkunden (Kunden ohne Kundennummer) erscheinen nur bei customer=true,
    // daher wird der Abruf in zwei Durchläufen ausgeführt und per UUID zusammengeführt.
    var contactsById = {};
    var pageSize = 100;

    var fetchPasses = [
        { label: "alle Kontakte",  params: {} },
        { label: "Einmalkunden",   params: { customer: true } }
    ];

    for (var p = 0; p < fetchPasses.length; p++) {
        var pass = fetchPasses[p];
        var page = 0;
        var totalPages = 1;

        do {
            var result;
            try {
                var queryParams = { page: page, size: pageSize };
                var passKeys = Object.keys(pass.params);
                for (var k = 0; k < passKeys.length; k++) {
                    queryParams[passKeys[k]] = pass.params[passKeys[k]];
                }
                result = lexwareRequest("/contacts", queryParams);
            } catch (e) {
                Logger.log(
                    "Kunden: Kontakt-Abruf fehlgeschlagen (" + pass.label +
                    ", Seite " + page + "): " + e.message
                );
                break;
            }

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

            for (var ci = 0; ci < pageContacts.length; ci++) {
                var c = pageContacts[ci];
                var cid = String(c.id || "").trim();
                if (cid && !Object.prototype.hasOwnProperty.call(contactsById, cid)) {
                    contactsById[cid] = c;
                }
            }

            totalPages = (body && body.page && body.page.totalPages !== undefined)
                ? body.page.totalPages
                : (body && body.totalPages !== undefined
                    ? body.totalPages
                    : (pageContacts.length === pageSize ? page + 2 : page + 1));
            page++;
        } while (page < totalPages);

        Logger.log(
            "Kunden: Durchlauf '" + pass.label + "' abgeschlossen (" +
            totalPages + " Seite(n))."
        );
    }

    var contacts = Object.keys(contactsById).map(function (id) { return contactsById[id]; });

    Logger.log(
        "Kunden: " + contacts.length + " eindeutige Kontakt(e) nach " +
        fetchPasses.length + " Durchläufen abgerufen."
    );

    // Vorhandene Zeilen nach UUID indizieren (Upsert-Logik)
    var existingById = {};
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        var existingData = sheet.getRange(
            2, 1, lastRow - 1, LEXWARE_KUNDEN_HEADERS.length
        ).getValues();
        existingData.forEach(function (row, idx) {
            var id = String(row[0] || "").trim();
            if (id) existingById[id] = { rowIndex: idx + 2, data: row };
        });
    }

    var newRows = [];
    var updatedCount = 0;

    contacts.forEach(function (contact) {
        var uuid = String(contact.id || "").trim();
        if (!uuid) return;

        var nummer         = String(contact.number || "").trim();
        var kundennummer   = String(
            contact.roles && contact.roles.customer
                ? contact.roles.customer.number || ""
                : ""
        ).trim();
        var lieferantennum = String(
            contact.roles && contact.roles.vendor
                ? contact.roles.vendor.number || ""
                : ""
        ).trim();
        var name = getContactDisplayName_(contact);

        var hasCustomer = !!(contact.roles && contact.roles.customer);
        var hasVendor   = !!(contact.roles && contact.roles.vendor);
        var typ = hasCustomer && hasVendor ? "Kunde/Lieferant"
                : hasCustomer             ? "Kunde"
                : hasVendor               ? "Lieferant"
                :                          "Kontakt";
        var notiz = String(contact.note || "").trim();

        var row = [uuid, nummer, kundennummer, lieferantennum, name, typ, notiz];

        if (existingById[uuid]) {
            var existing = existingById[uuid].data;
            var changed = row.some(function (val, i) {
                return String(val) !== String(existing[i]);
            });
            if (changed) {
                sheet.getRange(existingById[uuid].rowIndex, 1, 1, row.length).setValues([row]);
                updatedCount++;
            }
        } else {
            newRows.push(row);
        }
    });

    if (newRows.length > 0) {
        sheet.getRange(
            sheet.getLastRow() + 1, 1, newRows.length, LEXWARE_KUNDEN_HEADERS.length
        ).setValues(newRows);
    }

    Logger.log(
        "Kunden: Sync abgeschlossen – total=" + contacts.length +
        ", eingefügt=" + newRows.length +
        ", aktualisiert=" + updatedCount
    );

    return {
        ok: true,
        sheet: sheetName,
        total: contacts.length,
        inserted: newRows.length,
        updated: updatedCount
    };
}

// ---- Sheet-based contact index -----------------------------

/**
 * Liest das Tabellenblatt "Lexware Kunden" und erstellt einen Index
 * Kontaktnummer → { id: UUID, ambiguous: boolean, source: string }.
 *
 * Indexiert werden:
 *   - Nummer (Spalte B)
 *   - Kundennummer (Spalte C)
 *   - Lieferantennummer (Spalte D)
 *
 * Wenn dieselbe Nummer mehrfach vorkommt und auf unterschiedliche UUIDs
 * zeigt, wird sie als nicht eindeutig markiert (ambiguous: true).
 *
 * Gibt null zurück, wenn das Blatt nicht vorhanden oder leer ist,
 * damit der Aufrufer auf API-Abfragen zurückfallen kann.
 *
 * @return {Object|null}  Index oder null wenn kein Blatt vorhanden / leer.
 */
function buildContactIndexFromKundenSheet_() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return null;

    var sheetName = getKundenSheetName_();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return null;

    var data = sheet.getRange(
        2, 1, sheet.getLastRow() - 1, LEXWARE_KUNDEN_HEADERS.length
    ).getValues();
    if (!data || data.length === 0) return null;

    var index = {};

    function addToIndex(numberValue, contactId, source) {
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

    for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var uuid           = String(row[0] || "").trim();
        var nummer         = String(row[1] || "").trim();
        var kundennummer   = String(row[2] || "").trim();
        var lieferantennum = String(row[3] || "").trim();

        if (!uuid) continue;

        addToIndex(nummer,         uuid, "Nummer");
        addToIndex(kundennummer,   uuid, "Kundennummer");
        addToIndex(lieferantennum, uuid, "Lieferantennummer");
    }

    Logger.log(
        "Kunden: Kontaktindex aus Sheet erstellt – " +
        Object.keys(index).length + " Einträge."
    );
    return index;
}

/**
 * Liest das Tabellenblatt "Lexware Kunden" und erstellt einen Index
 * UUID → Lieferantennummer.
 *
 * Wird verwendet, um aus der Kontakt-UUID eines Belegs die Lieferantennummer
 * zu ermitteln, da die Lexware-Voucher-API nur die Kontakt-UUID zurückgibt,
 * nicht aber die volle Kontakt-Rolle mit der Lieferantennummer.
 *
 * Gibt ein leeres Objekt zurück, wenn das Blatt nicht vorhanden oder leer ist.
 *
 * @return {Object}  { contactUUID → lieferantennummer }
 */
function buildContactIdToVendorNumberIndex_() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return {};

    var sheetName = getKundenSheetName_();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return {};

    var data = sheet.getRange(
        2, 1, sheet.getLastRow() - 1, LEXWARE_KUNDEN_HEADERS.length
    ).getValues();
    if (!data || data.length === 0) return {};

    var index = {};
    for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var uuid           = String(row[0] || "").trim();
        var lieferantennum = String(row[3] || "").trim();
        if (uuid && lieferantennum) {
            index[uuid] = lieferantennum;
        }
    }

    Logger.log(
        "Kunden: UUID→Lieferantennummer-Index erstellt – " +
        Object.keys(index).length + " Einträge."
    );
    return index;
}
