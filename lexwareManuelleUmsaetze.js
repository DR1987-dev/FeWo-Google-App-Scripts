// ============================================================
// Lexware Manuelle Umsätze – einmalige Einnahmen/Ausgaben erstellen
//
// Liest das Tabellenblatt "Lexware Manuelle Umsätze" und erstellt
// für jeden noch nicht gebuchten, aktiven Eintrag einmalig einen
// Beleg (Einnahme oder Ausgabe) in Lexware Office.
//
// Mehrere Positionen (lineItems) pro Beleg werden unterstützt:
// Alle Zeilen mit derselben Beleg_Ref bilden einen gemeinsamen
// Beleg. Die Belegkopf-Felder (Typ, Kontaktnummer, Datum etc.)
// werden von der ersten Zeile der Gruppe übernommen.
//
// Erforderliche Spalten im Blatt "Lexware Manuelle Umsätze":
//
//  Spalte A  Beleg_Ref          – Eindeutige Kennung des Belegs
//                                  (z. B. "RG-2024-001"). Alle Zeilen mit
//                                  derselben Beleg_Ref gehören zusammen.
//  Spalte B  Typ                – salesinvoice (Einnahme)
//                                  purchaseinvoice (Ausgabe)
//  Spalte C  Kontaktnummer      – Kundennummer / Lieferantennummer in Lexware
//  Spalte D  Kontakt            – Anzeigename (nur zur Übersicht, kein API-Lookup)
//  Spalte E  Belegdatum         – Belegdatum (JJJJ-MM-TT oder Datumszelle)
//  Spalte F  Fälligkeitsdatum   – Fälligkeitsdatum (optional)
//  Spalte G  Belegnummer        – Eigene Belegnummer (optional; wird sonst
//                                  automatisch aus Beleg_Ref generiert)
//  Spalte H  Notiz              – Freitext, wird als Remark in Lexware übernommen
//  Spalte I  Aktiv              – TRUE/FALSE – wird nur verarbeitet wenn TRUE
//  Spalte J  Kategorie          – Buchungskategoriename für diese Position
//                                  (z. B. "Reise MA")
//  Spalte K  Betrag_Brutto      – Bruttobetrag für diese Position
//  Spalte L  MwSt_Satz          – Steuersatz in % für diese Position (0, 7 oder 19)
//  Spalte M  Zuletzt_Gebucht    – wird vom Skript zurückgeschrieben (JJJJ-MM-TT)
//                                  nur in die erste Zeile der Gruppe
//  Spalte N  Lexware_Beleg_ID   – wird vom Skript zurückgeschrieben (UUID)
//                                  nur in die erste Zeile der Gruppe
//
// Mehrere Positionen für einen Beleg:
//   Trage jede Position in einer eigenen Zeile ein. Alle Zeilen mit
//   derselben Beleg_Ref (Spalte A) werden zu einem Beleg zusammengefasst.
//   Die Felder Typ, Kontaktnummer, Belegdatum usw. werden nur aus der
//   ersten Zeile der Gruppe gelesen.
//
// Script Properties (optional):
//   MANUELLE_UMSAETZE_SHEET_NAME  – Name des Tabellenblatts
//                                    (Standard: "Lexware Manuelle Umsätze")
// ============================================================

var MANUELLE_UMSAETZE_DEFAULT_SHEET_NAME = "Lexware Manuelle Umsätze";

var MANUELLE_UMSAETZE_HEADERS = [
    "Beleg_Ref",           // A  1
    "Typ",                 // B  2  (salesinvoice | purchaseinvoice)
    "Kontaktnummer",       // C  3
    "Kontakt",             // D  4  (Anzeigename, kein API-Lookup)
    "Belegdatum",          // E  5
    "Fälligkeitsdatum",    // F  6
    "Belegnummer",         // G  7  (optional, sonst automatisch)
    "Notiz",               // H  8
    "Aktiv",               // I  9
    "Kategorie",           // J  10
    "Betrag_Brutto",       // K  11
    "MwSt_Satz",           // L  12
    "Zuletzt_Gebucht",     // M  13
    "Lexware_Beleg_ID"     // N  14
];

// ---- Config ------------------------------------------------

function getManuelleUmsaetzeSheetName_() {
    var props = PropertiesService.getScriptProperties();
    return (
        props.getProperty("MANUELLE_UMSAETZE_SHEET_NAME") ||
        MANUELLE_UMSAETZE_DEFAULT_SHEET_NAME
    ).trim();
}

// ---- Sheet setup -------------------------------------------

/**
 * Erstellt das Tabellenblatt "Lexware Manuelle Umsätze" mit den
 * erforderlichen Spaltenüberschriften, falls es noch nicht existiert.
 * Kann manuell ausgeführt werden, um das Blatt vorzubereiten.
 */
function setupManuelleUmsaetzeSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("No active spreadsheet");

    var sheetName = getManuelleUmsaetzeSheetName_();
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        Logger.log("Manuelle Umsätze: Blatt '" + sheetName + "' erstellt.");
    }

    if (sheet.getLastRow() === 0) {
        sheet.appendRow(MANUELLE_UMSAETZE_HEADERS);
        sheet.getRange(1, 1, 1, MANUELLE_UMSAETZE_HEADERS.length).setFontWeight("bold");
        Logger.log("Manuelle Umsätze: Spaltenüberschriften gesetzt.");
    } else {
        Logger.log("Manuelle Umsätze: Blatt existiert bereits, keine Änderung an der Kopfzeile.");
    }

    return { ok: true, sheet: sheetName };
}

// ---- Date helper -------------------------------------------

/**
 * Wandelt einen Zellwert (Date-Objekt oder String) in "JJJJ-MM-TT" um.
 * Gibt "" zurück wenn der Wert leer oder ungültig ist.
 *
 * @param  {*} value  Zellwert aus getValues()
 * @return {string}
 */
function toDateString_(value) {
    if (value === "" || value === null || value === undefined) return "";
    if (value instanceof Date) {
        if (isNaN(value.getTime())) return "";
        return formatDate_(value);
    }
    var str = String(value).trim();
    if (!str) return "";
    var d = new Date(str);
    if (!isNaN(d.getTime())) return formatDate_(d);
    // If the string is already in JJJJ-MM-TT format, return as-is
    if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
    return "";
}

// ---- Voucher number generation -----------------------------

/**
 * Generiert eine Belegnummer aus Beleg_Ref oder Datums-/Kontaktteilen.
 *
 * @param  {Object} params
 * @param  {string} params.belegRef      Beleg_Ref aus dem Sheet
 * @param  {string} params.voucherDate   Belegdatum (JJJJ-MM-TT)
 * @param  {string} params.kontaktnummer Kontaktnummer
 * @param  {string} params.typ           Belegtyp
 * @return {string}
 */
function buildManuelleUmsaetzeVoucherNumber_(params) {
    // Prefer a sanitised Beleg_Ref as voucher number
    var refPart = String(params.belegRef || "")
        .replace(/[^A-Za-z0-9\-_]/g, "")
        .slice(0, 60);
    if (refPart) return refPart;

    var typPart = String(params.typ || "").toLowerCase() === "salesinvoice" ? "ER" : "AR";
    var datePart = String(params.voucherDate || "").replace(/[^0-9]/g, "");
    var contactPart = String(params.kontaktnummer || "")
        .replace(/[^A-Za-z0-9]/g, "")
        .toUpperCase();

    return [typPart, datePart || "DATE", contactPart || "NA"].join("-").slice(0, 60);
}

// ---- Voucher creation --------------------------------------

/**
 * Erstellt einen Beleg (Einnahme oder Ausgabe) in Lexware Office.
 *
 * @param {Object}   params
 * @param {string}   params.typ           "salesinvoice" | "purchaseinvoice"
 * @param {string}   params.contactId     Lexware-UUID des Kontakts
 * @param {string}   params.kontaktnummer Kontaktnummer (für Belegnummer-Generierung)
 * @param {string}   params.belegRef      Beleg_Ref aus dem Sheet
 * @param {string}   params.voucherDate   Belegdatum (JJJJ-MM-TT)
 * @param {string}   [params.dueDate]     Fälligkeitsdatum (JJJJ-MM-TT, optional)
 * @param {string}   [params.belegnummer] Eigene Belegnummer (optional)
 * @param {string}   [params.notiz]       Freitext-Notiz (optional)
 * @param {Array}    params.lineItems     Positionen: [{kategorieName, categoryId,
 *                                          betragBrutto, mwstSatz}]
 * @return {string}  Lexware-UUID des erstellten Belegs
 */
function createLexwareManuellerUmsatz_(params) {
    var typ = String(params.typ || "").toLowerCase().trim();
    if (typ !== "salesinvoice" && typ !== "purchaseinvoice") {
        throw new Error(
            "Unbekannter Typ '" + params.typ +
            "'. Erlaubt: salesinvoice, purchaseinvoice."
        );
    }

    var voucherNumber = params.belegnummer && String(params.belegnummer).trim()
        ? String(params.belegnummer).trim()
        : buildManuelleUmsaetzeVoucherNumber_(params);

    var remark = params.notiz ? String(params.notiz).trim() : "";

    var voucherItems = [];
    var totalGrossAmount = 0;
    var totalTaxAmount = 0;

    for (var i = 0; i < params.lineItems.length; i++) {
        var item = params.lineItems[i];
        var grossAmount = round2(Number(item.betragBrutto) || 0);
        var taxRatePercent = Number(item.mwstSatz) || 0;
        var taxAmount = round2(grossAmount - grossAmount / (1 + taxRatePercent / 100));

        var voucherItem = {
            amount: grossAmount,
            taxAmount: taxAmount,
            taxRatePercent: taxRatePercent
        };
        if (item.categoryId) {
            voucherItem.categoryId = item.categoryId;
        }

        voucherItems.push(voucherItem);
        totalGrossAmount = round2(totalGrossAmount + grossAmount);
        totalTaxAmount = round2(totalTaxAmount + taxAmount);
    }

    var payload = {
        type: typ,
        voucherNumber: voucherNumber,
        voucherDate: params.voucherDate,
        totalGrossAmount: totalGrossAmount,
        totalTaxAmount: totalTaxAmount,
        taxType: "gross",
        contactId: params.contactId,
        voucherItems: voucherItems
    };

    if (params.dueDate) {
        payload.dueDate = params.dueDate;
    }
    if (remark) {
        payload.remark = remark;
    }

    var result = lexwarePostRequest_("/vouchers", payload);
    var body = result.body;

    var voucherId = String(
        (body && (body.id || body.voucherId || body.uuid)) || ""
    );

    Logger.log(
        "Manuelle Umsätze: Beleg erstellt – Typ=" + typ +
        ", Ref=" + params.belegRef +
        ", Brutto=" + totalGrossAmount +
        ", Positionen=" + voucherItems.length +
        ", ID=" + voucherId
    );

    return voucherId;
}

// ---- Main entry point --------------------------------------

/**
 * Liest das Tabellenblatt "Lexware Manuelle Umsätze" und erstellt für
 * jeden noch nicht gebuchten, aktiven Beleg einmalig einen Eintrag
 * in Lexware Office.
 *
 * Alle Zeilen mit derselben Beleg_Ref bilden einen gemeinsamen Beleg
 * (mehrere Positionen). Die Belegkopf-Felder werden von der ersten Zeile
 * der jeweiligen Gruppe übernommen.
 *
 * Zurückgeschrieben werden (nur in die erste Zeile der Gruppe):
 *   - Spalte M (Zuletzt_Gebucht)  – heutiges Datum
 *   - Spalte N (Lexware_Beleg_ID) – UUID des erstellten Belegs
 *
 * @return {{ok:boolean, created:number, skipped:number, errors:number}}
 */
function createLexwareManuelleUmsaetze() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("No active spreadsheet");

    // Kategorien-Sheet aktualisieren, damit findLexwarePostingCategoryId_()
    // die aktuellen UUIDs aus dem lokalen Sheet lesen kann.
    try {
        importLexwareKategorien();
    } catch (e) {
        Logger.log(
            "Manuelle Umsätze: Kategorien-Import fehlgeschlagen (wird fortgesetzt): " +
            e.message
        );
    }

    var sheetName = getManuelleUmsaetzeSheetName_();
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
        Logger.log(
            "Manuelle Umsätze: Tabellenblatt '" + sheetName +
            "' nicht gefunden – setupManuelleUmsaetzeSheet() ausführen."
        );
        return {
            ok: false, created: 0, skipped: 0, errors: 0,
            error: "Sheet not found: " + sheetName
        };
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
        Logger.log("Manuelle Umsätze: Keine Datenzeilen im Blatt '" + sheetName + "'.");
        return { ok: true, created: 0, skipped: 0, errors: 0 };
    }

    var colMap = buildColMap_(sheet);
    Logger.log("Manuelle Umsätze: Spalten-Map: " + JSON.stringify(colMap));

    var numCols = sheet.getLastColumn();
    var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    var now = new Date();

    // Pre-compute write-back column indices (1-based)
    var zuletztGebuchtCol = writeCol_(colMap, "Zuletzt_Gebucht");
    var lexwareBelegIdCol  = writeCol_(colMap, "Lexware_Beleg_ID");

    // Caches to avoid redundant API calls
    var contactCache  = {};
    var contactIndex  = null;
    var categoryCache = {};

    // ---- Group rows by Beleg_Ref (in order of first occurrence) ----
    var groupOrder = [];  // Beleg_Ref values in insertion order
    var groups = {};      // { belegRef: { firstRowNum, headerRow, rows: [{rowNum, row}] } }

    for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var rowNum = i + 2; // 1-based sheet row
        var belegRef = String(readCell_(row, colMap, "Beleg_Ref") || "").trim();

        if (!belegRef) {
            Logger.log(
                "Manuelle Umsätze: Zeile " + rowNum +
                " – Beleg_Ref fehlt, übersprungen."
            );
            continue;
        }

        if (!Object.prototype.hasOwnProperty.call(groups, belegRef)) {
            groupOrder.push(belegRef);
            groups[belegRef] = {
                firstRowNum: rowNum,
                headerRow: row,
                rows: []
            };
        }
        groups[belegRef].rows.push({ rowNum: rowNum, row: row });
    }

    // ---- Process each voucher group ----------------------------
    var created = 0;
    var skipped = 0;
    var errors  = 0;

    for (var g = 0; g < groupOrder.length; g++) {
        var ref   = groupOrder[g];
        var group = groups[ref];
        var headerRow   = group.headerRow;
        var firstRowNum = group.firstRowNum;

        // Read voucher-level fields from header row
        var typ               = String(readCell_(headerRow, colMap, "Typ")              || "").trim().toLowerCase();
        var kontaktnummer     = String(readCell_(headerRow, colMap, "Kontaktnummer")    || "").trim();
        var kontakt           = String(readCell_(headerRow, colMap, "Kontakt")          || "").trim();
        var belegdatum        = readCell_(headerRow, colMap, "Belegdatum");
        var faelligkeitsdatum = readCell_(headerRow, colMap, "Fälligkeitsdatum");
        var belegnummer       = String(readCell_(headerRow, colMap, "Belegnummer")      || "").trim();
        var notiz             = String(readCell_(headerRow, colMap, "Notiz")            || "").trim();
        var aktiv             = readCell_(headerRow, colMap, "Aktiv");
        var lexwareBelegId    = String(readCell_(headerRow, colMap, "Lexware_Beleg_ID") || "").trim();

        // Skip if already booked
        if (lexwareBelegId) {
            Logger.log(
                "Manuelle Umsätze: Beleg_Ref '" + ref +
                "' – bereits gebucht (ID=" + lexwareBelegId + "), übersprungen."
            );
            skipped++;
            continue;
        }

        // Skip inactive
        if (aktiv === false || String(aktiv).toUpperCase() === "FALSE" || aktiv === 0) {
            Logger.log(
                "Manuelle Umsätze: Beleg_Ref '" + ref + "' – inaktiv, übersprungen."
            );
            skipped++;
            continue;
        }

        // Validate type
        if (typ !== "salesinvoice" && typ !== "purchaseinvoice") {
            Logger.log(
                "Manuelle Umsätze: Beleg_Ref '" + ref +
                "' – ungültiger Typ '" + typ +
                "' (erlaubt: salesinvoice, purchaseinvoice), übersprungen."
            );
            errors++;
            continue;
        }

        // Validate contact number
        if (!kontaktnummer) {
            Logger.log(
                "Manuelle Umsätze: Beleg_Ref '" + ref +
                "' – Kontaktnummer fehlt, übersprungen."
            );
            errors++;
            continue;
        }

        // Format dates
        var voucherDateStr = toDateString_(belegdatum) || formatDate_(now);
        var dueDateStr     = toDateString_(faelligkeitsdatum);

        // ---- Collect line items from all rows of this group --------
        var lineItems = [];
        for (var r = 0; r < group.rows.length; r++) {
            var lineRow      = group.rows[r].row;
            var lineRowNum   = group.rows[r].rowNum;
            var kategorieName = String(readCell_(lineRow, colMap, "Kategorie")    || "").trim();
            var betragBrutto  = Number(readCell_(lineRow, colMap, "Betrag_Brutto")) || 0;
            var mwstSatz      = Number(readCell_(lineRow, colMap, "MwSt_Satz"))    || 0;

            if (!betragBrutto || betragBrutto === 0) {
                Logger.log(
                    "Manuelle Umsätze: Beleg_Ref '" + ref +
                    "', Zeile " + lineRowNum +
                    " – Betrag_Brutto fehlt oder 0, Position übersprungen."
                );
                continue;
            }

            // Look up posting category UUID (cached per name)
            var categoryId = null;
            if (kategorieName) {
                if (Object.prototype.hasOwnProperty.call(categoryCache, kategorieName)) {
                    categoryId = categoryCache[kategorieName];
                } else {
                    categoryId = findLexwarePostingCategoryId_(kategorieName);
                    categoryCache[kategorieName] = categoryId;
                }
                if (!categoryId) {
                    Logger.log(
                        "Manuelle Umsätze: Beleg_Ref '" + ref +
                        "' – Buchungskategorie '" + kategorieName +
                        "' nicht gefunden, Position wird ohne Kategorie erstellt."
                    );
                }
            }

            lineItems.push({
                kategorieName: kategorieName,
                categoryId:    categoryId,
                betragBrutto:  betragBrutto,
                mwstSatz:      mwstSatz
            });
        }

        if (lineItems.length === 0) {
            Logger.log(
                "Manuelle Umsätze: Beleg_Ref '" + ref +
                "' – keine gültigen Positionen, übersprungen."
            );
            skipped++;
            continue;
        }

        // ---- Look up contact (cached per Kontaktnummer) ------------
        var contactId = contactCache[kontaktnummer];
        if (!contactId) {
            if (!contactIndex) contactIndex = buildLexwareContactNumberIndex_();
            contactId = findLexwareContactIdInIndex_(kontaktnummer, contactIndex);
        }
        if (!contactId) {
            contactId = findLexwareContactIdByNumber_(kontaktnummer);
            if (contactId) contactCache[kontaktnummer] = contactId;
        }
        if (!contactId) {
            Logger.log(
                "Manuelle Umsätze: Beleg_Ref '" + ref +
                "' – Kontaktnummer '" + kontaktnummer +
                "' (" + (kontakt || "?") + ") nicht in Lexware gefunden, übersprungen."
            );
            errors++;
            continue;
        }

        // ---- Create voucher ----------------------------------------
        try {
            var voucherId = createLexwareManuellerUmsatz_({
                typ:           typ,
                contactId:     contactId,
                kontaktnummer: kontaktnummer,
                belegRef:      ref,
                voucherDate:   voucherDateStr,
                dueDate:       dueDateStr || undefined,
                belegnummer:   belegnummer || undefined,
                notiz:         notiz || undefined,
                lineItems:     lineItems
            });

            // Write back to the first row of this group only
            if (zuletztGebuchtCol > 0) {
                sheet.getRange(firstRowNum, zuletztGebuchtCol).setValue(formatDate_(now));
            }
            if (lexwareBelegIdCol > 0) {
                sheet.getRange(firstRowNum, lexwareBelegIdCol).setValue(voucherId);
            }

            Logger.log(
                "Manuelle Umsätze: ✅ Beleg_Ref '" + ref +
                "' – Beleg erstellt: " + voucherId
            );
            created++;
        } catch (e) {
            Logger.log(
                "Manuelle Umsätze: ❌ Beleg_Ref '" + ref +
                "' – Fehler beim Erstellen: " + e.message
            );
            errors++;
        }
    }

    Logger.log(
        "Manuelle Umsätze abgeschlossen: erstellt=" + created +
        ", übersprungen=" + skipped +
        ", Fehler=" + errors
    );

    return { ok: errors === 0, created: created, skipped: skipped, errors: errors };
}
