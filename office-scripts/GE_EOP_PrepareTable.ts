/**
 * GE_EOP_PrepareTable  (variant A — data passed in, not read from workbook)
 * -----------------------------------------------------------------------------
 * Builds the flat EOP_DATA lookup table in the EOP_DATA.xlsx workbook, from two
 * JSON arrays the match-sync flow supplies via "List rows present in a table":
 *   headcountJson : rows of Headcount_data in the new EOP source workbook
 *   addressJson   : rows of Address in the new EOP source workbook
 *
 * The source tables live in a DIFFERENT workbook, so this script canNOT read
 * them directly — the flow reads them with List rows and passes the JSON in.
 * This script runs inside EOP_DATA.xlsx and only writes EOP_DATA.
 *
 * Source column names (NEW source, verified from the file 2026-10-21):
 *   Headcount (table Headcount_data): Employee_ID, Employee_Name, Email,
 *                                     Birth_date, Manager_ID, ...
 *   Address   (table Address):        Employee_ID, Line_data, Municiplaity,
 *                                     postel_code, Country, country_name, ...
 *   NOTE: "Municiplaity" and "postel_code" are the source's own spellings
 *   (typos in the file) — matched here verbatim so the JSON keys line up.
 *   NOTE: source "Country" is the CODE (e.g. DE), source "country_name" is the
 *   full NAME (e.g. Deutschland) — the labels are the opposite of what they
 *   look like, so the mapping is spelled out explicitly below.
 *
 * Join: Headcount.Employee_ID = Address.Employee_ID = Personalnummer.
 * LEFT from Headcount. Personalnummer stored NORMALIZED (upper + ws stripped).
 *
 * Germany normalization: source uses 3-letter country codes (Germany = "DEU").
 * When the code is DEU but the name is something else, force CountryName to
 * "Germany" — the client standard, used as-is in the letters. Germany only —
 * nothing is built for other countries yet.
 *
 * BirthDate: Birth_date arrives from "List rows" as an Excel DATE SERIAL (e.g.
 * "32648") — converted here to DD.MM.YYYY (German letter format) so the dispatch
 * phase reads a ready value, not a serial.
 *
 * ManagerEmail: the line manager's e-mail, resolved WITHIN Headcount. Each row's
 * Manager_ID is another employee's Employee_ID, so we build an Employee_ID->Email
 * map over Headcount and look the manager up in it. Empty when the row has no
 * Manager_ID or the manager is not in Headcount. Used by the N3 escalation mail
 * for employees matched via the DUMP branch (pool-matched drivers get their line
 * manager straight from the pool file instead).
 *
 * Target table EOP_DATA (10 cols, this order):
 *   Personalnummer, EmployeeName, EmployeeEmail,
 *   AddressLine, PostalCode, Municipality, CountryName, CountryCode, BirthDate,
 *   ManagerEmail
 */

function main(
    workbook: ExcelScript.Workbook,
    headcountJson: string,
    addressJson: string
): void {
    const TARGET_TABLE = "EOP_DATA";
    const MIN_EXPECTED_ROWS = 7000; // guard vs a truncated source (live Headcount ~9800)

    // ---- helpers --------------------------------------------------------------
    const normKey = (v: string | number | boolean | null | undefined): string =>
        (v === null || v === undefined) ? "" : ("" + v).replace(/\s+/g, "").toUpperCase();

    const txt = (v: string | number | boolean | null | undefined): string =>
        (v === null || v === undefined) ? "" : ("" + v);

    // Source quirk: Line_data encodes address parts with "&&" separators, e.g.
    // "44&&Bahnhofstrasse" or "SAN ANTONIO&&3&&3ºIZQ." (other addresses use
    // plain spaces). Rule: split on "&&", drop empty segments (this clears any
    // leading/trailing/doubled "&&"), rejoin inner parts with a single space,
    // then collapse stray whitespace. Only Line_data has this; other fields pass
    // through txt() unchanged.
    const cleanLine = (v: string | number | boolean | null | undefined): string =>
        txt(v)
            .split("&&")
            .map(s => s.trim())
            .filter(s => s.length > 0)
            .join(" ")
            .replace(/\s+/g, " ")
            .trim();

    // Germany-only name normalization. Source country code is 3-letter; Germany
    // is "DEU". When code is DEU, force the name to the agreed value regardless
    // of what the source name says (the Normal-fines letter is German).
    const DE_CODE = "DEU";
    const DE_NAME = "Germany"; // agreed target name for Germany (client standard — used in letters)
    const normCountryName = (code: string, name: string): string =>
        code.trim().toUpperCase() === DE_CODE ? DE_NAME : name;

    // Birth_date arrives from "List rows" as an Excel DATE SERIAL (e.g. "32648"),
    // not a date string — the source cell is a real date shown as dd/mm/yyyy.
    // Convert serial -> DD.MM.YYYY (German letter format). Excel's epoch is
    // 1899-12-30 (accounts for the fake 1900 leap year); compute in UTC so the
    // day never shifts by a timezone offset. Non-numeric/blank input passes
    // through unchanged (defensive — should not happen for this column).
    const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30); // 1899-12-30
    const MS_PER_DAY = 86400000;
    const formatBirthDate = (v: string | number | null | undefined): string => {
        const s = (v === null || v === undefined) ? "" : ("" + v).trim();
        if (s === "") return "";
        const serial = Number(s);
        if (!isFinite(serial) || serial <= 0) return s; // not a serial — leave as-is
        const d = new Date(EXCEL_EPOCH_UTC + Math.round(serial) * MS_PER_DAY);
        const dd = ("0" + d.getUTCDate()).slice(-2);
        const mm = ("0" + (d.getUTCMonth() + 1)).slice(-2);
        const yyyy = d.getUTCFullYear();
        return dd + "." + mm + "." + yyyy;
    };

    // ---- parse inputs ---------------------------------------------------------
    // Keys MUST match the NEW source column headers exactly (List rows returns
    // the header text as the JSON key).
    interface HeadRow {
        Employee_ID?: string | number | null;
        Employee_Name?: string | null;
        Email?: string | null;
        Birth_date?: string | number | null;
        Manager_ID?: string | number | null; // source Employee_ID of this person's line manager
    }
    interface AddrRow {
        Employee_ID?: string | number | null;
        Line_data?: string | null;
        Municiplaity?: string | null;   // sic — source spelling
        postel_code?: string | number | null; // sic — source spelling
        Country?: string | null;        // source CODE  (e.g. DE) -> CountryCode
        country_name?: string | null;   // source NAME  (e.g. Deutschland) -> CountryName
    }

    let headcount: HeadRow[] = [];
    let address: AddrRow[] = [];
    try {
        headcount = JSON.parse(headcountJson) as HeadRow[];
    } catch (e) {
        throw new Error("headcountJson is not valid JSON: " + e);
    }
    try {
        address = JSON.parse(addressJson) as AddrRow[];
    } catch (e) {
        throw new Error("addressJson is not valid JSON: " + e);
    }

    // Guard: a truncated Headcount must not silently overwrite a good EOP_DATA.
    if (headcount.length < MIN_EXPECTED_ROWS) {
        throw new Error(
            "Headcount has only " + headcount.length + " rows (< " +
            MIN_EXPECTED_ROWS + "). Aborting — source looks truncated."
        );
    }

    // ---- manager lookup: normKey(Employee_ID) -> Email ------------------------
    // The manager is another employee in the same Headcount; Manager_ID points at
    // their Employee_ID. Build the map once, then resolve each row's manager.
    const empEmailMap = new Map<string, string>();
    for (const h of headcount) {
        const k = normKey(h.Employee_ID);
        if (k && !empEmailMap.has(k)) empEmailMap.set(k, txt(h.Email));
    }

    // ---- address lookup: normKey(Employee_ID) -> address fields ---------------
    interface Addr {
        line: string; plz: string; city: string;
        countryName: string; countryCode: string;
    }
    const addrMap = new Map<string, Addr>();
    for (const a of address) {
        const key = normKey(a.Employee_ID);
        if (!key) continue;
        if (!addrMap.has(key)) { // first occurrence wins
            const code = txt(a.Country);          // source Country     = code (DEU…)
            const name = txt(a.country_name);     // source country_name = full name
            addrMap.set(key, {
                line: cleanLine(a.Line_data),     // strip/space the "&&" separators
                plz: txt(a.postel_code),
                city: txt(a.Municiplaity),
                countryName: normCountryName(code, name), // Germany forced to DE_NAME
                countryCode: code
            });
        }
    }

    // ---- LEFT join from Headcount --------------------------------------------
    const out: string[][] = [];
    for (const h of headcount) {
        const pn = normKey(h.Employee_ID);
        if (!pn) continue; // no id -> can't match or mail; skip

        const a = addrMap.get(pn);
        const mgrEmail = empEmailMap.get(normKey(h.Manager_ID)) || ""; // "" if no/unknown manager
        out.push([
            pn,                        // Personalnummer (normalized)
            txt(h.Employee_Name),      // EmployeeName
            txt(h.Email),              // EmployeeEmail
            a ? a.line : "",           // AddressLine
            a ? a.plz : "",            // PostalCode
            a ? a.city : "",           // Municipality
            a ? a.countryName : "",    // CountryName (DEU normalized to DE_NAME)
            a ? a.countryCode : "",    // CountryCode (source Country)
            formatBirthDate(h.Birth_date), // BirthDate: serial -> DD.MM.YYYY
            mgrEmail                   // ManagerEmail (resolved within Headcount)
        ]);
    }

    // ---- rewrite EOP_DATA -----------------------------------------------------
    const target = workbook.getTable(TARGET_TABLE);
    if (!target) {
        throw new Error("Target table '" + TARGET_TABLE + "' not found in the workbook.");
    }

    // Clear existing body before appending — a full rebuild, not an append.
    //
    // Two failure modes to tell apart:
    //  - Real data present (thousands of rows): a single range delete is robust;
    //    if it genuinely fails we THROW, so stale rows are never left under the
    //    fresh ones (the old code doubled EOP_DATA by swallowing this).
    //  - Empty table: Excel keeps one blank "phantom" row, so getRowCount() is 1
    //    with no real data. delete() on that lone row throws a Conflict. That is
    //    harmless — nothing to clear — so we skip the delete when the body holds
    //    only blanks, instead of blindly deleting (which hit this exact case
    //    right after a manual clear).
    const existing = target.getRangeBetweenHeaderAndTotal();
    if (existing && existing.getRowCount() > 0) {
        // Is there any real content, or just the phantom blank row?
        const cells = existing.getTexts();
        let hasData = false;
        for (let r = 0; r < cells.length && !hasData; r++) {
            for (let c = 0; c < cells[r].length; c++) {
                if (("" + cells[r][c]).trim() !== "") { hasData = true; break; }
            }
        }
        if (hasData) {
            try {
                existing.delete(ExcelScript.DeleteShiftDirection.up);
            } catch (e) {
                throw new Error(
                    "Failed to clear EOP_DATA body (" + existing.getRowCount() +
                    " rows): " + e + ". Aborting so stale rows are not left behind."
                );
            }
        }
        // else: only the phantom blank row — leave it; addRows writes into it.
    }

    if (out.length > 0) {
        // Set Text format on the identifier/date columns BEFORE addRows. Doing it
        // after lets Excel coerce BirthDate into a date serial as the rows land,
        // and a later "@" cannot undo an already-baked number (same lesson as the
        // Fines ReferenceNumber: format first, write second). Formatting the
        // table columns up front makes the new rows inherit Text.
        const pnCol = target.getColumnByName("Personalnummer").getRange();
        pnCol.setNumberFormat("@");
        const bdCol = target.getColumnByName("BirthDate").getRange();
        bdCol.setNumberFormat("@");

        target.addRows(-1, out);
    } else {
        console.log("No Headcount rows produced output — EOP_DATA left empty.");
    }

    console.log("EOP_DATA rebuilt: " + out.length + " rows.");
}