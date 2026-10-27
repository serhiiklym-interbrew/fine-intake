/**
 * sk_Dump_FindByPlate
 * -----------------------------------------------------------------------------
 * Looks up a vehicle plate in the DUMP_DATA table (the raw SAP dump, ~60 cols,
 * built by GE_Dump_PrepareTable which wraps the whole A1:BF range as-is) and
 * returns the driver's Personalnummer — BUT only if the vehicle was already
 * assigned to that driver on the offence date.
 *
 * Runs INSIDE the dump workbook (DUMP_INPUT_TEST.xlsx / prod DUMP_INPUT.xlsx),
 * where DUMP_DATA lives — so it reads the table directly, no data passed in.
 *
 * DATE CHECK (the reason offenceDate is a parameter):
 *   A row counts only if its FAHRZEUGUEBERNAHME (the date the vehicle was handed
 *   over to that driver) is ON or BEFORE the offence date:
 *       FAHRZEUGUEBERNAHME <= OffenceDate
 *   i.e. the driver is responsible only for fines dated after they got the car.
 *   The dump holds one row per vehicle, but if several ever match the plate, the
 *   row with the GREATEST FAHRZEUGUEBERNAHME that is still <= OffenceDate wins
 *   (the most recent handover before the offence). A plate that exists in the
 *   dump but whose handover is AFTER the offence -> not this driver -> the caller
 *   falls back to the pool file.
 *
 * Quality filter: a row counts only if Personalnummer is non-empty. That alone
 * drops the dump's junk (TANKKARTE, FEUERWEHR, STAPLER numbers, OHNE) and the
 * duplicate plate rows, which carry no Personalnummer. FAHRZEUGSTATUS is IGNORED
 * on purpose — its values are a zoo (AKTIV / INACTIVE / INAKTIV / RÜCKGABE
 * LEASING / VERKAUFT), and filtering by it would drop fines on recently-retired
 * vehicles whose driver still needs the letter (confirmed with the client).
 *
 * Plate matching is normalization-tolerant: both the incoming plate and the
 * dump's KENNZEICHEN are upper-cased with spaces and hyphens stripped.
 *
 * Dates (FAHRZEUGUEBERNAHME and the offenceDate parameter) are read with getTexts
 * and may arrive as an Excel serial ("45900"), US "M/D/YYYY" (how DUMP_DATA
 * actually stores FAHRZEUGUEBERNAHME — confirmed 2026-10-27: month first, day
 * second across the whole column), German "DD.MM.YYYY", or ISO "YYYY-MM-DD" —
 * toSerial() accepts all four and compares in serial space.
 */

interface DumpLookupResult {
    found: boolean;
    personalnummer: string;
    vehicleFrom: number | string;
    vehicleFromText: string;
    plateInDump: boolean;   // plate seen in the dump at all (with OR without a PN)
    plateNoPn: boolean;     // plate seen, but every matching row had an empty PN
    dateFiltered: boolean;
    multiple: boolean;
    matchCount: number;
}

function main(
    workbook: ExcelScript.Workbook,
    plateRaw: string,
    offenceDate: string
): DumpLookupResult {
    const TABLE_NAME = "DUMP_DATA";
    const PLATE_COL = "KENNZEICHEN";
    const PN_COL = "Personalnummer";
    const FROM_COL = "FAHRZEUGUEBERNAHME";

    // Plate normalization for COMPARISON only (upper, strip spaces AND hyphens).
    const normPlate = (v: string | number | boolean | null | undefined): string =>
        (v === null || v === undefined) ? "" : ("" + v).toUpperCase().replace(/[\s-]/g, "");

    // Personalnummer normalization (upper + strip whitespace) — MUST match the
    // rule used when EOP_DATA was built, so the downstream PN->EOP lookup works.
    const normPn = (v: string | number | boolean | null | undefined): string =>
        (v === null || v === undefined) ? "" : ("" + v).replace(/\s+/g, "").toUpperCase();

    // ---- date helpers (serial space; Excel epoch 1899-12-30, UTC) --------------
    const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);
    const MS_PER_DAY = 86400000;
    const dmyToSerial = (d: number, mo: number, y: number): number =>
        (Date.UTC(y, mo - 1, d) - EXCEL_EPOCH_UTC) / MS_PER_DAY;

    // Accept a serial, US M/D/YYYY, German DD.MM.YYYY, or ISO YYYY-MM-DD -> Excel
    // serial; NaN if unparseable. The German and US branches differ by separator
    // ('.' vs '/') so they never collide. DUMP_DATA.FAHRZEUGUEBERNAHME arrives as
    // US M/D/YYYY (month first); the offenceDate parameter arrives as a serial
    // (Excel Date column from List rows) — both covered here.
    const toSerial = (v: string | number | null | undefined): number => {
        const s = (v === null || v === undefined) ? "" : ("" + v).trim();
        if (!s) return NaN;
        if (/^\d+(\.\d+)?$/.test(s)) return Number(s);          // already a serial
        let m = s.match(/^(\d{2})\.(\d{2})\.(\d{4})/);          // DD.MM.YYYY (German)
        if (m) return dmyToSerial(Number(m[1]), Number(m[2]), Number(m[3]));
        m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);          // M/D/YYYY (US — DUMP_DATA)
        if (m) return dmyToSerial(Number(m[2]), Number(m[1]), Number(m[3]));
        m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);                // YYYY-MM-DD (ISO)
        if (m) return dmyToSerial(Number(m[3]), Number(m[2]), Number(m[1]));
        return NaN;
    };

    const serialToText = (serial: number): string => {
        if (!isFinite(serial)) return "";
        const d = new Date(EXCEL_EPOCH_UTC + Math.round(serial) * MS_PER_DAY);
        const dd = ("0" + d.getUTCDate()).slice(-2);
        const mm = ("0" + (d.getUTCMonth() + 1)).slice(-2);
        return dd + "." + mm + "." + d.getUTCFullYear();
    };

    const empty: DumpLookupResult = {
        found: false, personalnummer: "", vehicleFrom: "", vehicleFromText: "",
        plateInDump: false, plateNoPn: false, dateFiltered: false,
        multiple: false, matchCount: 0
    };

    const targetPlate = normPlate(plateRaw);
    if (!targetPlate) return empty;

    const table = workbook.getTable(TABLE_NAME);
    if (!table) {
        throw new Error("Table '" + TABLE_NAME + "' not found in the dump workbook.");
    }

    // Locate the three columns by header name (order-independent across dumps).
    const headerVals = table.getHeaderRowRange().getValues()[0];
    let plateIdx = -1, pnIdx = -1, fromIdx = -1;
    for (let c = 0; c < headerVals.length; c++) {
        const h = "" + headerVals[c];
        if (h === PLATE_COL) plateIdx = c;
        if (h === PN_COL) pnIdx = c;
        if (h === FROM_COL) fromIdx = c;
    }
    if (plateIdx < 0) throw new Error("Column '" + PLATE_COL + "' not found in DUMP_DATA.");
    if (pnIdx < 0) throw new Error("Column '" + PN_COL + "' not found in DUMP_DATA.");
    if (fromIdx < 0) throw new Error("Column '" + FROM_COL + "' not found in DUMP_DATA.");

    const body = table.getRangeBetweenHeaderAndTotal();
    if (!body || body.getRowCount() === 0) return empty;

    const vals = body.getTexts(); // keep PN leading zeros / plate & date as shown

    // Collect rows that carry this plate. plateInDump is set for ANY matching row
    // (even one without a Personalnummer), so the diagnostics tell the truth; only
    // PN-bearing rows become candidates for the date check. A plate that exists in
    // the dump but only in PN-less rows is a pool car by the client's rule
    // (plate present, no PN -> go to the pool), reported via plateNoPn.
    interface Cand { pn: string; fromSerial: number; }
    const cands: Cand[] = [];
    let plateInDump = false;
    let plateWithPn = false;
    for (let r = 0; r < vals.length; r++) {
        if (normPlate(vals[r][plateIdx]) !== targetPlate) continue;
        plateInDump = true; // plate is here, regardless of PN
        const pnRaw = ("" + vals[r][pnIdx]).trim();
        if (!pnRaw) continue; // no Personalnummer -> not a usable driver row (pool car / junk)
        plateWithPn = true;
        cands.push({ pn: normPn(pnRaw), fromSerial: toSerial(vals[r][fromIdx]) });
    }

    if (!plateInDump) return empty; // plate not in the dump at all -> caller goes to pool

    // Plate is in the dump, but no row carried a Personalnummer -> pool car.
    if (!plateWithPn) {
        return {
            found: false, personalnummer: "", vehicleFrom: "", vehicleFromText: "",
            plateInDump: true, plateNoPn: true, dateFiltered: false,
            multiple: false, matchCount: 0
        };
    }

    // Can't date-check without a usable offence date -> no safe match here (pool).
    const offenceSerial = toSerial(offenceDate);
    if (!isFinite(offenceSerial)) {
        return {
            found: false, personalnummer: "", vehicleFrom: "", vehicleFromText: "",
            plateInDump: true, plateNoPn: false, dateFiltered: false,
            multiple: false, matchCount: 0
        };
    }

    // Keep rows whose handover is on/before the offence (and whose date parsed).
    const passed = cands.filter(c => isFinite(c.fromSerial) && c.fromSerial <= offenceSerial);
    if (passed.length === 0) {
        // plate exists, but it was not this person's car yet on the offence date
        return {
            found: false, personalnummer: "", vehicleFrom: "", vehicleFromText: "",
            plateInDump: true, plateNoPn: false, dateFiltered: true,
            multiple: false, matchCount: 0
        };
    }

    // Pick the most recent handover that is still <= offence date.
    passed.sort((a, b) => b.fromSerial - a.fromSerial);
    const chosen = passed[0];
    return {
        found: true,
        personalnummer: chosen.pn,
        vehicleFrom: chosen.fromSerial,
        vehicleFromText: serialToText(chosen.fromSerial),
        plateInDump: true,
        plateNoPn: false,
        dateFiltered: false,
        multiple: passed.length > 1,
        matchCount: passed.length
    };
}