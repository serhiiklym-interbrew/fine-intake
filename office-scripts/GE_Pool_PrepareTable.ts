/**
 * GE_Pool_PrepareTable  (variant A — pool rows passed in as JSON)
 * -----------------------------------------------------------------------------
 * Builds the clean POOL_DATA lookup table in POOL_DATA.xlsx from the manual pool
 * tracker, so the match flow reads OUR copy instead of touching the pool file's
 * foreign site on every run.
 *
 * The pool tracker is a MANUAL file in a different site:
 *   .../RewardsEURteam559/.../Fleet GE PRP/Pool Car Tracker - CURRENT.xlsx
 *   sheet "Car Requests Pipeline_NEW", Excel table "Table1"
 * The POOL-sync flow reads Table1 with "List rows present in a table" and passes
 * the row array (string(value)) in as poolJson. This script runs INSIDE
 * POOL_DATA.xlsx and only writes POOL_DATA.
 *
 * Source headers are long bilingual labels; matched NORMALIZED (trim + collapse
 * internal whitespace) so the file's trailing/double spaces don't break mapping.
 *
 * Transforms:
 *   - PlateNumber : kept raw (as in the file) — the lookup normalizes on compare.
 *   - PickupDate / ReturnDate : pool shows US M/D/YYYY; converted to an Excel
 *     serial here so the lookup reads ready dates. Empty Return stays "" (open
 *     window = car still held).
 *   - PersonalId : normalized (upper + ws stripped); may be malformed in source —
 *     non-critical, kept for audit only (the plate+window already identify the row).
 *   - Rows with no License Plate are skipped (nothing to match on).
 *
 * Guard: an empty source (0 rows) does NOT wipe a good POOL_DATA — it logs and
 * leaves the existing table untouched (the pool can legitimately be small, so
 * there is no high MIN_EXPECTED_ROWS threshold like DUMP/EOP).
 *
 * Target table POOL_DATA (7 cols, this order):
 *   PlateNumber, DriverName, DriverEmail, PersonalId, LineManagerEmail,
 *   PickupDate, ReturnDate
 */

function main(workbook: ExcelScript.Workbook, poolJson: string): void {
    const TARGET_TABLE = "POOL_DATA";

    // Source header labels (matched normalized below).
    const LABEL = {
        plate: "License Plate",
        name: "Name of the Driver - Name des Fahrers",
        email: "E-mail address of Driver - E-Mail Adresse des Fahrers",
        pid: "Personal ID number of the driver - Personal ID nummer des Fahrers",
        mgr: "E-mail address of Line Manager - E-Mail Adresse des Vorgesetzten",
        pickup: "Pick up date - Abholungsdatum (Datum, ab dem Sie das Auto benötigen)",
        ret: "Return date - Abgabedatum (wann werden Sie das Auto retournieren)"
    };

    // ---- helpers --------------------------------------------------------------
    const normLabel = (s: string | null | undefined): string =>
        ("" + (s === null || s === undefined ? "" : s)).trim().replace(/\s+/g, " ");
    const txt = (v: string | number | boolean | null | undefined): string =>
        (v === null || v === undefined) ? "" : ("" + v);
    const normPn = (v: string | number | boolean | null | undefined): string =>
        (v === null || v === undefined) ? "" : ("" + v).replace(/\s+/g, "").toUpperCase();

    // Date -> Excel serial. Accept serial | US M/D/YYYY | ISO YYYY-MM-DD |
    // German D.M.YYYY. Return "" when empty/unparseable (keeps the cell blank).
    const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);
    const MS_PER_DAY = 86400000;
    const dmyToSerial = (d: number, mo: number, y: number): number =>
        (Date.UTC(y, mo - 1, d) - EXCEL_EPOCH_UTC) / MS_PER_DAY;
    const toSerial = (v: string | number | null | undefined): number | string => {
        const s = (v === null || v === undefined) ? "" : ("" + v).trim();
        if (!s) return "";
        if (/^\d+(\.\d+)?$/.test(s)) return Number(s);              // serial
        let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);                // ISO
        if (m) return dmyToSerial(Number(m[3]), Number(m[2]), Number(m[1]));
        m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);              // US M/D/YYYY
        if (m) return dmyToSerial(Number(m[2]), Number(m[1]), Number(m[3]));
        m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/);              // German D.M.YYYY
        if (m) return dmyToSerial(Number(m[1]), Number(m[2]), Number(m[3]));
        return "";
    };

    // ---- parse input ----------------------------------------------------------
    let rows: { [k: string]: string | number | boolean | null }[] = [];
    try {
        rows = JSON.parse(poolJson) as { [k: string]: string | number | boolean | null }[];
    } catch (e) {
        throw new Error("poolJson is not valid JSON: " + e);
    }

    const want = {
        plate: normLabel(LABEL.plate), name: normLabel(LABEL.name),
        email: normLabel(LABEL.email), pid: normLabel(LABEL.pid),
        mgr: normLabel(LABEL.mgr), pickup: normLabel(LABEL.pickup), ret: normLabel(LABEL.ret)
    };
    const getField = (
        row: { [k: string]: string | number | boolean | null }, wantNorm: string
    ): string | number | boolean | null => {
        for (const k of Object.keys(row)) {
            if (normLabel(k) === wantNorm) return row[k];
        }
        return "";
    };

    // ---- build output rows (skip rows with no plate) --------------------------
    const out: (string | number)[][] = [];
    for (const row of rows) {
        const plate = txt(getField(row, want.plate)).trim();
        if (!plate) continue; // nothing to match on
        out.push([
            plate,                                 // PlateNumber (raw)
            txt(getField(row, want.name)).trim(),  // DriverName
            txt(getField(row, want.email)).trim(), // DriverEmail
            normPn(getField(row, want.pid)),       // PersonalId (normalized)
            txt(getField(row, want.mgr)).trim(),   // LineManagerEmail
            toSerial(getField(row, want.pickup)),  // PickupDate (serial or "")
            toSerial(getField(row, want.ret))      // ReturnDate (serial or "")
        ]);
    }

    // Guard: do NOT wipe a good POOL_DATA with an empty/failed read.
    if (out.length === 0) {
        console.log("No pool rows with a plate — POOL_DATA left untouched.");
        return;
    }

    // ---- rewrite POOL_DATA ----------------------------------------------------
    const target = workbook.getTable(TARGET_TABLE);
    if (!target) {
        throw new Error("Target table '" + TARGET_TABLE + "' not found in the workbook.");
    }

    // Clear existing body (real data -> delete + throw-on-fail; phantom blank ->
    // skip, addRows writes into it). Same robust pattern as EOP.
    const existing = target.getRangeBetweenHeaderAndTotal();
    if (existing && existing.getRowCount() > 0) {
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
                    "Failed to clear POOL_DATA body (" + existing.getRowCount() +
                    " rows): " + e + ". Aborting so stale rows are not left behind."
                );
            }
        }
    }

    // Format identifier/date columns BEFORE addRows (same lesson as EOP BirthDate:
    // format first, write second — otherwise Excel re-coerces the serials/text).
    target.getColumnByName("PersonalId").getRange().setNumberFormat("@");
    target.getColumnByName("PickupDate").getRange().setNumberFormat("dd.mm.yyyy");
    target.getColumnByName("ReturnDate").getRange().setNumberFormat("dd.mm.yyyy");

    target.addRows(-1, out);

    console.log("POOL_DATA rebuilt: " + out.length + " rows.");
}