/**
 * sk_Pool_FindByPlateAndDate  (reads POOL_DATA directly)
 * -----------------------------------------------------------------------------
 * Fallback driver lookup for fines whose plate was NOT resolved in the DUMP
 * (plate absent, or FAHRZEUGUEBERNAHME after the offence). Finds the pool-car
 * driver who HAD the vehicle on the offence date.
 *
 * Runs INSIDE POOL_DATA.xlsx, where the clean POOL_DATA table lives (built by the
 * POOL-sync via GE_Pool_PrepareTable) — so it reads the table directly, no data
 * passed in, same pattern as sk_Dump_FindByPlate over DUMP_DATA.
 *
 * POOL_DATA columns (resolved by header name): PlateNumber, DriverName,
 * DriverEmail, PersonalId, LineManagerEmail, PickupDate, ReturnDate. Pickup/Return
 * are already Excel serials (the sync converted them), read via getValues().
 *
 * WINDOW CHECK: a row matches when the plate matches AND the offence date falls
 * inside the assignment window:
 *       Pickup <= OffenceDate <= Return
 * An empty Return means the car is still held -> open-ended window, so
 * Pickup <= OffenceDate alone suffices. Windows don't overlap (client-confirmed);
 * if several ever match, the row with the NEWEST Pickup wins.
 *
 * offenceDate (parameter) is parsed by toSerial (serial / ISO / US M/D/YYYY /
 * German D.M.YYYY); pass it as a serial or ISO so month/day are never ambiguous.
 *
 * Returns:
 *   found            : plate + window check passed (usable driver)
 *   driverName       : for the "Dear <name>" letter
 *   driverEmail      : recipient
 *   personalId       : from POOL_DATA (already normalized), stored to tblFines
 *   lineManagerEmail : pool's own Vorgesetzter column — used for N3 escalation
 *   vehicleFrom      : chosen Pickup as an Excel serial ("" if none) -> VehicleFrom
 *   vehicleFromText  : chosen Pickup as DD.MM.YYYY, for logs ("" if none)
 *   plateInPool      : at least one row carried this plate (diagnostics)
 *   windowFiltered   : plate present, but offence fell outside every window
 *   multiple         : more than one row passed the window check
 *   matchCount       : rows that passed the window check
 */

interface PoolLookupResult {
    found: boolean;
    driverName: string;
    driverEmail: string;
    personalId: string;
    lineManagerEmail: string;
    vehicleFrom: number | string;
    vehicleFromText: string;
    plateInPool: boolean;
    windowFiltered: boolean;
    multiple: boolean;
    matchCount: number;
}

function main(
    workbook: ExcelScript.Workbook,
    plateRaw: string,
    offenceDate: string
): PoolLookupResult {
    const TABLE_NAME = "POOL_DATA";

    // ---- helpers --------------------------------------------------------------
    const normPlate = (v: string | number | boolean | null | undefined): string =>
        (v === null || v === undefined) ? "" : ("" + v).toUpperCase().replace(/[\s-]/g, "");
    const txt = (v: string | number | boolean | null | undefined): string =>
        (v === null || v === undefined) ? "" : ("" + v);

    // ---- date helpers (serial space; Excel epoch 1899-12-30, UTC) --------------
    const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);
    const MS_PER_DAY = 86400000;
    const dmyToSerial = (d: number, mo: number, y: number): number =>
        (Date.UTC(y, mo - 1, d) - EXCEL_EPOCH_UTC) / MS_PER_DAY;

    // Parse a text date (serial / ISO / US M/D/YYYY / German D.M.YYYY) -> serial.
    const toSerial = (v: string | number | null | undefined): number => {
        const s = (v === null || v === undefined) ? "" : ("" + v).trim();
        if (!s) return NaN;
        if (/^\d+(\.\d+)?$/.test(s)) return Number(s);
        let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (m) return dmyToSerial(Number(m[3]), Number(m[2]), Number(m[1]));
        m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (m) return dmyToSerial(Number(m[2]), Number(m[1]), Number(m[3]));
        m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/);
        if (m) return dmyToSerial(Number(m[1]), Number(m[2]), Number(m[3]));
        return NaN;
    };

    // POOL_DATA dates are already serials (number). getValues() returns a number
    // for a date cell, "" for an empty cell. Turn either into a serial or NaN —
    // NEVER let an empty cell become 0 (Number("") === 0), which would corrupt the
    // window check (an empty Return must read as "open", i.e. NaN, not 0).
    const cellToSerial = (v: string | number | boolean | null | undefined): number => {
        if (v === null || v === undefined || v === "") return NaN;
        if (typeof v === "number") return v;
        return toSerial("" + v);
    };

    const serialToText = (serial: number): string => {
        if (!isFinite(serial)) return "";
        const d = new Date(EXCEL_EPOCH_UTC + Math.round(serial) * MS_PER_DAY);
        const dd = ("0" + d.getUTCDate()).slice(-2);
        const mm = ("0" + (d.getUTCMonth() + 1)).slice(-2);
        return dd + "." + mm + "." + d.getUTCFullYear();
    };

    const empty: PoolLookupResult = {
        found: false, driverName: "", driverEmail: "", personalId: "",
        lineManagerEmail: "", vehicleFrom: "", vehicleFromText: "",
        plateInPool: false, windowFiltered: false, multiple: false, matchCount: 0
    };

    const targetPlate = normPlate(plateRaw);
    if (!targetPlate) return empty;

    const table = workbook.getTable(TABLE_NAME);
    if (!table) {
        throw new Error("Table '" + TABLE_NAME + "' not found in the POOL workbook.");
    }

    // Column indexes by header name.
    const headerVals = table.getHeaderRowRange().getValues()[0];
    const idx: { [k: string]: number } = {};
    for (let c = 0; c < headerVals.length; c++) {
        idx["" + headerVals[c]] = c;
    }
    const need = ["PlateNumber", "DriverName", "DriverEmail", "PersonalId",
        "LineManagerEmail", "PickupDate", "ReturnDate"];
    for (const n of need) {
        if (idx[n] === undefined) {
            throw new Error("Column '" + n + "' not found in POOL_DATA.");
        }
    }

    const body = table.getRangeBetweenHeaderAndTotal();
    if (!body || body.getRowCount() === 0) return empty;

    const vals = body.getValues(); // dates as serials (number), text as strings

    // Collect rows carrying this plate.
    interface Cand {
        name: string; email: string; pid: string; mgr: string;
        pickupS: number; retS: number;
    }
    const cands: Cand[] = [];
    let plateInPool = false;
    for (let r = 0; r < vals.length; r++) {
        if (normPlate(vals[r][idx["PlateNumber"]]) !== targetPlate) continue;
        plateInPool = true;
        cands.push({
            name: txt(vals[r][idx["DriverName"]]).trim(),
            email: txt(vals[r][idx["DriverEmail"]]).trim(),
            pid: txt(vals[r][idx["PersonalId"]]).trim(),
            mgr: txt(vals[r][idx["LineManagerEmail"]]).trim(),
            pickupS: cellToSerial(vals[r][idx["PickupDate"]]),
            retS: cellToSerial(vals[r][idx["ReturnDate"]]) // NaN if empty -> open window
        });
    }

    if (!plateInPool) return empty; // plate not in pool -> caller reports MatchFailed

    const offenceSerial = toSerial(offenceDate);
    if (!isFinite(offenceSerial)) {
        return {
            found: false, driverName: "", driverEmail: "", personalId: "",
            lineManagerEmail: "", vehicleFrom: "", vehicleFromText: "",
            plateInPool: true, windowFiltered: false, multiple: false, matchCount: 0
        };
    }

    // Window: Pickup <= offence <= Return (open Return = still held).
    const passed = cands.filter(c =>
        isFinite(c.pickupS) && c.pickupS <= offenceSerial &&
        (!isFinite(c.retS) || offenceSerial <= c.retS)
    );
    if (passed.length === 0) {
        return {
            found: false, driverName: "", driverEmail: "", personalId: "",
            lineManagerEmail: "", vehicleFrom: "", vehicleFromText: "",
            plateInPool: true, windowFiltered: true, multiple: false, matchCount: 0
        };
    }

    // Newest Pickup wins (windows shouldn't overlap; safety for edge data).
    passed.sort((a, b) => b.pickupS - a.pickupS);
    const ch = passed[0];
    return {
        found: true,
        driverName: ch.name,
        driverEmail: ch.email,
        personalId: ch.pid,
        lineManagerEmail: ch.mgr,
        vehicleFrom: ch.pickupS,
        vehicleFromText: serialToText(ch.pickupS),
        plateInPool: true,
        windowFiltered: false,
        multiple: passed.length > 1,
        matchCount: passed.length
    };
}