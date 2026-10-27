/**
 * sk_Fines_UpdateMatchState
 * -----------------------------------------------------------------------------
 * Phase-1 (match) writer for the GE Fines dispatch flow. Finds the ONE tblFines
 * row whose RowId equals the given key and writes the match result into it:
 * Personalnummer, EmployeeName, EmployeeEmail, EmployeeAddress, DriverSource,
 * VehicleFrom, MailStatus, MatchNote. Does NOT touch the notify/payment columns
 * (LastNotifyStage, LastNotifyDate, DateSent, ConfirmedDate, PaymentDate,
 * ConfirmedBy) nor RowId (the search key).
 *
 * Runs INSIDE the Fines workbook, where tblFines lives. Columns are resolved by
 * HEADER NAME, not fixed position, so the 32-column reorder does not shift writes.
 *
 * Params (positional, as "Run script" binds them):
 *   rowId           : the unique key written by sk_Fines_WriteToExcel.
 *   personalnummer  : matched PN (normalized by sk_Dump_FindByPlate / pool);
 *                     written as TEXT so the key never turns into a number.
 *   employeeName    : from sk_Eop_FindByPn (DUMP branch) or the pool (pool branch).
 *   employeeEmail   : from sk_Eop_FindByPn or the pool.
 *   employeeAddress : pre-joined, empty-safe address (EOP branch) or "" (pool).
 *   driverSource    : "DUMP" | "Pool" | "" — which source resolved the driver;
 *                     the N3 step uses it to decide where the line manager comes
 *                     from (Headcount self-lookup vs the pool column).
 *   vehicleFrom     : the handover/pickup date as an Excel serial (string ok),
 *                     "" when unknown; written to the Date column VehicleFrom.
 *   mailStatus      : "Matched" | "MatchFailed" (phase-1 verdict from the flow).
 *   matchNote       : human-readable reason (e.g. "plate not in DUMP",
 *                     "offence before handover", "not in pool window") — "" when
 *                     fully matched.
 *
 * All fields are written exactly as passed, including empty strings, so a
 * MatchFailed row shows how far the match got. The Matched/MatchFailed decision
 * belongs to the flow; this script only records what it is given.
 *
 * Throws if rowId is empty or no row carries it — the flow catches this and logs
 * it, because a match result with nowhere to land must not pass silently.
 *
 * Returns: { updated:true, rowIndex } on success.
 */

function main(
    workbook: ExcelScript.Workbook,
    rowId: string,
    personalnummer: string,
    employeeName: string,
    employeeEmail: string,
    employeeAddress: string,
    driverSource: string,
    vehicleFrom: string,
    mailStatus: string,
    matchNote: string
): { updated: boolean; rowIndex: number } {

    const TABLE_NAME = "tblFines";

    const key = (rowId === null || rowId === undefined) ? "" : ("" + rowId).trim();
    if (!key) {
        throw new Error("sk_Fines_UpdateMatchState: empty rowId — nothing to match against.");
    }

    const table = workbook.getTable(TABLE_NAME);
    if (!table) {
        throw new Error("Table '" + TABLE_NAME + "' not found in the Fines workbook.");
    }

    // Resolve the columns we read/write by header name.
    const headerVals = table.getHeaderRowRange().getValues()[0];
    const idx: { [k: string]: number } = {};
    for (let c = 0; c < headerVals.length; c++) {
        idx["" + headerVals[c]] = c;
    }
    const need = ["RowId", "Personalnummer", "EmployeeName", "EmployeeEmail",
        "EmployeeAddress", "DriverSource", "VehicleFrom", "MailStatus", "MatchNote"];
    for (const n of need) {
        if (idx[n] === undefined) {
            throw new Error("Column '" + n + "' not found in tblFines.");
        }
    }

    const body = table.getRangeBetweenHeaderAndTotal();
    if (!body || body.getRowCount() === 0) {
        throw new Error("sk_Fines_UpdateMatchState: tblFines is empty — RowId '" + key + "' not found.");
    }

    // Find the row by exact RowId (text compare, same string that was written).
    const texts = body.getTexts();
    let target = -1;
    for (let r = 0; r < texts.length; r++) {
        if (("" + texts[r][idx["RowId"]]).trim() === key) {
            target = r;
            break;
        }
    }
    if (target < 0) {
        throw new Error("sk_Fines_UpdateMatchState: RowId '" + key + "' not found in tblFines.");
    }

    const str = (v: string): string => (v === null || v === undefined) ? "" : "" + v;

    // Personalnummer is an identifier — force Text before writing so a numeric PN
    // keeps leading zeros and never compares wrong later.
    const pnCell = body.getCell(target, idx["Personalnummer"]);
    pnCell.setNumberFormat("@");
    pnCell.setValue(str(personalnummer));

    body.getCell(target, idx["EmployeeName"]).setValue(str(employeeName));
    body.getCell(target, idx["EmployeeEmail"]).setValue(str(employeeEmail));
    body.getCell(target, idx["EmployeeAddress"]).setValue(str(employeeAddress));
    body.getCell(target, idx["DriverSource"]).setValue(str(driverSource));

    // VehicleFrom is a Date column: write the Excel serial as a number and set the
    // date format; leave it blank when no handover/pickup date was resolved.
    const vf = str(vehicleFrom).trim();
    const vfNum = Number(vf);
    const vfCell = body.getCell(target, idx["VehicleFrom"]);
    if (vf !== "" && isFinite(vfNum)) {
        vfCell.setNumberFormat("dd.mm.yyyy");
        vfCell.setValue(vfNum);
    } else {
        vfCell.setValue("");
    }

    body.getCell(target, idx["MailStatus"]).setValue(str(mailStatus));
    body.getCell(target, idx["MatchNote"]).setValue(str(matchNote));

    return { updated: true, rowIndex: target };
}