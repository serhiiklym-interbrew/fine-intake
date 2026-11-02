/**
 * sk_Fines_GetForConfirm
 * -----------------------------------------------------------------------------
 * Read-only companion to sk_Fines_ConfirmPayment. The capture flow
 * ("GE_FineIntake_ConfirmPayment_Capture", HTTP GET) calls this FIRST, before
 * any date is submitted, to render the confirmation page: it locates the row by
 * RowId, checks the token + state exactly as ConfirmPayment will, and hands back
 * the fine's details already converted to display text so the flow's HTML can
 * just drop them in. It writes NOTHING — the actual transition happens only when
 * the driver submits the date and the flow calls sk_Fines_ConfirmPayment.
 *
 * Why a separate read: the GET page needs to show which fine is being confirmed
 * (client wants details on the page, "Варіант Б"), and it must refuse a
 * forged/stale/already-done link the same way the write path does — so the page
 * never offers a date-picker for a link that could not confirm anyway.
 *
 * MODEL (confirmed 2026-10-27): a row is confirmable only once it has been
 * notified -> MailStatus "Notified". Confirmed = already done (friendly page).
 * Anything else (Pending/Matched/MatchFailed/SendFailed) = not_actionable.
 *
 * Parameters:
 *   rowId        : the tblFines RowId from the confirm-link URL (col RowId).
 *   confirmToken : the one-time nonce from the URL, compared to the row's
 *                  ConfirmToken (col ConfirmToken).
 *
 * Returns (read by the flow to choose + fill the page):
 *   ok      : true  -> show the date-picker page (outcome "ready") or the
 *                      friendly "already recorded" page (already_confirmed).
 *             false -> show an error/notice page; details are blank.
 *   outcome : "ready"             -> Notified + token ok: render the date-picker.
 *             "already_confirmed" -> row already Confirmed: friendly no-op page
 *                                    (details filled so the page can show them).
 *             "row_not_found"     -> no row with that RowId.
 *             "invalid_token"     -> token missing/!= the row's ConfirmToken
 *                                    (details left blank — don't leak a fine to a
 *                                    forged link).
 *             "not_actionable"    -> row in a state we don't confirm from;
 *                                    previousStatus tells the flow which.
 *   previousStatus : the row's MailStatus (for tailoring the page text).
 *   plateNumber, fineType, vendorName, referenceNumber : row text fields.
 *   offenceDateText, dueDateText : serials -> "DD.MM.YYYY" ("" if empty).
 *   amountText : amount -> German "1.234,56 €" ("" if empty) — same display the
 *                MatchFailed mail uses, so the page matches the letters.
 *   reason     : short free text, mirrors outcome for logging.
 *
 * Columns are resolved by NAME (getColumnByName), so column reordering in
 * tblFines is safe here (unlike the positional addRow in sk_Fines_WriteToExcel).
 */

interface GetForConfirmResult {
    ok: boolean;
    outcome: string;
    previousStatus: string;
    plateNumber: string;
    fineType: string;
    vendorName: string;
    referenceNumber: string;
    offenceDateText: string;
    dueDateText: string;
    amountText: string;
    reason: string;
}

function main(
    workbook: ExcelScript.Workbook,
    rowId: string,
    confirmToken: string
): GetForConfirmResult {
    const table = workbook.getTable("tblFines");
    if (!table) {
        // A missing table is a real failure, not an expected user outcome.
        throw new Error("Table 'tblFines' not found in the workbook.");
    }

    // The only state the confirm page offers a date-picker FROM: a notified row.
    // Mirrors SOURCE_STATES in sk_Fines_ConfirmPayment so the page never promises
    // what the write path would reject.
    const SOURCE_STATES = ["Notified"];

    // ---- display helpers (serial space; Excel epoch 1899-12-30, UTC) -----------
    const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);
    const MS_PER_DAY = 86400000;

    // Excel serial -> "dd.MM.yyyy"; "" for empty/non-numeric (e.g. a Check row
    // with no OffenceDate/DueDate). Dates were stored as midnight UTC, read UTC.
    const serialToDMY = (v: string | number | boolean): string => {
        if (v === "" || v === null || v === undefined) return "";
        const n = typeof v === "number" ? v : Number(v);
        if (!isFinite(n) || n <= 0) return "";
        const d = new Date(EXCEL_EPOCH_UTC + Math.round(n) * MS_PER_DAY);
        const dd = ("0" + d.getUTCDate()).slice(-2);
        const mm = ("0" + (d.getUTCMonth() + 1)).slice(-2);
        return dd + "." + mm + "." + d.getUTCFullYear();
    };

    // Amount serial/number -> German "1.234,56 €"; "" for empty. Same format the
    // MatchFailed mail renders (formatNumber N2 de-DE + " €"), so page == letters.
    const amountToText = (v: string | number | boolean): string => {
        if (v === "" || v === null || v === undefined) return "";
        const n = typeof v === "number" ? v : Number(v);
        if (!isFinite(n)) return "";
        const fixed = n.toFixed(2);                 // "1234.56"
        const dot = fixed.indexOf(".");
        const intPart = fixed.slice(0, dot);
        const fracPart = fixed.slice(dot + 1);
        const grouped = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, "."); // thousands '.'
        return grouped + "," + fracPart + " \u20AC";
    };

    const str = (v: string | number | boolean): string =>
        (v === null || v === undefined) ? "" : ("" + v);

    // ---- result factory -------------------------------------------------------
    const emptyDetails = {
        plateNumber: "", fineType: "", vendorName: "", referenceNumber: "",
        offenceDateText: "", dueDateText: "", amountText: ""
    };
    const res = (
        ok: boolean, outcome: string, previousStatus: string,
        details: {
            plateNumber: string; fineType: string; vendorName: string;
            referenceNumber: string; offenceDateText: string;
            dueDateText: string; amountText: string;
        }
    ): GetForConfirmResult => ({
        ok, outcome, previousStatus,
        plateNumber: details.plateNumber,
        fineType: details.fineType,
        vendorName: details.vendorName,
        referenceNumber: details.referenceNumber,
        offenceDateText: details.offenceDateText,
        dueDateText: details.dueDateText,
        amountText: details.amountText,
        reason: outcome
    });

    const wantRowId = ("" + (rowId || "")).trim();
    const wantToken = ("" + (confirmToken || "")).trim();

    // ---- locate the row by RowId (columns resolved by NAME, not position) ------
    const colIdx = (name: string): number => {
        const c = table.getColumnByName(name);
        if (!c) throw new Error("Column '" + name + "' not found in tblFines.");
        return c.getIndex(); // 0-based index within the table's columns
    };
    const iRowId = colIdx("RowId");
    const iToken = colIdx("ConfirmToken");
    const iStatus = colIdx("MailStatus");
    const iPlate = colIdx("PlateNumber");
    const iFineType = colIdx("FineType");
    const iVendor = colIdx("VendorName");
    const iRef = colIdx("ReferenceNumber");
    const iOffence = colIdx("OffenceDate");
    const iDue = colIdx("DueDate");
    const iAmount = colIdx("Amount");

    const body = table.getRangeBetweenHeaderAndTotal();
    if (!body || body.getRowCount() === 0) {
        return res(false, "row_not_found", "", emptyDetails);
    }

    const values = body.getValues();
    let r = -1;
    for (let i = 0; i < values.length; i++) {
        if (("" + values[i][iRowId]).trim() === wantRowId) { r = i; break; }
    }
    if (r === -1) {
        return res(false, "row_not_found", "", emptyDetails);
    }

    const prevStatus = ("" + values[r][iStatus]).trim();
    const storedToken = ("" + values[r][iToken]).trim();

    // ---- token check (before revealing any fine detail) -----------------------
    // Empty stored token = row never sent. Mismatch = forged/guessed/stale link.
    // Details stay blank so a bad link learns nothing about the fine.
    if (storedToken === "" || wantToken === "" || wantToken !== storedToken) {
        return res(false, "invalid_token", prevStatus, emptyDetails);
    }

    // Token is valid from here on — safe to build the fine's display details.
    const details = {
        plateNumber: str(values[r][iPlate]),
        fineType: str(values[r][iFineType]),
        vendorName: str(values[r][iVendor]),
        referenceNumber: str(values[r][iRef]),
        offenceDateText: serialToDMY(values[r][iOffence]),
        dueDateText: serialToDMY(values[r][iDue]),
        amountText: amountToText(values[r][iAmount])
    };

    // ---- idempotent repeat: already confirmed ---------------------------------
    // Friendly page; show the details so the driver sees which fine is settled.
    if (prevStatus === "Confirmed") {
        return res(true, "already_confirmed", prevStatus, details);
    }

    // ---- guard: only a Notified row gets the date-picker -----------------------
    // Pending/Matched (not notified yet) and MatchFailed/SendFailed (need
    // attention) come back not_actionable with the state, so the flow can explain.
    if (SOURCE_STATES.indexOf(prevStatus) === -1) {
        return res(false, "not_actionable", prevStatus, details);
    }

    // ---- ready: token ok, row Notified -> render the date-picker page ----------
    return res(true, "ready", prevStatus, details);
}