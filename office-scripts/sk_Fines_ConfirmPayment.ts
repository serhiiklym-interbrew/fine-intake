/**
 * sk_Fines_ConfirmPayment
 * -----------------------------------------------------------------------------
 * Records a driver's response to a payment-confirmation Teams card into the
 * tblFines row it belongs to. This is the ONLY place the payment-confirmation
 * state transition happens, so all of the safety (token check, idempotency,
 * guarded transition) lives here — the capture flow just passes the inputs
 * through and renders a page from the returned outcome.
 *
 * Called from the "GE_FineIntake_ConfirmPayment_Capture" flow (HTTP-triggered)
 * via Run script. The flow takes rowId / token / action from the clicked
 * button's URL and hands them here unchanged.
 *
 * Parameters:
 *   rowId        : the tblFines RowId the card was built for (col RowId).
 *   confirmToken : the one-time nonce the send flow stamped on that row
 *                  (col ConfirmToken) and embedded in the button URL.
 *   action       : "paid"  -> driver confirms payment  -> status Confirmed
 *                  "dispute"-> "not my car / contest"   -> status Disputed
 *   confirmedBy  : who responded, if known. The HTTP trigger is anonymous, so
 *                  this is normally "" in this variant; written as given.
 *
 * Returns (read by the flow to pick which HTML page to show):
 *   ok      : true  -> a transition was made OR it was a safe idempotent repeat.
 *             false -> nothing was written; see outcome/reason.
 *   outcome : "recorded"          -> status set to Confirmed/Disputed now.
 *             "already_confirmed" -> row was already Confirmed (no-op).
 *             "already_disputed"  -> row was already Disputed  (no-op).
 *             "row_not_found"     -> no row with that RowId.
 *             "invalid_token"     -> token missing/!= the row's ConfirmToken.
 *             "not_actionable"    -> row in a state we don't transition from
 *                                    (previousStatus tells the flow which).
 *             "bad_action"        -> action not "paid"/"dispute".
 *   action         : the normalized action acted on ("paid"/"dispute"/"").
 *   newStatus      : the status written ("Confirmed"/"Disputed"/"").
 *   previousStatus : the row's MailStatus before this call (for the flow's page).
 *   reason         : short free text, mirrors outcome for logging.
 *
 * Idempotency: a second click, a forwarded card, or a double-fired request is
 * safe — once the row is Confirmed/Disputed the guard returns already_* without
 * rewriting. The token is deliberately NOT cleared after use, so a legitimate
 * repeat still reads a friendly "already recorded" rather than "invalid".
 */

interface ConfirmResult {
    ok: boolean;
    outcome: string;
    action: string;
    newStatus: string;
    previousStatus: string;
    reason: string;
}

function main(
    workbook: ExcelScript.Workbook,
    rowId: string,
    confirmToken: string,
    action: string,
    confirmedBy: string
): ConfirmResult {
    const table = workbook.getTable("tblFines");
    if (!table) {
        // A missing table is a real failure, not an expected user outcome.
        throw new Error("Table 'tblFines' not found in the workbook.");
    }

    // States this script is allowed to transition FROM. Anything else is either
    // an idempotent repeat (Confirmed/Disputed, handled below) or not-actionable.
    const SOURCE_STATES = ["AwaitingConfirmation", "Reminded"];

    const res = (
        ok: boolean, outcome: string, action: string,
        newStatus: string, previousStatus: string
    ): ConfirmResult => ({
        ok, outcome, action, newStatus, previousStatus, reason: outcome
    });

    // ---- normalize the action first (cheap, no table work) --------------------
    const act = ("" + (action || "")).trim().toLowerCase();
    let targetStatus = "";
    if (act === "paid") {
        targetStatus = "Confirmed";
    } else if (act === "dispute") {
        targetStatus = "Disputed";
    } else {
        return res(false, "bad_action", act, "", "");
    }

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
    const iConfDate = colIdx("ConfirmedDate");
    const iConfBy = colIdx("ConfirmedBy");

    const body = table.getRangeBetweenHeaderAndTotal();
    if (!body || body.getRowCount() === 0) {
        return res(false, "row_not_found", act, "", "");
    }

    const values = body.getValues();
    let r = -1;
    for (let i = 0; i < values.length; i++) {
        if (("" + values[i][iRowId]).trim() === wantRowId) { r = i; break; }
    }
    if (r === -1) {
        return res(false, "row_not_found", act, "", "");
    }

    const prevStatus = ("" + values[r][iStatus]).trim();
    const storedToken = ("" + values[r][iToken]).trim();

    // ---- token check ----------------------------------------------------------
    // Empty stored token = the row was never sent (no card, no nonce) -> reject.
    // Mismatch = forged/guessed/stale link -> reject. No shared secret needed:
    // the per-row GUID nonce is unguessable and lives only in the driver's chat.
    if (storedToken === "" || wantToken === "" || wantToken !== storedToken) {
        return res(false, "invalid_token", act, "", prevStatus);
    }

    // ---- idempotent repeat ----------------------------------------------------
    if (prevStatus === "Confirmed") {
        return res(true, "already_confirmed", act, "Confirmed", prevStatus);
    }
    if (prevStatus === "Disputed") {
        return res(true, "already_disputed", act, "Disputed", prevStatus);
    }

    // ---- guarded transition ---------------------------------------------------
    // Only AwaitingConfirmation / Reminded proceed. Pending, Matched, Sent,
    // SendFailed, Expired, etc. are returned as not_actionable with the state so
    // the flow can tailor the page (e.g. Expired -> contact Fleet.Germany).
    if (SOURCE_STATES.indexOf(prevStatus) === -1) {
        return res(false, "not_actionable", act, "", prevStatus);
    }

    // ---- write the transition -------------------------------------------------
    // now -> Excel serial (local, with time fraction), same convention as
    // ProcessedAt in sk_Fines_WriteToExcel.
    const n = new Date();
    const localMs = Date.UTC(
        n.getFullYear(), n.getMonth(), n.getDate(),
        n.getHours(), n.getMinutes(), n.getSeconds()
    );
    const confirmedSerial = (localMs - Date.UTC(1899, 11, 30)) / 86400000;

    const rowRange = body.getRow(r);

    rowRange.getCell(0, iStatus).setValue(targetStatus);

    // ConfirmedDate: set the display format, then the serial value.
    rowRange.getCell(0, iConfDate).setNumberFormat("dd.mm.yyyy hh:mm");
    rowRange.getCell(0, iConfDate).setValue(confirmedSerial);

    // ConfirmedBy is a text identifier (email/name) — keep the cell Text, then
    // write it (may be "" in the anonymous-HTTP variant).
    rowRange.getCell(0, iConfBy).setNumberFormat("@");
    rowRange.getCell(0, iConfBy).setValue("" + (confirmedBy || ""));

    return res(true, "recorded", act, targetStatus, prevStatus);
}