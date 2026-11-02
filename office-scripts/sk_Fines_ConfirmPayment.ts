/**
 * sk_Fines_ConfirmPayment
 * -----------------------------------------------------------------------------
 * Records a driver's payment confirmation into the tblFines row it belongs to.
 * This is the ONLY place the payment-confirmation state transition happens, so
 * all of the safety (token check, date validation, idempotency, guarded
 * transition) lives here — the capture flow just passes the inputs through and
 * renders a page from the returned outcome.
 *
 * Called from the "GE_FineIntake_ConfirmPayment_Capture" flow (HTTP-triggered)
 * via Run script. The confirm page's date-picker + Submit hand rowId / token /
 * paymentDate here unchanged.
 *
 * MODEL (confirmed 2026-10-27, email + 3-stage notifications):
 *   MailStatus lifecycle: Pending -> Matched -> Notified -> Confirmed
 *                                              (SendFailed on a failed notify)
 *   The escalation stage (N1/N2/N3) lives in LastNotifyStage, NOT in MailStatus.
 *   A row becomes confirmable once ANY notification has gone out -> MailStatus
 *   is "Notified". There is no dispute path: the driver can only confirm payment
 *   (client decision) and must supply the payment date.
 *
 * Parameters:
 *   rowId        : the tblFines RowId the notification was built for (col RowId).
 *   confirmToken : the one-time nonce the send flow stamped on that row
 *                  (col ConfirmToken) and embedded in the confirm-link URL.
 *   paymentDate  : the date the driver actually paid, from the page's date-picker.
 *                  <input type="date"> always yields ISO "YYYY-MM-DD" regardless
 *                  of browser locale; serial and "DD.MM.YYYY" are also accepted
 *                  defensively. Written to the PaymentDate column as a real date.
 *
 *   NB: the former `confirmedBy` parameter was removed 2026-10-29 together with
 *   the ConfirmedBy column — the anonymous HTTP link cannot verify who actually
 *   clicked, so it only ever duplicated the matched driver's EmployeeEmail.
 *
 * Returns (read by the flow to pick which HTML page to show):
 *   ok      : true  -> a transition was made OR it was a safe idempotent repeat.
 *             false -> nothing was written; see outcome/reason.
 *   outcome : "recorded"          -> status set to Confirmed now, PaymentDate written.
 *             "already_confirmed" -> row was already Confirmed (no-op, friendly).
 *             "row_not_found"     -> no row with that RowId.
 *             "invalid_token"     -> token missing/!= the row's ConfirmToken.
 *             "bad_date"          -> paymentDate missing/unparseable/implausible.
 *             "not_actionable"    -> row in a state we don't transition from
 *                                    (previousStatus tells the flow which, e.g.
 *                                    Pending/Matched = not notified yet,
 *                                    MatchFailed/SendFailed = needs attention).
 *   newStatus      : the status written ("Confirmed" or "").
 *   previousStatus : the row's MailStatus before this call (for the flow's page).
 *   paymentDateText: the accepted payment date as "DD.MM.YYYY" (for the page/log).
 *   reason         : short free text, mirrors outcome for logging.
 *
 * Idempotency: a second click or a double-fired request is safe — once the row
 * is Confirmed the guard returns already_confirmed without rewriting (the first
 * PaymentDate stands). The token is deliberately NOT cleared after use, so a
 * legitimate repeat still reads a friendly "already recorded" rather than
 * "invalid".
 */

interface ConfirmResult {
    ok: boolean;
    outcome: string;
    newStatus: string;
    previousStatus: string;
    paymentDateText: string;
    reason: string;
}

function main(
    workbook: ExcelScript.Workbook,
    rowId: string,
    confirmToken: string,
    paymentDate: string
): ConfirmResult {
    const table = workbook.getTable("tblFines");
    if (!table) {
        // A missing table is a real failure, not an expected user outcome.
        throw new Error("Table 'tblFines' not found in the workbook.");
    }

    // The only state a confirmation may transition FROM: a row that has been
    // notified at least once. Anything else is either an idempotent repeat
    // (Confirmed, handled below) or not-actionable (Pending/Matched/MatchFailed/
    // SendFailed) — the flow tailors the page from previousStatus.
    const SOURCE_STATES = ["Notified"];

    const res = (
        ok: boolean, outcome: string,
        newStatus: string, previousStatus: string, paymentDateText: string
    ): ConfirmResult => ({
        ok, outcome, newStatus, previousStatus, paymentDateText, reason: outcome
    });

    // ---- date helpers (serial space; Excel epoch 1899-12-30, UTC) --------------
    const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);
    const MS_PER_DAY = 86400000;
    const dmyToSerial = (d: number, mo: number, y: number): number =>
        (Date.UTC(y, mo - 1, d) - EXCEL_EPOCH_UTC) / MS_PER_DAY;

    // Accept ISO YYYY-MM-DD (what <input type="date"> sends), a serial, or German
    // DD.MM.YYYY -> Excel serial; NaN if unparseable.
    const toSerial = (v: string): number => {
        const s = ("" + (v || "")).trim();
        if (!s) return NaN;
        if (/^\d+(\.\d+)?$/.test(s)) return Number(s);          // already a serial
        let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);            // ISO YYYY-MM-DD
        if (m) return dmyToSerial(Number(m[3]), Number(m[2]), Number(m[1]));
        m = s.match(/^(\d{2})\.(\d{2})\.(\d{4})/);              // German DD.MM.YYYY
        if (m) return dmyToSerial(Number(m[1]), Number(m[2]), Number(m[3]));
        return NaN;
    };

    const serialToText = (serial: number): string => {
        if (!isFinite(serial)) return "";
        const d = new Date(EXCEL_EPOCH_UTC + Math.round(serial) * MS_PER_DAY);
        const dd = ("0" + d.getUTCDate()).slice(-2);
        const mm = ("0" + (d.getUTCMonth() + 1)).slice(-2);
        return dd + "." + mm + "." + d.getUTCFullYear();
    };

    // ---- validate the payment date first (cheap, no table work) ----------------
    const paySerial = toSerial(paymentDate);
    if (!isFinite(paySerial)) {
        return res(false, "bad_date", "", "", "");
    }
    // Plausibility window: not before 2000-01-01, not more than ~2 days in the
    // future (a bit of slack for time zones). An out-of-range date is a typo or a
    // tampered link, not a real payment.
    const MIN_SERIAL = dmyToSerial(1, 1, 2000);
    const nowUtc = new Date();
    const todaySerial =
        (Date.UTC(nowUtc.getUTCFullYear(), nowUtc.getUTCMonth(), nowUtc.getUTCDate())
            - EXCEL_EPOCH_UTC) / MS_PER_DAY;
    if (paySerial < MIN_SERIAL || paySerial > todaySerial + 2) {
        return res(false, "bad_date", "", "", "");
    }
    const payText = serialToText(paySerial);

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
    const iPayDate = colIdx("PaymentDate");
    const iConfDate = colIdx("ConfirmedDate");

    const body = table.getRangeBetweenHeaderAndTotal();
    if (!body || body.getRowCount() === 0) {
        return res(false, "row_not_found", "", "", payText);
    }

    const values = body.getValues();
    let r = -1;
    for (let i = 0; i < values.length; i++) {
        if (("" + values[i][iRowId]).trim() === wantRowId) { r = i; break; }
    }
    if (r === -1) {
        return res(false, "row_not_found", "", "", payText);
    }

    const prevStatus = ("" + values[r][iStatus]).trim();
    const storedToken = ("" + values[r][iToken]).trim();

    // ---- token check ----------------------------------------------------------
    // Empty stored token = the row was never sent (no link, no nonce) -> reject.
    // Mismatch = forged/guessed/stale link -> reject. No shared secret needed:
    // the per-row GUID nonce is unguessable and lives only in the driver's inbox.
    if (storedToken === "" || wantToken === "" || wantToken !== storedToken) {
        return res(false, "invalid_token", "", prevStatus, payText);
    }

    // ---- idempotent repeat ----------------------------------------------------
    if (prevStatus === "Confirmed") {
        return res(true, "already_confirmed", "Confirmed", prevStatus, payText);
    }

    // ---- guarded transition ---------------------------------------------------
    // Only a Notified row proceeds. Pending/Matched (not notified yet),
    // MatchFailed/SendFailed (need attention) are returned as not_actionable with
    // the state so the flow can tailor the page.
    if (SOURCE_STATES.indexOf(prevStatus) === -1) {
        return res(false, "not_actionable", "", prevStatus, payText);
    }

    // ---- write the transition -------------------------------------------------
    const rowRange = body.getRow(r);

    // MailStatus -> Confirmed
    rowRange.getCell(0, iStatus).setValue("Confirmed");

    // PaymentDate: the driver-supplied payment date (date only), format then serial.
    rowRange.getCell(0, iPayDate).setNumberFormat("dd.mm.yyyy");
    rowRange.getCell(0, iPayDate).setValue(paySerial);

    // ConfirmedDate: when the click landed (now, local, with time fraction), same
    // convention as ProcessedAt in sk_Fines_WriteToExcel.
    const n = new Date();
    const localMs = Date.UTC(
        n.getFullYear(), n.getMonth(), n.getDate(),
        n.getHours(), n.getMinutes(), n.getSeconds()
    );
    const confirmedSerial = (localMs - EXCEL_EPOCH_UTC) / MS_PER_DAY;
    rowRange.getCell(0, iConfDate).setNumberFormat("dd.mm.yyyy hh:mm");
    rowRange.getCell(0, iConfDate).setValue(confirmedSerial);

    return res(true, "recorded", "Confirmed", prevStatus, payText);
}