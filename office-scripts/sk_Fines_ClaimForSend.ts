/**
 * sk_Fines_ClaimForSend
 * -----------------------------------------------------------------------------
 * Atomically "claims" one tblFines row for payment-confirmation dispatch, so the
 * Teams card can be posted exactly once. This is the guard that prevents
 * double-posting: the moment a row is claimed its MailStatus flips to
 * AwaitingConfirmation, so the next recurrence's List-rows filter
 * (MailStatus in {Matched, SendFailed}) no longer picks it up.
 *
 * Called from the "GE_FineIntake_ConfirmPayment_Send" flow via Run script, once
 * per row inside the Apply-to-each over the filtered rows. The flow then posts
 * the card ONLY when claimed = true, using the returned card fields + token.
 *
 * Within a single Run script execution the read-check-write is one synchronous
 * pass against the workbook, so the claim is atomic for that run; the status flip
 * is what protects against a second recurrence re-selecting the same row.
 *
 * Parameters:
 *   rowId : the tblFines RowId to claim (col RowId).
 *
 * Returns (read by the flow):
 *   ok            : false only on a real failure (row not found). true otherwise.
 *   claimed       : true  -> row flipped to AwaitingConfirmation; POST the card.
 *                   false -> nothing changed; skip this row (see reason).
 *   reason        : "" on claim; otherwise why not:
 *                   "already_<status>" -> not in {Matched, SendFailed}
 *                                         (stale list / already claimed).
 *                   "not_passed"       -> OCR Status != "Passed" (a Check row);
 *                                         a last-gate safety so an un-reviewed
 *                                         fine never gets a driver card, even if
 *                                         the phase-1 filter ever lets it through.
 *                   "no_recipient"     -> EmployeeEmail empty, cannot address a
 *                                         card; left untouched for review.
 *                   "row_not_found"    -> no row with that RowId (ok=false).
 *   confirmToken   : the one-time nonce written to the row and to embed in the
 *                    button URL (only meaningful when claimed = true).
 *   recipientEmail : EmployeeEmail to address the Teams card to.
 *   card           : display-ready strings for the card body (serials already
 *                    converted): employeeName, plateNumber, referenceNumber,
 *                    fineType, vendorName, amountText, offenceDateText.
 *
 * On a post failure the flow sets the row back to SendFailed (a claimable source
 * state); the next claim overwrites ConfirmToken + DateSent and re-claims, so a
 * failed send retries cleanly and no stale token survives.
 */

interface ClaimCard {
    employeeName: string;
    plateNumber: string;
    referenceNumber: string;
    fineType: string;
    vendorName: string;
    amountText: string;
    offenceDateText: string;
}

interface ClaimResult {
    ok: boolean;
    claimed: boolean;
    reason: string;
    confirmToken: string;
    recipientEmail: string;
    card: ClaimCard;
}

function main(
    workbook: ExcelScript.Workbook,
    rowId: string
): ClaimResult {
    const table = workbook.getTable("tblFines");
    if (!table) {
        throw new Error("Table 'tblFines' not found in the workbook.");
    }

    // Source states a row may be claimed FROM: fresh match, or a prior send that
    // failed and is being retried. Reminded is the sweeper's job, not this flow.
    const SOURCE_STATES = ["Matched", "SendFailed"];

    const emptyCard = (): ClaimCard => ({
        employeeName: "", plateNumber: "", referenceNumber: "",
        fineType: "", vendorName: "", amountText: "", offenceDateText: ""
    });

    const res = (
        ok: boolean, claimed: boolean, reason: string,
        confirmToken: string, recipientEmail: string, card: ClaimCard
    ): ClaimResult => ({ ok, claimed, reason, confirmToken, recipientEmail, card });

    // ---- helpers --------------------------------------------------------------
    // A v4-style GUID from Math.random — a nonce, not a cryptographic secret
    // (crypto isn't guaranteed in the Office Scripts runtime). Unguessability
    // plus the status guard in sk_Fines_ConfirmPayment is what secures confirm.
    function makeToken(): string {
        const hex = (n: number): string => {
            let s = "";
            for (let i = 0; i < n; i++) s += Math.floor(Math.random() * 16).toString(16);
            return s;
        };
        const y = (8 + Math.floor(Math.random() * 4)).toString(16); // 8..b
        return hex(8) + "-" + hex(4) + "-4" + hex(3) + "-" + y + hex(3) + "-" + hex(12);
    }

    // Excel serial -> "dd.MM.yyyy"; "" for empty/non-numeric (e.g. a Check row
    // with no OffenceDate). Dates were stored as midnight UTC, so read back UTC.
    function serialToDMY(v: string | number | boolean): string {
        if (v === "" || v === null || v === undefined) return "";
        const n = typeof v === "number" ? v : Number(v);
        if (!isFinite(n) || n <= 0) return "";
        const ms = n * 86400000 + Date.UTC(1899, 11, 30);
        const dt = new Date(ms);
        const dd = ("0" + dt.getUTCDate()).slice(-2);
        const mm = ("0" + (dt.getUTCMonth() + 1)).slice(-2);
        const yyyy = dt.getUTCFullYear();
        return dd + "." + mm + "." + yyyy;
    }

    // Amount serial/number -> "123.45"; "" for empty.
    function amountToText(v: string | number | boolean): string {
        if (v === "" || v === null || v === undefined) return "";
        const n = typeof v === "number" ? v : Number(v);
        if (!isFinite(n)) return "";
        return n.toFixed(2);
    }

    const str = (v: string | number | boolean): string =>
        (v === null || v === undefined) ? "" : ("" + v);

    // ---- locate the row by RowId (columns by NAME) ----------------------------
    const colIdx = (name: string): number => {
        const c = table.getColumnByName(name);
        if (!c) throw new Error("Column '" + name + "' not found in tblFines.");
        return c.getIndex();
    };
    const iRowId = colIdx("RowId");
    const iStatus = colIdx("MailStatus");
    const iOcrStatus = colIdx("Status"); // OCR status (Passed/Check) — NOT MailStatus
    const iToken = colIdx("ConfirmToken");
    const iDateSent = colIdx("DateSent");
    const iEmail = colIdx("EmployeeEmail");
    const iEmpName = colIdx("EmployeeName");
    const iPlate = colIdx("PlateNumber");
    const iRef = colIdx("ReferenceNumber");
    const iFineType = colIdx("FineType");
    const iVendor = colIdx("VendorName");
    const iAmount = colIdx("Amount");
    const iOffence = colIdx("OffenceDate");

    const want = ("" + (rowId || "")).trim();

    const body = table.getRangeBetweenHeaderAndTotal();
    if (!body || body.getRowCount() === 0) {
        return res(false, false, "row_not_found", "", "", emptyCard());
    }

    const values = body.getValues();
    let r = -1;
    for (let i = 0; i < values.length; i++) {
        if (("" + values[i][iRowId]).trim() === want) { r = i; break; }
    }
    if (r === -1) {
        return res(false, false, "row_not_found", "", "", emptyCard());
    }

    const prevStatus = ("" + values[r][iStatus]).trim();

    // ---- guard: only claim from an allowed source state -----------------------
    if (SOURCE_STATES.indexOf(prevStatus) === -1) {
        return res(true, false, "already_" + (prevStatus || "unknown"), "", "", emptyCard());
    }

    // ---- guard: OCR status must be Passed (last gate before a live card) -------
    // Check rows are un-reviewed fines (bad plate / missing amount) — a human must
    // fix and promote them to Passed first. The phase-1 filter already excludes
    // Check, so reaching here means the filter was changed; this is the safety net
    // that keeps a driver from ever getting a card for an un-reviewed fine.
    if (("" + values[r][iOcrStatus]).trim() !== "Passed") {
        return res(true, false, "not_passed", "", "", emptyCard());
    }

    // ---- guard: must have a recipient to address the card ----------------------
    const recipientEmail = str(values[r][iEmail]).trim();
    if (recipientEmail === "") {
        return res(true, false, "no_recipient", "", "", emptyCard());
    }

    // ---- claim: write the nonce + flip the state ------------------------------
    const token = makeToken();

    const n = new Date();
    const localMs = Date.UTC(
        n.getFullYear(), n.getMonth(), n.getDate(),
        n.getHours(), n.getMinutes(), n.getSeconds()
    );
    const sentSerial = (localMs - Date.UTC(1899, 11, 30)) / 86400000;

    const rowRange = body.getRow(r);

    // ConfirmToken is a text identifier compared exactly on confirm — keep Text.
    rowRange.getCell(0, iToken).setNumberFormat("@");
    rowRange.getCell(0, iToken).setValue(token);

    rowRange.getCell(0, iDateSent).setNumberFormat("dd.mm.yyyy hh:mm");
    rowRange.getCell(0, iDateSent).setValue(sentSerial);

    rowRange.getCell(0, iStatus).setValue("AwaitingConfirmation");

    // ---- build the card payload from the row ----------------------------------
    const card: ClaimCard = {
        employeeName: str(values[r][iEmpName]),
        plateNumber: str(values[r][iPlate]),
        referenceNumber: str(values[r][iRef]),
        fineType: str(values[r][iFineType]),
        vendorName: str(values[r][iVendor]),
        amountText: amountToText(values[r][iAmount]),
        offenceDateText: serialToDMY(values[r][iOffence])
    };

    return res(true, true, "", token, recipientEmail, card);
}