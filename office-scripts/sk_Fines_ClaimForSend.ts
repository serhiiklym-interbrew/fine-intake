/**
 * sk_Fines_ClaimForSend
 * -----------------------------------------------------------------------------
 * Decides whether a tblFines row is DUE for the next notification stage and, if
 * so, writes ONLY the one-time confirm token + returns everything the flow needs
 * to build and send the mail(s). It deliberately does NOT commit the "sent"
 * state — MailStatus / LastNotifyStage / LastNotifyDate / DateSent are written
 * AFTER the send succeeds, by sk_Fines_SetSendResult. (Two-phase design.)
 *
 * WHY TWO PHASES (important): if this script flipped the row to Notified +
 * stamped LastNotifyStage up front and the send then FAILED, the row would lie
 * ("N2 sent") while nothing went out, and the next sweep would see the stage as
 * already-sent and never retry. By committing the stage only after a successful
 * send, LastNotifyStage reflects ONLY stages that truly went out, so:
 *   - a failed send leaves LastNotifyStage untouched -> next sweep recomputes the
 *     SAME due stage, rank(stage) > rank(lastStage) passes naturally -> RETRY,
 *     with no guard hack and no stage rollback.
 *   - SendFailed is in SOURCE_STATES, so the row keeps being picked up until it
 *     is Confirmed or the key user fixes it.
 *
 * Called from "GE_FineIntake_Notify_Scheduled" via Run script, once per row
 * inside Apply-to-each. The flow proceeds to send ONLY when claimed = true, then
 * calls sk_Fines_SetSendResult(rowId, "sent", stage, "") on success, or
 * sk_Fines_SetSendResult(rowId, "failed", stage, <error>) from its Scope catch.
 *
 * STAGE MODEL (confirmed matrix):
 *   Based on days-to-DueDate (diff = DueDate - today, in whole days):
 *     diff  > 3   -> "N1"   (> 3 days before due)
 *     1..3        -> "N2"   (<= 3 days before due, not yet past)
 *     diff <= 0   -> "N3"   (due date today or already past)
 *   FineType "Cost Notification" overrides to "CN" regardless of DueDate.
 *   Empty/unreadable DueDate -> "N1" (notify once, no escalation).
 *
 *   Recipients per stage are the FLOW's job; this script only reports the stage,
 *   plus two flags the flow acts on:
 *     includeAuthority  -> Normal + stage N1 + a non-empty AuthorityEmail
 *                          (the one stage the authority Fahrerbenennung goes out)
 *     escalate          -> stage N3 or CN (add line manager + security manager)
 *   For manager/security and the Normal-authority BirthDate the flow calls
 *   sk_Eop_FindByPn(personalnummer) itself — those are transient, not in tblFines.
 *
 * ATTACHMENT: fileUrl is the server-relative path (decoded) of the fine's own
 *   PDF, read from the hyperlink in the "File" column (List rows only returns
 *   the hyperlink DisplayText, so the path must come from a script). The flow
 *   feeds it to "Get file content using path" (Site = .../sites/FleetEurope)
 *   and attaches the result to the DRIVER mail only (authority already has it).
 *   Populated only on a real claim; "" on every skip.
 *
 * IDEMPOTENCY: a stage is "due" only if it is NEWER than LastNotifyStage
 *   (rank N1<N2<N3; CN is one-shot, ranked with N3). A same-or-older stage comes
 *   back claimed=false reason="already_<stage>". Because LastNotifyStage is now
 *   only stamped on a SUCCESSFUL send, this guard never blocks a retry of a
 *   failed stage. Re-claiming overwrites ConfirmToken so each attempt gets a
 *   fresh nonce and no stale token survives.
 *
 * Parameters:
 *   rowId   : the tblFines RowId to claim (col RowId).
 *   todayIso: the sweep's "today" as "YYYY-MM-DD" (pass utcNow('yyyy-MM-dd') or a
 *             CET-local date from the flow, so the stage math is deterministic and
 *             not tied to the script runtime's zone). Empty -> script uses now.
 *
 * Returns (read by the flow):
 *   ok            : false only on a real failure (row not found). true otherwise.
 *   claimed       : true  -> row is DUE; token written; SEND the stage, then
 *                            call sk_Fines_SetSendResult with the outcome.
 *                   false -> nothing to send; skip (see reason).
 *   reason        : "" on claim; else why not:
 *                   "already_<stage>" -> computed stage not newer than LastNotifyStage.
 *                   "not_passed"      -> OCR Status != "Passed" (un-reviewed fine).
 *                   "no_recipient"    -> EmployeeEmail empty, cannot address driver.
 *                   "row_not_found"   -> no row with that RowId (ok=false).
 *   stage            : "N1" | "N2" | "N3" | "CN" (the due stage; PASS THIS to
 *                      sk_Fines_SetSendResult so it stamps the right stage).
 *   recipientEmail   : EmployeeEmail (the driver) to address.
 *   includeAuthority : flow adds the authority Fahrerbenennung mail (see above).
 *   authorityEmail   : AuthorityEmail (only meaningful when includeAuthority).
 *   escalate         : flow adds line manager + security manager (N3 / CN).
 *   personalnummer   : for the flow's EOP lookup (manager / BirthDate).
 *   driverSource     : "DUMP" | "Pool" — which lookup the flow should use.
 *   confirmToken     : the one-time nonce written to the row + embedded in the
 *                      confirm-link URL (meaningful only when claimed = true).
 *   fileUrl          : server-relative path of the fine's PDF for the driver
 *                      attachment (meaningful only when claimed = true; "" else).
 *   mail             : display-ready strings for the letter body (serials already
 *                      converted): employeeName, employeeAddress, plateNumber,
 *                      referenceNumber, fineType, vendorName, location,
 *                      amountText, offenceDateText, dueDateText.
 */

interface ClaimMail {
    employeeName: string;
    employeeAddress: string;
    plateNumber: string;
    referenceNumber: string;
    fineType: string;
    vendorName: string;
    location: string;
    amountText: string;
    offenceDateText: string;
    dueDateText: string;
}

interface ClaimResult {
    ok: boolean;
    claimed: boolean;
    reason: string;
    stage: string;
    recipientEmail: string;
    includeAuthority: boolean;
    authorityEmail: string;
    escalate: boolean;
    personalnummer: string;
    driverSource: string;
    confirmToken: string;
    fileUrl: string;
    mail: ClaimMail;
}

function main(
    workbook: ExcelScript.Workbook,
    rowId: string,
    todayIso: string
): ClaimResult {
    const table = workbook.getTable("tblFines");
    if (!table) {
        throw new Error("Table 'tblFines' not found in the workbook.");
    }

    // Rows the sweeper may claim FROM: a fresh match, a row already notified but
    // not yet confirmed (escalation), or a prior send that failed (retry).
    const SOURCE_STATES = ["Matched", "Notified", "SendFailed"];

    const emptyMail = (): ClaimMail => ({
        employeeName: "", employeeAddress: "", plateNumber: "", referenceNumber: "",
        fineType: "", vendorName: "", location: "", amountText: "",
        offenceDateText: "", dueDateText: ""
    });

    const res = (
        ok: boolean, claimed: boolean, reason: string, stage: string,
        recipientEmail: string, includeAuthority: boolean, authorityEmail: string,
        escalate: boolean, personalnummer: string, driverSource: string,
        confirmToken: string, fileUrl: string, mail: ClaimMail
    ): ClaimResult => ({
        ok, claimed, reason, stage, recipientEmail, includeAuthority,
        authorityEmail, escalate, personalnummer, driverSource, confirmToken,
        fileUrl, mail
    });

    // ---- date / display helpers (Excel epoch 1899-12-30, UTC) ------------------
    const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);
    const MS_PER_DAY = 86400000;

    const isoToSerial = (iso: string): number => {
        const m = ("" + (iso || "")).match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (!m) return NaN;
        return (Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])) - EXCEL_EPOCH_UTC) / MS_PER_DAY;
    };

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
    // confirm page + MatchFailed mail use, so letters and pages agree.
    const amountToText = (v: string | number | boolean): string => {
        if (v === "" || v === null || v === undefined) return "";
        const n = typeof v === "number" ? v : Number(v);
        if (!isFinite(n)) return "";
        const fixed = n.toFixed(2);
        const dot = fixed.indexOf(".");
        const grouped = fixed.slice(0, dot).replace(/\B(?=(\d{3})+(?!\d))/g, ".");
        return grouped + "," + fixed.slice(dot + 1) + " €";
    };

    const str = (v: string | number | boolean): string =>
        (v === null || v === undefined) ? "" : ("" + v);

    // SharePoint hyperlink address -> decoded server-relative path for
    // "Get file content using path" (Site = .../sites/FleetEurope).
    // "https://....sharepoint.com/sites/FleetEurope/Shared%20Documents/.../x.pdf"
    //   -> "/Shared Documents/.../x.pdf"
    const toServerRelative = (addr: string): string => {
        const a = ("" + (addr || "")).trim();
        if (a === "") return "";
        const marker = "/sites/FleetEurope";
        const i = a.indexOf(marker);
        const rel = i === -1 ? a : a.slice(i + marker.length);
        try { return decodeURIComponent(rel); } catch (e) { return rel; }
    };

    // GUID-style nonce from Math.random (crypto not guaranteed in Office Scripts).
    function makeToken(): string {
        const hex = (n: number): string => {
            let s = "";
            for (let i = 0; i < n; i++) s += Math.floor(Math.random() * 16).toString(16);
            return s;
        };
        const y = (8 + Math.floor(Math.random() * 4)).toString(16);
        return hex(8) + "-" + hex(4) + "-4" + hex(3) + "-" + y + hex(3) + "-" + hex(12);
    }

    // Stage rank for the "newer than LastNotifyStage" guard. CN is one-shot,
    // ranked with N3 so nothing escalates past it.
    const rank = (s: string): number => {
        const u = ("" + s).trim().toUpperCase();
        if (u === "N1") return 1;
        if (u === "N2") return 2;
        if (u === "N3") return 3;
        if (u === "CN") return 3;
        return 0; // empty / unknown -> never sent yet
    };

    // ---- locate the row by RowId (columns by NAME) ----------------------------
    const colIdx = (name: string): number => {
        const c = table.getColumnByName(name);
        if (!c) throw new Error("Column '" + name + "' not found in tblFines.");
        return c.getIndex();
    };
    const iRowId = colIdx("RowId");
    const iMailStatus = colIdx("MailStatus");
    const iOcrStatus = colIdx("Status");       // OCR status (Passed/Check), NOT MailStatus
    const iFineType = colIdx("FineType");
    const iDueDate = colIdx("DueDate");
    const iToken = colIdx("ConfirmToken");
    const iLastStage = colIdx("LastNotifyStage");
    const iEmpEmail = colIdx("EmployeeEmail");
    const iEmpName = colIdx("EmployeeName");
    const iEmpAddr = colIdx("EmployeeAddress");
    const iAuthEmail = colIdx("AuthorityEmail");
    const iPN = colIdx("Personalnummer");
    const iDriverSrc = colIdx("DriverSource");
    const iFile = colIdx("File");              // hyperlink -> the fine's PDF
    const iPlate = colIdx("PlateNumber");
    const iRef = colIdx("ReferenceNumber");
    const iVendor = colIdx("VendorName");
    const iLocation = colIdx("Location");
    const iAmount = colIdx("Amount");
    const iOffence = colIdx("OffenceDate");

    const want = ("" + (rowId || "")).trim();

    const body = table.getRangeBetweenHeaderAndTotal();
    if (!body || body.getRowCount() === 0) {
        return res(false, false, "row_not_found", "", "", false, "", false, "", "", "", "", emptyMail());
    }

    const values = body.getValues();
    let r = -1;
    for (let i = 0; i < values.length; i++) {
        if (("" + values[i][iRowId]).trim() === want) { r = i; break; }
    }
    if (r === -1) {
        return res(false, false, "row_not_found", "", "", false, "", false, "", "", "", "", emptyMail());
    }

    const prevStatus = ("" + values[r][iMailStatus]).trim();

    // ---- guard: only sweep from an allowed source state -----------------------
    if (SOURCE_STATES.indexOf(prevStatus) === -1) {
        return res(true, false, "already_" + (prevStatus || "unknown"), "", "", false, "", false, "", "", "", "", emptyMail());
    }

    // ---- guard: OCR status must be Passed (never mail an un-reviewed fine) -----
    if (("" + values[r][iOcrStatus]).trim() !== "Passed") {
        return res(true, false, "not_passed", "", "", false, "", false, "", "", "", "", emptyMail());
    }

    // ---- guard: driver recipient required -------------------------------------
    const recipientEmail = str(values[r][iEmpEmail]).trim();
    if (recipientEmail === "") {
        return res(true, false, "no_recipient", "", "", false, "", false, "", "", "", "", emptyMail());
    }

    // ---- compute the due stage ------------------------------------------------
    const fineType = str(values[r][iFineType]).trim();
    let stage: string;
    if (fineType.toLowerCase() === "cost notification") {
        stage = "CN";
    } else {
        const dueSerial = Number(values[r][iDueDate]);
        const todaySerial = isFinite(isoToSerial(todayIso))
            ? isoToSerial(todayIso)
            : Math.floor((Date.now() - EXCEL_EPOCH_UTC) / MS_PER_DAY);
        if (!isFinite(dueSerial) || dueSerial <= 0) {
            stage = "N1";                       // no readable DueDate -> notify once
        } else {
            const diff = dueSerial - todaySerial;
            stage = diff > 3 ? "N1" : (diff > 0 ? "N2" : "N3");
        }
    }

    // ---- idempotency: only send a stage newer than the last one SUCCESSFULLY
    //      sent. Because LastNotifyStage is stamped only after a successful send
    //      (sk_Fines_SetSendResult), a SendFailed row still has the previous
    //      stage here, so the due stage is "newer" and this guard lets the retry
    //      through naturally — no special-casing of SendFailed needed.
    const lastStage = str(values[r][iLastStage]).trim();
    if (rank(stage) <= rank(lastStage)) {
        return res(true, false, "already_" + (lastStage || stage), stage,
            recipientEmail, false, "", false, "", "", "", "", emptyMail());
    }

    // ---- claim: write ONLY the fresh confirm nonce ----------------------------
    //      State (MailStatus/LastNotifyStage/LastNotifyDate/DateSent) is written
    //      AFTER a successful send by sk_Fines_SetSendResult. Do not stamp it here.
    const token = makeToken();
    const rowRange = body.getRow(r);

    // Read the PDF path from the "File" hyperlink (List rows can't; a script can).
    let fileUrl = "";
    try {
        const link = rowRange.getCell(0, iFile).getHyperlink();
        if (link && link.address) fileUrl = toServerRelative(link.address);
    } catch (e) { fileUrl = ""; }

    rowRange.getCell(0, iToken).setNumberFormat("@");
    rowRange.getCell(0, iToken).setValue(token);

    // ---- recipients the flow needs to resolve ---------------------------------
    // Authority Fahrerbenennung goes out once: for Normal on its first stage (N1),
    // and for a Cost Notification (its own one-shot stage). Fair never mails an
    // authority, and Normal's later stages (N2/N3) don't repeat the authority mail.
    const authorityEmail = str(values[r][iAuthEmail]).trim();
    const ftLower = fineType.toLowerCase();
    const includeAuthority =
        authorityEmail !== "" &&
        ((ftLower === "normal" && stage === "N1") || ftLower === "cost notification");
    const escalate = stage === "N3" || stage === "CN";

    const mail: ClaimMail = {
        employeeName: str(values[r][iEmpName]),
        employeeAddress: str(values[r][iEmpAddr]),
        plateNumber: str(values[r][iPlate]),
        referenceNumber: str(values[r][iRef]),
        fineType: fineType,
        vendorName: str(values[r][iVendor]),
        location: str(values[r][iLocation]),
        amountText: amountToText(values[r][iAmount]),
        offenceDateText: serialToDMY(values[r][iOffence]),
        dueDateText: serialToDMY(values[r][iDueDate])
    };

    return res(
        true, true, "", stage, recipientEmail, includeAuthority, authorityEmail,
        escalate, str(values[r][iPN]), str(values[r][iDriverSrc]), token,
        fileUrl, mail
    );
}