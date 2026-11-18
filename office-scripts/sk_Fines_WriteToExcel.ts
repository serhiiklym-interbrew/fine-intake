/**
 * sk_Fines_WriteToExcel
 * -----------------------------------------------------------------------------
 * Appends one recognized German fine to the tblFines table in Fines_Data.xlsx,
 * unless its reference number already exists (dedup) — in which case no row is
 * written and the caller is told it was a duplicate.
 *
 * Called from the "GE_FineIntake_*_Auto" Power Automate flows via Run script
 * (Normal / Fair / Cost Notification — all share this script).
 *
 * Parameters (fileName FIRST, per the house clone convention):
 *   fileName   : human file name as archived (flow-supplied).
 *   fineType   : "Normal" | "Fair" | "Cost Notification" (from the source folder).
 *   status     : "Passed" or "Check" ("Failed" rows are never written).
 *   fileLink   : SharePoint "Link to item" of the archived PDF (may be empty).
 *   folderPath : SharePoint link to the submission folder (may be empty).
 *   dataJson   : the GPT JSON string. 9 recognized fields + 3 DueDate inputs:
 *                VendorName, PlateNumber, Aktenzeichen, OffenceReason,
 *                OffenceDate, Location, LetterDate, Amount, AuthorityEmail,
 *                StampDate, DeadlineDays, DueDateExplicit.
 *                (The JSON key stays "Aktenzeichen" — the document's language;
 *                 it maps to the ReferenceNumber column.)
 *
 * DueDate (col 8) is computed HERE, not by the LLM (date arithmetic in the LLM
 * is unreliable). Rule:
 *     base    = StampDate ?? LetterDate        (incoming-stamp date preferred,
 *                                               it is closest to legal "Zugang")
 *     DueDate = DueDateExplicit ?? (base + DeadlineDays)
 * If nothing resolves, DueDate is left empty.
 *
 * OCRPercent (col 11) is computed HERE — the share of the recognized columns
 * EXPECTED for this fine type that came back non-empty (variant C: fields that
 * are absent by design are left out of the denominator so each type can honestly
 * reach 100%). It is a completeness signal, NOT an OCR confidence: it says how
 * many expected columns were filled, not whether any value is correct. Expected
 * columns (denominator) per type:
 *     Normal (9): PlateNumber, VendorName, ReferenceNumber, OffenceReason,
 *                 OffenceDate, Location, LetterDate, Amount, AuthorityEmail
 *     Fair   (8): the above minus AuthorityEmail
 *                 (private operators usually have no contact email)
 *     CN     (6): PlateNumber, VendorName, ReferenceNumber, Location,
 *                 LetterDate, Amount
 *                 (reminders may restate neither the offence reason nor its
 *                  date, and the issuer email is inconsistent — all three out)
 *
 * Check highlighting: when a row is written with Status = "Check", the key cells
 * that came back empty are filled yellow (#FFEB9C) so the specialist sees at a
 * glance what to complete from the file. The flagged set is per fine type, so a
 * field that is legitimately empty for a type is never falsely flagged:
 *     Normal: ReferenceNumber, OffenceDate, DueDate, AuthorityEmail,
 *             VendorName, LetterDate, Location
 *     Fair  : the above minus AuthorityEmail (private operators have no email)
 *     CN    : ReferenceNumber, VendorName, LetterDate, Location
 *             (offence date not restated; stage is fixed CN, not DueDate-driven;
 *              issuer email inconsistent)
 * Amount and OffenceReason are informational and not flagged. PlateNumber is
 * flagged only when it is present but malformed (the implausible-plate guard
 * that downgraded the row to Check) — an empty plate is Failed, never a row.
 *
 * Returns (read by the flow):
 *   { inserted: true,  reason: "", status }  -> a new row was appended. `status`
 *                                               is the FINAL row status (Passed or
 *                                               Check). The flow gates its Check
 *                                               Teams post on THIS, not on its own
 *                                               varStatus, so a script-side Check
 *                                               downgrade (implausible plate, or an
 *                                               unresolved DueDate on Normal/Fair)
 *                                               is posted too.
 *   { inserted: false, reason: "duplicate", status: "" } -> NormalizedReferenceNumber
 *                                               already present; no row written.
 *   The file is archived by the flow regardless, so a duplicate is logged as
 *   Skipped (RowsInserted = 0), not Failed.
 *
 * Column order written (must match tblFines exactly — 34 columns).
 * NB: this grid follows the EXECUTED row[] array below, which is the authoritative
 * physical order — not an earlier header/handoff. OffenceReason was inserted at
 * index 4 (before OffenceDate) 2026-11-05, shifting every column from OffenceDate
 * onward down by one. (ConfirmedBy was removed 2026-10-29.)
 *   0  PlateNumber        9  Amount           18 VehicleFrom       27 FileName
 *   1  FineType          10  Status           19 MailStatus        28 NormalizedReferenceNumber
 *   2  VendorName        11  OCRPercent       20 LastNotifyStage   29 RowId
 *   3  ReferenceNumber   12  AuthorityEmail   21 LastNotifyDate    30 ConfirmToken
 *   4  OffenceReason     13  Personalnummer   22 DateSent          31 ProcessedAt
 *   5  OffenceDate       14  EmployeeName     23 ConfirmedDate     32 File
 *   6  Location          15  EmployeeEmail    24 PaymentDate       33 FolderPath
 *   7  LetterDate        16  EmployeeAddress  25 MatchNote
 *   8  DueDate           17  DriverSource     26 ReviewNote
 *
 *   Intake fills the recognized fields (cols 0-11 + AuthorityEmail),
 *   NormalizedReferenceNumber, RowId, ProcessedAt, File, FolderPath, and seeds
 *   MailStatus = "Pending" for every row (Passed and Check alike). All match /
 *   notify / payment columns are left empty for the downstream flows:
 *     - match flow:   Personalnummer, EmployeeName, EmployeeEmail,
 *                     EmployeeAddress, DriverSource, VehicleFrom, MatchNote.
 *     - notify flow:  LastNotifyStage, LastNotifyDate, DateSent.
 *     - capture flow: ConfirmedDate, PaymentDate; the send flow stamps the
 *                     one-time ConfirmToken nonce.
 *   RowId is the unique text key the downstream flows address this exact row by
 *   ("R" + Date.now() + "-" + random suffix): FileName is not unique, ProcessedAt
 *   resolves only to the second, and Check rows carry no reference.
 *
 *   OffenceReason (col 4) is the short description of the alleged violation,
 *   pulled from dataJson — written as text, for the specialist to read in Excel;
 *   it is not a match / dedup key and affects nothing downstream.
 *   Location (col 6) is the violation location, pulled from dataJson — written as
 *   text, needed by the Normal/CN letters.
 */

interface WriteResult {
    inserted: boolean;
    reason: string;
    status: string; // final row status (Passed/Check) — the flow reads this
}

function main(
    workbook: ExcelScript.Workbook,
    fileName: string,
    fineType: string,
    status: string,
    fileLink: string,
    folderPath: string,
    dataJson: string
): WriteResult {
    const table = workbook.getTable("tblFines");
    if (!table) {
        throw new Error("Table 'tblFines' not found in the workbook.");
    }

    // ---- parse GPT output -----------------------------------------------------
    interface FineData {
        VendorName?: string | null;
        PlateNumber?: string | null;
        Aktenzeichen?: string | null;
        OffenceReason?: string | null;
        OffenceDate?: string | null;
        Location?: string | null;
        LetterDate?: string | null;
        Amount?: number | null;
        AuthorityEmail?: string | null;
        StampDate?: string | null;
        DeadlineDays?: number | string | null;
        DueDateExplicit?: string | null;
    }

    let d: FineData = {};
    try {
        d = JSON.parse(dataJson) as FineData;
    } catch (e) {
        throw new Error("dataJson is not valid JSON: " + e);
    }

    // ---- helpers --------------------------------------------------------------
    const txt = (v: string | null | undefined): string =>
        (v === null || v === undefined) ? "" : ("" + v);

    // Non-empty after trimming (used by the OCRPercent completeness count).
    const hasStr = (v: string): boolean =>
        v !== null && v !== undefined && ("" + v).trim() !== "";

    // Same rule as the stored NormalizedReferenceNumber: strip whitespace + upper.
    const normalizeRef = (v: string): string =>
        ("" + v).replace(/\s+/g, "").toUpperCase();

    // ISO date (YYYY-MM-DD...) -> Excel serial number; "" when empty/invalid.
    function isoDateToSerial(iso: string): number | string {
        if (!iso) return "";
        const m = iso.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (!m) return "";
        const ms = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
        const epoch = Date.UTC(1899, 11, 30);
        return (ms - epoch) / 86400000;
    }

    // Current local timestamp -> Excel serial (with time fraction).
    function nowToSerial(): number {
        const n = new Date();
        const localMs = Date.UTC(
            n.getFullYear(), n.getMonth(), n.getDate(),
            n.getHours(), n.getMinutes(), n.getSeconds()
        );
        const epoch = Date.UTC(1899, 11, 30);
        return (localMs - epoch) / 86400000;
    }

    // DueDate = DueDateExplicit, else (StampDate ?? LetterDate) + DeadlineDays.
    // Returns an Excel serial, or "" when nothing resolves. Because an Excel date
    // serial IS a day count, "base + N days" is just baseSerial + N.
    function computeDueSerial(
        explicitIso: string, stampIso: string, letterIso: string,
        deadlineDays: number | string | null | undefined
    ): number | string {
        // 1) explicit calendar deadline wins
        const exp = isoDateToSerial(explicitIso);
        if (exp !== "") return exp;
        // 2) base date + duration
        const base = stampIso || letterIso;
        const baseSerial = isoDateToSerial(base);
        const days = Number(deadlineDays);
        if (baseSerial !== "" && isFinite(days) && days > 0) {
            return (baseSerial as number) + days;
        }
        return "";
    }

    // Canonicalize a German plate to "AREA-LL DDD" (e.g. HB-AG 454) — same rule
    // as sk_Fleet_WriteToExcel, so tblFines and tblFleet store plates identically.
    // The region boundary is the first separator; the tail splits letters/digits.
    // A fully-glued input falls back to a 2-letter second group; an unexpected
    // shape is returned upper-cased + space-collapsed, never mangled.
    function canonicalizePlate(value: string): string {
        const raw = (value || "").trim().replace(/[-_\s]+/g, " ").toUpperCase();
        const parts = raw.split(" ");
        if (parts.length >= 2) {
            const region = parts[0];
            const tail = parts.slice(1).join("").replace(/\s+/g, "");
            const m = tail.match(/^([A-ZÄÖÜ]{1,2})(\d{1,4})$/);
            if (/^[A-ZÄÖÜ]{1,3}$/.test(region) && m) {
                return `${region}-${m[1]} ${m[2]}`;
            }
        }
        const glued = raw.replace(/\s+/g, "").match(/^([A-ZÄÖÜ]{1,3})([A-ZÄÖÜ]{2})(\d{1,4})$/);
        if (glued) {
            return `${glued[1]}-${glued[2]} ${glued[3]}`;
        }
        return raw;
    }

    // True if the value looks like a real German plate: region + ID letters then
    // digits (optional E/H suffix). A leading country letter ("D HB-...") is
    // stripped first. Catches OCR garbage from blurry scans (e.g. "HB 4444 FLA")
    // without flagging valid plates. Only catches wrong SHAPES — a plausible-but-
    // wrong misread (IB444 -> IB441) still passes; only OCR confidence or a human
    // Check guards against those.
    function isPlausiblePlate(value: string): boolean {
        let s = (value || "").toUpperCase().trim();
        s = s.replace(/^D[\s-]+/, "");   // drop a leading country marker "D HB-..."
        s = s.replace(/[\s-]/g, "");     // HB-AO 632 -> HBAO632
        return /^[A-ZÄÖÜ]{2,5}\d{1,4}[EH]?$/.test(s);
    }

    // ---- derive values --------------------------------------------------------
    const refRaw = txt(d.Aktenzeichen);
    const refNorm = normalizeRef(refRaw); // dedup key
    const reasonTxt = txt(d.OffenceReason);

    // ---- dedup check ----------------------------------------------------------
    // Only when we actually have a reference. An empty Aktenzeichen means a
    // Check row (date/amount missing) — those must always be written, and would
    // otherwise all collide on an empty key. Column resolved by NAME, so the
    // column reorder does not affect this.
    if (refNorm) {
        const normCol = table.getColumnByName("NormalizedReferenceNumber");
        const bodyRange = normCol.getRangeBetweenHeaderAndTotal();
        if (bodyRange && bodyRange.getRowCount() > 0) {
            const vals = bodyRange.getValues();
            for (let i = 0; i < vals.length; i++) {
                if (normalizeRef("" + vals[i][0]) === refNorm) {
                    // status "" — no row written, the flow never reads it here
                    // (Condition Inserted is false on a duplicate).
                    return { inserted: false, reason: "duplicate", status: "" };
                }
            }
        }
    }

    const offenceSerial = isoDateToSerial(txt(d.OffenceDate));
    const letterSerial = isoDateToSerial(txt(d.LetterDate));
    const dueSerial = computeDueSerial(
        txt(d.DueDateExplicit), txt(d.StampDate), txt(d.LetterDate), d.DeadlineDays
    );
    const amountVal: number | string =
        (d.Amount === null || d.Amount === undefined) ? "" : d.Amount;
    const processedSerial = nowToSerial();

    // Unique row key for the downstream flows to address this exact row.
    const rowId = "R" + Date.now() + "-" + Math.floor(Math.random() * 1e6);

    // Store the plate in the uniform "AREA-LL DDD" form, not the raw GPT text.
    const plate = canonicalizePlate(txt(d.PlateNumber));

    // OCR-quality guard: a non-empty but implausible plate (e.g. a blurry scan
    // yielding "HB 4444 FLA") is unreliable — downgrade a Passed row to Check so
    // it still appears for the specialist to verify against the file, instead of
    // a wrong plate being trusted. The raw plate is kept in the row as a hint.
    let effectiveStatus = status;
    if (plate !== "" && !isPlausiblePlate(plate) &&
        effectiveStatus.toLowerCase() === "passed") {
        effectiveStatus = "Check";
    }

    // Normal/Fair: an unresolved DueDate silently defaults escalation to N1 ->
    // hold the row for review instead of mailing it as a fresh N1. (CN stage is
    // fixed, not DueDate-driven, so an empty DueDate is fine there.) dueSerial is
    // "" when neither an explicit deadline nor base-date+DeadlineDays resolved.
    {
        const ftLower = txt(fineType).toLowerCase();
        const isCNtype = ftLower.indexOf("cost") !== -1 || ftLower.indexOf("notification") !== -1;
        if (effectiveStatus.toLowerCase() === "passed" && !isCNtype && dueSerial === "") {
            effectiveStatus = "Check";
        }
    }

    // ---- OCRPercent (variant C) -----------------------------------------------
    // Presence of each recognized column (a valid parsed date counts, an
    // unparseable one does not). The denominator is the subset EXPECTED for this
    // fine type, so by-design-absent fields don't drag the score below 100%.
    const present: { [k: string]: boolean } = {
        plate: plate !== "",
        vendor: hasStr(txt(d.VendorName)),
        reference: refRaw.trim() !== "",
        reason: hasStr(reasonTxt),
        offenceDate: offenceSerial !== "",
        location: hasStr(txt(d.Location)),
        letterDate: letterSerial !== "",
        amount: amountVal !== "",
        authEmail: hasStr(txt(d.AuthorityEmail))
    };

    const ft = txt(fineType).toLowerCase();
    let expected: boolean[];
    if (ft.indexOf("cost") !== -1 || ft.indexOf("notification") !== -1) {
        // Cost Notification (6): offence reason, offence date and issuer email
        // are all excluded (see header note).
        expected = [present.plate, present.vendor, present.reference,
        present.location, present.letterDate, present.amount];
    } else if (ft.indexOf("fair") !== -1) {
        // Fair parking (8): AuthorityEmail excluded (private operators).
        expected = [present.plate, present.vendor, present.reference,
        present.reason, present.offenceDate, present.location,
        present.letterDate, present.amount];
    } else {
        // Normal / authority fines (9): all recognized columns expected.
        expected = [present.plate, present.vendor, present.reference,
        present.reason, present.offenceDate, present.location,
        present.letterDate, present.amount, present.authEmail];
    }

    let ocrPercent: number | string = "";
    if (expected.length > 0) {
        let got = 0;
        for (let i = 0; i < expected.length; i++) {
            if (expected[i]) got++;
        }
        ocrPercent = Math.round((got / expected.length) * 100);
    }

    // ---- build the row (exact tblFines column order, 34 columns) ---------------
    const row: (string | number)[] = [
        plate,                  // 0  PlateNumber (canonicalized "AREA-LL DDD")
        txt(fineType),          // 1  FineType
        txt(d.VendorName),      // 2  VendorName
        refRaw,                 // 3  ReferenceNumber
        reasonTxt,              // 4  OffenceReason (violation description, from OCR)
        offenceSerial,          // 5  OffenceDate
        txt(d.Location),        // 6  Location (violation location, from OCR)
        letterSerial,           // 7  LetterDate
        dueSerial,              // 8  DueDate (computed)
        amountVal,              // 9  Amount
        txt(effectiveStatus),   // 10 Status (Passed downgraded to Check on bad plate)
        ocrPercent,             // 11 OCRPercent (computed completeness %, per type)
        txt(d.AuthorityEmail),  // 12 AuthorityEmail
        "",                     // 13 Personalnummer   (match flow)
        "",                     // 14 EmployeeName      (match flow)
        "",                     // 15 EmployeeEmail     (match flow)
        "",                     // 16 EmployeeAddress   (match flow)
        "",                     // 17 DriverSource      (match flow: DUMP/Pool)
        "",                     // 18 VehicleFrom       (match flow)
        "Pending",              // 19 MailStatus (intake seeds Pending)
        "",                     // 20 LastNotifyStage   (notify flow)
        "",                     // 21 LastNotifyDate    (notify flow)
        "",                     // 22 DateSent          (notify flow)
        "",                     // 23 ConfirmedDate     (capture flow)
        "",                     // 24 PaymentDate       (capture flow)
        "",                     // 25 MatchNote         (match flow)
        "",                     // 26 ReviewNote        (filled manually)
        txt(fileName),          // 27 FileName
        refNorm,                // 28 NormalizedReferenceNumber
        rowId,                  // 29 RowId (unique key, text)
        "",                     // 30 ConfirmToken (send flow stamps the nonce)
        processedSerial,        // 31 ProcessedAt
        "",                     // 32 File (hyperlink set below)
        ""                      // 33 FolderPath (hyperlink set below)
    ];

    table.addRow(-1, row);

    // ---- format + hyperlinks on the new row -----------------------------------
    const body = table.getRangeBetweenHeaderAndTotal();
    const newRow = body.getRow(body.getRowCount() - 1);

    // Identifier columns -> Text, so an all-digit reference (e.g. 6014213508638)
    // is not turned into 6.01421E+12 and a long one doesn't lose precision. Set
    // the format FIRST, then re-write the value — after addRow the cell already
    // holds a value, so it must be replaced.
    newRow.getColumn(3).setNumberFormat("@");                 // ReferenceNumber
    newRow.getColumn(3).setValue(refRaw);
    newRow.getColumn(28).setNumberFormat("@");                // NormalizedReferenceNumber
    newRow.getColumn(28).setValue(refNorm);

    // Personalnummer / ConfirmToken are text identifiers the downstream flows
    // write and compare as exact strings — keep the cells Text even though intake
    // leaves them empty, so a later write isn't coerced to a number if the sheet
    // is ever recreated.
    newRow.getColumn(13).setNumberFormat("@");                // Personalnummer
    newRow.getColumn(13).setValue("");
    newRow.getColumn(29).setNumberFormat("@");                // RowId
    newRow.getColumn(29).setValue(rowId);
    newRow.getColumn(30).setNumberFormat("@");                // ConfirmToken
    newRow.getColumn(30).setValue("");

    // Date / number formats on the columns intake actually fills. The date cells
    // left empty here (LastNotifyDate, DateSent, ConfirmedDate, PaymentDate) stay
    // unformatted — each downstream flow sets its own format when it writes.
    newRow.getColumn(5).setNumberFormat("dd.mm.yyyy");        // OffenceDate
    newRow.getColumn(7).setNumberFormat("dd.mm.yyyy");        // LetterDate
    newRow.getColumn(8).setNumberFormat("dd.mm.yyyy");        // DueDate
    newRow.getColumn(9).setNumberFormat("0.00");              // Amount
    newRow.getColumn(31).setNumberFormat("dd.mm.yyyy hh:mm"); // ProcessedAt

    if (fileLink) {
        newRow.getColumn(32).setHyperlink({ address: fileLink, textToDisplay: "Open" });
    }
    if (folderPath) {
        newRow.getColumn(33).setHyperlink({ address: folderPath, textToDisplay: "Open folder" });
    }

    // ---- clear inherited fill before (re)highlighting --------------------------
    // Excel propagates the previous row's formatting to a newly added table row,
    // so a prior Check row's yellow would otherwise bleed down — onto filled cells
    // that must stay clean, and onto rows that are not Check at all. Clear the
    // whole new row's direct fill first (table-style banding is a separate layer
    // and is unaffected); the block below re-applies yellow only where this row
    // actually needs it.
    newRow.getFormat().getFill().clear();

    // ---- highlight empty key cells on Check rows ------------------------------
    // On a Check row, paint the key cells that came back empty so the specialist
    // sees at a glance what to complete from the file. The flagged set is per
    // fine type, so a field that is legitimately empty for a type is never
    // falsely flagged (AuthorityEmail on Fair; OffenceDate/DueDate on CN).
    // PlateNumber is never flagged (empty plate -> Failed, never a written row);
    // Amount and OffenceReason are informational and excluded by choice.
    // Reuses the presence flags already computed for OCRPercent; DueDate is
    // computed separately, so its serial is checked directly.
    if (effectiveStatus.toLowerCase() === "check") {
        const CHECK_FILL = "#FFFF00"; // bright yellow — must stand out in the sheet

        // Common to every type: VendorName, ReferenceNumber, Location, LetterDate.
        const keyCells: { col: number; filled: boolean }[] = [
            { col: 2, filled: present.vendor },     // VendorName
            { col: 3, filled: present.reference },  // ReferenceNumber
            { col: 6, filled: present.location },   // Location
            { col: 7, filled: present.letterDate }  // LetterDate
        ];

        const isCN = ft.indexOf("cost") !== -1 || ft.indexOf("notification") !== -1;
        const isFair = ft.indexOf("fair") !== -1;

        if (!isCN) {
            // Normal + Fair also expect OffenceDate and the computed DueDate.
            keyCells.push({ col: 5, filled: present.offenceDate }); // OffenceDate
            keyCells.push({ col: 8, filled: dueSerial !== "" });    // DueDate
            if (!isFair) {
                // Normal only: AuthorityEmail (Fair private operators have none).
                keyCells.push({ col: 12, filled: present.authEmail }); // AuthorityEmail
            }
        }

        for (let i = 0; i < keyCells.length; i++) {
            if (!keyCells[i].filled) {
                newRow.getColumn(keyCells[i].col).getFormat().getFill().setColor(CHECK_FILL);
            }
        }

        // Plate-caused Check: when the plate is present but malformed (the guard
        // above downgraded Passed -> Check, or the flow already sent Check with a
        // bad plate), none of the key cells above may be empty, leaving the row
        // Check with nothing painted. Flag the plate itself so the cause is
        // always visible. (An empty plate never reaches here — it is Failed.)
        if (plate !== "" && !isPlausiblePlate(plate)) {
            newRow.getColumn(0).getFormat().getFill().setColor(CHECK_FILL); // PlateNumber
        }
    }

    return { inserted: true, reason: "", status: effectiveStatus };
}