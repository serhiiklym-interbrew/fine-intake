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
 *   dataJson   : the GPT JSON string. 8 recognized fields + 3 DueDate inputs:
 *                VendorName, PlateNumber, Aktenzeichen, OffenceDate, Location,
 *                LetterDate, Amount, AuthorityEmail,
 *                StampDate, DeadlineDays, DueDateExplicit.
 *                (The JSON key stays "Aktenzeichen" — the document's language;
 *                 it maps to the ReferenceNumber column.)
 *
 * DueDate (col 7) is computed HERE, not by the LLM (date arithmetic in the LLM
 * is unreliable). Rule:
 *     base    = StampDate ?? LetterDate        (incoming-stamp date preferred,
 *                                               it is closest to legal "Zugang")
 *     DueDate = DueDateExplicit ?? (base + DeadlineDays)
 * If nothing resolves, DueDate is left empty.
 *
 * Returns (read by the flow to decide the log entry):
 *   { inserted: true,  reason: "" }          -> a new row was appended.
 *   { inserted: false, reason: "duplicate" } -> NormalizedReferenceNumber already
 *                                               present; no row written.
 *   The file is archived by the flow regardless, so a duplicate is logged as
 *   Skipped (RowsInserted = 0), not Failed.
 *
 * Column order written (must match tblFines exactly — 33 columns).
 * NB: this grid follows the EXECUTED row[] array below (MailStatus at 18), which
 * is the authoritative physical order — not an earlier header/handoff that had
 * MailStatus at 26. The ConfirmedBy column was removed 2026-10-29 (it only ever
 * duplicated EmployeeEmail, since the anonymous confirm link cannot verify who
 * actually clicked), shifting every column after PaymentDate down by one:
 *   0  PlateNumber        9  Status            18 MailStatus        27 NormalizedReferenceNumber
 *   1  FineType          10  OCRPercent        19 LastNotifyStage   28 RowId
 *   2  VendorName        11  AuthorityEmail    20 LastNotifyDate    29 ConfirmToken
 *   3  ReferenceNumber   12  Personalnummer    21 DateSent          30 ProcessedAt
 *   4  OffenceDate       13  EmployeeName      22 ConfirmedDate     31 File
 *   5  Location          14  EmployeeEmail     23 PaymentDate       32 FolderPath
 *   6  LetterDate        15  EmployeeAddress   24 MatchNote
 *   7  DueDate           16  DriverSource      25 ReviewNote
 *   8  Amount            17  VehicleFrom       26 FileName
 *
 *   Intake fills the recognized fields (cols 0-9 + AuthorityEmail),
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
 *   Location (col 5) is the violation location, pulled from dataJson (the 8th
 *   recognized field) — written as text, needed by the Normal/CN letters.
 *   OCRPercent (col 10) is reserved for a future OCR-recognition score; intake
 *   leaves it empty for now (no format set).
 */

interface WriteResult {
    inserted: boolean;
    reason: string;
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
        const raw = (value || "").trim().replace(/[-\s]+/g, " ").toUpperCase();
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
                    return { inserted: false, reason: "duplicate" };
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

    // ---- build the row (exact tblFines column order, 33 columns) ---------------
    const row: (string | number)[] = [
        plate,                  // 0  PlateNumber (canonicalized "AREA-LL DDD")
        txt(fineType),          // 1  FineType
        txt(d.VendorName),      // 2  VendorName
        refRaw,                 // 3  ReferenceNumber
        offenceSerial,          // 4  OffenceDate
        txt(d.Location),        // 5  Location (violation location, from OCR)
        letterSerial,           // 6  LetterDate
        dueSerial,              // 7  DueDate (computed)
        amountVal,              // 8  Amount
        txt(effectiveStatus),   // 9  Status (Passed downgraded to Check on bad plate)
        "",                     // 10 OCRPercent (reserved; empty for now)
        txt(d.AuthorityEmail),  // 11 AuthorityEmail
        "",                     // 12 Personalnummer   (match flow)
        "",                     // 13 EmployeeName      (match flow)
        "",                     // 14 EmployeeEmail     (match flow)
        "",                     // 15 EmployeeAddress   (match flow)
        "",                     // 16 DriverSource      (match flow: DUMP/Pool)
        "",                     // 17 VehicleFrom       (match flow)
        "Pending",              // 18 MailStatus (intake seeds Pending)
        "",                     // 19 LastNotifyStage   (notify flow)
        "",                     // 20 LastNotifyDate    (notify flow)
        "",                     // 21 DateSent          (notify flow)
        "",                     // 22 ConfirmedDate     (capture flow)
        "",                     // 23 PaymentDate       (capture flow)
        "",                     // 24 MatchNote         (match flow)
        "",                     // 25 ReviewNote        (filled manually)
        txt(fileName),          // 26 FileName
        refNorm,                // 27 NormalizedReferenceNumber
        rowId,                  // 28 RowId (unique key, text)
        "",                     // 29 ConfirmToken (send flow stamps the nonce)
        processedSerial,        // 30 ProcessedAt
        "",                     // 31 File (hyperlink set below)
        ""                      // 32 FolderPath (hyperlink set below)
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
    newRow.getColumn(27).setNumberFormat("@");                // NormalizedReferenceNumber
    newRow.getColumn(27).setValue(refNorm);

    // Personalnummer / ConfirmToken are text identifiers the downstream flows
    // write and compare as exact strings — keep the cells Text even though intake
    // leaves them empty, so a later write isn't coerced to a number if the sheet
    // is ever recreated.
    newRow.getColumn(12).setNumberFormat("@");                // Personalnummer
    newRow.getColumn(12).setValue("");
    newRow.getColumn(28).setNumberFormat("@");                // RowId
    newRow.getColumn(28).setValue(rowId);
    newRow.getColumn(29).setNumberFormat("@");                // ConfirmToken
    newRow.getColumn(29).setValue("");

    // Date / number formats on the columns intake actually fills. The date cells
    // left empty here (LastNotifyDate, DateSent, ConfirmedDate, PaymentDate) stay
    // unformatted — each downstream flow sets its own format when it writes.
    newRow.getColumn(4).setNumberFormat("dd.mm.yyyy");        // OffenceDate
    newRow.getColumn(6).setNumberFormat("dd.mm.yyyy");        // LetterDate
    newRow.getColumn(7).setNumberFormat("dd.mm.yyyy");        // DueDate
    newRow.getColumn(8).setNumberFormat("0.00");              // Amount
    newRow.getColumn(30).setNumberFormat("dd.mm.yyyy hh:mm"); // ProcessedAt

    if (fileLink) {
        newRow.getColumn(31).setHyperlink({ address: fileLink, textToDisplay: "Open" });
    }
    if (folderPath) {
        newRow.getColumn(32).setHyperlink({ address: folderPath, textToDisplay: "Open folder" });
    }

    return { inserted: true, reason: "" };
}