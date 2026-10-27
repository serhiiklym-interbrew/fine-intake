/**
 * sk_Fines_WriteToExcel
 * -----------------------------------------------------------------------------
 * Appends one recognized German fine to the tblFines table in Fines_Data.xlsx,
 * unless its reference number already exists (dedup) — in which case no row is
 * written and the caller is told it was a duplicate.
 *
 * Called from the "GE_FineIntake_Normal_Auto" Power Automate flow via Run script.
 *
 * Parameters (fileName FIRST, per the house clone convention):
 *   fileName   : human file name as archived (flow-supplied).
 *   fineType   : "Normal" or "Fair" (derived from the source folder by the flow).
 *   status     : "Passed" or "Check" ("Failed" rows are never written).
 *   fileLink   : SharePoint "Link to item" of the archived PDF (may be empty).
 *   folderPath : SharePoint link to the submission folder (may be empty).
 *   dataJson   : the GPT JSON string with the 7 extracted fields:
 *                VendorName, PlateNumber, Aktenzeichen, OffenceDate,
 *                LetterDate, Amount, AuthorityEmail.
 *                (The JSON key stays "Aktenzeichen" — the document's language;
 *                 it maps to the ReferenceNumber column.)
 *
 * Returns (read by the flow to decide the log entry):
 *   { inserted: true,  reason: "" }          -> a new row was appended.
 *   { inserted: false, reason: "duplicate" } -> NormalizedReferenceNumber already
 *                                               present; no row written.
 *   The file is archived by the flow regardless, so a duplicate is logged as
 *   Skipped (RowsInserted = 0), not Failed.
 *
 * Column order written (must match tblFines exactly — 26 columns):
 *   FileName, FineType, VendorName, PlateNumber, ReferenceNumber, OffenceDate,
 *   LetterDate, Amount, AuthorityEmail, Status, NormalizedReferenceNumber,
 *   ReviewNote, ProcessedAt, File, FolderPath,
 *   Personalnummer, EmployeeName, EmployeeEmail, EmployeeAddress,
 *   MailStatus, MatchNote, DateSent, RowId,
 *   ConfirmToken, ConfirmedDate, ConfirmedBy
 *
 *   Intake writes MailStatus = "Pending" for every row it appends (Passed and
 *   Check alike); the six send-state columns (Personalnummer, Employee*,
 *   MatchNote, DateSent) are left empty for the match/dispatch flow to fill.
 *   Intake also generates RowId — a unique text key the match/dispatch flow
 *   addresses this exact row by. FileName is not unique, ProcessedAt resolves
 *   only to the second, and Check rows carry no reference, so none of those is a
 *   safe key; RowId is "R" + Date.now() (real ms) + "-" + a random suffix.
 *
 *   The three payment-confirmation columns (ConfirmToken, ConfirmedDate,
 *   ConfirmedBy) are likewise left empty here — the send flow stamps a one-time
 *   ConfirmToken (nonce) when it posts the driver's Teams card, and the capture
 *   flow fills ConfirmedDate/ConfirmedBy when the driver confirms. Intake never
 *   touches them beyond reserving the (Text-formatted) cells.
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
        LetterDate?: string | null;
        Amount?: number | null;
        AuthorityEmail?: string | null;
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
    // otherwise all collide on an empty key.
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
    const amountVal: number | string =
        (d.Amount === null || d.Amount === undefined) ? "" : d.Amount;
    const processedSerial = nowToSerial();

    // Unique row key for the match/dispatch flow to address this exact row.
    // Date.now() gives real milliseconds (nowToSerial() only resolves to the
    // second), and the random suffix covers two rows written in the same ms.
    const rowId = "R" + Date.now() + "-" + Math.floor(Math.random() * 1e6);

    // Store the plate in the uniform "AREA-LL DDD" form, not the raw GPT text.
    const plate = canonicalizePlate(txt(d.PlateNumber));

    // OCR-quality guard: a non-empty but implausible plate (e.g. a blurry scan
    // yielding "HB 4444 FLA") is unreliable — downgrade a Passed row to Check so
    // it still appears for the specialist to verify against the file, instead of
    // a wrong plate being trusted. The raw plate is kept in the row as a hint.
    // ("Check" capitalised to match the tblFines Status convention.)
    let effectiveStatus = status;
    if (plate !== "" && !isPlausiblePlate(plate) &&
        effectiveStatus.toLowerCase() === "passed") {
        effectiveStatus = "Check";
    }

    // ---- build the row (exact tblFines column order, 26 columns) --------------
    const row: (string | number)[] = [
        txt(fileName),          // 0  FileName
        txt(fineType),          // 1  FineType
        txt(d.VendorName),      // 2  VendorName
        plate,                  // 3  PlateNumber (canonicalized "AREA-LL DDD")
        refRaw,                 // 4  ReferenceNumber
        offenceSerial,          // 5  OffenceDate
        letterSerial,           // 6  LetterDate
        amountVal,              // 7  Amount
        txt(d.AuthorityEmail),  // 8  AuthorityEmail
        txt(effectiveStatus),   // 9  Status (Passed downgraded to Check on bad plate)
        refNorm,                // 10 NormalizedReferenceNumber
        "",                     // 11 ReviewNote (filled manually)
        processedSerial,        // 12 ProcessedAt
        "",                     // 13 File (hyperlink set below)
        "",                     // 14 FolderPath (hyperlink set below)
        "",                     // 15 Personalnummer (match flow)
        "",                     // 16 EmployeeName (match flow)
        "",                     // 17 EmployeeEmail (match flow)
        "",                     // 18 EmployeeAddress (match flow)
        "Pending",              // 19 MailStatus (intake seeds Pending)
        "",                     // 20 MatchNote (match flow)
        "",                     // 21 DateSent (dispatch flow)
        rowId,                  // 22 RowId (unique key, text)
        "",                     // 23 ConfirmToken (send flow stamps the nonce)
        "",                     // 24 ConfirmedDate (capture flow fills on confirm)
        ""                      // 25 ConfirmedBy (capture flow fills on confirm)
    ];

    table.addRow(-1, row);

    // ---- format + hyperlinks on the new row -----------------------------------
    const body = table.getRangeBetweenHeaderAndTotal();
    const newRow = body.getRow(body.getRowCount() - 1);

    // Reference columns are identifiers, not numbers: force Text so an all-digit
    // case number (e.g. 6014213508638) is not turned into 6.01421E+12 and a long
    // one doesn't lose precision. Set the format FIRST, then re-write the value —
    // after addRow the cell already holds a number, so it must be replaced.
    newRow.getColumn(4).setNumberFormat("@");                 // ReferenceNumber
    newRow.getColumn(4).setValue(refRaw);
    newRow.getColumn(10).setNumberFormat("@");                // NormalizedReferenceNumber
    newRow.getColumn(10).setValue(refNorm);

    // Personalnummer is a text identifier too (leading zeros / long) — keep the
    // column Text even though intake leaves it empty, so the match flow's later
    // write can't be coerced to a number if the sheet is ever recreated.
    newRow.getColumn(15).setNumberFormat("@");                // Personalnummer
    newRow.getColumn(15).setValue("");

    // RowId is an identifier, not a number — keep the column Text so the key the
    // match flow searches by always compares as the exact string written here.
    newRow.getColumn(22).setNumberFormat("@");                // RowId
    newRow.getColumn(22).setValue(rowId);

    // ConfirmToken is a GUID nonce the capture flow compares as an exact string,
    // and ConfirmedBy is an email/name — keep both columns Text even though intake
    // leaves them empty, so the send/capture flow's later write isn't coerced to a
    // number if the sheet is ever recreated. ConfirmedDate stays unformatted here;
    // the capture flow sets its dd.mm.yyyy hh:mm format when it writes the value.
    newRow.getColumn(23).setNumberFormat("@");                // ConfirmToken
    newRow.getColumn(23).setValue("");
    newRow.getColumn(25).setNumberFormat("@");                // ConfirmedBy
    newRow.getColumn(25).setValue("");

    newRow.getColumn(5).setNumberFormat("dd.mm.yyyy");        // OffenceDate
    newRow.getColumn(6).setNumberFormat("dd.mm.yyyy");        // LetterDate
    newRow.getColumn(7).setNumberFormat("0.00");              // Amount
    newRow.getColumn(12).setNumberFormat("dd.mm.yyyy hh:mm"); // ProcessedAt

    if (fileLink) {
        newRow.getColumn(13).setHyperlink({ address: fileLink, textToDisplay: "Open" });
    }
    if (folderPath) {
        newRow.getColumn(14).setHyperlink({ address: folderPath, textToDisplay: "Open folder" });
    }

    return { inserted: true, reason: "" };
}