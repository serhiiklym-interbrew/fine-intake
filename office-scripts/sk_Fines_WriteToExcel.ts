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
 * Column order written (must match tblFines exactly):
 *   FileName, FineType, VendorName, PlateNumber, ReferenceNumber, OffenceDate,
 *   LetterDate, Amount, AuthorityEmail, Status, NormalizedReferenceNumber,
 *   ReviewNote, ProcessedAt, File, FolderPath
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

    // ---- build the row (exact tblFines column order) --------------------------
    const row: (string | number)[] = [
        txt(fileName),          // 0  FileName
        txt(fineType),          // 1  FineType
        txt(d.VendorName),      // 2  VendorName
        txt(d.PlateNumber),     // 3  PlateNumber
        refRaw,                 // 4  ReferenceNumber
        offenceSerial,          // 5  OffenceDate
        letterSerial,           // 6  LetterDate
        amountVal,              // 7  Amount
        txt(d.AuthorityEmail),  // 8  AuthorityEmail
        txt(status),            // 9  Status
        refNorm,                // 10 NormalizedReferenceNumber
        "",                     // 11 ReviewNote (filled manually)
        processedSerial,        // 12 ProcessedAt
        "",                     // 13 File (hyperlink set below)
        ""                      // 14 FolderPath (hyperlink set below)
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