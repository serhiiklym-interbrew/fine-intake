/**
 * sk_Fines_SetSendResult
 * -----------------------------------------------------------------------------
 * Phase 2 of the two-phase send design (pairs with sk_Fines_ClaimForSend).
 *
 * ClaimForSend writes ONLY the ConfirmToken and returns the stage/flags; it does
 * NOT commit any "sent" state. This script commits the outcome AFTER the actual
 * send action in the flow has run:
 *
 *   outcome = "sent"   -> MailStatus = "Notified"
 *                         LastNotifyStage = <stage>
 *                         LastNotifyDate  = now
 *                         DateSent        = now
 *
 *   outcome = "failed" -> MailStatus = "SendFailed"
 *                         (LastNotifyStage is intentionally LEFT UNTOUCHED, so the
 *                          next sweeper run recomputes the SAME stage and retries
 *                          naturally: rank(stage) > rank(lastStage) still passes)
 *
 * Row is located by RowId (exact text match). All columns are addressed BY NAME
 * (getColumnByName), so physical column order is irrelevant here.
 *
 * Table: tblFines in Fines_Data.xlsx
 * -----------------------------------------------------------------------------
 */

interface SetSendResultOutput {
    ok: boolean;
    outcome: string;        // echoed back: "sent" | "failed" | "row_not_found" | "bad_input"
    rowId: string;
    newStatus: string;      // MailStatus written (empty when not written)
    stage: string;          // stage written (empty for failed)
    reason: string;         // human-readable note / error
}

function main(
    workbook: ExcelScript.Workbook,
    rowId: string,
    outcome: string,
    stage: string,
    note: string
): SetSendResultOutput {

    const TABLE_NAME = "tblFines";

    const fail = (oc: string, reason: string): SetSendResultOutput => ({
        ok: false, outcome: oc, rowId: rowId || "", newStatus: "", stage: "", reason
    });

    // ---- validate input -----------------------------------------------------
    if (!rowId || rowId.trim() === "") {
        return fail("bad_input", "rowId is empty");
    }
    const oc = (outcome || "").trim().toLowerCase();
    if (oc !== "sent" && oc !== "failed") {
        return fail("bad_input", "outcome must be 'sent' or 'failed' (got '" + outcome + "')");
    }

    // ---- locate table -------------------------------------------------------
    const table = workbook.getTable(TABLE_NAME);
    if (!table) {
        return fail("bad_input", "Table " + TABLE_NAME + " not found");
    }

    // column index helper (by name) ------------------------------------------
    const headerVals = table.getHeaderRowRange().getValues()[0] as string[];
    const colIdx = (name: string): number => {
        const i = headerVals.indexOf(name);
        if (i < 0) { throw new Error("Column not found: " + name); }
        return i;
    };

    const iRowId = colIdx("RowId");
    const iStatus = colIdx("MailStatus");
    const iStage = colIdx("LastNotifyStage");
    const iNotifyDate = colIdx("LastNotifyDate");
    const iDateSent = colIdx("DateSent");

    // ---- find the row by RowId ---------------------------------------------
    const bodyRange = table.getRangeBetweenHeaderAndTotal();
    if (!bodyRange) {
        return fail("row_not_found", "Table has no data rows");
    }
    const rowCount = bodyRange.getRowCount();
    const values = bodyRange.getValues();

    const target = (rowId || "").trim();
    let found = -1;
    for (let r = 0; r < rowCount; r++) {
        const cell = values[r][iRowId];
        if (cell !== null && cell !== undefined && String(cell).trim() === target) {
            found = r;
            break;
        }
    }
    if (found < 0) {
        return fail("row_not_found", "RowId '" + target + "' not found in " + TABLE_NAME);
    }

    // ---- compute "now" as Excel serial (date-time) --------------------------
    // Excel serial epoch = 1899-12-30; 1 day = 86400000 ms.
    const now = new Date();
    const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
    const serialNow =
        (Date.UTC(
            now.getFullYear(), now.getMonth(), now.getDate(),
            now.getHours(), now.getMinutes(), now.getSeconds()
        ) - EXCEL_EPOCH) / 86400000;

    const rowRange = bodyRange.getRow(found);

    if (oc === "sent") {
        const st = (stage || "").trim();

        // MailStatus = Notified (text)
        const cStatus = rowRange.getCell(0, iStatus);
        cStatus.setNumberFormatLocal("@");
        cStatus.setValue("Notified");

        // LastNotifyStage = stage (text)
        if (st !== "") {
            const cStage = rowRange.getCell(0, iStage);
            cStage.setNumberFormatLocal("@");
            cStage.setValue(st);
        }

        // LastNotifyDate = now (date-time)
        const cND = rowRange.getCell(0, iNotifyDate);
        cND.setNumberFormatLocal("dd.mm.yyyy hh:mm");
        cND.setValue(serialNow);

        // DateSent = now (date-time)
        const cDS = rowRange.getCell(0, iDateSent);
        cDS.setNumberFormatLocal("dd.mm.yyyy hh:mm");
        cDS.setValue(serialNow);

        return {
            ok: true,
            outcome: "sent",
            rowId: target,
            newStatus: "Notified",
            stage: st,
            reason: ""
        };
    }

    // oc === "failed"
    // MailStatus = SendFailed; LastNotifyStage intentionally NOT touched so retry
    // recomputes the same stage on the next sweeper run.
    const cStatusF = rowRange.getCell(0, iStatus);
    cStatusF.setNumberFormatLocal("@");
    cStatusF.setValue("SendFailed");

    return {
        ok: true,
        outcome: "failed",
        rowId: target,
        newStatus: "SendFailed",
        stage: "",
        reason: (note || "").toString()
    };
}
