const SHEET_NAME = "";              // "" = first worksheet
const LAST_COLUMN = "BF";           // last column to include in the table
const TABLE_NAME = "DUMP_DATA";     // MUST match the fines flows
const MIN_EXPECTED_ROWS = 500;      // alarm threshold - a real dump has ~3400

interface PrepareResult {
    tableName: string;
    dataRows: number;
    rangeAddress: string;
    deletedTables: number;
}

function main(workbook: ExcelScript.Workbook): PrepareResult {

    const sheet = SHEET_NAME
        ? workbook.getWorksheet(SHEET_NAME)
        : workbook.getWorksheets()[0];

    if (!sheet) {
        throw new Error("Worksheet not found: " + SHEET_NAME);
    }

    // 1) drop every existing table so the name DUMP_DATA is always free
    const existing = workbook.getTables();
    const deletedTables = existing.length;
    for (const t of existing) {
        t.delete();
    }

    // 2) clear the worksheet-level AutoFilter (clearCriteria throws when none exists)
    try {
        const af = sheet.getAutoFilter();
        af.clearCriteria();
        af.remove();
    } catch {
        /* no autofilter on this sheet */
    }

    // 3) find the real extent of the data
    const used = sheet.getUsedRange(true); // values only
    if (!used) {
        throw new Error("DUMP_INPUT.xlsx is empty after the copy - aborting.");
    }

    const lastRow = used.getRowIndex() + used.getRowCount(); // 1-based
    const dataRows = lastRow - 1;                            // minus header

    if (dataRows < MIN_EXPECTED_ROWS) {
        throw new Error(
            "Only " + dataRows + " data rows found (expected at least " +
            MIN_EXPECTED_ROWS + "). DumpFile.xlsx looks truncated - check the source."
        );
    }

    const address = "A1:" + LAST_COLUMN + lastRow;

    // 4) unhide anything hidden manually in the source file
    try {
        sheet.getRange(address).setRowHidden(false);
    } catch {
        /* not supported / nothing hidden */
    }

    // 5) create the table with the stable name
    const table = workbook.addTable(sheet.getRange(address), true);
    table.setName(TABLE_NAME);

    return {
        tableName: table.getName(),
        dataRows: dataRows,
        rangeAddress: address,
        deletedTables: deletedTables
    };
}