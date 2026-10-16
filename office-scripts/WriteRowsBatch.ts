/**
 * Batch-write rows to a named Excel table.
 * Called from Power Automate to append extracted PDF data in one operation.
 * 
 * @param workbook - injected by Office Scripts runtime
 * @param tableName - target table (e.g., "tblDamages" or "tblLog")
 * @param rowsJson - JSON string of a 2D array: [[val, val, ...], [val, val, ...]]
 * @returns count of rows inserted
 */
function main(
    workbook: ExcelScript.Workbook,
    tableName: string,
    rowsJson: string
): { inserted: number; tableName: string } {
    // Validate table exists
    const table = workbook.getTable(tableName);
    if (!table) {
        throw new Error(`Table '${tableName}' not found in workbook.`);
    }

    // Parse rows
    let rows: (string | number | boolean)[][];
    try {
        rows = JSON.parse(rowsJson);
    } catch (e) {
        throw new Error(`Invalid JSON in rowsJson: ${e}`);
    }

    if (!Array.isArray(rows) || rows.length === 0) {
        throw new Error("rowsJson must be a non-empty array of arrays.");
    }

    // Validate column count matches table
    const expectedCols = table.getHeaderRowRange().getColumnCount();
    for (let i = 0; i < rows.length; i++) {
        if (!Array.isArray(rows[i])) {
            throw new Error(`Row ${i} is not an array.`);
        }
        if (rows[i].length !== expectedCols) {
            throw new Error(
                `Row ${i} has ${rows[i].length} columns, but table '${tableName}' expects ${expectedCols}.`
            );
        }
    }

    // Batch insert at the end of the table (-1 = append)
    table.addRows(-1, rows);

    return {
        inserted: rows.length,
        tableName: tableName
    };
}
