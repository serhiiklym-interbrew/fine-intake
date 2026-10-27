/**
 * sk_Eop_FindByPn
 * -----------------------------------------------------------------------------
 * Looks up a (normalized) Personalnummer in the EOP_DATA table and returns the
 * employee's name, email, address and line-manager e-mail — both as atomic
 * fields and as one pre-joined address string.
 *
 * Runs INSIDE the EOP workbook (EOP_DATA.xlsx), where EOP_DATA lives, so it
 * reads the table directly.
 *
 * Input: personalnummer — already normalized by sk_Dump_FindByPlate (upper +
 * whitespace stripped). EOP_DATA.Personalnummer was stored with the SAME rule,
 * so comparison is a plain equals.
 *
 * Address may be partly/fully empty (LEFT join from Headcount — an employee with
 * no Address row keeps empty address fields; Fair letters don't need it). The
 * joined string skips empty segments so an addressless employee yields "" rather
 * than ", , ".
 *
 * Country: EOP_DATA carries CountryName (full name, e.g. Deutschland) and
 * CountryCode (e.g. DE) as separate columns. The `country` output field keeps
 * returning the NAME (reads CountryName) so existing callers and the joined
 * address are unchanged; `countryCode` is a separate output.
 *
 * ManagerEmail: the line manager's e-mail, pre-resolved within Headcount by
 * GE_EOP_PrepareTable. Used by the N3 escalation mail for DUMP-matched drivers.
 * May be empty (no/unknown manager).
 *
 * Returns:
 *   found:false  -> PN not in EOP_DATA
 *   found:true   -> name/email/address populated (address fields may be empty)
 */

interface EopLookupResult {
    found: boolean;
    employeeName: string;
    employeeEmail: string;
    addressLine: string;
    postalCode: string;
    municipality: string;
    country: string;       // full name (reads CountryName) — unchanged behaviour
    countryCode: string;   // 3-letter code (reads CountryCode)
    birthDate: string;     // as-is from EOP_DATA.BirthDate (already DD.MM.YYYY)
    managerEmail: string;  // line manager e-mail (reads ManagerEmail) — may be empty
    employeeAddress: string; // pre-joined, empty-safe
}

function main(workbook: ExcelScript.Workbook, personalnummer: string): EopLookupResult {
    const TABLE_NAME = "EOP_DATA";

    const empty: EopLookupResult = {
        found: false, employeeName: "", employeeEmail: "",
        addressLine: "", postalCode: "", municipality: "", country: "",
        countryCode: "", birthDate: "", managerEmail: "", employeeAddress: ""
    };

    // Same normalization as sk_Dump_FindByPlate / GE_EOP_PrepareTable.
    const normPn = (v: string | number | boolean | null | undefined): string =>
        (v === null || v === undefined) ? "" : ("" + v).replace(/\s+/g, "").toUpperCase();

    const target = normPn(personalnummer);
    if (!target) return empty;

    const table = workbook.getTable(TABLE_NAME);
    if (!table) {
        throw new Error("Table '" + TABLE_NAME + "' not found in the EOP workbook.");
    }

    // Column indexes by header name (EOP_DATA fixed order, but resolve by name).
    const headerVals = table.getHeaderRowRange().getValues()[0];
    const idx: { [k: string]: number } = {};
    for (let c = 0; c < headerVals.length; c++) {
        idx["" + headerVals[c]] = c;
    }
    const need = ["Personalnummer", "EmployeeName", "EmployeeEmail",
        "AddressLine", "PostalCode", "Municipality", "CountryName", "CountryCode",
        "BirthDate", "ManagerEmail"];
    for (const n of need) {
        if (idx[n] === undefined) {
            throw new Error("Column '" + n + "' not found in EOP_DATA.");
        }
    }

    const body = table.getRangeBetweenHeaderAndTotal();
    if (!body || body.getRowCount() === 0) return empty;

    const vals = body.getTexts();
    for (let r = 0; r < vals.length; r++) {
        if (normPn(vals[r][idx["Personalnummer"]]) === target) {
            const line = ("" + vals[r][idx["AddressLine"]]).trim();
            const plz = ("" + vals[r][idx["PostalCode"]]).trim();
            const city = ("" + vals[r][idx["Municipality"]]).trim();
            const countryName = ("" + vals[r][idx["CountryName"]]).trim();
            const countryCode = ("" + vals[r][idx["CountryCode"]]).trim();
            const birthDate = ("" + vals[r][idx["BirthDate"]]).trim();
            const managerEmail = ("" + vals[r][idx["ManagerEmail"]]).trim();

            // empty-safe join: line / "plz city" / country name, commas only between present parts
            const seg2 = [plz, city].filter(s => s.length > 0).join(" ");
            const joined = [line, seg2, countryName].filter(s => s.length > 0).join(", ");

            return {
                found: true,
                employeeName: ("" + vals[r][idx["EmployeeName"]]).trim(),
                employeeEmail: ("" + vals[r][idx["EmployeeEmail"]]).trim(),
                addressLine: line,
                postalCode: plz,
                municipality: city,
                country: countryName,   // unchanged output: the full name
                countryCode: countryCode,
                birthDate: birthDate,
                managerEmail: managerEmail,
                employeeAddress: joined
            };
        }
    }

    return empty;
}