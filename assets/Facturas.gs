/**
 * Formats all non-empty sheets according to predefined rules.
 *
 * For each sheet, this function:
 * - Freezes the first column.
 * - Sets specific formats (date, currency, power) for designated columns.
 * - Hides the "Fecha de factura" column.
 * - Sorts the sheet by the "Inicio del periodo" column.
 * - Hides power columns ("P1" to "P6") if they are empty.
 *
 */
function formatAllSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const sheets = spreadsheet.getSheets();
  sheets.forEach(sheet => {
    Logger.log("Processing sheet '" + sheet.getName() + "' ...");
    if (sheet.getName().startsWith('Simulación') ||
      ['Loads'].includes(sheet.getName())) {
      // Skipping...
    } else {
      FormatElectricBillSheet.formatSheet(sheet)
    }
  });
}
