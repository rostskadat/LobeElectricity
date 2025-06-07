/**
 * Formats all non-empty sheets in the active Google Spreadsheet according to predefined rules.
 *
 * For each sheet, this function:
 * - Freezes the first column.
 * - Sets specific formats (date, currency, power) for designated columns.
 * - Hides the "Fecha de factura" column.
 * - Sorts the sheet by the "Inicio del periodo" column.
 * - Hides power columns ("P1" to "P6") if they are empty.
 *
 * @see freezeFirstColumn
 * @see setColumnFormat
 * @see hideColumn
 * @see sortSheetByColumn
 */
function formatAllSheets() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  sheets.forEach(sheet => {

    Logger.log("Processing sheet '" + sheet.getName() + "' ...");

    const lastRow = sheet.getLastRow();
    if (lastRow === 0) return; // Skip empty sheets

    // get a map of column indices
    const columnIndices = getColumnIndices(sheet)

    // Setting format in order to allow for calculations
    setColumnFormat(sheet, columnIndices["Fecha de factura"], "date");
    setColumnFormat(sheet, columnIndices["Inicio del periodo"], "date");
    setColumnFormat(sheet, columnIndices["Fin del periodo"], "date");
    setColumnFormat(sheet, columnIndices["Potencia"], "currency");
    setColumnFormat(sheet, columnIndices["Energía"], "currency");
    setColumnFormat(sheet, columnIndices["Importe a pagar"], "currency");
    setColumnFormat(sheet, columnIndices["Importe facturado"], "currency");
    setColumnFormat(sheet, columnIndices["P1"], "power");
    setColumnFormat(sheet, columnIndices["P2"], "power");
    setColumnFormat(sheet, columnIndices["P3"], "power");
    setColumnFormat(sheet, columnIndices["P4"], "power");
    setColumnFormat(sheet, columnIndices["P5"], "power");
    setColumnFormat(sheet, columnIndices["P6"], "power");

    // calculate the kWh average price
    const powerIndices = ["P1", "P2", "P3", "P4", "P5", "P6"].map(p => columnIndices[p]);
    const averagePrice = "€/kWh"
    createAveragePriceColumn(sheet, columnIndices["Energía"], powerIndices, columnIndices[averagePrice] , averagePrice)
    setColumnFormat(sheet, columnIndices[averagePrice], "power_consumption");

    createLinks(sheet, columnIndices["Nº de factura"], columnIndices["Fichero"])

    freezeColumn(sheet, 3)


    sortByColumn(sheet, columnIndices["Inicio del periodo"])

    hideColumn(sheet, columnIndices["Fecha de factura"])
    hideColumn(sheet, columnIndices["Rectificación"])
    hideColumn(sheet, columnIndices["Rectificación"])
    hideColumn(sheet, columnIndices["Fichero"])

    powerColumns = ["P1", "P2", "P3", "P4", "P5", "P6"]
    powerColumns.forEach(function (columnName) {
      if (isColumnEmpty(sheet, columnIndices[columnName])) {
        hideColumn(sheet, columnIndices[columnName])
      }
    });
  });
}

function createAveragePriceColumn(sheet, billedAmountColumnIndex, powerColumnIndices, columnIndex, columnName) {
  if (billedAmountColumnIndex === -1 || powerColumnIndices.some(i => i === -1)) {
    throw new Error("One or more required columns ('Importe facturado', P1 to P5) not found.");
  }

  // Add a new header for the calculated column
  if (columnIndex === undefined || columnIndex === -1) {
    sheet.getRange(1, sheet.getLastColumn()+1).setValue(columnName);
    columnIndex = sheet.getLastColumn();
  }

  // Calculate the average...
  for (let i = 2; i <= sheet.getLastRow(); i++) { // skip headers
    const clBilledAmount = columnToLetter(billedAmountColumnIndex)
    const clPowers = powerColumnIndices.map(c => `${columnToLetter(c)}${i}`).join("+")
    const formulaValue = `${clBilledAmount}${i}/(${clPowers})`
    const formula = `=IFERROR(IF(${formulaValue} > 0; ${formulaValue}; 0); "")`;
    sheet.getRange(i, columnIndex).setFormula(formula);
  }
}

function columnToLetter(columnIndex) {
  let columnLetter = '';
  while (columnIndex > 0) {
    let temp = (columnIndex - 1) % 26;
    columnLetter = String.fromCharCode(temp + 65) + columnLetter;
    columnIndex = Math.floor((columnIndex - 1) / 26);
  }
  return columnLetter;
}

/**
 * Freezes the specified number of columns in the given Google Sheets sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet where columns will be frozen.
 * @param {number} columnIndex - The number of columns to freeze, starting from the left.
 */
function freezeColumn(sheet, columnIndex) {
  sheet.setFrozenColumns(columnIndex);
}

/**
 * Applies a specific number format to a column in a Google Sheets sheet based on the given column type.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet where the formatting will be applied.
 * @param {number} columnIndex - The 1-based index of the column to format.
 * @param {string} columnType - The type of formatting to apply. Supported values: "power", "currency", "number", "date".
 * @throws {Error} Throws an error if an unsupported column type is provided.
 */
function setColumnFormat(sheet, columnIndex, columnType) {
  const numRows = sheet.getLastRow();
  const columnRange = sheet.getRange(2, columnIndex, numRows - 1); // Exclude header row

  // Apply formatting based on the columnType
  switch (columnType.toLowerCase()) {
    case "power":
      columnRange.setNumberFormat("#,##0.00\" kWh\"");
      break;
    case "power_consumption":
      columnRange.setNumberFormat("#,##0.00000\" €/kWh\"");
      break;
    case "currency":
      columnRange.setNumberFormat("#,##0.00 €");
      break;
    case "number":
      columnRange.setNumberFormat("0.00");
      break;
    case "date":
        columnRange.setNumberFormat("yyyy-MM-dd");
      break;
    default:
      throw new Error(`Unsupported column type: ${columnType}`);
  }

  try {
    SpreadsheetApp.flush();
  } catch (e) {
    // when the sheet has been converted to a formatted table,
    // this fails
  }

}

/**
 * Hides a specific column in the given Google Sheets sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet where the column will be hidden.
 * @param {number} columnIndex - The 1-based index of the column to hide.
 */
function hideColumn(sheet, columnIndex) {
  const range = sheet.getRange(1, columnIndex);
  sheet.hideColumn(range);
}

/**
 * Checks if a specified column in a Google Sheets sheet is empty (excluding the header row).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to check.
 * @param {number} columnIndex - The 1-based index of the column to check.
 * @returns {boolean} True if the column is empty (excluding the header), false otherwise.
 */
function isColumnEmpty(sheet, columnIndex) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return true; // Only header exists
  const values = sheet.getRange(2, columnIndex, lastRow - 1).getValues(); // Exclude header
  return values.every(row => row[0] === "" || row[0] === null);
}

/**
 * Sorts the data in a given Google Sheets sheet by a specified column.
 * The header row (first row) is excluded from sorting.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to sort.
 * @param {number} columnIndex - The 1-based index of the column to sort by.
 * @param {boolean} [ascending=true] - Whether to sort in ascending order (default is true).
 */
function sortByColumn(sheet, columnIndex, ascending = true) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  // Sort full data range (excluding header)
  const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
  dataRange.sort({ column: columnIndex, ascending });
}

/**
 * Returns a mapping of column header names to their 1-based column indices for the given sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object to extract headers from.
 * @returns {Object.<string, number>} An object mapping header names to their corresponding 1-based column indices.
 */
function getColumnIndices(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headerMap = {};
  headers.forEach((header, index) => {
    headerMap[header] = index + 1;  // 1-based column index
  });
  return headerMap;
}

/**
 * Creates hyperlinks in a Google Sheet by searching for PDF files in Google Drive
 * and inserting a HYPERLINK formula in the specified column.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet where links will be created.
 * @param {number} noFacturaColumnIndex - The column index where the hyperlink will be set (1-based).
 * @param {number} pdfFileColumnIndex - The column index containing the PDF file names (1-based).
 */
function createLinks(sheet, noFacturaColumnIndex, pdfFileColumnIndex) {
  const data = sheet.getDataRange().getValues();
  for (let i = 2; i <= data.length; i++) {
    const noFactura = sheet.getRange(i, noFacturaColumnIndex).getValue();
    const fileName = sheet.getRange(i, pdfFileColumnIndex).getValue();
    const files = DriveApp.getFilesByName(fileName);
    if (files.hasNext()) {
      const url = files.next().getUrl();
      sheet.getRange(i, noFacturaColumnIndex).setFormula(`=HYPERLINK("${url}"; "${noFactura}")`);
    } else {
      Logger.log("No file found with the name: " + fileName);
    }
  }
}


/** **********************************************************************************
 *
 * NOT USED
 *
 */
// function addNoteToColumn() {
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
//   const column = 3; // Column C (1 = A, 2 = B, 3 = C)
//   const numRows = sheet.getMaxRows();

//   const range = sheet.getRange(1, column, numRows, 1);

//   // Set the same note for all cells in the column
//   const note = "This is a note for column C.";
//   const notesArray = Array.from({ length: numRows }, () => [note]);

//   range.setNotes(notesArray);
// }

// function highlightPaidRows() {
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
//   const numCols = sheet.getLastColumn();
//   const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols); // All rows below header

//   const rule = SpreadsheetApp.newConditionalFormatRule()
//     .whenFormulaSatisfied('=$C2="Paid"')
//     .setBackground("#d9ead3") // Light green
//     .setRanges([dataRange])
//     .build();

//   const rules = sheet.getConditionalFormatRules();
//   rules.push(rule);
//   sheet.setConditionalFormatRules(rules);
// }


// function addNumberValidation() {
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
//   const range = sheet.getRange("C2:C");

//   const rule = SpreadsheetApp.newDataValidation()
//     .requireNumber()
//     .setAllowInvalid(false)
//     .build();

//   range.setDataValidation(rule);
// }

