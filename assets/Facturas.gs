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
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const sheets = spreadsheet.getSheets();
  sheets.forEach(sheet => {
    Logger.log("Processing sheet '" + sheet.getName() + "' ...");
    if (['Simulación'].includes(sheet.getName())) {
      addNamedRange(spreadsheet, sheet, 'SimulatedPriceP1', 'B')
      addNamedRange(spreadsheet, sheet, 'SimulatedPriceP2', 'C')
      addNamedRange(spreadsheet, sheet, 'SimulatedPriceP3', 'D')
      addNamedRange(spreadsheet, sheet, 'SimulatedPriceP4', 'E')
      addNamedRange(spreadsheet, sheet, 'SimulatedPriceP5', 'F')
      addNamedRange(spreadsheet, sheet, 'SimulatedPriceP6', 'G')
    } else {
      const lastRow = sheet.getLastRow();
      if (lastRow === 0) return; // Skip empty sheets

      // get a map of column indices
      const columnIndices = getColumnIndices(sheet)
      const powerIndices = ["P1", "P2", "P3", "P4", "P5", "P6"].map(p => columnIndices[p]);

      resetAllCharts(sheet)

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
      const averagePrice = "AVG €/kWh"
      createAveragePriceColumn(sheet, columnIndices["Energía"], powerIndices, columnIndices[averagePrice], averagePrice)
      setColumnFormat(sheet, columnIndices[averagePrice], "power_consumption");

      // Simulate the energy cost with the prices in the different named range.
      const simulatedEnergy = "Energía € (Simulación) "
      createSimulatedEnergyColumn(sheet, powerIndices, columnIndices[simulatedEnergy], simulatedEnergy)
      setColumnFormat(sheet, columnIndices[simulatedEnergy], "currency");

      // Link with the files in Drive
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

      createBilledAmountChart(sheet, columnIndices["Inicio del periodo"], columnIndices["Importe facturado"])
      createPowerDistributionChart(sheet, columnIndices["Inicio del periodo"], powerIndices)
    }

  });
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
 * Applies a specific number format to a column in a Google Sheets sheet based on the given column type.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet where the formatting will be applied.
 * @param {number} columnIndex - The 1-based index of the column to format.
 * @param {string} columnType - The type of formatting to apply. Supported values: "power", "currency", "number", "date".
 * @throws {Error} Throws an error if an unsupported column type is provided.
 */
function setColumnFormat(sheet, columnIndex, columnType) {
  if (columnIndex == null) {
    return
  }
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
 * Adds a named range to the given spreadsheet if it does not already exist.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet to add the named range to.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet containing the range.
 * @param {string} rangeName - The name to assign to the range.
 * @param {string} columnLetter - The column letter (e.g., "A", "B") for the range starting at row 2.
 */
function addNamedRange(spreadsheet, sheet, rangeName, columnLetter) {
  if (spreadsheet.getRangeByName(rangeName) === null) {
    const range = sheet.getRange(`${sheet.getName()}!${columnLetter}2`);
    spreadsheet.setNamedRange(rangeName, range);
  }
}


/**
 * Adds a new column to the given sheet that calculates the average price based
 * on the billed amount and power columns. It gives a idea of the price of the
 * kWh independeantly from the price in any specific time period
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to which the average price column will be added.
 * @param {number} ciBilledAmount - The column index of the 'Importe facturado' (billed amount) column (1-based).
 * @param {number[]} powerColumnIndices - An array of column indices (1-based) for the power columns (P1 to P5).
 * @param {number} columnIndex - The column index (1-based) where the new column should be inserted. If undefined or -1, the column is added at the end.
 * @param {string} columnName - The name of the new column to be added as a header.
 * @throws {Error} If any required column index is not found.
 */
function createAveragePriceColumn(sheet, ciBilledAmount, powerColumnIndices, columnIndex, columnName) {
  if (ciBilledAmount === -1 || powerColumnIndices.some(i => i === -1)) {
    throw new Error("One or more required columns ('Importe facturado', P1 to P6) not found.");
  }

  // Add a new header for the calculated column
  if (columnIndex === undefined || columnIndex === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue(columnName);
    columnIndex = sheet.getLastColumn();
  }

  // Calculate the average...
  for (let i = 2; i <= sheet.getLastRow(); i++) { // skip headers
    const clBilledAmount = index2Letter(ciBilledAmount)
    const clPowers = powerColumnIndices.map(c => `${index2Letter(c)}${i}`).join("+")
    const formulaValue = `${clBilledAmount}${i}/(${clPowers})`
    const formula = `=IFERROR(IF(${formulaValue} > 0; ${formulaValue}; 0); "")`;
    sheet.getRange(i, columnIndex).setFormula(formula);
  }
}


/**
 * Adds a new column to the given sheet that calculates the simulated energy cost
 * based on provided power column indices and simulated price named range
 * (SimulatedPriceP1 to SimulatedPriceP6).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to modify.
 * @param {number[]} powerColumnIndices - Array of column indices for power values (P1 to P6).
 * @param {number} columnIndex - The index at which to insert the new calculated column. If undefined or -1, appends to the end.
 * @param {string} columnName - The header name for the new calculated column.
 * @throws {Error} If any of the required power columns (P1 to P6) are not found.
 */
function createSimulatedEnergyColumn(sheet, powerColumnIndices, columnIndex, columnName) {
  if (powerColumnIndices.some(i => i === -1)) {
    throw new Error("One or more required columns (P1 to P6) not found.");
  }

  // Add a new header for the calculated column
  if (columnIndex === undefined || columnIndex === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue(columnName);
    columnIndex = sheet.getLastColumn();
  }

  // Calculate the simulated Energy cost...
  for (let row = 2; row <= sheet.getLastRow(); row++) { // skip headers
    let fragments = []
    const clPowers = powerColumnIndices.map(ci => `${index2Letter(ci)}${row}`)
    for (let j = 0 ; j < 6; j++) { // for each Px
      fragments.push(`(SimulatedPriceP${j+1} * ${clPowers[j]})`);
    }
    sheet.getRange(row, columnIndex).setFormula("=" + fragments.join(" + "));
  }
}


/**
 * Converts a 1-based column index to its corresponding Excel-style column letter(s).
 *
 * @param {number} columnIndex - The 1-based index of the column (e.g., 1 for 'A', 27 for 'AA').
 * @returns {string} The column letter(s) corresponding to the given index.
 */
function index2Letter(columnIndex) {
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
  if (columnIndex == null) {
    return
  }
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  // Sort full data range (excluding header)
  const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
  dataRange.sort({ column: columnIndex, ascending });
}


/**
 * Creates hyperlinks in a Google Sheet by searching for PDF files in Google Drive
 * and inserting a HYPERLINK formula in the specified column.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet where links will be created.
 * @param {number} ciBillId - The column index where the hyperlink will be set (1-based).
 * @param {number} ciPdfFile - The column index containing the PDF file names (1-based).
 */
function createLinks(sheet, ciBillId, ciPdfFile) {
  if (ciBillId == null) {
    return
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 2; i <= data.length; i++) {
    const noFactura = sheet.getRange(i, ciBillId).getValue();
    const fileName = sheet.getRange(i, ciPdfFile).getValue();
    const files = DriveApp.getFilesByName(fileName);
    if (files.hasNext()) {
      const url = files.next().getUrl();
      sheet.getRange(i, ciBillId).setFormula(`=HYPERLINK("${url}"; "${noFactura}")`);
    } else {
      Logger.log("No file found with the name: " + fileName);
    }
  }
}


/**
 * Removes all charts from the given Google Sheets sheet, except for sheets named "Hoja 1".
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet from which to remove all charts.
 * @returns {void|null} Returns null if there are no charts to remove; otherwise, returns nothing.
 */
function resetAllCharts(sheet) {
  if (sheet.getName() in ["Hoja 1"]) {
    return
  }

  const charts = sheet.getCharts();
  if (charts.length === 0) {
    return null;
  }

  for (const i in charts) {
    sheet.removeChart(charts[i]);
  }

}


/**
 * Creates and inserts a column chart in the given Google Sheets sheet, visualizing billed amounts per billing period.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet where the chart will be inserted.
 * @param {number} ciBillingPeriodStart - The column index (1-based) for the billing period start data.
 * @param {number} ciBilledAmount - The column index (1-based) for the billed amount data.
 *
 * @returns {void}
 */
function createBilledAmountChart(sheet, ciBillingPeriodStart, ciBilledAmount) {
  const clDomain = index2Letter(ciBillingPeriodStart)
  const clData = index2Letter(ciBilledAmount)

  const numRows = sheet.getLastRow();
  const domainRange = sheet.getRange(`${clDomain}1:${clDomain}${numRows}`); // skip headers
  const dataRange = sheet.getRange(`${clData}1:${clData}${numRows}`);

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(domainRange)
    .addRange(dataRange)
    .setPosition(5, 4, 0, 0)
    .setNumHeaders(1)
    .setOption('isStacked', true)
    .setOption('title', 'Importe facturado')
    .setOption('legend.position', 'bottom')
    .setOption('hAxis.showTextEvery', 1)
    .build();

  sheet.insertChart(chart);
}


/**
 * Creates and inserts a stacked column chart displaying monthly power consumption
 * into the given Google Sheets sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet where the chart will be inserted.
 * @param {number} ciBillingPeriodStart - The column index for the billing period start (domain axis).
 * @param {number[]} powerIndices - Array of column indices representing power data series.
 */
function createPowerDistributionChart(sheet, ciBillingPeriodStart, powerIndices) {
  const clDomain = index2Letter(ciBillingPeriodStart)
  const clDataStart = index2Letter(powerIndices[0])
  const clDataEnd = index2Letter(powerIndices[powerIndices.length - 1])

  const numRows = sheet.getLastRow();
  const domainRange = sheet.getRange(`${clDomain}1:${clDomain}${numRows}`); // skip headers
  const dataRange = sheet.getRange(`${clDataStart}1:${clDataEnd}${numRows}`);

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(domainRange)
    .addRange(dataRange)
    .setPosition(6, 5, 0, 0)
    .setNumHeaders(1)
    .setOption('isStacked', true)
    .setOption('title', 'Consumo Mensual')
    .setOption('legend.position', 'bottom')
    .setOption('hAxis.showTextEvery', 1)
    .build();

  sheet.insertChart(chart);
}
