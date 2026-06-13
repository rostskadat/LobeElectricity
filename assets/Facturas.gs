// KEEP IT SYNC WITH default.yaml

const cl_bill_id = "Nº de factura"
const cl_billing_date = "Fecha de factura"
const cl_billing_period_start = "Inicio del periodo"
const cl_billing_period_end = "Fin del periodo"
const cl_billed_power = "Potencia"
const cl_billed_energy = "Energía"
const cl_gross_amount = "Importe bruto"
const cl_billed_amount = "Importe facturado"
const cl_P1 = "P1"
const cl_P2 = "P2"
const cl_P3 = "P3"
const cl_P4 = "P4"
const cl_P5 = "P5"
const cl_P6 = "P6"
const cl_CP1 = "CP1"
const cl_CP2 = "CP2"
const cl_CP3 = "CP3"
const cl_CP4 = "CP4"
const cl_CP5 = "CP5"
const cl_CP6 = "CP6"
const cl_contract_type = "Tipo de contrato"
const cl_file = "Fichero"
const cl_simulated_power = "Simulación - Potencia"
const cl_simulated_energy = "Simulación - Energía"
const cl_simulated_gross_amount = "Simulación - Importe bruto"
const cl_simulated_variation = "Simulación - Variation"

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
    if (sheet.getName().startsWith('Simulación') ||
      ['Loads'].includes(sheet.getName())) {
      // Skipping...
    } else {
      const lastRow = sheet.getLastRow();
      if (lastRow === 0) return; // Skip empty sheets

      // get a map of column indices
      const columns = getColumns(sheet)
      const columnsIndices = getColumnIndices(sheet)

      // Do some cleanup
      removeAllCharts(sheet)
      sheet.setConditionalFormatRules([]);

      // Setting format in order to allow for calculations
      setColumnFormat(sheet, columns[cl_billing_date], "date")
      setColumnFormat(sheet, columns[cl_billing_period_start], "date")
      setColumnFormat(sheet, columns[cl_billing_period_end], "date")
      setColumnFormat(sheet, columns[cl_billed_power], "currency")
      setColumnFormat(sheet, columns[cl_billed_energy], "currency")
      setColumnFormat(sheet, columns[cl_gross_amount], "currency")
      setColumnFormat(sheet, columns[cl_billed_amount], "currency")
      const energyColumns = [cl_P1, cl_P2, cl_P3, cl_P4, cl_P5, cl_P6].map(cl => columns[cl]);
      const energyIndices = [cl_P1, cl_P2, cl_P3, cl_P4, cl_P5, cl_P6].map(cl => columnsIndices[cl]);
      energyColumns.forEach((c, i) => setColumnFormat(sheet, c, "energy"))

      const powerColumns = [cl_CP1, cl_CP2, cl_CP3, cl_CP4, cl_CP5, cl_CP6].map(cl => columns[cl]);
      powerColumns.forEach((c, i) => setColumnFormat(sheet, c, "power"))

      // creating simulated columns...
      createSimulatedColumns(sheet, columns, columnsIndices, energyColumns, powerColumns)
      setColumnFormat(sheet, columns[cl_simulated_power], "currency");
      setColumnFormat(sheet, columns[cl_simulated_energy], "currency");
      setColumnFormat(sheet, columns[cl_simulated_gross_amount], "currency");
      setColumnFormat(sheet, columns[cl_simulated_variation], "percent");

      // Link with the files in Drive
      createLinks(sheet, columnsIndices[cl_bill_id], columnsIndices[cl_file])

      freezeColumn(sheet, 3)

      sortByColumn(sheet, columnsIndices[cl_billing_period_start])


      hideColumn(sheet, columnsIndices[cl_billing_date])
      hideColumn(sheet, columnsIndices[cl_P1])
      hideColumn(sheet, columnsIndices[cl_P2])
      hideColumn(sheet, columnsIndices[cl_P3])
      hideColumn(sheet, columnsIndices[cl_P4])
      hideColumn(sheet, columnsIndices[cl_P5])
      hideColumn(sheet, columnsIndices[cl_P6])
      hideColumn(sheet, columnsIndices[cl_CP1])
      hideColumn(sheet, columnsIndices[cl_CP2])
      hideColumn(sheet, columnsIndices[cl_CP3])
      hideColumn(sheet, columnsIndices[cl_CP4])
      hideColumn(sheet, columnsIndices[cl_CP5])
      hideColumn(sheet, columnsIndices[cl_CP6])
      hideColumn(sheet, columnsIndices[cl_contract_type])
      hideColumn(sheet, columnsIndices[cl_file])
      hideColumn(sheet, columnsIndices[cl_simulated_power])
      hideColumn(sheet, columnsIndices[cl_simulated_energy])

      if (sheet.getLastRow() > 5) {
        createBilledAmountChart(sheet, columns[cl_billing_period_start], columns[cl_billed_amount])
        createPowerDistributionChart(sheet, columns[cl_billing_period_start], energyIndices)
      } else {
        Logger.log("Skipping graphs for sheet '" + sheet.getName() + "': not enough data");
      }
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

function getColumns(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headerMap = {};
  headers.forEach((header, index) => {
    headerMap[header] = index2Column(index + 1);  // 1-based column index
  });
  return headerMap;

}

/**
 * Applies a specific number format to a column in a Google Sheets sheet based on the given column type.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet where the formatting will be applied.
 * @param {string} column - The column definition (i.e. C2:C)
 * @param {string} columnType - The type of formatting to apply. Supported values: "power", "currency", "number", "date".
 * @throws {Error} Throws an error if an unsupported column type is provided.
 */
function setColumnFormat(sheet, column, columnType) {
  const range = sheet.getRange(column);

  // Apply formatting based on the columnType
  switch (columnType.toLowerCase()) {
    case "energy":
      range.setNumberFormat("#,##0.00\" kWh\"");
      break;
    case "power":
      range.setNumberFormat("#,##0.000\" kWh\"");
      break;
    case "currency":
      range.setNumberFormat("#,##0.00 €");
      break;
    case "number":
      range.setNumberFormat("0.00");
      break;
    case "datetime":
      range.setNumberFormat("yyyy-MM-dd HH:00");
      break;
    case "date":
      range.setNumberFormat("yyyy-MM-dd");
      break;
    case "percent":
      range.setNumberFormat("0.00%");
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
 * Create the simulated columns...
 */
function createSimulatedColumns(sheet, columns, columnsIndices, energyColumns, powerColumns) {
  // Add a new header for the each simulated columns
  for (const columnName of [cl_simulated_power, cl_simulated_energy, cl_simulated_gross_amount, cl_simulated_variation]) {
    let column = columns[columnName]
    if (column === undefined || column === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(columnName);
      columnIndex = sheet.getLastColumn();
      columnsIndices[columnName] = columnIndex;
      columns[columnName] = index2Column(columnIndex);
    }
  }

  const columnBPS = columns[cl_billing_period_start]
  const columnBPE = columns[cl_billing_period_end]

  const columnCT = columns[cl_contract_type]
  const contractType = sheet.getRange(`${columnCT}2`).getValue();

  const nrP = contractType === "2.0TD" ? "PP2" : "PP3"
  const formulaSimulatedPower = `=ARRAYFORMULA(
    IF(${columnBPS}="";;
      LET(
        days; ${columnBPE} - ${columnBPS} + 1;
        days*${nrP}P1*${powerColumns[0]} + 
        days*${nrP}P2*${powerColumns[1]} + 
        days*${nrP}P3*${powerColumns[2]} + 
        days*${nrP}P4*${powerColumns[3]} + 
        days*${nrP}P5*${powerColumns[4]} + 
        days*${nrP}P6*${powerColumns[5]}
      )
    )
  )
  `
  const nrE = contractType === "2.0TD" ? "EP2" : "EP3"
  const formulaSimulatedEnergy = `=ARRAYFORMULA(
    ${nrE}P1*${energyColumns[0]} + 
    ${nrE}P2*${energyColumns[1]} + 
    ${nrE}P3*${energyColumns[2]} + 
    ${nrE}P4*${energyColumns[3]} + 
    ${nrE}P5*${energyColumns[4]} + 
    ${nrE}P6*${energyColumns[5]}
  )
  `

  const columnSP = columns[cl_simulated_power];
  const columnSE = columns[cl_simulated_energy];
  const formulaSimulatedGrossAmount = `=ARRAYFORMULA(
    ${columnSP} + ${columnSE}
  )
  `
  const columnSGA = columns[cl_simulated_gross_amount]
  const columnGA = columns[cl_gross_amount]
  const formulaSimulatedVariation = `=ARRAYFORMULA(
    IFERROR((${columnSGA} / ${columnGA}) - 1; "")
  )
  `
  const columnSV = columns[cl_simulated_variation]
  sheet.getRange(`${columnSP}2`).setFormula(formulaSimulatedPower);
  sheet.getRange(`${columnSE}2`).setFormula(formulaSimulatedEnergy);
  sheet.getRange(`${columnSGA}2`).setFormula(formulaSimulatedGrossAmount);
  sheet.getRange(`${columnSV}2`).setFormula(formulaSimulatedVariation);

  range = sheet.getRange(columnSV);
  range.setNote("Indica la variation del importe bruto de la factura entre la antigua tarifa y la nueva");
  setConditionalFormatting(sheet, columnSV)
}

/**
 * Add conditional formatting to the given column.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet where the formatting will be applied.
 * @param {string} column - The column specification (i.e. C2:C)
 */
function setConditionalFormatting(sheet, column) {
  const range = sheet.getRange(column);

  // Clear existing rules
  const rules = sheet.getConditionalFormatRules().filter(rule => {
    const ruleRanges = rule.getRanges();
    return !ruleRanges.some(r => r.getA1Notation() === range.getA1Notation());
  });

  // Rule for positive numbers - RED background
  const positiveRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0.01) // 0.xx% is still considered being 0
    .setBackground("#F4CCCC") // Light Red 3
    .setRanges([range])
    .build();

  // Rule for negative numbers - GREEN background
  const negativeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(-0.01)
    .setBackground("#D9EAD3") // Light Green 3
    .setRanges([range])
    .build();

  rules.push(positiveRule);
  rules.push(negativeRule);
  sheet.setConditionalFormatRules(rules);
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

function index2Column(columnIndex) {
  const columnLetter = index2Letter(columnIndex);
  const column = `${columnLetter}2:${columnLetter}`
  return column
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
 * @param {string} columnBillId - The column where the hyperlink will be set (1-based).
 * @param {string} columnPdfFile - The column containing the PDF file names (1-based).
 */
function createLinks(sheet, ciBillId, ciPdfFile) {
  if (ciBillId == null) {
    return
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 2; i <= data.length; i++) {
    cell = sheet.getRange(i, ciBillId)
    if (cell.getFormula().startsWith("=HYPERLINK(")) {
      continue;
    }
    const billId = cell.getValue();
    const pdfFile = sheet.getRange(i, ciPdfFile).getValue();
    const pdfFiles = DriveApp.getFilesByName(pdfFile);
    if (pdfFiles.hasNext()) {
      const url = pdfFiles.next().getUrl();
      cell.setFormula(`=HYPERLINK("${url}"; "${billId}")`);
    } else {
      Logger.log("No file found with the name: " + pdfFile);
    }
  }
}


/**
 * Removes all charts from the given Google Sheets sheet, except for sheets named "Hoja 1".
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet from which to remove all charts.
 * @returns {void} returns nothing.
 */
function removeAllCharts(sheet) {
  const charts = sheet.getCharts();
  if (charts.length === 0) {
    return
  }

  for (const i in charts) {
    sheet.removeChart(charts[i]);
  }
}

/**
 * Creates and inserts a column chart in the given Google Sheets sheet, visualizing billed amounts per billing period.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet where the chart will be inserted.
 * @param {string} clBillingPeriodStart - The column for the billing period start data (i.e. C2:C).
 * @param {string} clBilledAmount - The column for the billed amount data (i.e. C2:C).
 *
 * @returns {void}
 */
function createBilledAmountChart(sheet, clBillingPeriodStart, clBilledAmount) {
  const domainRange = sheet.getRange(clBillingPeriodStart); // skip headers
  const dataRange = sheet.getRange(clBilledAmount);

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
 * @param {string} ciBillingPeriodStart - The column for the billing period start (domain axis) (i.e. C2:C).
 * @param {number[]} energyIndices - Array of column indices representing power data series.
 */
function createPowerDistributionChart(sheet, clBillingPeriodStart, energyIndices) {
  const domainRange = sheet.getRange(clBillingPeriodStart); // skip headers

  const columnStart = index2Letter(energyIndices[0])
  const columnEnd = index2Letter(energyIndices[energyIndices.length - 1])
  const rowStart = 1;
  const rowEnd = sheet.getLastRow();
  const dataRange = sheet.getRange(`${columnStart}${rowStart}:${columnEnd}${rowEnd}`);

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

