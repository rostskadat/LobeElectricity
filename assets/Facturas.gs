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
    if (['Constants'].includes(sheet.getName()) ||
        ['Loads'].includes(sheet.getName())) {
      // Skipping...
    } else {
      const lastRow = sheet.getLastRow();
      if (lastRow === 0) return; // Skip empty sheets

      // get a map of column indices
      const columnIndices = getColumnIndices(sheet)
      const energyIndices = ["P1", "P2", "P3", "P4", "P5", "P6"].map(p => columnIndices[p]);
      const powerIndices = ["CP1", "CP2", "CP3", "CP4", "CP5", "CP6"].map(p => columnIndices[p]);

      removeAllCharts(sheet)

      // Setting format in order to allow for calculations
      setColumnFormat(sheet, columnIndices["Fecha de factura"], "date");
      setColumnFormat(sheet, columnIndices["Inicio del periodo"], "date");
      setColumnFormat(sheet, columnIndices["Fin del periodo"], "date");
      setColumnFormat(sheet, columnIndices["Potencia"], "currency");
      setColumnFormat(sheet, columnIndices["Energía"], "currency");
      setColumnFormat(sheet, columnIndices["Importe bruto"], "currency");
      setColumnFormat(sheet, columnIndices["Importe facturado"], "currency");
      setColumnFormat(sheet, columnIndices["P1"], "energy");
      setColumnFormat(sheet, columnIndices["P2"], "energy");
      setColumnFormat(sheet, columnIndices["P3"], "energy");
      setColumnFormat(sheet, columnIndices["P4"], "energy");
      setColumnFormat(sheet, columnIndices["P5"], "energy");
      setColumnFormat(sheet, columnIndices["P6"], "energy");
      setColumnFormat(sheet, columnIndices["CP1"], "power");
      setColumnFormat(sheet, columnIndices["CP2"], "power");
      setColumnFormat(sheet, columnIndices["CP3"], "power");
      setColumnFormat(sheet, columnIndices["CP4"], "power");
      setColumnFormat(sheet, columnIndices["CP5"], "power");
      setColumnFormat(sheet, columnIndices["CP6"], "power");

      // calculate the simulated gross amount
      const simulatedPower = "Simulación - Potencia"
      const simulatedEnergy = "Simulación - Energía"
      const simulatedGrossAmount = "Simulación - Importe bruto"
      createSimulatedColumns(sheet, energyIndices, powerIndices, columnIndices, simulatedPower, simulatedEnergy, simulatedGrossAmount)
      setColumnFormat(sheet, columnIndices[simulatedPower], "currency");
      setColumnFormat(sheet, columnIndices[simulatedEnergy], "currency");
      setColumnFormat(sheet, columnIndices[simulatedGrossAmount], "currency");

      const simulatedVariation = "Simulación - Variation"
      const ciGrossAmount = columnIndices["Importe bruto"]
      const ciSimulatedGrossAmount = columnIndices[simulatedGrossAmount]
      createSimulatedVariationColumn(sheet, ciGrossAmount, ciSimulatedGrossAmount, columnIndices,  simulatedVariation)
      setColumnFormat(sheet, columnIndices[simulatedVariation], "percent");

      // Link with the files in Drive
      createLinks(sheet, columnIndices["Nº de factura"], columnIndices["Fichero"])

      freezeColumn(sheet, 3)

      sortByColumn(sheet, columnIndices["Inicio del periodo"])

      hideColumn(sheet, columnIndices["Fecha de factura"])
      hideColumn(sheet, columnIndices["P1"])
      hideColumn(sheet, columnIndices["P2"])
      hideColumn(sheet, columnIndices["P3"])
      hideColumn(sheet, columnIndices["P4"])
      hideColumn(sheet, columnIndices["P5"])
      hideColumn(sheet, columnIndices["P6"])
      hideColumn(sheet, columnIndices["CP1"])
      hideColumn(sheet, columnIndices["CP2"])
      hideColumn(sheet, columnIndices["CP3"])
      hideColumn(sheet, columnIndices["CP4"])
      hideColumn(sheet, columnIndices["CP5"])
      hideColumn(sheet, columnIndices["CP6"])
      hideColumn(sheet, columnIndices["Tipo de contrato"])
      hideColumn(sheet, columnIndices["Fichero"])
      hideColumn(sheet, columnIndices[simulatedPower])
      hideColumn(sheet, columnIndices[simulatedEnergy])

      // const contractType = sheet.getRange(2, columnIndices["Tipo de contrato"]).getValue();
      // if (contractType == 2) {
      //   ["P4", "P5", "P6"].forEach(function (columnName) {
      //     hideColumn(sheet, columnIndices[columnName])
      //   });
      // }

      if (sheet.getLastRow() > 5) {
        createBilledAmountChart(sheet, columnIndices["Inicio del periodo"], columnIndices["Importe facturado"])
        createPowerDistributionChart(sheet, columnIndices["Inicio del periodo"], energyIndices)
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


/**
 * Applies a specific number format to a column in a Google Sheets sheet based on the given column type.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet where the formatting will be applied.
 * @param {number} columnIndex - The 1-based index of the column to format.
 * @param {string} columnType - The type of formatting to apply. Supported values: "power", "currency", "number", "date".
 * @throws {Error} Throws an error if an unsupported column type is provided.
 */
function setColumnFormat(sheet, columnIndex, columnType) {
  // if (columnIndex == null) {
  //   return
  // }
  const numRows = sheet.getLastRow();
  const columnRange = sheet.getRange(2, columnIndex, numRows - 1); // Exclude header row

  // Apply formatting based on the columnType
  switch (columnType.toLowerCase()) {
    case "energy":
      columnRange.setNumberFormat("#,##0.00\" kWh\"");
      break;
    case "power":
      columnRange.setNumberFormat("#,##0.000\" kWh\"");
      break;
    case "currency":
      columnRange.setNumberFormat("#,##0.00 €");
      break;
    case "number":
      columnRange.setNumberFormat("0.00");
      break;
    case "datetime":
      columnRange.setNumberFormat("yyyy-MM-dd HH:00");
      break;
    case "date":
      columnRange.setNumberFormat("yyyy-MM-dd");
      break;
    case "percent":
      columnRange.setNumberFormat("0.00%");
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
function createSimulatedColumns(sheet, energyIndices, powerIndices, columnIndices, simulatedPower, simulatedEnergy, simulatedGrossAmount) {
  if (energyIndices.some(i => i === -1)) {
    throw new Error("One or more required columns (P1 to P6) not found.");
  }
  if (powerIndices.some(i => i === -1)) {
    throw new Error("One or more required columns (CP1 to CP6) not found.");
  }

  // Add a new header for the each column
  for (const columnName of [simulatedPower, simulatedEnergy, simulatedGrossAmount]) {
    let columnIndex = columnIndices[columnName]
    if (columnIndex === undefined || columnIndex === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(columnName);
      columnIndex = sheet.getLastColumn();
      columnIndices[columnName] = columnIndex
    }
  }

  const clContractType = index2Letter(columnIndices["Tipo de contrato"])
  // we assume the contract type does not change from one bill to the next
  const contractType = sheet.getRange(`${clContractType}2`).getValue();
  const ciSimulatedPower = columnIndices[simulatedPower]
  const ciSimulatedEnergy = columnIndices[simulatedEnergy]
  const ciSimulatedGrossAmount = columnIndices[simulatedGrossAmount]

  // Calculate the simulated value ...
  for (let row = 2; row <= sheet.getLastRow(); row++) { // skip headers
    const nrP = contractType === "2.0TD" ? "PP2" : "PP3"
    const nrE = contractType === "2.0TD" ? "EP2" : "EP3"
    let powerFragments = []
    let energyFragments = []
    const clPowers = powerIndices.map(ci => ci)
    const clEnergies = energyIndices.map(ci => ci)
    for (let j = 0 ; j < 6; j++) { // for each Px
      powerFragments.push(`days*${nrP}P${j+1}*${index2Letter(clPowers[j])}${row}`);
      energyFragments.push(`${nrE}P${j+1}*${index2Letter(clEnergies[j])}${row}`);
    }
    // power: number of days * price per kWh per day
    sheet.getRange(row, ciSimulatedPower).setFormula(`=LET(
      days; D${row}-C${row}+1;
      ${powerFragments.join(" + ")}
    )`);
    // energy
    sheet.getRange(row, ciSimulatedEnergy).setFormula(`=${energyFragments.join(" + ")}`);
    // sum of both
    sheet.getRange(row, ciSimulatedGrossAmount).setFormula(`=${index2Letter(ciSimulatedPower)}${row} + ${index2Letter(ciSimulatedEnergy)}${row}`);
  }
}

function createSimulatedVariationColumn(sheet, ciGrossAmount, ciSimulatedGrossAmount, columnIndices, columnName) {
  // Add a new header for the calculated column
  let columnIndex = columnIndices[columnName]
  if (columnIndex === undefined || columnIndex === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue(columnName);
    columnIndex = sheet.getLastColumn();
    columnIndices[columnName] = columnIndex
  }

  // Calculate the variation of the Energy cost...
  for (let row = 2; row <= sheet.getLastRow(); row++) { // skip headers
    const clGrossAmount = index2Letter(ciGrossAmount)
    const clSimulatedGrossAmount = index2Letter(ciSimulatedGrossAmount)
    const formulaValue = `(${clSimulatedGrossAmount}${row} / ${clGrossAmount}${row}) - 1`
    const formula = `=IFERROR(${formulaValue}; "")`;
    sheet.getRange(row, columnIndex).setFormula(formula);
  }

  range = sheet.getRange(`${index2Letter(columnIndex)}2:${index2Letter(columnIndex)}${sheet.getLastRow()}`); // skip headers
  range.setNote("Indica la variation del importe bruto de la factura entre la antigua tarifa y la nueva");

  // apply conditional formating
  setConditionalFormatting(sheet, columnIndex)
}

function setConditionalFormatting(sheet, columnIndex) {
  const cl = index2Letter(columnIndex)
  const range = sheet.getRange(`${cl}2:${cl}`); // skip headers

  // Clear existing rules
  const rules = sheet.getConditionalFormatRules().filter(rule => {
    const ruleRanges = rule.getRanges();
    return !ruleRanges.some(r => r.getA1Notation() === range.getA1Notation());
  });

  // Rule for positive numbers - RED background
  const positiveRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground("#F4CCCC") // Light Red 3
    .setRanges([range])
    .build();

  // Rule for negative numbers - GREEN background
  const negativeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
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
 * @param {number} ciBillId - The column index where the hyperlink will be set (1-based).
 * @param {number} ciPdfFile - The column index containing the PDF file names (1-based).
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
 * @param {number[]} energyIndices - Array of column indices representing power data series.
 */
function createPowerDistributionChart(sheet, ciBillingPeriodStart, energyIndices) {
  const clDomain = index2Letter(ciBillingPeriodStart)
  const clDataStart = index2Letter(energyIndices[0])
  const clDataEnd = index2Letter(energyIndices[energyIndices.length - 1])

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

