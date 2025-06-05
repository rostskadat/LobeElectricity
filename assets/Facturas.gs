function getColumnIndex(sheet, columnName) {
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  // Find the column index by name (1-based)
  return headerRow.indexOf(columnName) + 1;
}

function setColumnFormatByName(sheet, columnName, columnType) {
  const columnIndex = getColumnIndex(sheet, columnName)

  if (columnIndex === 0) {
    Logger.log(`[ERROR] Column "${columnName}" not found.`);
    return
  }

  const numRows = sheet.getLastRow();
  const columnRange = sheet.getRange(2, columnIndex, numRows - 1); // Exclude header row

  // Apply formatting based on the columnType
  switch (columnType.toLowerCase()) {
    case "power":
      columnRange.setNumberFormat("#,##0.00\" kWh\"");
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

  SpreadsheetApp.flush();
}

function applyColumnFormats(sheet) {
  setColumnFormatByName(sheet, "Fecha de factura", "date");
  setColumnFormatByName(sheet, "Inicio del periodo", "date");
  setColumnFormatByName(sheet, "Fin del periodo", "date");
  setColumnFormatByName(sheet, "Potencia", "currency");
  setColumnFormatByName(sheet, "Energía", "currency");
  setColumnFormatByName(sheet, "Importe a pagar", "currency");
  setColumnFormatByName(sheet, "Importe facturado", "currency");
  setColumnFormatByName(sheet, "P1", "power");
  setColumnFormatByName(sheet, "P2", "power");
  setColumnFormatByName(sheet, "P3", "power");
  setColumnFormatByName(sheet, "P4", "power");
  setColumnFormatByName(sheet, "P5", "power");
  setColumnFormatByName(sheet, "P6", "power");
  setColumnFormatByName(sheet, "Pleno", "power");
  setColumnFormatByName(sheet, "Llano", "power");
  setColumnFormatByName(sheet, "Valle", "power");
}

function hideColumn(sheet, columnName) {
  const columnIndex = getColumnIndex(sheet, columnName)

  if (columnIndex === 0) {
    Logger.log(`[ERROR] Column "${columnName}" not found.`);
    return
  }

  const range = sheet.getRange(1, columnIndex);
  sheet.hideColumn(range);
}

function isColumnEmpty(sheet, columnName) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return true; // Only header exists

  const columnIndex = getColumnIndex(sheet, columnName)

  if (columnIndex === 0) {
    Logger.log(`[ERROR] Column "${columnName}" not found.`);
    return false;
  }

  const values = sheet.getRange(2, columnIndex, lastRow - 1).getValues(); // Exclude header
  return values.every(row => row[0] === "" || row[0] === null);
}

function sortSheetByColumn(sheet, columnName, ascending = true) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  const columnIndex = getColumnIndex(sheet, columnName)

  if (columnIndex === 0) {
    Logger.log(`[ERROR] Column "${columnName}" not found.`);
    return;
  }

  // Sort full data range (excluding header)
  const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
  dataRange.sort({ column: columnIndex, ascending });
}


function freezeFirstColumn(sheet) {
  sheet.setFrozenColumns(1);
}

function formatAllSheets() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  sheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    if (lastRow === 0) return; // Skip empty sheets

    // get a map of column indices


    applyColumnFormats(sheet)
    hideColumn(sheet, "Fecha de factura")
    freezeFirstColumn(sheet)
//    sortSheetByColumn("Inicio del periodo")
/*
    powerColumns = ["P1", "P2", "P3", "P4", "P5", "P6", "Pleno", "Llano", "Valle"]
    powerColumns.forEach(function(columnName) {
      if (isColumnEmpty(sheet, columnName)) {
        hideColumn(sheet, columnName)
      }
    });
    */
  });
}

function highlightPaidRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const numCols = sheet.getLastColumn();
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols); // All rows below header

  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$C2="Paid"')
    .setBackground("#d9ead3") // Light green
    .setRanges([dataRange])
    .build();

  const rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}

