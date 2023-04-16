function main(): void {
  const data = [
    {
      "keys": {
        "k1": "key1",
        "k2": "key2",
        "k3": "key3"
      },
      "values": {
        "v1": "value1",
        "v2": "value2"
      }
    },
    {
      "keys": {
        "k1": "key4",
        "k2": "key5",
        "k3": "key6"
      },
      "values": {
        "v1": "value3",
        "v2": "value4"
      }
    }
  ];

  const sheet = switchSheet("destination");
  updateDestinationSheet(data, sheet);
}

/**
 * Updates the destination sheet with the provided data.
 * If a row with the same values for keys exists, updates the existing row.
 * Otherwise, appends a new row with the keys and values.
 * 
 * @param data An array of objects containing "keys" and "values" properties.
 * @param destinationSheet The sheet to update.
 */
function updateDestinationSheet(data: {keys: {[key: string]: any}, values: {[key: string]: any}}[], destinationSheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  // Get the header row of the sheet
  const headerRow = destinationSheet.getRange(1, 1, 1, destinationSheet.getLastColumn()).getValues()[0];
  
  // Get the columns corresponding to the keys in the data
  const keyColumns = headerRow
    .map((header, index) => data[0].keys[header] ? index + 1 : null)
    .filter(i => i !== null) as number[];

  // Loop through the data and update the sheet
  data.forEach(datum => {
    // Find the row in the sheet corresponding to the keys in the datum
    const rowRange = getRowRangeByValues(destinationSheet, keyColumns, Object.values(datum.keys));
    
    // Get the existing values in the row, or an empty array if it doesn't exist yet
    const valuesRow = rowRange.getRow() === 0 ? Array(headerRow.length).fill("") : rowRange.getValues()[0];

    // Update the keys in the values row
    Object.entries(datum.keys).forEach(([keyHeader, key]) => {
      const keyColumn = headerRow.findIndex(header => header === keyHeader);
      if (keyColumn !== -1) {
        valuesRow[keyColumn] = key;
      }
    });

    // Update the values in the values row
    Object.entries(datum.values).forEach(([valueHeader, value]) => {
      const valueColumn = headerRow.findIndex(header => header === valueHeader);
      if (valueColumn !== -1) {
        valuesRow[valueColumn] = value;
      }
    });

    // If the row doesn't exist yet, append a new row with the keys and values
    if (rowRange.getRow() === 0) {
      destinationSheet.appendRow([...Object.values(datum.keys), ...valuesRow]);
    } else { // Otherwise, update the existing row with the new values
      rowRange.setValues([valuesRow]);
    }
  });
}

function switchSheet(sheetName: string): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }

  return sheet;
}

function getRowRangeByValues(sheet: GoogleAppsScript.Spreadsheet.Sheet, columns: number[], values: any[]): GoogleAppsScript.Spreadsheet.Range {
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (columns.every((colIndex, index) => row[colIndex - 1] === values[index])) {
      return sheet.getRange(i + 1, 1, 1, sheet.getLastColumn());
    }
  }
  return sheet.getRange(sheet.getLastRow() + 1, 1, 1, sheet.getLastColumn());
}
