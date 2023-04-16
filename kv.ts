function main(): void {
  // Define data to be updated in the destination sheet
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

  // Get the destination sheet by name
  const sheet = switchSheet("destination");

  // Update the destination sheet with the data
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

/**
 * Gets a Sheet object by its name
 * 
 * @param sheetName Name of the sheet to get
 * @returns Sheet object with the given name
 * @throws Error if the sheet with the given name does not exist
 */
function switchSheet(sheetName: string): GoogleAppsScript.Spreadsheet.Sheet {
  // Get the active spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get the sheet by its name
  const sheet = ss.getSheetByName(sheetName);

  // Throw an error if the sheet does not exist
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }

  // Return the sheet object
  return sheet;
}

/**
 * Get the range of a row that matches given values in specified columns.
 * If no such row is found, returns a range for a new row at the bottom of the sheet.
 * @param sheet - The sheet to search
 * @param columns - The 1-indexed columns to match values
 * @param values - The values to match in corresponding columns
 * @returns The range of the matched row or a new row at the bottom of the sheet
 */
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
