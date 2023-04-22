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

  const sheetNames: SheetNames = [
    {
      "sheet_id": "kv1",
      "sheet_name": "destination"
    }
  ];
  
  const columnNamesList: SheetColumnNames = [
    {
      "sheet_id": "kv1",
      "col_id": "k1",
      "col_name": "col_name1"
    },
    {
      "sheet_id": "kv1",
      "col_id": "k2",
      "col_name": "col_name2"
    },
    {
      "sheet_id": "kv1",
      "col_id": "k3",
      "col_name": "col_name3"
    },
    {
      "sheet_id": "kv1",
      "col_id": "v1",
      "col_name": "col_name4"
    },
    {
      "sheet_id": "kv1",
      "col_id": "v2",
      "col_name": "col_name5"
    }
  ];

  const columnNames: ColumnNames = columnNamesList
    .filter(col => col.sheet_id === "kv1")
    .reduce((obj, {col_id, col_name}) => {
          obj[col_id] = col_name;
          return obj;
        }, {});

  // Get the destination sheet by name
  const sheet = switchSheet("destination");

  // Update the destination sheet with the data
  // updateDestinationSheet(data, sheet);
  updateDestinationSheet(sheet, columnNames, data);
}

/**
 * Update the specified sheet with the given data.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to update.
 * @param {ColumnNames} columnNames - The column names mapping object.
 * @param {Array<{keys: {[key: string]: any}, values: {[key: string]: any}}>} data - The data to update the sheet with.
 * @throws {Error} Throws an error if any column name is not found in the sheet header row.
 */
function updateDestinationSheet(sheet, columnNames, data) {
  // Get the header row.
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Check if all column names exist in header row.
  Object.values(columnNames).forEach(columnName => {
    if (!headerRow.includes(columnName)) {
      throw new Error(`Column ${columnName} not found in sheet header row`);
    }
  });

  // Get the key columns.
  const firstDatum = data[0];
  const keyColumns = Object.entries(columnNames)
    .filter(([col_id, col_name]) => Object.keys(firstDatum.keys).includes(col_id))
    .map(([col_id, col_name]) => headerRow.indexOf(col_name) + 1);

  // Update the sheet with the data.
  data.forEach(datum => {
    // Get the row range.
    const rowRange = getRowRangeByValues(sheet, keyColumns, Object.values(datum.keys));
    // Get the values row.
    const valuesRow = rowRange.getRow() === 0 ? Array(headerRow.length).fill("") : rowRange.getValues()[0];

    // Update the keys in the values row.
    Object.entries(datum.keys).forEach(([keyHeader, key]) => {
      const keyColumn = headerRow.indexOf(columnNames[keyHeader]);
      if (keyColumn !== -1) {
        valuesRow[keyColumn] = key;
      }
    });

    // Update the values in the values row.
    Object.entries(datum.values).forEach(([valueHeader, value]) => {
      const valueColumn = headerRow.indexOf(columnNames[valueHeader]);
      if (valueColumn !== -1) {
        valuesRow[valueColumn] = value;
      }
    });

    // Append a new row if the data row was not found, otherwise update the existing row.
    if (rowRange.getRow() === 0) {
      // If the data row was not found, append a new row with the keys and values.
      const keys = Object.entries(datum.keys).reduce((arr, [keyHeader, key]) => {
        const keyColumn = headerRow.indexOf(columnNames[keyHeader]);
        if (keyColumn !== -1) {
          arr[keyColumn] = key;
        }
        return arr;
      }, Array(headerRow.length).fill(""));
      sheet.appendRow([...keys, ...Object.values(datum.values)]);
    } else {
      // If the data row was found, update the values in the row.
      rowRange.setValues([valuesRow]);
    }
  });
  // Note: This function does not return anything, instead it updates the given sheet directly.
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
