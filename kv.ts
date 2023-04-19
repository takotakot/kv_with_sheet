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
  
  const columnNamesList = [
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

function updateDestinationSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet, columnNames: ColumnNames, data: Array<{keys: {[key: string]: any}, values: {[key: string]: any}}>): void {
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // check if all column names exist in header row
  Object.values(columnNames).forEach(columnName => {
    if (!headerRow.includes(columnName)) {
      throw new Error(`Column ${columnName} not found in sheet header row`);
    }
  });

  const firstDatum = data[0];
  const keyColumns = Object.entries(columnNames)
    .filter(([col_id, col_name]) => Object.keys(firstDatum.keys).includes(col_id))
    .map(([col_id, col_name]) => headerRow.indexOf(col_name) + 1);

  data.forEach(datum => {
    const rowRange = getRowRangeByValues(sheet, keyColumns, Object.values(datum.keys));
    const valuesRow = rowRange.getRow() === 0 ? Array(headerRow.length).fill("") : rowRange.getValues()[0];

    Object.entries(datum.keys).forEach(([keyHeader, key]) => {
      const keyColumn = headerRow.indexOf(columnNames[keyHeader]);
      if (keyColumn !== -1) {
        valuesRow[keyColumn] = key;
      }
    });

    Object.entries(datum.values).forEach(([valueHeader, value]) => {
      const valueColumn = headerRow.indexOf(columnNames[valueHeader]);
      if (valueColumn !== -1) {
        valuesRow[valueColumn] = value;
      }
    });

    if (rowRange.getRow() === 0) {
      const keys = Object.entries(datum.keys).reduce((arr, [keyHeader, key]) => {
        const keyColumn = headerRow.indexOf(columnNames[keyHeader]);
        if (keyColumn !== -1) {
          arr[keyColumn] = key;
        }
        return arr;
      }, Array(headerRow.length).fill(""));
      sheet.appendRow([...keys, ...Object.values(datum.values)]);
    } else {
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
