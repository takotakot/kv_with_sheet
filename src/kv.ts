/**
 * Update ShreadSheet from Dictionary.
 *
 * @param {destination: string, data: [keys: any, values: any]} dict - The dictionary to update the sheet with.
 */
function updateUsingDictionary(dict: {
  destination: string;
  data: [keys: any, values: any];
}) {
  const kvConfig = kvConfigFactory();
  const sheetNames: SheetNames = kvConfig.getSheetNames();
  const sheetColumnNames: SheetColumnNames = kvConfig.getSheetColumnNames();

  const destinationId = dict.destination;
  const destinationSheetName = sheetNames.filter(
    sheetPointer => sheetPointer.sheetId === destinationId
  )[0].sheetName;
  const sheet = switchSheet(destinationSheetName);

  const columnNames: ColumnNames = sheetColumnNames
    .filter(col => col.sheetId === destinationId)
    .reduce((obj, { colId, colName }) => {
      obj[colId] = colName;
      return obj;
    }, {});

  updateDestinationSheet(sheet, columnNames, dict.data);
}

/**
 * KvConfig factory method.
 *
 * @returns {KvConfig} The KvConfig object.
 */
function kvConfigFactory(): KvConfig {
  return new KvConfig('kv_config');
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
  const headerRow = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  // Check if all column names exist in header row.
  Object.values(columnNames).forEach(columnName => {
    if (!headerRow.includes(columnName)) {
      throw new Error(`Column ${columnName} not found in sheet header row`);
    }
  });

  // Get the key columns.
  const firstDatum = data[0];
  const keyColumns = Object.entries(columnNames)
    .filter(([colId, colName]) => Object.keys(firstDatum.keys).includes(colId))
    .map(([colId, colName]) => headerRow.indexOf(colName) + 1);

  // Update the sheet with the data.
  data.forEach(datum => {
    // Get the row range.
    const rowRange = getRowRangeByValues(
      sheet,
      keyColumns,
      Object.values(datum.keys)
    );
    // Get the values row.
    const valuesRow =
      rowRange.getRow() === 0
        ? Array(headerRow.length).fill('')
        : rowRange.getValues()[0];

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
      const keys = Object.entries(datum.keys).reduce(
        (arr, [keyHeader, key]) => {
          const keyColumn = headerRow.indexOf(columnNames[keyHeader]);
          if (keyColumn !== -1) {
            arr[keyColumn] = key;
          }
          return arr;
        },
        Array(headerRow.length).fill('')
      );
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
function getRowRangeByValues(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  columns: number[],
  values: any[]
): GoogleAppsScript.Spreadsheet.Range {
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (columns.every((colIndex, index) => valueEquals(row[colIndex - 1], values[index]))) {
      return sheet.getRange(i + 1, 1, 1, sheet.getLastColumn());
    }
  }
  return sheet.getRange(sheet.getLastRow() + 1, 1, 1, sheet.getLastColumn());
}

/**
 * Compares two values and returns true if they are equal. 
 * If the values are Dates, they are compared as UTC timestamps.
 *
 * @param {any} lhs - The first value to compare. This value is assumed to be the value from the sheet.
 * @param {any} rhs - The second value to compare. This value is assumed to be the value from the API.
 * @returns True if the values are equal, false otherwise.
 */
function valueEquals(lhs: any, rhs: any): boolean {
  if (lhs instanceof Date) {
    const rhsDate = new Date(rhs);
    const rhsUTC = new Date(
      rhsDate.toLocaleString('en-US', {
        timeZone: 'UTC',
      }),
    );

    return lhs.getTime() == rhsUTC.getTime();
  }
  return lhs == rhs;
}
