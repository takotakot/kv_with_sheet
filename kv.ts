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

function updateDestinationSheet(data: {keys: {[key: string]: any}, values: {[key: string]: any}}[], sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const keyColumns = headers.slice(0, 3);
  const valueColumns = headers.slice(3);

  for (const row of data) {
    const rowValues = keyColumns.map(key => row.keys[key]);
    const rowRange = getRowRangeByValues(sheet, rowValues);

    if (rowRange) {
      const valueRange = sheet.getRange(rowRange.getRow(), 4, 1, 2);
      const valueArray = valueColumns.map(column => row.values[column]);
      valueRange.setValues([valueArray]);
    } else {
      const newRow = keyColumns.concat(valueColumns.map(column => row.values[column]));
      sheet.appendRow(newRow);
    }
  }
}

function switchSheet(sheetName: string): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }

  return sheet;
}

function getRowRangeByValues(sheet: GoogleAppsScript.Spreadsheet.Sheet, values: any[]): GoogleAppsScript.Spreadsheet.Range | null {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.join() === values.join()) {
      return sheet.getRange(i + 1, 1, 1, row.length);
    }
  }
  return null;
}
