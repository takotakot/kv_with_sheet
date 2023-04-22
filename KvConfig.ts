class KvConfig {
  private sheet: GoogleAppsScript.Spreadsheet.Sheet;
  private sheetNames: SheetNames = [];
  private sheetColumnNames: SheetColumnNames = [];

  constructor(sheetName: string) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    this.sheet = ss.getSheetByName(sheetName);
    this.readFromSheet();
  }

  private readFromSheet() {
    const dataRange = this.sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values.shift();
    const sheetIdIndex = headers.indexOf('sheet_id');
    const sheetNameIndex = headers.indexOf('sheet_name');
    const colIdIndex = headers.indexOf('col_id');
    const colNameIndex = headers.indexOf('col_name');
    for (const row of values) {
      if (row[sheetIdIndex] && row[sheetNameIndex]) {
        this.sheetNames.push({
          sheet_id: row[sheetIdIndex],
          sheet_name: row[sheetNameIndex],
        });
      }
      if (row[sheetIdIndex] && row[colIdIndex] && row[colNameIndex]) {
        this.sheetColumnNames.push({
          sheet_id: row[sheetIdIndex],
          col_id: row[colIdIndex],
          col_name: row[colNameIndex],
        });
      }
    }
  }

  getSheetNames(): SheetNames {
    return this.sheetNames;
  }

  getSheetColumnNames(): SheetColumnNames {
    return this.sheetColumnNames;
  }
}
