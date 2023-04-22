class KvConfig {
  private sheet: GoogleAppsScript.Spreadsheet.Sheet;
  private sheetNames: SheetNames = [];
  private sheetColumnNames: SheetColumnNames = [];

  constructor(sheetName: string) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    this.sheet = ss.getSheetByName(sheetName);
    this.readFromSheet();
  }

  private readFromSheet(): void {
    const blocks = this.splitIntoBlocks();
    for (const block of blocks) {
      if (this.isSheetNamesBlock(block)) {
        this.processSheetNamesBlock(block);
      } else if (this.isSheetColumnNamesBlock(block)) {
        this.processSheetColumnNamesBlock(block);
      }
    }
  }
  
  private processSheetNamesBlock(rows: string[][]): void {
    const headerRow = rows[0];
    const sheetNames = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const sheetName = row[headerRow.indexOf("sheet_name")];
      const sheetId = row[headerRow.indexOf("sheet_id")];
      sheetNames.push({ sheet_id: sheetId, sheet_name: sheetName });
    }
    this.sheetNames = sheetNames;
  }
  
  private processSheetColumnNamesBlock(rows: string[][]): void {
    const headerRow = rows[0];
    const sheetColumns: SheetColumnNames = [];
  
    for (let i = 2; i < rows.length; i++) {
      const row = rows[i];
      const sheetId = row[headerRow.indexOf("sheet_id")];
      const colId = row[headerRow.indexOf("col_id")];
      const colName = row[headerRow.indexOf("col_name")];
      sheetColumns.push({ sheet_id: sheetId, col_id: colId, col_name: colName });
    }
  
    this.sheetColumnNames = sheetColumns;
  }

  private splitIntoBlocks(): string[][][] {
    const blocks = [];
    const numRows = this.sheet.getLastRow();
    const numCols = this.sheet.getLastColumn();
    let currentBlock = [];
    for (let j = 1; j <= numCols; j++) {
      const column = [];
      for (let i = 1; i <= numRows; i++) {
        const cellValue = this.sheet.getRange(i, j).getValue().toString();
        column.push(cellValue);
      }
      if (column.some((cellValue) => cellValue)) {
        // Column has at least one non-empty cell
        currentBlock.push(column);
      } else if (currentBlock.length > 0) {
        // End of block
        blocks.push(currentBlock);
        currentBlock = [];
      }
    }
    if (currentBlock.length > 0) {
      // Add last block
      blocks.push(currentBlock);
    }
    return blocks;
  }

  private isSheetNamesBlock(block: string[][]): boolean {
    const expectedColumns = ["sheet_id", "sheet_name"];
    const headerRow = block[0];
  
    // 期待するカラム名がすべて含まれているか確認する
    const includesAllColumns = expectedColumns.every((col) =>
      headerRow.includes(col)
    );
  
    // 期待するカラム名がすべて含まれている場合は true を返す
    return includesAllColumns;
  }

  private isSheetColumnNamesBlock(block: string[][]): boolean {
    const expectedColumns = ["sheet_id", "col_id", "col_name"];
    const headerRow = block[0];
  
    // 期待するカラム名がすべて含まれているか確認する
    const includesAllColumns = expectedColumns.every((col) =>
      headerRow.includes(col)
    );
  
    // 期待するカラム名がすべて含まれている場合は true を返す
    return includesAllColumns;
  }

  getSheetNames(): SheetNames {
    return this.sheetNames;
  }

  getSheetColumnNames(): SheetColumnNames {
    return this.sheetColumnNames;
  }
}
