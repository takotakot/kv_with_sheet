class KvConfig {
  // This class represents the key-value configuration
  // stored in a Google Sheets spreadsheet.

  // The sheet object representing the active spreadsheet.
  private sheet: GoogleAppsScript.Spreadsheet.Sheet;

  // An array of objects representing the sheet names and their IDs.
  private sheetNames: SheetNames = [];

  // An array of objects representing the sheet column names, IDs and names.
  private sheetColumnNames: SheetColumnNames = [];

  // Constructs a KvConfig object with the specified sheet name.
  constructor(sheetName: string) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    this.sheet = ss.getSheetByName(sheetName);
    this.readFromSheet();
  }

  // Reads the configuration from the sheet.
  private readFromSheet(): void {
    // Splits the sheet into blocks, and processes each block.
    const blocks = this.splitIntoBlocks();
    for (const block of blocks) {
      if (this.isSheetNamesBlock(block)) {
        this.processSheetNamesBlock(block);
      } else if (this.isSheetColumnNamesBlock(block)) {
        this.processSheetColumnNamesBlock(block);
      }
    }
  }

  // Processes a block of rows representing sheet names and IDs.
  private processSheetNamesBlock(rows: string[][]): void {
    // The header row contains the names of the columns.
    const headerRow = rows[0];
    const sheetNames = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      // Skip empty rows
      if (row.every((cellValue) => !cellValue)) {
        continue;
      }
      // Extract the sheet name and ID from the row.
      const sheetName = row[headerRow.indexOf("sheet_name")];
      const sheetId = row[headerRow.indexOf("sheet_id")];
      sheetNames.push({ sheet_id: sheetId, sheet_name: sheetName });
    }
    this.sheetNames = sheetNames;
  }

  // Processes a block of rows representing sheet column names, IDs and names.
  private processSheetColumnNamesBlock(rows: string[][]): void {
    // The header row contains the names of the columns.
    const headerRow = rows[0];
    const sheetColumns: SheetColumnNames = [];

    for (let i = 2; i < rows.length; i++) {
      const row = rows[i];
      // Skip empty rows
      if (row.every((cellValue) => !cellValue)) {
        continue;
      }
      // Extract the sheet ID, column ID and column name from the row.
      const sheetId = row[headerRow.indexOf("sheet_id")];
      const colId = row[headerRow.indexOf("col_id")];
      const colName = row[headerRow.indexOf("col_name")];
      sheetColumns.push({ sheet_id: sheetId, col_id: colId, col_name: colName });
    }

    this.sheetColumnNames = sheetColumns;
  }

  /**
   * Split the sheet into blocks, where each block consists of consecutive columns
   * containing at least one non-empty cell. Each block is represented as a 2D array
   * of strings, where the outer array represents columns and the inner array represents
   * rows within a column.
   * @returns The blocks of non-empty columns and rows in the sheet
   */
  private splitIntoBlocks(): string[][][] {
    const blocks = [];
    const numRows = this.sheet.getLastRow();
    const numCols = this.sheet.getLastColumn();
    let currentBlock = [];
  
    for (let j = 1; j <= numCols; j++) {
      let currentColumn = [];
  
      // Get all cell values in the current column
      for (let i = 1; i <= numRows; i++) {
        const cellValue = this.sheet.getRange(i, j).getValue().toString();
        currentColumn.push(cellValue);
      }
  
      // Check if current column has at least one non-empty cell
      if (currentColumn.some((cellValue) => cellValue)) {
        // Column has at least one non-empty cell, add it to the current block
        currentBlock.push(currentColumn);
      } else if (currentBlock.length > 0) {
        // Current column has no non-empty cells, but there are non-empty columns
        // in the current block, so add the current block to the list of blocks
        blocks.push(this.transpose(currentBlock));
        currentBlock = [];
      }
    }
  
    // Add the last block if there are any columns in it
    if (currentBlock.length > 0) {
      blocks.push(this.transpose(currentBlock));
    }
  
    return blocks;
  }

  /**
   * Transpose a 2D array (i.e., rows become columns and columns become rows).
   * @param array - The array to transpose
   * @returns The transposed array
   */
  private transpose(array: any[][]): any[][] {
    return array[0].map((_, colIndex) => array.map((row) => row[colIndex]));
  }

  /**
   * Check if a block of data contains sheet names.
   * @param block - The block of data to check
   * @returns True if the block contains sheet names, false otherwise
   */
  private isSheetNamesBlock(block: string[][]): boolean {
    const expectedColumns = ["sheet_id", "sheet_name"];
    const headerRow = block[0];
    Logger.log(headerRow);

    // Check if all expected columns are present in the header row
    const includesAllColumns = expectedColumns.every((col) =>
      headerRow.includes(col)
    );

    // Return true if all expected columns are present in the header row
    return includesAllColumns;
  }

  /**
   * Check if a block of data contains column names for a sheet.
   * @param block - The block of data to check
   * @returns True if the block contains column names, false otherwise
   */
  private isSheetColumnNamesBlock(block: string[][]): boolean {
    const expectedColumns = ["sheet_id", "col_id", "col_name"];
    const headerRow = block[0];

    // Check if all expected columns are present in the header row
    const includesAllColumns = expectedColumns.every((col) =>
      headerRow.includes(col)
    );

    // Return true if all expected columns are present in the header row
    return includesAllColumns;
  }

  /**
   * Returns an object that maps sheet IDs to sheet names.
   * @returns Object that maps sheet IDs to sheet names.
   */
  getSheetNames(): SheetNames {
    return this.sheetNames;
  }

  /**
   * Returns an object that maps sheet IDs to column IDs to column names.
   * @returns Object that maps sheet IDs to column IDs to column names.
   */
  getSheetColumnNames(): SheetColumnNames {
    return this.sheetColumnNames;
  }
}
