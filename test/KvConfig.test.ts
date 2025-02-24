import { KvConfig } from '../src/KvConfig';

describe('KvConfig', () => {
  let sheetMock: GoogleAppsScript.Spreadsheet.Sheet;
  let spreadsheetAppMock: GoogleAppsScript.Spreadsheet.SpreadsheetApp;

  beforeEach(() => {
    sheetMock = {
      getLastRow: jest.fn().mockReturnValue(3),
      getLastColumn: jest.fn().mockReturnValue(3),
      getRange: jest.fn().mockImplementation((row, col) => ({
        getValue: jest.fn().mockReturnValue(`value${row}${col}`),
      })),
    } as unknown as GoogleAppsScript.Spreadsheet.Sheet;

    spreadsheetAppMock = {
      getActiveSpreadsheet: jest.fn().mockReturnValue({
        getSheetByName: jest.fn().mockReturnValue(sheetMock),
      }),
    } as unknown as GoogleAppsScript.Spreadsheet.SpreadsheetApp;

    global.SpreadsheetApp = spreadsheetAppMock;
  });

  it('should construct KvConfig and read from sheet', () => {
    const kvConfig = new KvConfig('TestSheet');
    expect(spreadsheetAppMock.getActiveSpreadsheet).toHaveBeenCalled();
    expect(
      spreadsheetAppMock.getActiveSpreadsheet().getSheetByName
    ).toHaveBeenCalledWith('TestSheet');
  });

  it('should split sheet into blocks', () => {
    const kvConfig = new KvConfig('TestSheet');
    const blocks = kvConfig['splitIntoBlocks']();
    expect(blocks.length).toBeGreaterThan(0);
  });

  it('should process sheet names block', () => {
    const kvConfig = new KvConfig('TestSheet');
    const rows = [
      ['sheet_id', 'sheet_name'],
      ['1', 'Sheet1'],
      ['2', 'Sheet2'],
    ];
    kvConfig['processSheetNamesBlock'](rows);
    expect(kvConfig.getSheetNames()).toEqual([
      { sheetId: '1', sheetName: 'Sheet1' },
      { sheetId: '2', sheetName: 'Sheet2' },
    ]);
  });

  it('should process sheet column names block', () => {
    const kvConfig = new KvConfig('TestSheet');
    const rows = [
      ['sheet_id', 'col_id', 'col_name'],
      ['1', 'col1', 'Column1'],
      ['1', 'col2', 'Column2'],
    ];
    kvConfig['processSheetColumnNamesBlock'](rows);
    expect(kvConfig.getSheetColumnNames()).toEqual([
      { sheetId: '1', colId: 'col1', colName: 'Column1' },
      { sheetId: '1', colId: 'col2', colName: 'Column2' },
    ]);
  });

  it('should identify sheet names block', () => {
    const kvConfig = new KvConfig('TestSheet');
    const block = [
      ['sheet_id', 'sheet_name'],
      ['1', 'Sheet1'],
    ];
    expect(kvConfig['isSheetNamesBlock'](block)).toBe(true);
  });

  it('should identify sheet column names block', () => {
    const kvConfig = new KvConfig('TestSheet');
    const block = [
      ['sheet_id', 'col_id', 'col_name'],
      ['1', 'col1', 'Column1'],
    ];
    expect(kvConfig['isSheetColumnNamesBlock'](block)).toBe(true);
  });

  describe('KvConfig', () => {
    let sheetMock: GoogleAppsScript.Spreadsheet.Sheet;
    let spreadsheetAppMock: GoogleAppsScript.Spreadsheet.SpreadsheetApp;

    beforeEach(() => {
      sheetMock = {
        getLastRow: jest.fn().mockReturnValue(3),
        getLastColumn: jest.fn().mockReturnValue(3),
        getRange: jest.fn().mockImplementation((row, col) => ({
          getValue: jest.fn().mockReturnValue(`value${row}${col}`),
        })),
      } as unknown as GoogleAppsScript.Spreadsheet.Sheet;

      spreadsheetAppMock = {
        getActiveSpreadsheet: jest.fn().mockReturnValue({
          getSheetByName: jest.fn().mockReturnValue(sheetMock),
        }),
      } as unknown as GoogleAppsScript.Spreadsheet.SpreadsheetApp;

      global.SpreadsheetApp = spreadsheetAppMock;
    });

    it('should construct KvConfig and read from sheet', () => {
      const kvConfig = new KvConfig('TestSheet');
      expect(spreadsheetAppMock.getActiveSpreadsheet).toHaveBeenCalled();
      expect(
        spreadsheetAppMock.getActiveSpreadsheet().getSheetByName
      ).toHaveBeenCalledWith('TestSheet');
    });

    it('should split sheet into blocks', () => {
      const kvConfig = new KvConfig('TestSheet');
      const blocks = kvConfig['splitIntoBlocks']();
      expect(blocks.length).toBeGreaterThan(0);
    });

    it('should process sheet names block', () => {
      const kvConfig = new KvConfig('TestSheet');
      const rows = [
        ['sheet_id', 'sheet_name'],
        ['1', 'Sheet1'],
        ['2', 'Sheet2'],
      ];
      kvConfig['processSheetNamesBlock'](rows);
      expect(kvConfig.getSheetNames()).toEqual([
        { sheetId: '1', sheetName: 'Sheet1' },
        { sheetId: '2', sheetName: 'Sheet2' },
      ]);
    });

    it('should process sheet column names block', () => {
      const kvConfig = new KvConfig('TestSheet');
      const rows = [
        ['sheet_id', 'col_id', 'col_name'],
        ['1', 'col1', 'Column1'],
        ['1', 'col2', 'Column2'],
      ];
      kvConfig['processSheetColumnNamesBlock'](rows);
      expect(kvConfig.getSheetColumnNames()).toEqual([
        { sheetId: '1', colId: 'col1', colName: 'Column1' },
        { sheetId: '1', colId: 'col2', colName: 'Column2' },
      ]);
    });

    it('should identify sheet names block', () => {
      const kvConfig = new KvConfig('TestSheet');
      const block = [
        ['sheet_id', 'sheet_name'],
        ['1', 'Sheet1'],
      ];
      expect(kvConfig['isSheetNamesBlock'](block)).toBe(true);
    });

    it('should identify sheet column names block', () => {
      const kvConfig = new KvConfig('TestSheet');
      const block = [
        ['sheet_id', 'col_id', 'col_name'],
        ['1', 'col1', 'Column1'],
      ];
      expect(kvConfig['isSheetColumnNamesBlock'](block)).toBe(true);
    });

    // it('should read from sheet and process blocks correctly', () => {
    //   const kvConfig = new KvConfig('TestSheet');
    //   const splitIntoBlocksSpy = jest.spyOn(kvConfig as any, 'splitIntoBlocks');
    //   const processSheetNamesBlockSpy = jest.spyOn(
    //     kvConfig as any,
    //     'processSheetNamesBlock'
    //   );
    //   const processSheetColumnNamesBlockSpy = jest.spyOn(
    //     kvConfig as any,
    //     'processSheetColumnNamesBlock'
    //   );

    //   kvConfig['readFromSheet']();

    //   expect(splitIntoBlocksSpy).toHaveBeenCalled();
    //   expect(processSheetNamesBlockSpy).toHaveBeenCalled();
    //   expect(processSheetColumnNamesBlockSpy).toHaveBeenCalled();
    // });

    it('should transpose a 2D array correctly', () => {
      const kvConfig = new KvConfig('TestSheet');
      const array = [
        ['a1', 'a2', 'a3'],
        ['b1', 'b2', 'b3'],
      ];
      const transposedArray = kvConfig['transpose'](array);
      expect(transposedArray).toEqual([
        ['a1', 'b1'],
        ['a2', 'b2'],
        ['a3', 'b3'],
      ]);
    });
  });
});
