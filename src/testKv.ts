function testKv(): void {
  const kvConfig = new KvConfig('kv_config');
  const sheetNames: SheetNames = kvConfig.getSheetNames();
  const sheetColumnNames: SheetColumnNames = kvConfig.getSheetColumnNames();

  Logger.log(sheetNames);
  Logger.log(sheetColumnNames);

  // Define data to be updated in the destination sheet
  const data = [
    {
      keys: {
        k1: 'key1',
        k2: 'key2',
        k3: 'key3',
      },
      values: {
        v1: 'value1',
        v2: 'value2',
      },
    },
    {
      keys: {
        k1: 'key4',
        k2: 'key5',
        k3: 'key6',
      },
      values: {
        v1: 'value3',
        v2: 'value4',
      },
    },
  ];

  const columnNames: ColumnNames = sheetColumnNames
    .filter(col => col.sheetId === 'kv1')
    .reduce((obj, { colId, colName }) => {
      obj[colId] = colName;
      return obj;
    }, {});

  // Get the destination sheet by name
  const sheet = switchSheet('destination');

  // Update the destination sheet with the data
  // updateDestinationSheet(data, sheet);
  updateDestinationSheet(sheet, columnNames, data);
}

/**
 * Update the sheet identified with "kv1" with the data.
 */
function testUpdateUsingDictionary() {
  const request: { destination: string; data: [keys: any, values: any] } = {
    destination: 'kv1',
    data: [
      {
        keys: {
          k1: 'key1',
          k2: 'key2',
          k3: 'key3',
        },
        values: {
          v1: 'value1',
          v2: 'value2',
        },
      },
      {
        keys: {
          k1: 'key4',
          k2: 'key5',
          k3: 'key6',
        },
        values: {
          v1: 'value3',
          v2: 'value4',
        },
      },
    ],
  };

  updateUsingDictionary(request);
}
