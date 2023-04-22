function main(): void {
  const kvConfig = new KvConfig('kv_config');
  const sheetNames = kvConfig.getSheetNames();
  const sheetColumnNames = kvConfig.getSheetColumnNames();

  Logger.log(sheetNames);
  Logger.log(sheetColumnNames);
  
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
