function testKvConfig(): void {
  const kvConfig = new KvConfig('kv_config');
  const sheetNames = kvConfig.getSheetNames();
  const sheetColumnNames = kvConfig.getSheetColumnNames();

  Logger.log(sheetNames);
  Logger.log(sheetColumnNames);
}
