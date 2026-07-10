import {KvConfig} from './KvConfig';

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function testKvConfig(): void {
  const kvConfig = new KvConfig('kv_config');
  const sheetNames = kvConfig.getSheetNames();
  const sheetColumnNames = kvConfig.getSheetColumnNames();

  Logger.log(sheetNames);
  Logger.log(sheetColumnNames);
}
