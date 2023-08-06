import { updateUsingDictionary } from './kv';

function doPost(e: GoogleAppsScript.Events.DoPost) {
  const data = JSON.parse(e.postData.contents);
  // Logger.log(data);

  updateUsingDictionary(data);

  return ContentService.createTextOutput(
    JSON.stringify({ result: 'success' })
  ).setMimeType(ContentService.MimeType.JSON);
}
