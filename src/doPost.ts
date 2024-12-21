import { updateUsingDictionary } from './kv';

/**
 * Handler for doPost request.
 *
 * @param e - The event object.
 * @returns The response object.
 */
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function doPost(
  e: GoogleAppsScript.Events.DoPost
): GoogleAppsScript.Content.TextOutput {
  const data = JSON.parse(e.postData.contents);
  // Logger.log(data);

  updateUsingDictionary(data);

  return ContentService.createTextOutput(
    JSON.stringify({ result: 'success' })
  ).setMimeType(ContentService.MimeType.JSON);
}
