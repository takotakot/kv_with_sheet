function doPost(e) {
  var data = JSON.parse(e.postData.getDataAsString());
  // Logger.log(data);

  updateUsingDictionary(data);

  return ContentService.createTextOutput(JSON.stringify({ "result": "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}
