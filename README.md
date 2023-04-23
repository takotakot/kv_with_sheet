# Key-Value with Google Sheets

A key-value store with Google Sheets using Google Apps Script

## Claim

This program was written with the assistance of ChatGPT. The transcript of the conversation is documented in [request_for_kv.md](request_for_kv.md).

## Usage

### Push scripts

1. Create Spreadsheet
2. Open Apps Script from Extension menu
3. Copy your scriptId and create `.clasp.json`
```JSON
{
  "scriptId": "1eqopA...PgwrvXh",
  "rootDir": "."
}
```
4. Install clasp by like `npm install -g @google/clasp`
5. `npx clasp setting`, `npx clasp pull` (if you have saved some files), and `npx clasp push`

### Publish as API

Deploy your app with "Web application" and copy the URL.
The request is handled by `doPost`.

```sh
curl -v -X POST --data @./data.json -H "Content-Type: application/json" https://script.google.com/macros/s/AK.../exec
```

### Sheet structure

kv_config sheet:

||sheet_id|sheet_name|memo||sheet_id|col_id|col_name|memo|
|--|--|--|--|--|--|--|--|--|
||kv1|destination|memo||kv1|k1|col_name1|memo1|
||||||kv1|k2|col_name2||
||||||kv1|k3|col_name3||
||||||kv1|v1|col_name4||
||||||kv1|v2|col_name5|memo5|
||||||||||

destination sheet:

|col_name1|col_name2|col_name3|col_name4|col_name5|
|--|--|--|--|--|
|key1|key2|key3|a|b|

## License

These codes are licensed under CC0 or MIT.

[![CC0](http://i.creativecommons.org/p/zero/1.0/88x31.png "CC0")](http://creativecommons.org/publicdomain/zero/1.0/deed.ja)  
[MIT](https://opensource.org/licenses/MIT) (If you need, use `Copyright (c) 2023 takotakot`)

## Notice

Use at your own risk.
