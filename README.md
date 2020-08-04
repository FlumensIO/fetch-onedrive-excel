# Fetch OneDrive Excel file as JSON

Node.js helper function that fetches one sheet from an excel file as a JSON object.

## How to use it

```
const fetchSheet = require('@flumens/fetch-onedrive-excel');

const fileId = '013SAXWCB2VHYCCDY76FF3KGKPN7T55EU2';
const sheetName = 'My Sheet 1';

const sheet = await fetchSheet(fileId, sheetName);
```

the script requires you to authenticate to your account by setting `MS_TOKEN` env variable. One way to get your token is by visiting https://developer.microsoft.com/en-us/graph/graph-explorer/preview and opening `Access token` tab where you can copy the token.

