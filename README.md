# Fetch OneDrive Excel file as JSON

Node.js helper function that fetches one sheet from an excel file as a JSON object.

## How to use it

```js
const fetchSheet = require('@flumens/fetch-onedrive-excel');

const file = '013SAXWCB2VHYCCDY76FF3KGKPN7T55EU2';
const sheet = 'My Sheet 1';
// const drive = '/me/drive'; // optional, see https://docs.microsoft.com/en-us/graph/api/resources/onedrive?view=graph-rest-1.0#commonly-accessed-resources

const sheet = await fetchSheet({file, sheet});
```

the script requires you to authenticate to your account by setting `MS_TOKEN` env variable. One way to get your token is by visiting https://developer.microsoft.com/en-us/graph/graph-explorer/preview and opening `Access token` tab where you can copy the token.

