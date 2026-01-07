# @instago/sharepoint

Upload files to SharePoint/OneDrive via Microsoft Graph and append rows to Excel tables. Handy for saving mail attachments (and their AI-parsed JSON) into a document library.

- Version: 1.0.0
- Entry: index.js
- Exports: `{ uploadSmall, saveInvoices, ensureAccessToken, excelAppendRow }`
- Dependencies: `graph@1.0.0`, `b64@1.0.0`, `json@1.0.0`, `path@1.0.0`, `fmt@1.0.0`, `sys.storage.get`, `sys.http.fetch`
- Auth: via Graph (`accessToken` or refresh/client credentials)

## API
- `uploadSmall({ siteId, driveId, drivePath?, filename, path?, dataBase64?, accessToken, contentType?, workflow? })` – simple upload (<4MB recommended).
- `saveInvoices({ attachments, aiResults?, sp: { siteId, driveId, drivePath?, jsonSubdir? }, auth, workflow? })` – upload PDF attachments and optional AI JSON to SharePoint.
- `excelAppendRow({ driveId, itemId|path|link, table, values|row, auth, debug? })` – append a row to an Excel table.
- `ensureAccessToken(auth)` – resolve an access token using Graph helper and optional stored tokens.

## Example
```js
const sp = require('sharepoint@1.0.0');
const token = await sp.ensureAccessToken({ refreshToken, clientId, clientSecret, tenant: 'common' });

await sp.uploadSmall({
  siteId: 'site-id',
  driveId: 'drive-id',
  drivePath: 'Invoices',
  filename: 'invoice.pdf',
  path: 'mail/123/invoice.pdf',
  accessToken: token
});
```

## Notes
- Requires Graph permissions to write to the target drive/site.
- Uses module storage for reading data to upload and optionally persisting tokens/JSON; pass `workflow: true` to opt out.
