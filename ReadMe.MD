module/sharepoint - V8 module

Move files and invoice data into SharePoint or OneDrive with a small, workflow-friendly helper. It is designed for common automation tasks like archiving mail attachments and keeping a structured trail of extracted invoice data.

Learn more about Instago at https://instago.ai.

## Technical reference

Exports:
- `uploadSmall`
- `saveInvoices`
- `ensureAccessToken`
- `excelAppendRow`

Usage:
- `uploadSmall({ siteId, driveId, drivePath?, filename, path?, dataBase64?, accessToken, contentType? })`
  - Uploads a small file (simple upload). Provide either `path` or `dataBase64`.
- `saveInvoices({ attachments, aiResults?, sp: { siteId, driveId, drivePath?, jsonSubdir? }, auth, workflow? })`
  - Uploads PDF attachments and optional AI JSON into SharePoint, returning `{ ok, data, error }`.
- `ensureAccessToken(auth)`
  - Resolves an access token from `auth` or stored tokens. Accepts `accessToken`, or `refreshToken` plus client details.
- `excelAppendRow({ driveId, itemId?, path?, link?, table, values?, row?, auth, debug? })`
  - Appends a row to an Excel table. `values` is an array row; `row` can map by column name or letter.

Configuration:
- No required or optional `env` keys are defined in `env.json`.
- Pass credentials and target IDs via function arguments.

Examples:
```js
const sp = require('sharepoint@latest');

const token = await sp.ensureAccessToken({
  refreshToken,
  clientId,
  clientSecret,
  tenant: 'common'
});

await sp.uploadSmall({
  siteId: 'site-id',
  driveId: 'drive-id',
  drivePath: 'Invoices',
  filename: 'invoice.pdf',
  path: 'mail/123/invoice.pdf',
  accessToken: token,
  contentType: 'application/pdf'
});
```

```js
const sp = require('sharepoint@latest');

await sp.excelAppendRow({
  driveId: 'drive-id',
  path: '/Documents/Book.xlsx',
  table: 'Table1',
  row: { A: '2024-01-01', B: 'Acme', C: '120.00' },
  auth: { accessToken }
});
```
