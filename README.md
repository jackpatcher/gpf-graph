# gpf-graph

## Data Flow (GAS Only)

This project now uses Google Apps Script and Google Sheet as the single data source.

- Frontend -> GAS Web App
- GAS Web App -> Google Sheet (sync/read)
- Frontend saves fetched records into IndexedDB for local cache

No runtime proxy is required.

## Required Setup

1. Deploy `gas-sheet.js` and `gas-connect.js` in the same GAS project.
2. Set script property `SHEET_ID` in GAS.
3. Deploy GAS as Web App.
4. In `index.html`, set:

```javascript
const GAS_WEB_APP_URL = 'YOUR_GAS_WEB_APP_URL';
const GAS_SYNC_TOKEN = 'YOUR_SYNC_TOKEN_OR_EMPTY';
```

5. Call once:

`GET ?action=init`

## Sync Behavior

- When chart data is requested, frontend pulls incremental data from GAS and upserts to IndexedDB.
- "Sync all" in DB panel triggers GAS sync first, then updates IndexedDB.

## Optional Local Proxy

`proxy-server.js` is optional and no longer required for the primary flow.

```powershell
node proxy-server.js
```