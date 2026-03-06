# gpf-graph

## Run locally (with CORS-safe proxy)

1. Start static file server as usual (for example Live Server at `http://127.0.0.1:5500`).
2. Start local proxy in another terminal:

	```powershell
	node proxy-server.js
	```

3. Open the web page and fetch data normally.

The app will call `http://127.0.0.1:8787/api/nav?month=MM&year=YYYY` first in local development, then fallback to public proxies and direct endpoint if needed.