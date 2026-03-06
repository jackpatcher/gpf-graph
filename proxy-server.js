const http = require('http');
const https = require('https');
const { URL } = require('url');

const HOST = '127.0.0.1';
const PORT = 8787;

function setCorsHeaders(res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

function sendJson(res, statusCode, payload) {
  setCorsHeaders(res);
  res.writeHead(statusCode, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify(payload));
}

function requestUpstream(month, year) {
  return new Promise((resolve, reject) => {
    const monthStr = String(month).padStart(2, '0');
    const yearStr = String(year);
    const upstreamUrl = `https://www.gpf.or.th/thai2019/About/memberfund-api.php?pageName=NAVBottom_${monthStr}_${yearStr}`;

    const req = https.get(
      upstreamUrl,
      {
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
          Accept: 'application/json,text/plain,*/*',
          Referer: 'https://www.gpf.or.th/'
        },
        timeout: 15000
      },
      (upstreamRes) => {
        let body = '';

        upstreamRes.setEncoding('utf8');
        upstreamRes.on('data', (chunk) => {
          body += chunk;
        });

        upstreamRes.on('end', () => {
          if (upstreamRes.statusCode < 200 || upstreamRes.statusCode >= 300) {
            reject(new Error(`Upstream status ${upstreamRes.statusCode}`));
            return;
          }

          resolve(body);
        });
      }
    );

    req.on('timeout', () => {
      req.destroy(new Error('Upstream timeout'));
    });

    req.on('error', reject);
  });
}

const server = http.createServer(async (req, res) => {
  try {
    if (req.method === 'OPTIONS') {
      setCorsHeaders(res);
      res.writeHead(204);
      res.end();
      return;
    }

    const url = new URL(req.url, `http://${HOST}:${PORT}`);

    if (url.pathname === '/health') {
      sendJson(res, 200, { ok: true });
      return;
    }

    if (url.pathname !== '/api/nav' || req.method !== 'GET') {
      sendJson(res, 404, { error: 'Not found' });
      return;
    }

    const month = Number(url.searchParams.get('month'));
    const year = Number(url.searchParams.get('year'));

    if (!Number.isInteger(month) || month < 1 || month > 12 || !Number.isInteger(year) || year < 1998 || year > 2100) {
      sendJson(res, 400, { error: 'Invalid month/year' });
      return;
    }

    const upstreamBody = await requestUpstream(month, year);

    setCorsHeaders(res);
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(upstreamBody);
  } catch (error) {
    sendJson(res, 502, { error: error.message || 'Proxy failed' });
  }
});

server.listen(PORT, HOST, () => {
  console.log(`GPF local proxy running at http://${HOST}:${PORT}`);
  console.log(`Try: http://${HOST}:${PORT}/api/nav?month=3&year=2026`);
});
