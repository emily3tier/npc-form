const https = require('https');

const CLIENT_ID = 'd6dd88e1-a49e-4350-8339-f0f42c4b3b2e';
const TENANT_ID = '667afa82-1126-4a78-8f76-0918c7f2a845';
const BASE_FOLDER = 'UPC Submissions Automated';

function getRawBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', chunk => chunks.push(chunk));
    req.on('end', () => resolve(Buffer.concat(chunks)));
    req.on('error', reject);
  });
}

function enc(s) { return encodeURIComponent(s); }

function uploadViaSession(token, fullPath, data) {
  return new Promise(async (resolve, reject) => {
    try {
      const sessionBody = JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace' } });
      const sessionResp = await new Promise((res, rej) => {
        const r = https.request({
          hostname: 'graph.microsoft.com',
          path: `/v1.0/me/drive/root:/${fullPath}:/createUploadSession`,
          method: 'POST',
          headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(sessionBody) },
        }, (response) => {
          let d = '';
          response.on('data', c => d += c);
          response.on('end', () => response.statusCode >= 400 ? rej(new Error(`Session error ${response.statusCode}: ${d}`)) : res(JSON.parse(d)));
        });
        r.on('error', rej);
        r.write(sessionBody);
        r.end();
      });

      const uploadUrl = new URL(sessionResp.uploadUrl);
      const req2 = https.request({
        hostname: uploadUrl.hostname,
        path: uploadUrl.pathname + uploadUrl.search,
        method: 'PUT',
        headers: { 'Content-Length': data.length, 'Content-Range': `bytes 0-${data.length - 1}/${data.length}` },
      }, (response) => {
        let d = '';
        response.on('data', c => d += c);
        response.on('end', () => response.statusCode >= 400 ? reject(new Error(`Upload error ${response.statusCode}: ${d}`)) : resolve(d));
      });
      req2.on('error', reject);
      req2.write(data);
      req2.end();
    } catch (e) { reject(e); }
  });
}

function uploadDirect(token, fullPath, data) {
  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'graph.microsoft.com',
      path: `/v1.0/me/drive/root:/${fullPath}:/content`,
      method: 'PUT',
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/octet-stream', 'Content-Length': data.length },
    }, (res) => {
      let d = '';
      res.on('data', c => d += c);
      res.on('end', () => res.statusCode >= 400 ? reject(new Error(`Upload error ${res.statusCode}: ${d}`)) : resolve(d));
    });
    req.on('error', reject);
    req.write(data);
    req.end();
  });
}

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  if (req.method === 'OPTIONS') return res.status(200).end();

  try {
    const rawBody = await getRawBody(req);
    const { token, folderName, fileName, data } = JSON.parse(rawBody.toString('utf8'));

    const buf = Buffer.from(data, 'base64');
    const fullPath = `${BASE_FOLDER}/${folderName}/${fileName}`.split('/').map(enc).join('/');

    if (buf.length > 4 * 1024 * 1024) {
      await uploadViaSession(token, fullPath, buf);
    } else {
      await uploadDirect(token, fullPath, buf);
    }

    res.status(200).json({ ok: true });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
};
