const https = require('https');
const querystring = require('querystring');

const CLIENT_ID = 'd6dd88e1-a49e-4350-8339-f0f42c4b3b2e';
const TENANT_ID = '667afa82-1126-4a78-8f76-0918c7f2a845';

function httpsRequest(options, body) {
  return new Promise((resolve, reject) => {
    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        if (res.statusCode >= 400) reject(new Error('HTTP ' + res.statusCode + ': ' + data.substring(0, 300)));
        else resolve(data);
      });
    });
    req.on('error', reject);
    if (body) { if (Buffer.isBuffer(body)) req.write(body); else req.write(body); }
    req.end();
  });
}

async function getAccessToken(refreshToken) {
  const body = querystring.stringify({ client_id: CLIENT_ID, grant_type: 'refresh_token', refresh_token: refreshToken, scope: 'Files.ReadWrite offline_access User.Read Mail.Send' });
  const data = await httpsRequest({ hostname: 'login.microsoftonline.com', path: '/' + TENANT_ID + '/oauth2/v2.0/token', method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) } }, body);
  return JSON.parse(data).access_token;
}

function enc(s) { return encodeURIComponent(s); }

function parseMultipart(body, boundary) {
  const parts = [];
  const boundaryBuf = Buffer.from('--' + boundary);
  let start = 0;
  while (start < body.length) {
    const bIdx = body.indexOf(boundaryBuf, start);
    if (bIdx === -1) break;
    const after = bIdx + boundaryBuf.length;
    if (body[after] === 45 && body[after + 1] === 45) break; // --
    const headerStart = after + 2;
    const headerEnd = body.indexOf(Buffer.from('\r\n\r\n'), headerStart);
    if (headerEnd === -1) break;
    const headerStr = body.slice(headerStart, headerEnd).toString();
    const contentStart = headerEnd + 4;
    const nextB = body.indexOf(boundaryBuf, contentStart);
    const contentEnd = nextB === -1 ? body.length : nextB - 2;
    const nameMatch = headerStr.match(/name="([^"]+)"/);
    const filenameMatch = headerStr.match(/filename="([^"]+)"/);
    const ctMatch = headerStr.match(/Content-Type:\s*([^\r\n]+)/i);
    parts.push({ name: nameMatch ? nameMatch[1] : null, filename: filenameMatch ? filenameMatch[1] : null, contentType: ctMatch ? ctMatch[1].trim() : 'image/jpeg', data: body.slice(contentStart, contentEnd) });
    start = nextB === -1 ? body.length : nextB;
  }
  return parts;
}

// Increase body size limit to 50MB for image uploads
module.exports.config = {
  api: {
    bodyParser: false,
    sizeLimit: '50mb',
    responseLimit: false,
  }
};

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  if (req.method === 'OPTIONS') return res.status(200).end();
  try {
    const refreshToken = process.env.MS_REFRESH_TOKEN;
    if (!refreshToken) throw new Error('MS_REFRESH_TOKEN not set');
    const chunks = [];
    await new Promise((resolve, reject) => { req.on('data', c => chunks.push(c)); req.on('end', resolve); req.on('error', reject); });
    const rawBody = Buffer.concat(chunks);
    const ct = req.headers['content-type'] || '';
    const bMatch = ct.match(/boundary=([^;]+)/);
    if (!bMatch) throw new Error('No boundary');
    const parts = parseMultipart(rawBody, bMatch[1].trim());
    const folderPart = parts.find(p => p.name === 'submissionFolder');
    const submissionFolder = folderPart ? folderPart.data.toString().trim() : '';
    if (!submissionFolder) throw new Error('No submissionFolder');
    const fileParts = parts.filter(p => p.filename && p.data.length > 0);
    if (fileParts.length === 0) return res.status(200).json({ ok: true, uploaded: 0 });
    const token = await getAccessToken(refreshToken);
    const results = [];
    for (const part of fileParts) {
      const [productFolder, ...rest] = (part.name || '').split('|||');
      const fname = rest.join('|||') || part.filename;
      const fullPath = (submissionFolder + '/' + productFolder + '/' + fname).split('/').map(enc).join('/');
      try {
        await httpsRequest({ hostname: 'graph.microsoft.com', path: '/v1.0/me/drive/root:/' + fullPath + ':/content', method: 'PUT', headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': part.contentType, 'Content-Length': part.data.length } }, part.data);
        results.push({ file: fname, ok: true });
        console.log('Uploaded:', fname, 'to', submissionFolder + '/' + productFolder);
      } catch (e) {
        console.error('Failed:', fname, e.message);
        results.push({ file: fname, ok: false, error: e.message });
      }
    }
    res.status(200).json({ ok: true, uploaded: results.filter(r => r.ok).length, total: results.length, results });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
};