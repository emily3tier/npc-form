const https = require('https');
const querystring = require('querystring');

const CLIENT_ID = 'd6dd88e1-a49e-4350-8339-f0f42c4b3b2e';
const TENANT_ID = '667afa82-1126-4a78-8f76-0918c7f2a845';
const BASE_FOLDER = 'UPC Submissions Automated';
const TEMPLATE_NAME = 'NPC Form 2026 1.xlsx';

function httpsRequest(options, body) {
  return new Promise((resolve, reject) => {
    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        if (res.statusCode >= 400) {
          reject(new Error(`HTTP ${res.statusCode}: ${data}`));
        } else {
          resolve(data);
        }
      });
    });
    req.on('error', reject);
    if (body) req.write(body);
    req.end();
  });
}

async function getAccessToken(refreshToken) {
  const body = querystring.stringify({
    client_id: CLIENT_ID,
    grant_type: 'refresh_token',
    refresh_token: refreshToken,
    scope: 'Files.ReadWrite offline_access User.Read',
  });
  const data = await httpsRequest({
    hostname: 'login.microsoftonline.com',
    path: `/${TENANT_ID}/oauth2/v2.0/token`,
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
      'Content-Length': Buffer.byteLength(body),
    },
  }, body);
  return JSON.parse(data).access_token;
}

async function graphRequest(token, method, path, body, contentType) {
  const isBuffer = Buffer.isBuffer(body);
  const bodyStr = body && !isBuffer ? JSON.stringify(body) : body;
  const options = {
    hostname: 'graph.microsoft.com',
    path: `/v1.0${path}`,
    method,
    headers: { 'Authorization': `Bearer ${token}` },
  };
  if (body) {
    options.headers['Content-Type'] = contentType || 'application/json';
    options.headers['Content-Length'] = isBuffer ? body.length : Buffer.byteLength(bodyStr);
  }
  try {
    const data = await httpsRequest(options, bodyStr || body);
    return data ? JSON.parse(data) : {};
  } catch (e) {
    if (e.message.includes('HTTP 409')) return {};
    throw e;
  }
}

async function graphDownload(token, path) {
  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'graph.microsoft.com',
      path: `/v1.0${path}`,
      method: 'GET',
      headers: { 'Authorization': `Bearer ${token}` },
    }, (res) => {
      const chunks = [];
      res.on('data', chunk => chunks.push(chunk));
      res.on('end', () => {
        if (res.statusCode >= 400) {
          reject(new Error(`HTTP ${res.statusCode}: ${Buffer.concat(chunks).toString()}`));
        } else {
          resolve(Buffer.concat(chunks));
        }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

function encodePathSegment(s) {
  return encodeURIComponent(s).replace(/'/g, "''");
}

async function ensureFolder(token, parentPath, folderName) {
  const url = parentPath
    ? `/me/drive/root:/${parentPath.split('/').map(encodePathSegment).join('/')}:/children`
    : '/me/drive/root/children';
  await graphRequest(token, 'POST', url, { name: folderName, folder: {} });
}

async function uploadFile(token, folderPath, fileName, data) {
  const fullPath = `${folderPath}/${fileName}`.split('/').map(encodePathSegment).join('/');
  await graphRequest(token, 'PUT', `/me/drive/root:/${fullPath}:/content`, data, 'application/octet-stream');
}

async function getTemplate(token) {
  return graphDownload(token, `/me/drive/root:/${encodePathSegment(TEMPLATE_NAME)}:/content`);
}

exports.handler = async (event) => {
  const cors = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
  };

  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers: cors, body: '' };
  }

  try {
    const refreshToken = process.env.MS_REFRESH_TOKEN;
    if (!refreshToken) throw new Error('MS_REFRESH_TOKEN not set');

    const { formData, products, files } = JSON.parse(event.body);

    const token = await getAccessToken(refreshToken);
    await ensureFolder(token, '', BASE_FOLDER);

    for (const p of products) {
      await ensureFolder(token, BASE_FOLDER, p.folderName);
    }

    for (const f of files) {
      const buf = Buffer.from(f.data, 'base64');
      await uploadFile(token, `${BASE_FOLDER}/${f.folderName}`, f.fileName, buf);
    }

    // Get template from OneDrive
    const templateBuf = await getTemplate(token);

    // We'll use the template as-is and populate via a simple XML patch
    // For now return the template with a note - full openpyxl-style editing needs a build step
    // Instead, build a simple xlsx from scratch using the existing template bytes
    const clientSafe = (formData.clientName || 'client').replace(/[^a-z0-9]/gi, '_');
    const date = new Date().toISOString().slice(0, 10);
    const excelName = `NPC_Form_${clientSafe}_${date}.xlsx`;

    // Upload template copy to OneDrive (it will be pre-populated by user or we just upload it)
    await uploadFile(token, BASE_FOLDER, excelName, templateBuf);

    return {
      statusCode: 200,
      headers: cors,
      body: JSON.stringify({
        excel: templateBuf.toString('base64'),
        excelName,
      }),
    };
  } catch (e) {
    console.error(e);
    return {
      statusCode: 500,
      headers: cors,
      body: JSON.stringify({ error: e.message }),
    };
  }
};
