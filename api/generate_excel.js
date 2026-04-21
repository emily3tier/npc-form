const https = require('https');
const querystring = require('querystring');

const CLIENT_ID = 'd6dd88e1-a49e-4350-8339-f0f42c4b3b2e';
const TENANT_ID = '667afa82-1126-4a78-8f76-0918c7f2a845';
const BASE_FOLDER = 'UPC Submissions Automated';
const TEMPLATE_NAME = 'NPC Form 2026 1.xlsx';

function getRawBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', chunk => chunks.push(chunk));
    req.on('end', () => resolve(Buffer.concat(chunks)));
    req.on('error', reject);
  });
}

function httpsRequest(options, body) {
  return new Promise((resolve, reject) => {
    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        if (res.statusCode >= 400) reject(new Error('HTTP ' + res.statusCode + ': ' + data));
        else resolve(data);
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
    path: '/' + TENANT_ID + '/oauth2/v2.0/token',
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) },
  }, body);
  return JSON.parse(data).access_token;
}

async function graphRequest(token, method, path, body, contentType) {
  const isBuffer = Buffer.isBuffer(body);
  const bodyStr = body && !isBuffer ? JSON.stringify(body) : body;
  const options = {
    hostname: 'graph.microsoft.com',
    path: '/v1.0' + path,
    method,
    headers: { 'Authorization': 'Bearer ' + token },
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
      path: '/v1.0' + path,
      method: 'GET',
      headers: { 'Authorization': 'Bearer ' + token },
    }, (res) => {
      const chunks = [];
      res.on('data', chunk => chunks.push(chunk));
      res.on('end', () => {
        if (res.statusCode >= 400) reject(new Error('HTTP ' + res.statusCode + ': ' + Buffer.concat(chunks).toString()));
        else resolve(Buffer.concat(chunks));
      });
    });
    req.on('error', reject);
    req.end();
  });
}

function enc(s) { return encodeURIComponent(s); }

async function ensureFolder(token, parentPath, folderName) {
  const url = parentPath
    ? '/me/drive/root:/' + parentPath.split('/').map(enc).join('/') + ':/children'
    : '/me/drive/root/children';
  await graphRequest(token, 'POST', url, { name: folderName, folder: {} });
}

async function uploadFile(token, folderPath, fileName, data) {
  const fullPath = (folderPath + '/' + fileName).split('/').map(enc).join('/');
  await graphRequest(token, 'PUT', '/me/drive/root:/' + fullPath + ':/content', data, 'application/octet-stream');
}

async function populateExcel(templateBuf, formData, products) {
  const ExcelJS = require('exceljs');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuf);
  const sheet = workbook.getWorksheet('New Product Coding Form') || workbook.worksheets[0];
  
  // Log all cells with values to find correct positions
  const cellMap = {};
  sheet.eachRow((row, rowNum) => {
    row.eachCell((cell, colNum) => {
      if (cell.value) cellMap[cell.address] = String(cell.value).substring(0, 50);
    });
  });
  console.log('Cell map:', JSON.stringify(cellMap));

  sheet.getCell('B4').value = formData.fromName || '';
  sheet.getCell('D4').value = formData.clientName || '';
  sheet.getCell('B5').value = formData.email || '';
  sheet.getCell('D5').value = formData.ownProduct || 'Yes';
  const coding = formData.codingOption || 'Code Immediately';
  sheet.getCell('A7').value = coding === 'Code Immediately' ? 'X' : '';
  sheet.getCell('A8').value = coding === 'Delay until Sales' ? 'X' : '';
  sheet.getCell('A9').value = coding === 'Code by Saturday Date' ? 'X' : '';
  if (coding === 'Code by Saturday Date' && formData.saturdayDate) {
    sheet.getCell('C9').value = formData.saturdayDate;
  }
  // Try both possible cell locations for container type/material
  sheet.getCell('D10').value = formData.containerType || '';
  sheet.getCell('E10').value = formData.containerMaterial || '';
  
  const addl = formData.additionalInfo || '';
  products.forEach((p, i) => {
    const row = 12 + i;
    sheet.getCell('A' + row).value = p.upc || '';
    sheet.getCell('B' + row).value = p.asin || '';
    sheet.getCell('C' + row).value = p.costco || '';
    sheet.getCell('D' + row).value = p.name + (addl ? ' | ' + addl : '');
  });
  const out = await workbook.xlsx.writeBuffer();
  return Buffer.from(out);
}

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  if (req.method === 'OPTIONS') return res.status(200).end();
  try {
    const refreshToken = process.env.MS_REFRESH_TOKEN;
    if (!refreshToken) throw new Error('MS_REFRESH_TOKEN not set');
    const rawBody = await getRawBody(req);
    const { formData, products } = JSON.parse(rawBody.toString('utf8'));
    const token = await getAccessToken(refreshToken);
    await ensureFolder(token, '', BASE_FOLDER);
    for (const p of products) {
      await ensureFolder(token, BASE_FOLDER, p.folderName);
    }
    const templateBuf = await graphDownload(token, '/me/drive/root:/' + enc(TEMPLATE_NAME) + ':/content');
    const excelBuf = await populateExcel(templateBuf, formData, products);
    const clientSafe = (formData.clientName || 'client').replace(/[^a-z0-9]/gi, '_');
    const date = new Date().toISOString().slice(0, 10);
    const excelName = 'NPC_Form_' + clientSafe + '_' + date + '.xlsx';
    await uploadFile(token, BASE_FOLDER, excelName, excelBuf);
    res.status(200).json({ token, excelData: excelBuf.toString('base64'), excelName });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
};