const https = require('https');
const querystring = require('querystring');

const CLIENT_ID = 'd6dd88e1-a49e-4350-8339-f0f42c4b3b2e';
const TENANT_ID = '667afa82-1126-4a78-8f76-0918c7f2a845';
const BASE_FOLDER = 'UPC Submissions Automated';
const TEMPLATE_NAME = 'NPC Form 2026 1.xlsx';
const MASTER_LOG = 'Submission Log.xlsx';
const MONDAY_BOARD_ID = '18419075542';
const MONDAY_GROUP_ID = 'group_mm4kqrh';

function getRawBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', chunk => chunks.push(chunk));
    req.on('end', () => resolve(Buffer.concat(chunks)));
    req.on('error', reject);
  });
}

function httpsGet(url) {
  return new Promise((resolve, reject) => {
    const follow = (u) => {
      https.get(u, (res) => {
        if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) { follow(res.headers.location); return; }
        const chunks = [];
        res.on('data', chunk => chunks.push(chunk));
        res.on('end', () => {
          const buf = Buffer.concat(chunks);
          if (res.statusCode >= 400) reject(new Error('HTTP ' + res.statusCode + ': ' + buf.toString().substring(0, 200)));
          else resolve(buf);
        });
      }).on('error', reject);
    };
    follow(url);
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
  const body = querystring.stringify({ client_id: CLIENT_ID, grant_type: 'refresh_token', refresh_token: refreshToken, scope: 'Files.ReadWrite offline_access User.Read' });
  const data = await httpsRequest({ hostname: 'login.microsoftonline.com', path: '/' + TENANT_ID + '/oauth2/v2.0/token', method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) } }, body);
  return JSON.parse(data).access_token;
}

async function graphRequest(token, method, path, body, contentType) {
  const isBuffer = Buffer.isBuffer(body);
  const bodyStr = body && !isBuffer ? JSON.stringify(body) : body;
  const options = { hostname: 'graph.microsoft.com', path: '/v1.0' + path, method, headers: { 'Authorization': 'Bearer ' + token } };
  if (body) { options.headers['Content-Type'] = contentType || 'application/json'; options.headers['Content-Length'] = isBuffer ? body.length : Buffer.byteLength(bodyStr); }
  try { const data = await httpsRequest(options, bodyStr || body); return data ? JSON.parse(data) : {}; }
  catch (e) { if (e.message.includes('HTTP 409')) return {}; throw e; }
}

async function graphDownload(token, path) {
  const metaData = await httpsRequest({ hostname: 'graph.microsoft.com', path: '/v1.0' + path.replace(':/content', ''), method: 'GET', headers: { 'Authorization': 'Bearer ' + token } });
  const item = JSON.parse(metaData);
  const downloadUrl = item['@microsoft.graph.downloadUrl'];
  if (!downloadUrl) throw new Error('No download URL found for template');
  return httpsGet(downloadUrl);
}

async function tryGraphDownload(token, filePath) {
  try {
    const metaData = await httpsRequest({ hostname: 'graph.microsoft.com', path: '/v1.0/me/drive/root:/' + filePath.split('/').map(encodeURIComponent).join('/'), method: 'GET', headers: { 'Authorization': 'Bearer ' + token } });
    const item = JSON.parse(metaData);
    const downloadUrl = item['@microsoft.graph.downloadUrl'];
    if (!downloadUrl) return null;
    return await httpsGet(downloadUrl);
  } catch (e) { return null; }
}

async function getShareLink(token, folderPath) {
  try {
    const encoded = folderPath.split('/').map(encodeURIComponent).join('/');
    const data = await graphRequest(token, 'POST', '/me/drive/root:/' + encoded + ':/createLink', { type: 'view', scope: 'anonymous' });
    return data.link ? data.link.webUrl : null;
  } catch (e) { return null; }
}

function enc(s) { return encodeURIComponent(s); }

async function ensureFolder(token, parentPath, folderName) {
  const url = parentPath ? '/me/drive/root:/' + parentPath.split('/').map(enc).join('/') + ':/children' : '/me/drive/root/children';
  await graphRequest(token, 'POST', url, { name: folderName, folder: {} });
}

async function uploadFile(token, folderPath, fileName, data) {
  const fullPath = (folderPath + '/' + fileName).split('/').map(enc).join('/');
  await graphRequest(token, 'PUT', '/me/drive/root:/' + fullPath + ':/content', data, 'application/octet-stream');
}

async function sendEmail(token, to, cc, subject, body) {
  const message = {
    message: {
      subject,
      body: { contentType: 'HTML', content: body },
      toRecipients: (Array.isArray(to) ? to : [to]).map(a => ({ emailAddress: { address: a } })),
      ccRecipients: cc ? (Array.isArray(cc) ? cc : [cc]).map(a => ({ emailAddress: { address: a } })) : []
    }
  };
  await graphRequest(token, 'POST', '/me/sendMail', message);
}

async function createMondayItem(mondayToken, clientName, submittedBy, email, numProducts, codingOption, date, folderLink, productLinksJson) {
  const colVals = JSON.stringify({
    text_mm4k71we: clientName,
    text_mm4ktt2k: submittedBy,
    email_mm4kwrkw: { email: email, text: email },
    numeric_mm4k42ma: String(numProducts),
    text_mm4kq9qw: codingOption,
    date_mm4kqf41: { date: date },
    link_mm4kw1j3: { url: folderLink || '', text: clientName + ' Folder' },
    long_text_mm4mhm3s: { text: productLinksJson || '' }
  });
  const escaped = colVals.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
  const query = `mutation { create_item(board_id: ${MONDAY_BOARD_ID}, group_id: "${MONDAY_GROUP_ID}", item_name: "${clientName.replace(/"/g, '')}", column_values: "${escaped}") { id } }`;
  const body = JSON.stringify({ query });
  const data = await httpsRequest({
    hostname: 'api.monday.com', path: '/v2', method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': mondayToken, 'API-Version': '2024-01', 'Content-Length': Buffer.byteLength(body) }
  }, body);
  const result = JSON.parse(data);
  return result.data?.create_item?.id;
}


async function populateExcel(templateBuf, formData, products) {
  const ExcelJS = require('exceljs');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuf);
  const sheet = workbook.getWorksheet('New Product Coding Form') || workbook.worksheets[0];
  sheet.getCell('B4').value = formData.fromName || '';
  sheet.getCell('D4').value = formData.clientName || '';
  sheet.getCell('B5').value = formData.email || '';
  sheet.getCell('D5').value = formData.ownProduct || 'Yes';
  const coding = formData.codingOption || 'Code Immediately';
  sheet.getCell('A7').value = coding === 'Code Immediately' ? 'X' : '';
  sheet.getCell('A8').value = coding === 'Delay until Sales' ? 'X' : '';
  sheet.getCell('A9').value = coding === 'Code by Saturday Date' ? 'X' : '';
  if (coding === 'Code by Saturday Date' && formData.saturdayDate) sheet.getCell('C9').value = formData.saturdayDate;
  sheet.getCell('C10').value = 'Please provide Container Type:  ' + (formData.containerType || '');
  sheet.getCell('D10').value = 'Please provide Container Material Substance:' + (formData.containerMaterial || '');
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

async function updateMasterLog(token, formData, products, date) {
  const ExcelJS = require('exceljs');
  const existingBuf = await tryGraphDownload(token, BASE_FOLDER + '/' + MASTER_LOG);
  const workbook = new ExcelJS.Workbook();
  if (existingBuf) {
    await workbook.xlsx.load(existingBuf);
  } else {
    const sheet = workbook.addWorksheet('Submissions');
    const headerRow = sheet.getRow(1);
    headerRow.values = ['Date', 'Client Name', 'Submitted By', 'Email', '# Products', 'Coding Option', 'Submitted to NIQ', 'Notes'];
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1B3A6B' } };
    sheet.columns = [{ width: 12 }, { width: 25 }, { width: 22 }, { width: 30 }, { width: 12 }, { width: 22 }, { width: 18 }, { width: 20 }];
  }
  const sheet = workbook.getWorksheet('Submissions') || workbook.worksheets[0];
  sheet.addRow([date, formData.clientName || '', formData.fromName || '', formData.email || '', products.length, formData.codingOption || 'Code Immediately', '', '']);
  const out = await workbook.xlsx.writeBuffer();
  await uploadFile(token, BASE_FOLDER, MASTER_LOG, Buffer.from(out));
}

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  if (req.method === 'OPTIONS') return res.status(200).end();
  try {
    const refreshToken = process.env.MS_REFRESH_TOKEN;
    const mondayToken = process.env.MONDAY_API_TOKEN;
    if (!refreshToken) throw new Error('MS_REFRESH_TOKEN not set');
    const rawBody = await getRawBody(req);
    const { formData, products } = JSON.parse(rawBody.toString('utf8'));
    const token = await getAccessToken(refreshToken);
    const date = new Date().toISOString().slice(0, 10);
    const clientSafe = (formData.clientName || 'Unknown').replace(/[^a-z0-9 ]/gi, '').trim();
    const timeStr = new Date().toISOString().slice(11,16).replace(':','');
    const submissionFolder = BASE_FOLDER + '/' + clientSafe + '/' + date + '_' + timeStr;

    // Create folders
    await ensureFolder(token, '', BASE_FOLDER);
    await ensureFolder(token, BASE_FOLDER, clientSafe);
    await ensureFolder(token, BASE_FOLDER + '/' + clientSafe, date);
    for (const p of products) await ensureFolder(token, submissionFolder, p.folderName);

    // Generate and upload Excel
    const templateBuf = await graphDownload(token, '/me/drive/root:/' + enc(TEMPLATE_NAME) + ':/content');
    const excelBuf = await populateExcel(templateBuf, formData, products);
    const timeStamp = new Date().toISOString().slice(11,16).replace(':','');
    const excelName = 'NPC_Form_' + clientSafe.replace(/ /g, '_') + '_' + date + '_' + timeStamp + '.xlsx';
    await uploadFile(token, submissionFolder, excelName, excelBuf);

    // Update master log
    await updateMasterLog(token, formData, products, date);

    // Get SharePoint links for each product folder
    const productLinks = [];
    for (const p of products) {
      const link = await getShareLink(token, submissionFolder + '/' + p.folderName);
      productLinks.push({ name: p.folderName, link });
    }
    const excelLink = await getShareLink(token, submissionFolder + '/' + excelName);

    // Create Monday.com item
    let mondayItemId = null;
    if (mondayToken) {
      try {
        const columnValues = {
          status: { label: 'Pending Review' },
          date4: { date: date },
          text_mm4kq87j: productLinks.map(pl => pl.link || '').join(', ')
        };
        const folderShareLink = productLinks.length > 0 ? productLinks[0].link : '';
        // Build per-product links JSON for email
        const productLinksData = products.map((p, i) => ({
          name: p.folderName,
          link: productLinks[i]?.link || ''
        }));
        const excelLinkData = excelLink || '';
        mondayItemId = await createMondayItem(mondayToken, formData.clientName || 'Unknown', formData.fromName || '', formData.email || '', products.length, formData.codingOption || 'Code Immediately', date, folderShareLink, JSON.stringify({ products: productLinksData, excel: excelLinkData }));
      } catch (e) { console.error('Monday error:', e.message); }
    }

    // Send review email to Emily
    const productListHtml = productLinks.map(pl =>
      '<p><strong>' + pl.name + '</strong><br>' + (pl.link ? '<a href="' + pl.link + '">' + pl.link + '</a>' : 'No link') + '</p>'
    ).join('') + '<p><strong>Excel Form:</strong> ' + (excelLink ? '<a href="' + excelLink + '">' + excelLink + '</a>' : 'No link') + '</p>';

    const reviewEmailBody = `
      <p>Hi Emily,</p>
      <p>A new NPC submission has come in from <strong>${formData.clientName}</strong> (${products.length} product(s)).</p>
      <p><strong>Submitted by:</strong> ${formData.fromName} (${formData.email})<br>
      <strong>Coding option:</strong> ${formData.codingOption}<br>
      <strong>Container:</strong> ${formData.containerType} / ${formData.containerMaterial}</p>
      <p><strong>Product folders:</strong></p>
      ${productListHtml}
      <p>Please review the images in OneDrive, then go to Monday.com and change the status to <strong>"Approved"</strong> to trigger the NIQ submission email.</p>
      <p>— NPC Form Auto-Notification</p>
    `;

    try {
      await sendEmail(token, 'emily@3tierbeverages.com', null, 'NPC Review: ' + formData.clientName + ' (' + products.length + ' products)', reviewEmailBody);
    } catch (e) { console.error('Review email error:', e.message); }

    res.status(200).json({ token, excelData: excelBuf.toString('base64'), excelName, submissionFolder, mondayItemId });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
};