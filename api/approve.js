const https = require('https');
const querystring = require('querystring');

const CLIENT_ID = 'd6dd88e1-a49e-4350-8339-f0f42c4b3b2e';
const TENANT_ID = '667afa82-1126-4a78-8f76-0918c7f2a845';
const MONDAY_BOARD_ID = '18419075542';

// SET TO true WHEN READY TO GO LIVE — sends to NIQ instead of Emily
const TESTING_MODE = true;
const NIQ_EMAIL = 'npcimages@nielseniq.com';
const NIQ_CC = 'james.augustine@nielseniq.com';
const TEST_EMAIL = 'emily@3tierbeverages.com';

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
    client_id: CLIENT_ID, grant_type: 'refresh_token', refresh_token: refreshToken,
    scope: 'Files.ReadWrite offline_access User.Read Mail.Send'
  });
  const data = await httpsRequest({
    hostname: 'login.microsoftonline.com', path: '/' + TENANT_ID + '/oauth2/v2.0/token', method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) }
  }, body);
  return JSON.parse(data).access_token;
}

async function graphRequest(token, method, path, body) {
  const bodyStr = body ? JSON.stringify(body) : null;
  const options = {
    hostname: 'graph.microsoft.com', path: '/v1.0' + path, method,
    headers: { 'Authorization': 'Bearer ' + token }
  };
  if (bodyStr) { options.headers['Content-Type'] = 'application/json'; options.headers['Content-Length'] = Buffer.byteLength(bodyStr); }
  const data = await httpsRequest(options, bodyStr);
  return data ? JSON.parse(data) : {};
}

async function getMondayItem(mondayToken, itemId) {
  const query = `{ items(ids: [${itemId}]) { name column_values { id text value } } }`;
  const body = JSON.stringify({ query });
  const data = await httpsRequest({
    hostname: 'api.monday.com', path: '/v2', method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': mondayToken, 'API-Version': '2024-01', 'Content-Length': Buffer.byteLength(body) }
  }, body);
  const result = JSON.parse(data);
  return result.data?.items?.[0];
}

async function updateMondayStatus(mondayToken, itemId, statusLabel) {
  const colVals = JSON.stringify({ status: { label: statusLabel } });
  const escaped = colVals.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
  const query = `mutation { change_multiple_column_values(board_id: ${MONDAY_BOARD_ID}, item_id: ${itemId}, column_values: "${escaped}") { id } }`;
  const body = JSON.stringify({ query });
  await httpsRequest({
    hostname: 'api.monday.com', path: '/v2', method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': mondayToken, 'API-Version': '2024-01', 'Content-Length': Buffer.byteLength(body) }
  }, body);
}

async function sendEmail(token, to, cc, subject, bodyHtml) {
  const message = {
    message: {
      subject,
      body: { contentType: 'HTML', content: bodyHtml },
      toRecipients: (Array.isArray(to) ? to : [to]).map(a => ({ emailAddress: { address: a } })),
      ccRecipients: cc ? (Array.isArray(cc) ? cc : [cc]).map(a => ({ emailAddress: { address: a } })) : []
    }
  };
  await graphRequest(token, 'POST', '/me/sendMail', message);
}

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  if (req.method === 'OPTIONS') return res.status(200).end();

  try {
    const rawBody = await getRawBody(req);
    const payload = JSON.parse(rawBody.toString('utf8'));

    // Monday.com webhook challenge handshake
    if (payload.challenge) return res.status(200).json({ challenge: payload.challenge });

    const event = payload.event;
    if (!event || event.type !== 'change_column_value') return res.status(200).json({ ok: true });

    // Only fire on status change to "Approved"
    const newValue = JSON.parse(event.value?.value || '{}');
    const label = newValue?.label?.text || newValue?.label;
    if (label !== 'Approved') return res.status(200).json({ ok: true, skipped: true });

    const itemId = event.pulseId;
    const mondayToken = process.env.MONDAY_API_TOKEN;
    const refreshToken = process.env.MS_REFRESH_TOKEN;
    const msToken = await getAccessToken(refreshToken);

    // Get item details from Monday
    const item = await getMondayItem(mondayToken, itemId);
    if (!item) return res.status(200).json({ ok: false, error: 'Item not found' });

    const getCol = (id) => item.column_values.find(c => c.id === id)?.text || '';
    const clientName = getCol('text_mm4k71we') || item.name;
    const submittedByEmail = getCol('email_mm4kwrkw');
    const numProducts = getCol('numeric_mm4k42ma');
    const codingOption = getCol('text_mm4kq9qw');
    const folderLink = getCol('link_mm4kw1j3');

    // Build NIQ email body — exact format from Emily's example
    const productLines = folderLink
      ? `<p><strong>${clientName}</strong><br><a href="${folderLink}">${folderLink}</a></p>`
      : `<p><strong>${clientName}</strong><br>(No folder link available)</p>`;

    const niqBody = `<p>Hi,</p>
<p>Please find the NPC product image submission for <strong>${clientName}</strong> below.</p>
${productLines}
<br>
<p>Emily Kessel, Consultant Analyst Intern, UPC Coordinator<br>
C: +925-984-6798<br>
Data-Driven Solutions, Dedicated Partnership, and Genuine Relationships</p>`;

    const toAddress = TESTING_MODE ? TEST_EMAIL : NIQ_EMAIL;
    const ccAddress = TESTING_MODE ? null : NIQ_CC;
    const subject = 'NPC Submission - ' + clientName + (TESTING_MODE ? ' [TEST]' : '');

    await sendEmail(msToken, toAddress, ccAddress, subject, niqBody);

    // Send client confirmation if we have their email
    if (submittedByEmail && submittedByEmail !== TEST_EMAIL) {
      const clientBody = `<p>Hi,</p>
<p>Your NPC UPC coding submission for <strong>${clientName}</strong> (${numProducts} product(s)) has been submitted to NIQ for processing.</p>
<p>Coding option: ${codingOption}</p>
<p>Thank you for submitting through 3 Tier Beverages!</p>
<p>Emily Kessel, Consultant Analyst Intern, UPC Coordinator<br>
C: +925-984-6798</p>`;
      await sendEmail(msToken, submittedByEmail, null, 'Your NPC Submission Has Been Sent to NIQ - ' + clientName, clientBody);
    }

    // Update Monday status to "Submitted to NIQ"
    await updateMondayStatus(mondayToken, itemId, TESTING_MODE ? 'Done' : 'Submitted to NIQ');

    res.status(200).json({ ok: true, sent: true, testing: TESTING_MODE, to: toAddress });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
};