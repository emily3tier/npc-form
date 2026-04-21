const https = require('https');

function getRawBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', chunk => chunks.push(chunk));
    req.on('end', () => resolve(Buffer.concat(chunks)));
    req.on('error', reject);
  });
}

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  if (req.method === 'OPTIONS') return res.status(200).end();
  try {
    const rawBody = await getRawBody(req);
    const { uploadUrl, data, start, end, total } = JSON.parse(rawBody.toString('utf8'));
    const buf = Buffer.from(data, 'base64');
    const urlObj = new URL(uploadUrl);
    await new Promise((resolve, reject) => {
      const r = https.request({
        hostname: urlObj.hostname,
        path: urlObj.pathname + urlObj.search,
        method: 'PUT',
        headers: { 'Content-Length': buf.length, 'Content-Range': 'bytes ' + start + '-' + end + '/' + total },
      }, (response) => {
        let d = '';
        response.on('data', c => d += c);
        response.on('end', () => response.statusCode >= 400 ? reject(new Error(response.statusCode + ': ' + d)) : resolve(d));
      });
      r.on('error', reject);
      r.write(buf);
      r.end();
    });
    res.status(200).json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
};
