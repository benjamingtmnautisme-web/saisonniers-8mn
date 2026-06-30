const https = require('https');

const TENANT_ID     = process.env.AZURE_TENANT_ID;
const CLIENT_ID     = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;

const SHAREPOINT_HOST = 'splgtm.sharepoint.com';
const SITE_PATH       = '/sites/ServiceNautisme77';

function httpsGetBinary(url, headers) {
  return new Promise((resolve, reject) => {
    https.get(url, { headers }, res => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        // Suivre la redirection
        return resolve(httpsGetBinary(res.headers.location, {}));
      }
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => resolve({ status: res.statusCode, body: Buffer.concat(chunks), headers: res.headers }));
    }).on('error', reject);
  });
}

function httpsPost(url, body, headers) {
  return new Promise((resolve, reject) => {
    const urlObj = new URL(url);
    const options = {
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      method: 'POST',
      headers: { ...headers, 'Content-Length': Buffer.byteLength(body) }
    };
    const req = https.request(options, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => resolve({ status: res.statusCode, body: data }));
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

async function getToken() {
  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default'
  }).toString();
  const res = await httpsPost(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    body,
    { 'Content-Type': 'application/x-www-form-urlencoded' }
  );
  const json = JSON.parse(res.body);
  if (!json.access_token) throw new Error('Token failed: ' + res.body);
  return json.access_token;
}

async function getSiteId(token) {
  const url = `https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_HOST}:${SITE_PATH}`;
  const res = await new Promise((resolve, reject) => {
    https.get(url, { headers: { Authorization: `Bearer ${token}` } }, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => resolve({ body: data }));
    }).on('error', reject);
  });
  const json = JSON.parse(res.body);
  if (!json.id) throw new Error('Site non trouvé');
  return json.id;
}

exports.handler = async (event) => {
  const itemId = event.queryStringParameters && event.queryStringParameters.id;
  if (!itemId) {
    return { statusCode: 400, body: JSON.stringify({ error: 'id manquant' }) };
  }

  try {
    const token  = await getToken();
    const siteId = await getSiteId(token);
    const contentUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${itemId}/content`;

    const result = await httpsGetBinary(contentUrl, { Authorization: `Bearer ${token}` });

    return {
      statusCode: 200,
      headers: {
        'Content-Type': 'application/pdf',
        'Content-Disposition': 'inline',
        'Access-Control-Allow-Origin': '*',
        'Cache-Control': 'public, max-age=300'
      },
      body: result.body.toString('base64'),
      isBase64Encoded: true
    };
  } catch (err) {
    return { statusCode: 500, body: JSON.stringify({ error: err.message }) };
  }
};
