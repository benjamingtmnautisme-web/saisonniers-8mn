const https = require('https');

// ── CONFIG ──
const TENANT_ID     = process.env.AZURE_TENANT_ID;
const CLIENT_ID     = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;

const SHAREPOINT_HOST = 'splgtm.sharepoint.com';
const SITE_PATH       = '/sites/ServiceNautisme77';
// Dossier contenant les PDF de planning (un fichier par semaine, toutes bases confondues)
const FOLDER_PATH     = 'Général/Plannings activités et équipes/PDF plannings Granville';

function httpsGet(url, headers) {
  return new Promise((resolve, reject) => {
    const req = https.get(url, { headers }, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => resolve({ status: res.statusCode, body: data }));
    });
    req.on('error', reject);
  });
}

function httpsPost(url, body, headers) {
  return new Promise((resolve, reject) => {
    const postData = body;
    const urlObj = new URL(url);
    const options = {
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      method: 'POST',
      headers: { ...headers, 'Content-Length': Buffer.byteLength(postData) }
    };
    const req = https.request(options, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => resolve({ status: res.statusCode, body: data }));
    });
    req.on('error', reject);
    req.write(postData);
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
  const res = await httpsGet(
    `https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_HOST}:${SITE_PATH}`,
    { Authorization: `Bearer ${token}` }
  );
  const json = JSON.parse(res.body);
  if (!json.id) throw new Error('Site non trouvé: ' + res.body);
  return json.id;
}

async function listPdfFiles(token, siteId) {
  const encodedPath = FOLDER_PATH.split('/').map(s => encodeURIComponent(s)).join('/');
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodedPath}:/children`;
  const res = await httpsGet(url, { Authorization: `Bearer ${token}` });
  const json = JSON.parse(res.body);
  if (!json.value) throw new Error('Dossier non trouvé: ' + res.body);
  return json.value;
}

// Extrait une date de début depuis un nom de fichier "Du 6 au 10 juillet 2026.pdf"
// pour permettre le tri chronologique
function extraireDateDebut(nomFichier) {
  const mois = {
    janvier: 0, février: 1, mars: 2, avril: 3, mai: 4, juin: 5,
    juillet: 6, août: 7, septembre: 8, octobre: 9, novembre: 10, décembre: 11
  };
  const m = nomFichier.toLowerCase().match(/du\s+(\d+)\s+au\s+(\d+)\s+([a-zéû]+)\s+(\d{4})/);
  if (!m) return null;
  const jourDebut = parseInt(m[1]);
  const nomMois = m[3];
  const annee = parseInt(m[4]);
  const moisIdx = mois[nomMois];
  if (moisIdx === undefined) return null;
  return new Date(annee, moisIdx, jourDebut).getTime();
}

exports.handler = async (event) => {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Content-Type': 'application/json'
  };

  try {
    const token  = await getToken();
    const siteId = await getSiteId(token);
    const files  = await listPdfFiles(token, siteId);

    const pdfs = files
      .filter(f => f.name && f.name.toLowerCase().endsWith('.pdf'))
      .map(f => ({
        nom: f.name.replace(/\.pdf$/i, ''),
        url: f['@microsoft.graph.downloadUrl'] || null,
        webUrl: f.webUrl || null,
        timestamp: extraireDateDebut(f.name)
      }))
      .sort((a, b) => (a.timestamp || 0) - (b.timestamp || 0));

    return { statusCode: 200, headers, body: JSON.stringify({ pdfs }) };

  } catch (err) {
    return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
  }
};
