const https = require('https');

// ── CONFIG ──
const TENANT_ID     = process.env.AZURE_TENANT_ID;
const CLIENT_ID     = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;

const SHAREPOINT_HOST = 'splgtm.sharepoint.com';
const SITE_PATH       = '/sites/ServiceNautisme77';
// Dossier contenant les PDF de planning (un fichier par semaine, toutes bases confondues)
const FOLDER_PATH     = 'Général/Plannings activités et équipes/PDF plannings Granville/Plannings CQP 2026';

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
    janvier: 0, fevrier: 1, mars: 2, avril: 3, mai: 4, juin: 5,
    juillet: 6, aout: 7, septembre: 8, octobre: 9, novembre: 10, decembre: 11
  };
  const normalise = nomFichier.toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');

  // Cas 1 : "du 27 juillet au 2 aout 2026" (deux mois différents)
  const m2 = normalise.match(/du\s+(\d+)\s+([a-z]+)\s+au\s+\d+\s+[a-z]+\s+(\d{4})/);
  if (m2) {
    const jourDebut = parseInt(m2[1]);
    const nomMois = m2[2];
    const annee = parseInt(m2[3]);
    const moisIdx = mois[nomMois];
    if (moisIdx !== undefined) return new Date(annee, moisIdx, jourDebut).getTime();
  }

  // Cas 2 : "du 6 au 12 juillet 2026" (même mois)
  const m1 = normalise.match(/du\s+(\d+)\s+au\s+\d+\s+([a-z]+)\s+(\d{4})/);
  if (m1) {
    const jourDebut = parseInt(m1[1]);
    const nomMois = m1[2];
    const annee = parseInt(m1[3]);
    const moisIdx = mois[nomMois];
    if (moisIdx !== undefined) return new Date(annee, moisIdx, jourDebut).getTime();
  }

  return null;
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
        id: f.id,
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
