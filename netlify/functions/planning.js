const https = require('https');

// ── CONFIG ──
const TENANT_ID     = process.env.AZURE_TENANT_ID;
const CLIENT_ID     = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;

const SITE_NAME   = 'Service Nautisme - Documents';
const FILE_PATH   = 'Général/Plannings activités et équipes/Saison 2026/Planning_moniteurs_CQP_2026.xlsm';

// Mapping colonnes -> jours (section Granville, feuille semaine)
const JOURS_COLS = [
  { jour: 'Lundi',    cPoste: 'C', cStageH: 'D', cStageNom: 'E', cSecteur: 'B' },
  { jour: 'Mardi',    cPoste: 'G', cStageH: 'H', cStageNom: 'I', cSecteur: 'F' },
  { jour: 'Mercredi', cPoste: 'K', cStageH: 'L', cStageNom: 'M', cSecteur: 'J' },
  { jour: 'Jeudi',    cPoste: 'O', cStageH: 'P', cStageNom: 'Q', cSecteur: 'N' },
  { jour: 'Vendredi', cPoste: 'S', cStageH: 'T', cStageNom: 'U', cSecteur: 'R' },
  { jour: 'Samedi',   cPoste: 'W', cStageH: 'X', cStageNom: 'Y', cSecteur: 'V' },
  { jour: 'Dimanche', cPoste: 'AA',cStageH: 'AB',cStageNom: 'AC',cSecteur: 'Z' },
];

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
    `https://graph.microsoft.com/v1.0/sites?search=Service+Nautisme`,
    { Authorization: `Bearer ${token}` }
  );
  const json = JSON.parse(res.body);
  const site = json.value && json.value[0];
  if (!site) throw new Error('Site non trouvé');
  return site.id;
}

async function getFileContent(token, siteId) {
  // Encoder le chemin
  const encoded = FILE_PATH.split('/').map(s => encodeURIComponent(s)).join('/');
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encoded}:/workbook/worksheets`;
  const res = await httpsGet(url, { Authorization: `Bearer ${token}` });
  return JSON.parse(res.body);
}

async function getSheetRange(token, siteId, sheetId, range) {
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${FILE_PATH.split('/').map(s => encodeURIComponent(s)).join('/')}:/workbook/worksheets/${sheetId}/range(address='${range}')`;
  const res = await httpsGet(url, { Authorization: `Bearer ${token}` });
  return JSON.parse(res.body);
}

function parseSheetData(values, sheetName) {
  // values est un tableau 2D de cellules
  // On cherche les lignes avec numéros en col A (index 0)
  const moniteurs = {};

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const cellA = String(row[0] || '').trim();
    if (!cellA || isNaN(parseInt(cellA))) continue;

    const num = parseInt(cellA);
    const row2 = values[i + 1] || [];

    // Colonne B = nom (remplacera le numéro plus tard)
    const nom = String(row[1] || '').trim() || String(num);

    const creneaux = [];
    for (const j of JOURS_COLS) {
      const colIdx = col => {
        if (col.length === 1) return col.charCodeAt(0) - 65;
        return 26 + col.charCodeAt(1) - 65;
      };

      const iPoste    = colIdx(j.cPoste);
      const iStageH   = colIdx(j.cStageH);
      const iStageNom = colIdx(j.cStageNom);
      const iSecteur  = colIdx(j.cSecteur);

      const matinPoste  = String(row[iPoste]    || '').trim();
      const matinStageH = String(row[iStageH]   || '').trim();
      const matinNom    = String(row[iStageNom] || '').trim();
      const matinSect   = String(row[iSecteur]  || '').trim();

      const amPoste  = String(row2[iPoste]    || '').trim();
      const amStageH = String(row2[iStageH]   || '').trim();
      const amNom    = String(row2[iStageNom] || '').trim();
      const amSect   = String(row2[iSecteur]  || '').trim();

      if (matinPoste || matinStageH) {
        creneaux.push({ jour: j.jour, session: 'Matin', secteur: matinSect, poste: matinPoste, stage_h: matinStageH, stage_nom: matinNom });
      }
      if (amPoste || amStageH) {
        creneaux.push({ jour: j.jour, session: 'A-M', secteur: amSect, poste: amPoste, stage_h: amStageH, stage_nom: amNom });
      }
      if (!matinPoste && !matinStageH && !amPoste && !amStageH) {
        creneaux.push({ jour: j.jour, session: 'Repos', secteur: '', poste: '', stage_h: '', stage_nom: '' });
      }
    }

    moniteurs[nom] = creneaux;
  }
  return moniteurs;
}

exports.handler = async (event) => {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Content-Type': 'application/json'
  };

  try {
    const token  = await getToken();
    const siteId = await getSiteId(token);

    // Récupérer la liste des feuilles
    const sheets = await getFileContent(token, siteId);
    if (!sheets.value) throw new Error('Pas de feuilles: ' + JSON.stringify(sheets));

    const planning = {}; // { semaine: { nom: [creneaux] } }

    for (const sheet of sheets.value) {
      const name = sheet.name;
      // On ignore les feuilles qui ne ressemblent pas à des semaines
      if (!name.match(/\d+/)) continue;

      // Lire une plage large (ex. A1:AE60)
      const range = await getSheetRange(token, siteId, encodeURIComponent(sheet.name), 'A1:AE60');
      if (!range.values) continue;

      const moniteurs = parseSheetData(range.values, name);
      planning[name] = moniteurs;
    }

    return { statusCode: 200, headers, body: JSON.stringify({ planning }) };

  } catch (err) {
    return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
  }
};
