const https = require('https');
const crypto = require('crypto');

const VAPID_PUBLIC_KEY = process.env.VAPID_PUBLIC_KEY;
const VAPID_PRIVATE_KEY = process.env.VAPID_PRIVATE_KEY;
const VAPID_EMAIL = process.env.VAPID_EMAIL || 'mailto:contact@8millesnautic.fr';

// Firebase config pour lire les subscriptions
const FIREBASE_DB_URL = 'https://milles-e4f69-default-rtdb.europe-west1.firebasedatabase.app';

function base64urlToBuffer(str) {
  const pad = str.length % 4;
  const b64 = str.replace(/-/g, '+').replace(/_/g, '/') + '=='.slice(0, pad ? 4 - pad : 0);
  return Buffer.from(b64, 'base64');
}

function bufferToBase64url(buf) {
  return buf.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
}

async function getVapidJwt(audience) {
  const header = bufferToBase64url(Buffer.from(JSON.stringify({ typ: 'JWT', alg: 'ES256' })));
  const now = Math.floor(Date.now() / 1000);
  const payload = bufferToBase64url(Buffer.from(JSON.stringify({
    aud: audience,
    exp: now + 12 * 3600,
    sub: VAPID_EMAIL
  })));
  const signingInput = `${header}.${payload}`;
  const privKeyDer = Buffer.concat([
    Buffer.from('308141020100301306072a8648ce3d020106082a8648ce3d030107042730250201010420', 'hex'),
    base64urlToBuffer(VAPID_PRIVATE_KEY)
  ]);
  const privateKey = crypto.createPrivateKey({ key: privKeyDer, format: 'der', type: 'pkcs8' });
  const sig = crypto.sign('SHA256', Buffer.from(signingInput), { key: privateKey, dsaEncoding: 'ieee-p1363' });
  return `${signingInput}.${bufferToBase64url(sig)}`;
}

async function sendPushNotification(subscription, payload) {
  const url = new URL(subscription.endpoint);
  const audience = `${url.protocol}//${url.host}`;
  const jwt = await getVapidJwt(audience);

  const body = JSON.stringify(payload);
  const options = {
    hostname: url.hostname,
    path: url.pathname + url.search,
    method: 'POST',
    headers: {
      'Authorization': `vapid t=${jwt},k=${VAPID_PUBLIC_KEY}`,
      'Content-Type': 'application/json',
      'Content-Length': Buffer.byteLength(body),
      'TTL': '86400'
    }
  };

  return new Promise((resolve) => {
    const req = https.request(options, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => resolve({ status: res.statusCode, body: data }));
    });
    req.on('error', () => resolve({ status: 0 }));
    req.write(body);
    req.end();
  });
}

async function getSubscriptions() {
  return new Promise((resolve, reject) => {
    https.get(`${FIREBASE_DB_URL}/push_subscriptions.json`, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try {
          const json = JSON.parse(data);
          if (!json) return resolve([]);
          resolve(Object.values(json));
        } catch { resolve([]); }
      });
    }).on('error', reject);
  });
}

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method Not Allowed' };
  }

  try {
    const { titre, texte, categorie } = JSON.parse(event.body || '{}');
    const subscriptions = await getSubscriptions();

    const payload = {
      title: `🏄 8 Milles Nautic — ${categorie || 'Info'}`,
      body: titre + (texte ? ' — ' + texte.slice(0, 80) : ''),
      icon: '/icon-192.png',
      badge: '/icon-192.png'
    };

    const results = await Promise.all(
      subscriptions.map(sub => sendPushNotification(sub, payload))
    );

    const ok = results.filter(r => r.status >= 200 && r.status < 300).length;
    return {
      statusCode: 200,
      headers: { 'Access-Control-Allow-Origin': '*', 'Content-Type': 'application/json' },
      body: JSON.stringify({ sent: ok, total: subscriptions.length })
    };
  } catch (err) {
    return { statusCode: 500, body: JSON.stringify({ error: err.message }) };
  }
};
