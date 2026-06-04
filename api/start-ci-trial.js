// ============================================================
//  start-ci-trial.js  —  Vercel serverless function
//  Generates a 7-day Contact Intelligence trial key and emails it
//  Calls Apps Script action=createTrial&product=ci
// ============================================================

const https = require('https');

const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyRGFMHPNCJW-tcYNr99jBiVCyPOai8Qli5vtwcAIE9HVYJ3gTgb7l-Zwi08ZMIXdDXSA/exec';

function fetchUrl(url) {
  return new Promise((resolve, reject) => {
    https.get(url, (res) => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        return fetchUrl(res.headers.location).then(resolve).catch(reject);
      }
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => resolve(data.trim()));
    }).on('error', reject);
  });
}

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ ok: false });

  try {
    let body = req.body;
    if (!body || typeof body === 'string') {
      body = await new Promise((resolve) => {
        let data = '';
        req.on('data', chunk => { data += chunk; });
        req.on('end', () => { try { resolve(JSON.parse(data)); } catch { resolve({}); } });
      });
    }

    const { name, email } = body || {};
    if (!name || !email) return res.status(400).json({ ok: false, message: 'Name and email are required.' });

    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return res.status(400).json({ ok: false, message: 'Invalid email address.' });
    }

    const params = new URLSearchParams({
      action:  'createTrial',
      product: 'ci',
      name:    name.trim().slice(0, 80),
      email:   email.trim().toLowerCase(),
    });

    const result = await fetchUrl(`${APPS_SCRIPT_URL}?${params.toString()}`);

    if (result === 'ALREADY_EXISTS') {
      return res.status(200).json({ ok: false, message: 'A trial was already issued to this email. Check your inbox.' });
    }
    if (result.startsWith('TRIAL_CREATED:')) {
      return res.status(200).json({ ok: true, message: 'Trial key sent to your email. Download the tool and enter your key to begin.' });
    }

    return res.status(500).json({ ok: false, message: 'Could not create trial. Please try again.' });

  } catch (err) {
    console.error('start-ci-trial error:', err.message);
    return res.status(500).json({ ok: false, message: 'Server error. Please try again.' });
  }
};
