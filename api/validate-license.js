// ============================================================
//  validate-license.js  —  Vercel serverless function
//  Validates CRM-XXXX-XXXX-XXXX license keys
//  Against same Google Apps Script server as Contact Intelligence
//  Note: Apps Script returns plain text responses (ACTIVE, INVALID,
//        EXPIRED, WRONG_PC, SUSPENDED) — not JSON.
// ============================================================

const https = require('https');
const { reportError } = require('./report-error');

const LICENSE_SERVER = 'https://script.google.com/macros/s/AKfycbyRGFMHPNCJW-tcYNr99jBiVCyPOai8Qli5vtwcAIE9HVYJ3gTgb7l-Zwi08ZMIXdDXSA/exec';

function isValidKeyFormat(key) {
  if (!key || typeof key !== 'string') return false;
  const k = key.trim().toUpperCase();
  // CRM-XXXX-XXXX-XXXX = 18 chars
  return k.length === 18 &&
    k.startsWith('CRM-') &&
    k[8] === '-' &&
    k[13] === '-';
}

function fetchUrl(url) {
  return new Promise((resolve, reject) => {
    https.get(url, (res) => {
      // Follow redirects (Google Apps Script redirects to googleusercontent.com)
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        return fetchUrl(res.headers.location).then(resolve).catch(reject);
      }
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => resolve(data.trim().toUpperCase()));
    }).on('error', reject);
  });
}

function fetchLicenseServer(url) {
  return fetchUrl(url);
}

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();

  try {
    const { key, hwid } = req.method === 'POST' ? req.body : req.query;

    if (!key) return res.status(400).json({ valid: false, message: 'No license key provided.' });
    if (!isValidKeyFormat(key)) return res.status(400).json({ valid: false, message: 'Invalid key format. Expected: CRM-XXXX-XXXX-XXXX' });

    const url = `${LICENSE_SERVER}?action=validate&key=${encodeURIComponent(key.trim().toUpperCase())}&hwid=${encodeURIComponent(hwid || '')}`;

    const result = await fetchLicenseServer(url);

    // Plain text response handling
    if (result === 'ACTIVE' || result === 'OK' || result === 'WRONG_PC' || result === 'OK_ACTIVATED') {
      // OK / ACTIVE = valid key, HWID matches
      // WRONG_PC = key exists and is valid, different machine — allow for web tool (no HWID binding)
      // OK_ACTIVATED = first-time activation, HWID now bound
      return res.status(200).json({ valid: true, name: 'Licensed User' });
    } else if (result === 'EXPIRED') {
      return res.status(200).json({ valid: false, message: 'License has expired. Please renew.' });
    } else if (result === 'SUSPENDED') {
      return res.status(200).json({ valid: false, message: 'License suspended. Contact support.' });
    } else {
      // INVALID or unknown
      return res.status(200).json({ valid: false, message: 'License key not found or invalid.' });
    }
  } catch (err) {
    // Offline grace — if server unreachable, allow if key format is valid
    console.error('License server unreachable:', err.message);
    const { key } = req.method === 'POST' ? req.body : req.query;

    // Report license server connectivity failures silently
    reportError({
      product: 'CRM Data Extractor',
      version: 'v2.0',
      message: `License server unreachable: ${err.message}`,
      key,
      crm:   '',
      extra: 'validate-license',
    }).catch(() => {});

    if (isValidKeyFormat(key)) {
      return res.status(200).json({ valid: true, offline: true, message: 'Offline mode — could not reach license server.' });
    }
    return res.status(500).json({ valid: false, message: 'License server unreachable.' });
  }
};
