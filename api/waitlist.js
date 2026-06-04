// api/waitlist.js
// Stores ApplyFlows (and future product) waitlist entries in Google Sheet

const https = require('https');

const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyRGFMHPNCJW-tcYNr99jBiVCyPOai8Qli5vtwcAIE9HVYJ3gTgb7l-Zwi08ZMIXdDXSA/exec';

function callAppsScript(params) {
  return new Promise((resolve, reject) => {
    const qs = new URLSearchParams(params).toString();
    const url = `${APPS_SCRIPT_URL}?${qs}`;

    const makeRequest = (targetUrl) => {
      https.get(targetUrl, (res) => {
        if (res.statusCode === 301 || res.statusCode === 302) {
          return makeRequest(res.headers.location);
        }
        let body = '';
        res.on('data', d => body += d);
        res.on('end', () => {
          try { resolve(JSON.parse(body)); }
          catch { resolve({ status: 'ERROR', message: body }); }
        });
      }).on('error', reject);
    };

    makeRequest(url);
  });
}

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ ok: false, message: 'Method not allowed' });

  const { name, email, product = 'applyflows' } = req.body || {};

  if (!name || !email) return res.status(400).json({ ok: false, message: 'Name and email are required.' });
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return res.status(400).json({ ok: false, message: 'Invalid email address.' });

  try {
    const result = await callAppsScript({ action: 'addWaitlist', name, email, product });

    if (result.status === 'ALREADY_EXISTS') {
      return res.json({ ok: true, message: "You're already on the waitlist! We'll be in touch." });
    }

    if (result.status === 'OK') {
      return res.json({ ok: true, message: "You're on the list! We'll email you when it launches." });
    }

    return res.status(500).json({ ok: false, message: result.message || 'Something went wrong. Please try again.' });
  } catch (err) {
    console.error('Waitlist error:', err);
    return res.status(500).json({ ok: false, message: 'Server error. Please try again.' });
  }
};
