// ============================================================
//  report-error.js  —  Silent error telemetry
//  Logs errors to Google Sheet via existing Apps Script backend
//  Never throws — always fails silently
// ============================================================

const https = require('https');

const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzl-W9lWLp3GTTXfyONpqS1Y9gZ6L7LG3DZfNRYI62O-pfm5nBj5fUBSJBhoSib_7Rh/exec';

/**
 * Follows redirects recursively (Google Apps Script always redirects)
 */
function fetchUrl(url) {
  return new Promise((resolve) => {
    https.get(url, (res) => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        return fetchUrl(res.headers.location).then(resolve).catch(() => resolve('ERROR'));
      }
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => resolve(data));
    }).on('error', () => resolve('ERROR'));
  });
}

/**
 * reportError — fire and forget, never throws
 *
 * @param {object} opts
 * @param {string} opts.product   - 'CRM Data Extractor' | 'Contact Intelligence'
 * @param {string} opts.version   - e.g. 'v2.0'
 * @param {string} opts.message   - error message
 * @param {string} [opts.key]     - license key (masked)
 * @param {string} [opts.crm]     - CRM type e.g. 'erpnext'
 * @param {string} [opts.extra]   - any extra context
 */
async function reportError({ product, version, message, key, crm, extra } = {}) {
  try {
    const maskedKey = key ? key.slice(0, 7) + '****' : '';
    const params = new URLSearchParams({
      action:  'logError',
      product: product  || 'CRM Data Extractor',
      version: version  || 'unknown',
      message: (message || 'No message').slice(0, 500), // cap at 500 chars
      key:     maskedKey,
      crm:     crm   || '',
      extra:   extra || '',
    });
    await fetchUrl(`${APPS_SCRIPT_URL}?${params.toString()}`);
  } catch (_) {
    // Completely silent — telemetry must never break the tool
  }
}

module.exports = { reportError };
