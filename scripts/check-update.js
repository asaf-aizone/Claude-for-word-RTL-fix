#!/usr/bin/env node
/**
 * Claude for Word RTL - update checker.
 *
 * Reads local version from ./package.json and compares to the latest GitHub
 * release. No side effects, no file writes, no new dependencies (uses the
 * built-in https module only).
 *
 * Exit codes:
 *   0 - up to date OR update available (informational only)
 *   1 - network/parse error or timeout
 */

'use strict';

const https = require('https');
const path = require('path');

const TIMEOUT_MS = 5000;
const REPO = 'asaf-aizone/Claude-for-word-RTL-fix';
const API_URL = 'https://api.github.com/repos/' + REPO + '/releases/latest';

let localVersion;
try {
  localVersion = require('./package.json').version;
} catch (e) {
  console.error('[ERROR] Could not read scripts/package.json: ' + e.message);
  process.exit(1);
}

function normalize(v) {
  return String(v || '').replace(/^v/i, '').trim();
}

// Proper semver-ish comparison: split on '.' and compare component-wise as
// integers. Non-numeric suffixes (e.g. "-beta") are stripped. Returns
// negative if a < b, positive if a > b, zero if equal.
function compareVersions(a, b) {
  var pa = String(a).split('-')[0].split('.').map(function (x) { return parseInt(x, 10) || 0; });
  var pb = String(b).split('-')[0].split('.').map(function (x) { return parseInt(x, 10) || 0; });
  var n = Math.max(pa.length, pb.length);
  for (var i = 0; i < n; i++) {
    var d = (pa[i] || 0) - (pb[i] || 0);
    if (d !== 0) return d;
  }
  return 0;
}

const opts = {
  headers: {
    'User-Agent': 'claude-word-rtl update-checker',
    'Accept': 'application/vnd.github+json'
  }
};

const req = https.get(API_URL, opts, function (res) {
  if (res.statusCode !== 200) {
    // Drain then report
    res.resume();
    console.error('[ERROR] Could not reach GitHub: HTTP ' + res.statusCode);
    process.exit(1);
  }
  let body = '';
  res.setEncoding('utf8');
  res.on('data', function (chunk) { body += chunk; });
  res.on('end', function () {
    let data;
    try { data = JSON.parse(body); }
    catch (e) {
      console.error('[ERROR] Could not reach GitHub: invalid JSON (' + e.message + ')');
      process.exit(1);
    }
    const latest = normalize(data.tag_name || data.name);
    const local = normalize(localVersion);
    if (!latest) {
      console.error('[ERROR] Could not reach GitHub: no tag_name in response');
      process.exit(1);
    }
    var cmp = compareVersions(latest, local);
    if (cmp === 0) {
      console.log('[UP TO DATE] Local version: ' + local);
    } else if (cmp > 0) {
      console.log('[UPDATE AVAILABLE] Local ' + local + ', latest ' + latest + '. Download: ' + (data.html_url || ('https://github.com/' + REPO + '/releases/latest')));
    } else {
      // Local is ahead (dev build). Treat as up to date.
      console.log('[UP TO DATE] Local version: ' + local + ' (ahead of latest release ' + latest + ')');
    }
  });
});

req.setTimeout(TIMEOUT_MS, function () {
  req.destroy();
  console.error('[ERROR] Timeout reaching GitHub');
  process.exit(1);
});

req.on('error', function (err) {
  console.error('[ERROR] Could not reach GitHub: ' + err.message);
  process.exit(1);
});
