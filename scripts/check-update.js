#!/usr/bin/env node
/**
 * Claude for Office RTL - update checker.
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
const url = require('url');
const path = require('path');

const TIMEOUT_MS = 5000;
const MAX_REDIRECTS = 3;
const REPO = 'asaf-aizone/Claude-for-Office-RTL-fix';
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
    'User-Agent': 'claude-office-rtl update-checker',
    'Accept': 'application/vnd.github+json'
  }
};

// Follows up to MAX_REDIRECTS HTTP 301/302 hops. Required because the
// repository was renamed in v0.2.1 (Claude-for-word-RTL-fix ->
// Claude-for-Office-RTL-fix); api.github.com returns 301 to old-name
// callers from v0.1.x and v0.2.0. Node's https.get does not follow
// redirects by default. Cap kept low to bound any future rename chain.
function fetchJson(targetUrl, hopsRemaining, callback) {
  const parsed = url.parse(targetUrl);
  const reqOptions = {
    protocol: parsed.protocol,
    hostname: parsed.hostname,
    port: parsed.port,
    path: parsed.path,
    method: 'GET',
    headers: opts.headers
  };
  const req = https.request(reqOptions, function (res) {
    if ((res.statusCode === 301 || res.statusCode === 302 ||
         res.statusCode === 307 || res.statusCode === 308) &&
        res.headers.location) {
      res.resume();
      if (hopsRemaining <= 0) {
        callback(new Error('too many redirects'));
        return;
      }
      const nextUrl = url.resolve(targetUrl, res.headers.location);
      fetchJson(nextUrl, hopsRemaining - 1, callback);
      return;
    }
    if (res.statusCode !== 200) {
      res.resume();
      callback(new Error('HTTP ' + res.statusCode));
      return;
    }
    let body = '';
    res.setEncoding('utf8');
    res.on('data', function (chunk) { body += chunk; });
    res.on('end', function () {
      let data;
      try { data = JSON.parse(body); }
      catch (e) {
        callback(new Error('invalid JSON (' + e.message + ')'));
        return;
      }
      callback(null, data);
    });
  });
  req.setTimeout(TIMEOUT_MS, function () {
    req.destroy(new Error('timeout'));
  });
  req.on('error', function (err) {
    callback(err);
  });
  req.end();
}

fetchJson(API_URL, MAX_REDIRECTS, function (err, data) {
  if (err) {
    if (err.message === 'timeout') {
      console.error('[ERROR] Timeout reaching GitHub');
    } else {
      console.error('[ERROR] Could not reach GitHub: ' + err.message);
    }
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
