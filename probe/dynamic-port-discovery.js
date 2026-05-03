#!/usr/bin/env node
/**
 * Dynamic-port discovery POC for Claude Office add-ins.
 *
 * Validates the architecture proposed in docs/OFFICE-EXPANSION-PLAN.md section 3.5:
 * when WebView2 is launched with --remote-debugging-port=0, each Office app picks
 * a different free port. This script discovers those ports without knowing them
 * in advance, by:
 *
 *   1. tasklist  - list all msedgewebview2.exe PIDs
 *   2. netstat   - map each PID to the TCP port(s) it is listening on
 *   3. /json/list - probe each candidate port for Claude targets
 *   4. _host_Info - identify which Office app the target belongs to
 *
 * Usage:
 *   node dynamic-port-discovery.js
 *
 * No deps. Windows only (relies on tasklist + netstat output format).
 */

const http = require('http');
const { execFileSync } = require('child_process');

const URL_PATTERN_PRIMARY = /\/\/([^\/]*\.)?pivot\.claude\.ai(\/|$|\?|#)/i;
const URL_PATTERN_FALLBACK = /\/\/([^\/]*\.)?claude\.ai(\/|$|\?|#)/i;
const HOST_INFO_REGEX = /[?&]_host_Info=([^&#]+)/i;

function listWebView2Pids() {
  const out = execFileSync('tasklist', ['/FI', 'IMAGENAME eq msedgewebview2.exe', '/FO', 'CSV', '/NH'], {
    encoding: 'utf8',
    windowsHide: true,
  });
  const pids = [];
  out.split(/\r?\n/).forEach(function (line) {
    const m = line.match(/^"[^"]*","(\d+)"/);
    if (m) pids.push(parseInt(m[1], 10));
  });
  return pids;
}

function listListeningPortsByPid() {
  const out = execFileSync('netstat', ['-ano'], {
    encoding: 'utf8',
    windowsHide: true,
  });
  const map = new Map();
  out.split(/\r?\n/).forEach(function (line) {
    const m = line.match(/^\s*TCP(?:v6)?\s+(\S+)\s+\S+\s+LISTENING\s+(\d+)/);
    if (!m) return;
    const local = m[1];
    const pid = parseInt(m[2], 10);
    const portMatch = local.match(/:(\d+)$/);
    if (!portMatch) return;
    const port = parseInt(portMatch[1], 10);
    if (!map.has(pid)) map.set(pid, new Set());
    map.get(pid).add(port);
  });
  return map;
}

function tryListTargetsOnHost(host, port) {
  return new Promise(function (resolve, reject) {
    const req = http.get({ host: host, port: port, path: '/json/list', timeout: 1500 }, function (res) {
      let body = '';
      res.on('data', function (c) { body += c; });
      res.on('end', function () {
        try { resolve(JSON.parse(body)); }
        catch (e) { reject(new Error('invalid JSON: ' + e.message)); }
      });
    });
    let settled = false;
    req.on('timeout', function () { if (!settled) { settled = true; req.destroy(new Error('timeout')); } });
    req.on('error', function (e) { if (!settled) { settled = true; reject(e); } });
  });
}

async function listTargets(port) {
  // Probe both IPv4 and IPv6 loopback - WebView2 historically binds [::1] only.
  // See CLAUDE.local.md section 4 (war story v0.1.3).
  try { return await tryListTargetsOnHost('127.0.0.1', port); }
  catch (e1) {
    try { return await tryListTargetsOnHost('::1', port); }
    catch (e2) { throw new Error('IPv4: ' + e1.message + '; IPv6: ' + e2.message); }
  }
}

function identifyApp(url) {
  const m = (url || '').match(HOST_INFO_REGEX);
  if (!m) return null;
  const decoded = decodeURIComponent(m[1]);
  const appName = decoded.split('$')[0];
  return appName || null;
}

async function main() {
  console.log('=== Dynamic-port discovery POC ===\n');

  console.log('Step 1: tasklist for msedgewebview2.exe...');
  let pids;
  try { pids = listWebView2Pids(); }
  catch (e) { console.error('  FAILED: ' + e.message); process.exit(1); }
  console.log('  Found ' + pids.length + ' WebView2 PID(s): ' + (pids.join(', ') || '(none)'));
  if (pids.length === 0) {
    console.log('\n  No WebView2 processes running.');
    console.log('  Open Word/Excel/PowerPoint with --remote-debugging-port=0 first.');
    console.log('  Use probe/launch-office-dynamic.bat to launch them.');
    return;
  }

  console.log('\nStep 2: netstat -ano for PID->port mapping...');
  let portMap;
  try { portMap = listListeningPortsByPid(); }
  catch (e) { console.error('  FAILED: ' + e.message); process.exit(1); }

  const candidates = [];
  pids.forEach(function (pid) {
    const ports = portMap.get(pid);
    if (ports && ports.size > 0) {
      ports.forEach(function (port) { candidates.push({ pid: pid, port: port }); });
    }
  });
  console.log('  Candidate (PID, port) pairs: ' + candidates.length);
  candidates.forEach(function (c) {
    console.log('    PID ' + c.pid + ' -> port ' + c.port);
  });
  if (candidates.length === 0) {
    console.log('\n  WebView2 is running but no PID is listening on TCP.');
    console.log('  This means --remote-debugging-port=0 was NOT in effect when these');
    console.log('  WebView2 instances started. Close Office, set the env var, retry.');
    return;
  }

  console.log('\nStep 3: probing each candidate port for Claude targets...');
  const results = [];
  for (const c of candidates) {
    let targets;
    try {
      targets = await listTargets(c.port);
    } catch (e) {
      console.log('  port ' + c.port + ' (PID ' + c.pid + '): ' + e.message + ' - skip');
      continue;
    }
    const pages = (targets || []).filter(function (t) { return t.type === 'page'; });
    const claudePages = pages.filter(function (t) {
      return URL_PATTERN_PRIMARY.test(t.url || '') || URL_PATTERN_FALLBACK.test(t.url || '');
    });
    if (claudePages.length === 0) {
      console.log('  port ' + c.port + ' (PID ' + c.pid + '): ' + pages.length + ' page(s), 0 Claude - skip');
      continue;
    }
    claudePages.forEach(function (t) {
      const app = identifyApp(t.url);
      results.push({ pid: c.pid, port: c.port, app: app, url: t.url, title: t.title });
    });
  }

  console.log('\n=== Results ===');
  if (results.length === 0) {
    console.log('No Claude targets found on any WebView2 debug port.');
    console.log('Possible causes:');
    console.log('  - Claude add-in is not open in any Office app.');
    console.log('  - Office was not launched with --remote-debugging-port=0.');
    console.log('  - WebView2 used a different mechanism (sandbox?).');
    return;
  }

  const byApp = {};
  results.forEach(function (r) {
    const key = r.app || '(unknown)';
    if (!byApp[key]) byApp[key] = [];
    byApp[key].push(r);
  });

  console.log('Apps detected: ' + Object.keys(byApp).join(', '));
  console.log('');
  Object.keys(byApp).forEach(function (app) {
    console.log('[' + app + ']');
    byApp[app].forEach(function (r) {
      console.log('  PID ' + r.pid + ', port ' + r.port);
      console.log('    title: ' + (r.title || '(empty)'));
      console.log('    url:   ' + (r.url || '(empty)'));
    });
    console.log('');
  });

  const distinctApps = Object.keys(byApp).filter(function (k) { return k !== '(unknown)'; });
  console.log('=== Verdict ===');
  console.log('Distinct Office apps with Claude open: ' + distinctApps.length);
  if (distinctApps.length >= 2) {
    console.log('SUCCESS: dynamic-port architecture works. Multiple Office apps');
    console.log('         simultaneously expose Claude on independent ports.');
  } else if (distinctApps.length === 1) {
    console.log('PARTIAL: only one app detected. Open another Office app with the');
    console.log('         Claude add-in to fully validate.');
  } else {
    console.log('UNCLEAR: targets found but app identification failed. Check the');
    console.log('         _host_Info= parameter in the URLs above.');
  }
}

main().catch(function (e) {
  console.error('\nFATAL: ' + e.message);
  process.exit(1);
});
