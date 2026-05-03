'use strict';

/**
 * Port discovery for Office WebView2 hosts.
 *
 * In v0.2.0 each Office app launches WebView2 with --remote-debugging-port=0,
 * so the port is dynamic and unknown until the process binds it. This module
 * discovers active CDP debug ports by:
 *
 *   1. tasklist  - find msedgewebview2.exe PIDs
 *   2. netstat   - map each PID to the TCP port(s) it is LISTENING on
 *   3. /json/list - probe each candidate port for Claude targets
 *
 * Backward compatibility: the legacy v0.1.x value
 * --remote-debugging-port=9222 is still probed alongside dynamic ports, so a
 * mid-migration user where install.bat hasn't run yet still gets RTL.
 *
 * The Claude target is identified via the _host_Info= URL parameter
 * (see lib/office-apps.js).
 *
 * Windows-only. Uses tasklist + netstat from the Windows distribution.
 * Both are present on every supported Windows version since Vista.
 */

const http = require('http');
const { execFile } = require('child_process');
const officeApps = require('../lib/office-apps');

const URL_PATTERN_PRIMARY = /\/\/([^\/]*\.)?pivot\.claude\.ai(\/|$|\?|#)/i;
const URL_PATTERN_FALLBACK = /\/\/([^\/]*\.)?claude\.ai(\/|$|\?|#)/i;
const LEGACY_PORT = 9222;
const PROBE_TIMEOUT_MS = 1500;

function execFileP(cmd, args) {
  return new Promise(function (resolve, reject) {
    execFile(cmd, args, { encoding: 'utf8', windowsHide: true, maxBuffer: 8 * 1024 * 1024 }, function (err, stdout, stderr) {
      if (err) return reject(err);
      resolve(stdout);
    });
  });
}

async function listWebView2Pids() {
  try {
    const out = await execFileP('tasklist', ['/FI', 'IMAGENAME eq msedgewebview2.exe', '/FO', 'CSV', '/NH']);
    const pids = [];
    out.split(/\r?\n/).forEach(function (line) {
      const m = line.match(/^"[^"]*","(\d+)"/);
      if (m) pids.push(parseInt(m[1], 10));
    });
    return pids;
  } catch (e) {
    return [];
  }
}

async function listListeningPortsByPid() {
  try {
    const out = await execFileP('netstat', ['-ano']);
    const map = new Map();
    out.split(/\r?\n/).forEach(function (line) {
      // Match both TCP (IPv4) and TCPv6 (IPv6). WebView2 on Windows often
      // binds IPv6-only ([::1]). The localhost-vs-loopback IPv4/IPv6 split
      // bit us in v0.1.3; that's why we accept both families.
      const m = line.match(/^\s*TCP(?:v6)?\s+(\S+)\s+\S+\s+LISTENING\s+(\d+)/);
      if (!m) return;
      const local = m[1];
      const pid = parseInt(m[2], 10);
      // Last :digits is the port (works for [::]:51244 and 127.0.0.1:51244).
      const portMatch = local.match(/:(\d+)$/);
      if (!portMatch) return;
      const port = parseInt(portMatch[1], 10);
      if (!map.has(pid)) map.set(pid, new Set());
      map.get(pid).add(port);
    });
    return map;
  } catch (e) {
    return new Map();
  }
}

function tryListTargetsOnHost(host, port) {
  return new Promise(function (resolve, reject) {
    const req = http.get({ host: host, port: port, path: '/json/list', timeout: PROBE_TIMEOUT_MS }, function (res) {
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

async function listTargetsAtPort(port) {
  // Probe both loopback families; WebView2 historically binds [::1] only.
  try { return await tryListTargetsOnHost('127.0.0.1', port); }
  catch (e1) {
    try { return await tryListTargetsOnHost('::1', port); }
    catch (e2) { return null; }
  }
}

function isClaudeUrl(url) {
  return URL_PATTERN_PRIMARY.test(url || '') || URL_PATTERN_FALLBACK.test(url || '');
}

/**
 * Discover all currently-listening CDP ports owned by msedgewebview2 processes,
 * plus the legacy 9222 fallback. Returns Set<number>.
 */
async function discoverPorts() {
  const ports = new Set();
  ports.add(LEGACY_PORT);
  const pids = await listWebView2Pids();
  if (pids.length === 0) return ports;
  const portMap = await listListeningPortsByPid();
  pids.forEach(function (pid) {
    const owned = portMap.get(pid);
    if (owned) owned.forEach(function (p) { ports.add(p); });
  });
  return ports;
}

/**
 * Probe each known port for Claude targets, identify the Office app for each,
 * and return a flat list. Used by inject.js on each tick.
 *
 * Returns: Array<{ port: number, target: {id, url, webSocketDebuggerUrl, ...}, app: {name, ...} | null }>
 */
async function discoverActiveTargets() {
  const ports = await discoverPorts();
  const probes = Array.from(ports).map(function (port) {
    return listTargetsAtPort(port).then(function (targets) { return { port: port, targets: targets }; });
  });
  const settled = await Promise.all(probes);
  const results = [];
  settled.forEach(function (entry) {
    if (!entry.targets || !Array.isArray(entry.targets)) return;
    entry.targets.forEach(function (t) {
      if (t.type !== 'page') return;
      if (!isClaudeUrl(t.url)) return;
      const app = officeApps.identifyAppFromUrl(t.url);
      results.push({ port: entry.port, target: t, app: app });
    });
  });
  return results;
}

module.exports = {
  discoverPorts: discoverPorts,
  discoverActiveTargets: discoverActiveTargets,
  isClaudeUrl: isClaudeUrl,
  LEGACY_PORT: LEGACY_PORT,
};
