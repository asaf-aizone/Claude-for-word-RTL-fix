#!/usr/bin/env node
/**
 * Probe script for Claude Office add-ins.
 *
 * Queries a WebView2 debug port and lists every page target it exposes,
 * so we can see (a) whether Claude actually loads inside Excel/PowerPoint,
 * (b) what URL it uses, and (c) whether it matches the patterns the Word
 * injector relies on (pivot.claude.ai / claude.ai).
 *
 * No dependencies - uses only Node's built-in http module.
 *
 * Usage:
 *   node probe.js              -> probes 9222 (Word), 9223 (Excel), 9224 (PowerPoint)
 *   node probe.js 9223         -> probes a single port
 *   node probe.js 9222 9224    -> probes multiple specific ports
 */

const http = require('http');

const URL_PATTERN_PRIMARY = /\/\/([^\/]*\.)?pivot\.claude\.ai(\/|$|\?|#)/i;
const URL_PATTERN_FALLBACK = /\/\/([^\/]*\.)?claude\.ai(\/|$|\?|#)/i;

const APP_NAMES = { 9222: 'Word', 9223: 'Excel', 9224: 'PowerPoint' };

function listTargets(port) {
  return new Promise(function (resolve, reject) {
    var req = http.get({ host: 'localhost', port: port, path: '/json/list', timeout: 3000 }, function (res) {
      var body = '';
      res.on('data', function (c) { body += c; });
      res.on('end', function () {
        try { resolve(JSON.parse(body)); }
        catch (e) { reject(new Error('Invalid JSON from port ' + port + ': ' + e.message)); }
      });
    });
    req.on('timeout', function () { req.destroy(new Error('Timeout')); });
    req.on('error', reject);
  });
}

function describeTarget(t) {
  return {
    type: t.type,
    title: t.title,
    url: t.url,
    matchesPivot: URL_PATTERN_PRIMARY.test(t.url || ''),
    matchesClaude: URL_PATTERN_FALLBACK.test(t.url || ''),
  };
}

async function probePort(port) {
  var app = APP_NAMES[port] || '(unknown)';
  console.log('\n=== Port ' + port + ' (' + app + ') ===');
  var targets;
  try {
    targets = await listTargets(port);
  } catch (e) {
    console.log('  [not reachable] ' + e.message);
    console.log('  - App may not be running with WebView2 debug port enabled.');
    console.log('  - Or no WebView2 panel is currently active (open the Claude add-in).');
    return;
  }
  if (!Array.isArray(targets) || targets.length === 0) {
    console.log('  [no targets] WebView2 debug is listening but nothing is attached.');
    return;
  }
  console.log('  Total targets: ' + targets.length);
  var pages = targets.filter(function (t) { return t.type === 'page'; });
  console.log('  Page targets: ' + pages.length);
  console.log('');
  pages.forEach(function (t, i) {
    var d = describeTarget(t);
    var tag = d.matchesPivot ? '[MATCH: pivot.claude.ai]'
           : d.matchesClaude ? '[MATCH: claude.ai (fallback)]'
           : '[no match]';
    console.log('  #' + (i + 1) + ' ' + tag);
    console.log('     title: ' + (d.title || '(empty)'));
    console.log('     url:   ' + (d.url || '(empty)'));
  });

  var nonPage = targets.filter(function (t) { return t.type !== 'page'; });
  if (nonPage.length > 0) {
    console.log('');
    console.log('  Non-page targets (' + nonPage.length + '): '
      + nonPage.map(function (t) { return t.type; }).join(', '));
  }

  var matched = pages.filter(function (t) {
    return URL_PATTERN_PRIMARY.test(t.url || '') || URL_PATTERN_FALLBACK.test(t.url || '');
  });
  console.log('');
  if (matched.length > 0) {
    console.log('  >> Injector WOULD attach to ' + matched.length + ' target(s) on port ' + port + '.');
  } else {
    console.log('  >> Injector would NOT attach - no Claude URL matched.');
    console.log('     (This means either Claude add-in is not open, or the add-in uses a different URL.)');
  }
}

async function main() {
  var args = process.argv.slice(2);
  var ports = args.length > 0 ? args.map(Number) : [9222, 9223, 9224];
  ports = ports.filter(function (p) { return Number.isFinite(p) && p > 0; });
  if (ports.length === 0) {
    console.error('No valid ports provided.');
    process.exit(1);
  }
  console.log('Probing ports: ' + ports.join(', '));
  for (var i = 0; i < ports.length; i++) {
    await probePort(ports[i]);
  }
  console.log('');
  console.log('Done. Share this output so the injector can be extended accordingly.');
}

main().catch(function (e) { console.error(e); process.exit(1); });
