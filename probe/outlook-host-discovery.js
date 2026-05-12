#!/usr/bin/env node
/**
 * Outlook host discovery probe (M0 for v0.3.0 Outlook expansion).
 *
 * Answers the 6 questions in docs/OUTLOOK-EXPANSION-PLAN.md section 2:
 *
 *   Q1. Does Claude in Outlook run inside WebView2 at all? (msedgewebview2.exe
 *       must be a descendant of OUTLOOK.EXE.)
 *   Q2. Is the URL the same as Word/Excel/PowerPoint (pivot.claude.ai)?
 *   Q3. What is the exact _host_Info= value?
 *   Q4. Does --remote-debugging-port=0 get honored by Outlook's WebView2?
 *   Q5. Classic OUTLOOK.EXE vs New Outlook (olk.exe) - which one(s) work?
 *   Q6. Does Outlook share a WebView2 host pool with Word/Excel/PowerPoint?
 *
 * Run AFTER outlook-host-discovery.bat has launched Outlook with the dynamic
 * debug port and the user has manually opened the Claude add-in pane.
 *
 * Usage:
 *   node outlook-host-discovery.js
 *
 * No deps. Windows only.
 */

const http = require('http');
const { execFileSync } = require('child_process');

const URL_PATTERN_PRIMARY = /\/\/([^\/]*\.)?pivot\.claude\.ai(\/|$|\?|#)/i;
const URL_PATTERN_FALLBACK = /\/\/([^\/]*\.)?claude\.ai(\/|$|\?|#)/i;
const HOST_INFO_REGEX = /[?&]_host_Info=([^&#]+)/i;

const OFFICE_PROCESS_NAMES = ['OUTLOOK.EXE', 'olk.exe', 'WINWORD.EXE', 'EXCEL.EXE', 'POWERPNT.EXE'];

function tasklistAll() {
  // Returns Map<pid, {name, pid}>.
  const out = execFileSync('tasklist', ['/FO', 'CSV', '/NH'], {
    encoding: 'utf8',
    windowsHide: true,
  });
  const map = new Map();
  out.split(/\r?\n/).forEach(function (line) {
    const m = line.match(/^"([^"]+)","(\d+)"/);
    if (!m) return;
    const name = m[1];
    const pid = parseInt(m[2], 10);
    map.set(pid, { name: name, pid: pid });
  });
  return map;
}

function parentMap() {
  // Get parent PID for every process via PowerShell. Returns Map<pid, parentPid>.
  // wmic is deprecated on Win11; using CIM is the modern path.
  const ps = 'Get-CimInstance Win32_Process | ForEach-Object { "$($_.ProcessId),$($_.ParentProcessId)" }';
  const out = execFileSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', ps], {
    encoding: 'utf8',
    windowsHide: true,
  });
  const map = new Map();
  out.split(/\r?\n/).forEach(function (line) {
    const m = line.match(/^(\d+),(\d+)\s*$/);
    if (!m) return;
    map.set(parseInt(m[1], 10), parseInt(m[2], 10));
  });
  return map;
}

function findAncestorOfficeProcess(pid, allProcs, parents) {
  // Walk up the parent chain looking for any Office host process.
  // Stops at depth 10 to defend against cycles or runaway loops.
  let cur = pid;
  for (let depth = 0; depth < 10; depth++) {
    const parentPid = parents.get(cur);
    if (parentPid === undefined || parentPid === 0) return null;
    const p = allProcs.get(parentPid);
    if (p) {
      const upper = p.name.toUpperCase();
      for (const want of OFFICE_PROCESS_NAMES) {
        if (upper === want.toUpperCase()) return { name: p.name, pid: parentPid, depth: depth + 1 };
      }
    }
    cur = parentPid;
  }
  return null;
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
  try { return await tryListTargetsOnHost('127.0.0.1', port); }
  catch (e1) {
    try { return await tryListTargetsOnHost('::1', port); }
    catch (e2) { throw new Error('IPv4: ' + e1.message + '; IPv6: ' + e2.message); }
  }
}

function decodeHostInfo(url) {
  const m = (url || '').match(HOST_INFO_REGEX);
  if (!m) return null;
  return decodeURIComponent(m[1]);
}

function appNameFromHostInfo(decoded) {
  if (!decoded) return null;
  return decoded.split('$')[0] || null;
}

async function main() {
  console.log('=== Outlook host discovery probe (M0) ===\n');

  console.log('Step 1: tasklist (all processes)...');
  const allProcs = tasklistAll();
  console.log('  ' + allProcs.size + ' processes total.');

  const webview2Pids = [];
  const officePidsByName = new Map();
  allProcs.forEach(function (p) {
    const upper = p.name.toUpperCase();
    if (upper === 'MSEDGEWEBVIEW2.EXE') webview2Pids.push(p.pid);
    for (const want of OFFICE_PROCESS_NAMES) {
      if (upper === want.toUpperCase()) {
        if (!officePidsByName.has(want)) officePidsByName.set(want, []);
        officePidsByName.get(want).push(p.pid);
      }
    }
  });

  console.log('  msedgewebview2.exe PIDs: ' + (webview2Pids.join(', ') || '(none)'));
  OFFICE_PROCESS_NAMES.forEach(function (n) {
    const pids = officePidsByName.get(n) || [];
    if (pids.length > 0) console.log('  ' + n + ' PIDs: ' + pids.join(', '));
  });

  const outlookClassicPids = officePidsByName.get('OUTLOOK.EXE') || [];
  const outlookNewPids = officePidsByName.get('olk.exe') || [];

  if (outlookClassicPids.length === 0 && outlookNewPids.length === 0) {
    console.log('\n  No Outlook is running. Launch via outlook-host-discovery.bat first,');
    console.log('  then open the Claude add-in pane, then run this script.');
    return;
  }

  console.log('\nStep 2: parent process map (PowerShell Get-CimInstance)...');
  let parents;
  try { parents = parentMap(); }
  catch (e) { console.error('  FAILED: ' + e.message); process.exit(1); }
  console.log('  Parent PID known for ' + parents.size + ' processes.');

  console.log('\nStep 3: classify each WebView2 PID by its Office ancestor...');
  const webview2ByOffice = new Map(); // officeName -> array of {webviewPid, officePid, depth}
  webview2Pids.forEach(function (wvPid) {
    const ancestor = findAncestorOfficeProcess(wvPid, allProcs, parents);
    if (!ancestor) {
      console.log('  WebView2 PID ' + wvPid + ' -> no Office ancestor (Teams/Edge/other host?)');
      return;
    }
    console.log('  WebView2 PID ' + wvPid + ' -> ancestor ' + ancestor.name + ' (PID ' + ancestor.pid + ', depth ' + ancestor.depth + ')');
    if (!webview2ByOffice.has(ancestor.name)) webview2ByOffice.set(ancestor.name, []);
    webview2ByOffice.get(ancestor.name).push({ webviewPid: wvPid, officePid: ancestor.pid, depth: ancestor.depth });
  });

  console.log('\nStep 4: netstat -ano for LISTENING ports per WebView2 PID...');
  const portMap = listListeningPortsByPid();
  const candidates = [];
  webview2Pids.forEach(function (wvPid) {
    const ports = portMap.get(wvPid);
    if (!ports || ports.size === 0) {
      console.log('  WebView2 PID ' + wvPid + ' -> no LISTENING TCP port');
      return;
    }
    ports.forEach(function (port) {
      console.log('  WebView2 PID ' + wvPid + ' -> port ' + port);
      candidates.push({ pid: wvPid, port: port });
    });
  });

  console.log('\nStep 5: probe each candidate port for Claude targets...');
  const results = [];
  for (const c of candidates) {
    let targets;
    try { targets = await listTargets(c.port); }
    catch (e) { console.log('  port ' + c.port + ' (PID ' + c.pid + '): ' + e.message + ' - skip'); continue; }
    const pages = (targets || []).filter(function (t) { return t.type === 'page'; });
    const claudePages = pages.filter(function (t) {
      return URL_PATTERN_PRIMARY.test(t.url || '') || URL_PATTERN_FALLBACK.test(t.url || '');
    });
    if (claudePages.length === 0) {
      console.log('  port ' + c.port + ' (PID ' + c.pid + '): ' + pages.length + ' page(s), 0 Claude');
      continue;
    }
    claudePages.forEach(function (t) {
      const hostInfo = decodeHostInfo(t.url);
      const appName = appNameFromHostInfo(hostInfo);
      const ancestor = findAncestorOfficeProcess(c.pid, allProcs, parents);
      const matchedPrimary = URL_PATTERN_PRIMARY.test(t.url || '');
      results.push({
        webviewPid: c.pid,
        port: c.port,
        officeAncestor: ancestor ? ancestor.name : null,
        url: t.url,
        title: t.title,
        hostInfo: hostInfo,
        hostInfoAppName: appName,
        urlPrimary: matchedPrimary,
      });
    });
  }

  console.log('\n=== Raw findings ===');
  if (results.length === 0) {
    console.log('No Claude targets found on any WebView2 debug port.');
  } else {
    results.forEach(function (r, i) {
      console.log('Target #' + (i + 1) + ':');
      console.log('  WebView2 PID:      ' + r.webviewPid);
      console.log('  Office ancestor:   ' + (r.officeAncestor || '(none)'));
      console.log('  Debug port:        ' + r.port);
      console.log('  URL primary match: ' + (r.urlPrimary ? 'pivot.claude.ai' : 'claude.ai (fallback)'));
      console.log('  Full URL:          ' + (r.url || '(empty)'));
      console.log('  Title:             ' + (r.title || '(empty)'));
      console.log('  _host_Info=:       ' + (r.hostInfo || '(missing)'));
      console.log('  App name from it:  ' + (r.hostInfoAppName || '(missing)'));
      console.log('');
    });
  }

  // Q1-Q6 answers
  console.log('=== Answers to M0 questions ===\n');

  // Q1: WebView2 host?
  const outlookWebview = webview2ByOffice.get('OUTLOOK.EXE') || [];
  const olkWebview = webview2ByOffice.get('olk.exe') || [];
  console.log('Q1. Does Claude in Outlook run inside WebView2?');
  if (outlookClassicPids.length > 0) {
    if (outlookWebview.length > 0) {
      console.log('    YES (classic): ' + outlookWebview.length + ' WebView2 process(es) descend from OUTLOOK.EXE.');
    } else {
      console.log('    NO (classic): OUTLOOK.EXE is running but no msedgewebview2.exe descends from it.');
      console.log('         Possible: Claude add-in not opened yet, or Outlook uses a different web host (IE legacy, embedded).');
    }
  }
  if (outlookNewPids.length > 0) {
    if (olkWebview.length > 0) {
      console.log('    YES (New): ' + olkWebview.length + ' WebView2 process(es) descend from olk.exe.');
    } else {
      console.log('    NEW OUTLOOK detected but no WebView2 descendant - olk.exe likely uses Edge WebView under its app container, or no Claude add-in is open.');
    }
  }
  console.log('');

  // Q2: URL same as Word/Excel/PowerPoint?
  console.log('Q2. Is the URL the same as Word/Excel/PowerPoint (pivot.claude.ai)?');
  const outlookResults = results.filter(function (r) { return r.officeAncestor === 'OUTLOOK.EXE' || r.officeAncestor === 'olk.exe'; });
  if (outlookResults.length === 0) {
    console.log('    UNKNOWN: no Claude target was found inside Outlook.');
  } else {
    const allPrimary = outlookResults.every(function (r) { return r.urlPrimary; });
    if (allPrimary) {
      console.log('    YES: all Outlook Claude targets match pivot.claude.ai exactly.');
    } else {
      console.log('    PARTIAL: some Outlook Claude targets did NOT match pivot.claude.ai. See URLs above.');
    }
  }
  console.log('');

  // Q3: exact _host_Info
  console.log('Q3. Exact _host_Info= value(s) for Outlook:');
  if (outlookResults.length === 0) {
    console.log('    UNKNOWN: no Outlook Claude target.');
  } else {
    const seen = new Set();
    outlookResults.forEach(function (r) {
      if (r.hostInfo && !seen.has(r.hostInfo)) { seen.add(r.hostInfo); console.log('    ' + r.hostInfo); }
      else if (!r.hostInfo) { console.log('    (target had no _host_Info= parameter!)'); }
    });
  }
  console.log('');

  // Q4: dynamic port honored?
  console.log('Q4. Does --remote-debugging-port=0 get honored by Outlook?');
  if (outlookWebview.length === 0 && olkWebview.length === 0) {
    console.log('    UNKNOWN: no Outlook-descended WebView2 to check.');
  } else {
    const outlookDescendantPids = new Set();
    outlookWebview.forEach(function (e) { outlookDescendantPids.add(e.webviewPid); });
    olkWebview.forEach(function (e) { outlookDescendantPids.add(e.webviewPid); });
    let listening = 0;
    outlookDescendantPids.forEach(function (pid) { if (portMap.has(pid) && portMap.get(pid).size > 0) listening++; });
    if (listening > 0) {
      console.log('    YES: ' + listening + ' of ' + outlookDescendantPids.size + ' Outlook-descended WebView2 PIDs are LISTENING on a dynamic TCP port.');
    } else {
      console.log('    NO: Outlook-descended WebView2 is running but no LISTENING port found.');
      console.log('         WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS likely did NOT reach the process.');
      console.log('         For classic, confirm you launched via outlook-host-discovery.bat (not taskbar/start menu).');
      console.log('         For New Outlook (olk.exe / Appx), the app container may strip the env var - that is the expected failure.');
    }
  }
  console.log('');

  // Q5: classic vs new
  console.log('Q5. Classic OUTLOOK.EXE vs New olk.exe:');
  console.log('    Classic running: ' + (outlookClassicPids.length > 0 ? 'yes (PIDs ' + outlookClassicPids.join(',') + ')' : 'no'));
  console.log('    New running:     ' + (outlookNewPids.length > 0 ? 'yes (PIDs ' + outlookNewPids.join(',') + ')' : 'no'));
  console.log('    Classic Claude target found:  ' + (results.some(function (r) { return r.officeAncestor === 'OUTLOOK.EXE'; }) ? 'yes' : 'no'));
  console.log('    New Outlook Claude target found: ' + (results.some(function (r) { return r.officeAncestor === 'olk.exe'; }) ? 'yes' : 'no'));
  console.log('');

  // Q6: shared host pool?
  console.log('Q6. Does Outlook share a WebView2 host pool with Word/Excel/PowerPoint?');
  const otherOfficeNames = ['WINWORD.EXE', 'EXCEL.EXE', 'POWERPNT.EXE'];
  const anyOtherRunning = otherOfficeNames.some(function (n) { return (officePidsByName.get(n) || []).length > 0; });
  if (!anyOtherRunning) {
    console.log('    UNKNOWN: no Word/Excel/PowerPoint is running. To answer Q6, open Word with the existing word-wrapper.bat');
    console.log('             (or launch all four via probe/launch-office-dynamic.bat plus this one), open Claude in both, rerun.');
  } else {
    const outlookWebviewPidSet = new Set();
    outlookWebview.forEach(function (e) { outlookWebviewPidSet.add(e.webviewPid); });
    olkWebview.forEach(function (e) { outlookWebviewPidSet.add(e.webviewPid); });
    const otherWebviewPidSet = new Set();
    otherOfficeNames.forEach(function (n) {
      (webview2ByOffice.get(n) || []).forEach(function (e) { otherWebviewPidSet.add(e.webviewPid); });
    });
    let overlap = false;
    outlookWebviewPidSet.forEach(function (p) { if (otherWebviewPidSet.has(p)) overlap = true; });
    if (overlap) {
      console.log('    YES: at least one msedgewebview2.exe PID has BOTH Outlook and another Office app as ancestors via different parent chains.');
      console.log('         (This is uncommon; double-check parents map.)');
    } else {
      // Check the inverse: do any two Office apps' Claude targets share a port? That indicates host-pool sharing
      // even when ancestor walks differ - host pool collapses children into one msedgewebview2 worker.
      const outlookTargetPorts = new Set();
      const otherTargetPorts = new Set();
      results.forEach(function (r) {
        if (r.officeAncestor === 'OUTLOOK.EXE' || r.officeAncestor === 'olk.exe') outlookTargetPorts.add(r.port);
        else if (r.officeAncestor) otherTargetPorts.add(r.port);
      });
      let portOverlap = false;
      outlookTargetPorts.forEach(function (p) { if (otherTargetPorts.has(p)) portOverlap = true; });
      if (portOverlap) {
        console.log('    YES (via shared port): an Outlook Claude target and a non-Outlook Claude target are on the SAME debug port.');
        console.log('         The host pool is shared. The injector must filter targets by URL, not by port-owner identity.');
      } else {
        console.log('    NO: Outlook WebView2 PIDs and other-Office WebView2 PIDs are disjoint, and Claude targets live on disjoint ports.');
      }
    }
  }
  console.log('');

  // Go/no-go
  console.log('=== Go / no-go suggestion ===');
  const q1Pass = outlookWebview.length > 0 || olkWebview.length > 0;
  const q2Pass = outlookResults.length > 0 && outlookResults.every(function (r) { return r.urlPrimary; });
  const q3Pass = outlookResults.length > 0 && outlookResults.every(function (r) { return r.hostInfo && r.hostInfoAppName; });
  if (q1Pass && q2Pass && q3Pass) {
    console.log('GO: Q1-Q3 all positive. Architecture from v0.2.x carries over for Outlook.');
    console.log('     Confirm Q4 also positive (dynamic port honored). Document Q5, Q6 for design notes.');
  } else {
    console.log('NO-GO or PARTIAL: at least one blocker. See per-question verdict above.');
    console.log('     Required for go: Q1=yes, Q2=yes (pivot.claude.ai), Q3=yes (_host_Info present and starts with "Outlook").');
  }
}

main().catch(function (e) {
  console.error('\nFATAL: ' + e.message);
  process.exit(1);
});
