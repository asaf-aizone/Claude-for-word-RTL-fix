'use strict';

/**
 * Single source of truth for Office app metadata.
 *
 * No port numbers here - in v0.2.0 ports are dynamic (--remote-debugging-port=0).
 * Apps are identified at runtime via the _host_Info= URL parameter that Office
 * appends to the Claude add-in URL.
 *
 * _host_Info format observed in M0 POC: <AppName>$Win32$<office-version>$<locale>$$$$<extra>
 *   Word        -> _host_Info=Word$Win32$16.01$he-IL$$$$19
 *   Excel       -> _host_Info=Excel$Win32$16.01$he-IL$$$$19
 *   PowerPoint  -> _host_Info=Powerpoint$Win32$16.01$he-IL$$$$19  (note: capital P only)
 */

const APPS = [
  { name: 'Word',       processName: 'WINWORD.EXE',  urlHostInfo: 'Word' },
  { name: 'Excel',      processName: 'EXCEL.EXE',    urlHostInfo: 'Excel' },
  { name: 'PowerPoint', processName: 'POWERPNT.EXE', urlHostInfo: 'Powerpoint' },
];

const HOST_INFO_REGEX = /[?&]_host_Info=([^&#]+)/i;

// Apps that are NEVER attached automatically. Each one requires an explicit
// per-launch opt-in (a flag file written by the matching wrapper). Lowercased.
// Outlook is here because email contents become panel DOM when the user runs
// "Summarize this email" or "Draft a reply", and silent CDP attach to that
// DOM is a higher-grade exposure than for Word/Excel/PowerPoint.
// See docs/OUTLOOK-EXPANSION-PLAN.md sections 3 and 4.1, and probe/README.md
// "silent CDP attach" for the M0 evidence that motivated this.
const BLOCKED_HOST_INFO_KEYS = new Set(['outlook']);

function getHostInfoKey(url) {
  if (!url) return null;
  const m = url.match(HOST_INFO_REGEX);
  if (!m) return null;
  let decoded;
  try { decoded = decodeURIComponent(m[1]); } catch (e) { return null; }
  const k = decoded.split('$')[0];
  return k ? k.toLowerCase() : null;
}

function identifyAppFromUrl(url) {
  const appKey = getHostInfoKey(url);
  if (!appKey) return null;
  return APPS.find(function (a) { return a.urlHostInfo.toLowerCase() === appKey; }) || null;
}

// Returns true if this URL should not be attached automatically. optInKeys is
// a Set<string> of lowercased app keys the user has explicitly consented to
// for the current injector session.
function isHostInfoBlocked(url, optInKeys) {
  const k = getHostInfoKey(url);
  if (!k) return false;
  if (!BLOCKED_HOST_INFO_KEYS.has(k)) return false;
  if (optInKeys && optInKeys.has(k)) return false;
  return true;
}

module.exports = {
  APPS: APPS,
  BLOCKED_HOST_INFO_KEYS: BLOCKED_HOST_INFO_KEYS,
  identifyAppFromUrl: identifyAppFromUrl,
  getHostInfoKey: getHostInfoKey,
  isHostInfoBlocked: isHostInfoBlocked,
};
