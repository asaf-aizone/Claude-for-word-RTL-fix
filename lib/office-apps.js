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

function identifyAppFromUrl(url) {
  if (!url) return null;
  const m = url.match(HOST_INFO_REGEX);
  if (!m) return null;
  let decoded;
  try { decoded = decodeURIComponent(m[1]); } catch (e) { return null; }
  const appKey = decoded.split('$')[0];
  if (!appKey) return null;
  return APPS.find(function (a) { return a.urlHostInfo.toLowerCase() === appKey.toLowerCase(); }) || null;
}

module.exports = {
  APPS: APPS,
  identifyAppFromUrl: identifyAppFromUrl,
};
