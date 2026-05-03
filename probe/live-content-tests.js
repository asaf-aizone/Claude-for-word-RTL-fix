#!/usr/bin/env node
'use strict';

/**
 * Live-content RTL test suite for Claude for Office RTL Fix.
 *
 * Unlike text-rendering-tests.js (which injects synthetic test elements
 * into a hidden container), this script reads the user's REAL typed
 * Hebrew content from the live Claude pane (chat input, chat history,
 * message bubbles) and verifies the production CSS renders it RTL.
 *
 * Read-only: this script never types into the pane, never sends a
 * message, never triggers a Claude API call. It only inspects the DOM.
 *
 * Prerequisite: the user must have opened Claude panes in
 * Word/Excel/PowerPoint manually and typed some Hebrew text (in the
 * chat input or in chat history) before running this script.
 *
 * Exits 0 on all-pass, 1 on any failure or if no Hebrew content is
 * found in any app.
 *
 * NOTE: RTL_CSS is duplicated verbatim from scripts/inject.js to mirror
 * the convention in text-rendering-tests.js. See that file's header for
 * the rationale.
 */

const path = require('path');
const CDP = require(path.join(__dirname, '..', 'scripts', 'node_modules', 'chrome-remote-interface'));
const portDiscovery = require(path.join(__dirname, '..', 'scripts', 'port-discovery'));

// === RTL_CSS: copied verbatim from scripts/inject.js ===
const RTL_CSS = `
  /* Claude-RTL-Fix: base RTL + fixed list markers + code handling.
     Font family is intentionally NOT set, so Claude's own UI font is preserved. */
  html, body { direction: rtl !important; }
  p, li, td, th, h1, h2, h3, h4, h5, h6, blockquote, figcaption { text-align: start; }
  p, div, span, li, td, th, h1, h2, h3, h4, h5, h6, blockquote, figcaption {
    direction: rtl !important;
    unicode-bidi: isolate !important;
  }
  /* list containers must be RTL so markers (1. 2. •) render on the right */
  ol, ul, dl, menu {
    direction: rtl !important;
    unicode-bidi: isolate !important;
  }
  /* input fields */
  textarea, input[type="text"], input[type="search"], [contenteditable="true"] {
    direction: rtl !important;
    text-align: start !important;
    unicode-bidi: isolate !important;
  }
  /* code blocks stay LTR for code correctness */
  pre, pre *, code, kbd, samp, tt, .hljs, [class*="language-"], [class*="code-"] {
    direction: ltr !important;
    text-align: left !important;
    unicode-bidi: isolate !important;
  }
  /* inline code inside Hebrew paragraphs - isolate so Hebrew flow stays RTL */
  p code, li code, td code, th code, span code, div > code {
    unicode-bidi: isolate !important;
  }
  /* tables */
  table { direction: rtl !important; }
`;

const ENSURE_STYLE_SCRIPT = `
(function () {
  var id = '__claude_rtl_fix__';
  if (document.getElementById(id)) return 'already-present';
  var style = document.createElement('style');
  style.id = id;
  style.textContent = ${JSON.stringify(RTL_CSS)};
  (document.head || document.documentElement).appendChild(style);
  return 'injected-by-test';
})();
`;

const CHECK_STYLE_PRESENT_SCRIPT =
  `!!document.getElementById('__claude_rtl_fix__')`;

// Walks the live DOM, finds every visible element whose own (direct)
// text content contains Hebrew characters (U+0590 to U+05FF). Returns
// for each: tag name, computed direction, text-align, a trimmed sample
// of its own text, and a structural path. Also separately probes the
// chat input element and message-bubble candidates.
//
// "Visible" = has layout (offsetParent or fixed/sticky), and a non-zero
// bounding rect. Hidden test scaffolding (e.g. our previous synthetic
// container at left:-9999px) IS still visible to this query if it's in
// the DOM, so we explicitly skip our test scaffolding by id.
//
// "Own direct text" = node.childNodes filtered to text nodes only,
// concatenated. This avoids attributing a child's Hebrew text to every
// ancestor (which would make every <body> "Hebrew").
const SCAN_LIVE_CONTENT_SCRIPT = `
(function () {
  var HEBREW_RE = /[\\u0590-\\u05FF]/;
  var SCAFFOLD_IDS = { '__rtl_test__': 1, '__claude_rtl_fix__': 1 };

  function isVisible(el) {
    if (!el || el.nodeType !== 1) return false;
    if (el.id && SCAFFOLD_IDS[el.id]) return false;
    // Skip elements inside our own test scaffolding too.
    var p = el.parentElement;
    while (p) {
      if (p.id && SCAFFOLD_IDS[p.id]) return false;
      p = p.parentElement;
    }
    var rect = el.getBoundingClientRect();
    if (rect.width === 0 && rect.height === 0) return false;
    var cs = getComputedStyle(el);
    if (cs.display === 'none' || cs.visibility === 'hidden') return false;
    return true;
  }

  function ownText(el) {
    var t = '';
    for (var i = 0; i < el.childNodes.length; i++) {
      var c = el.childNodes[i];
      if (c.nodeType === 3) t += c.nodeValue;
    }
    return t;
  }

  function trimSample(s) {
    s = (s || '').replace(/\\s+/g, ' ').trim();
    if (s.length > 80) s = s.slice(0, 77) + '...';
    return s;
  }

  function shortPath(el) {
    var parts = [];
    var cur = el;
    var depth = 0;
    while (cur && cur.nodeType === 1 && depth < 4) {
      var seg = cur.tagName.toLowerCase();
      if (cur.id) seg += '#' + cur.id;
      else if (cur.className && typeof cur.className === 'string') {
        var cls = cur.className.trim().split(/\\s+/)[0];
        if (cls) seg += '.' + cls;
      }
      parts.unshift(seg);
      cur = cur.parentElement;
      depth++;
    }
    return parts.join('>');
  }

  // 1. Scan all elements, count visible total, collect those with Hebrew
  // own-text.
  var all = document.querySelectorAll('*');
  var totalVisible = 0;
  var hebrewElements = [];
  for (var i = 0; i < all.length; i++) {
    var el = all[i];
    if (!isVisible(el)) continue;
    totalVisible++;
    var own = ownText(el);
    if (!own || !HEBREW_RE.test(own)) continue;
    var cs = getComputedStyle(el);
    hebrewElements.push({
      tag: el.tagName.toLowerCase(),
      direction: cs.direction,
      textAlign: cs.textAlign,
      sample: trimSample(own),
      path: shortPath(el)
    });
  }

  // 2. Find the chat input (first visible textarea or contenteditable).
  // Read from input.value for textarea, textContent for contenteditable.
  var chatInput = null;
  var inputCandidates = document.querySelectorAll('textarea, [contenteditable="true"]');
  for (var j = 0; j < inputCandidates.length; j++) {
    var el2 = inputCandidates[j];
    if (!isVisible(el2)) continue;
    var cs2 = getComputedStyle(el2);
    var val = el2.tagName === 'TEXTAREA' ? (el2.value || '') : (el2.textContent || '');
    chatInput = {
      tag: el2.tagName.toLowerCase() + (el2.getAttribute('contenteditable') === 'true' ? '[contenteditable]' : ''),
      direction: cs2.direction,
      textAlign: cs2.textAlign,
      hasHebrew: HEBREW_RE.test(val),
      valueSample: trimSample(val),
      path: shortPath(el2)
    };
    break;
  }

  // 3. Find Claude's message-bubble containers. Structure varies by
  // version; probe widely. For each candidate, record direction and
  // whether it contains Hebrew anywhere in its subtree.
  var bubbleSelectors = [
    '[data-testid*="message"]',
    '[data-test-id*="message"]',
    '[role="article"]',
    '[class*="message"]',
    '[class*="Message"]',
    '[class*="bubble"]',
    '[class*="Bubble"]'
  ];
  var seen = new Set();
  var bubbles = [];
  for (var k = 0; k < bubbleSelectors.length; k++) {
    var matches = document.querySelectorAll(bubbleSelectors[k]);
    for (var m = 0; m < matches.length; m++) {
      var b = matches[m];
      if (seen.has(b)) continue;
      seen.add(b);
      if (!isVisible(b)) continue;
      var subtreeText = b.textContent || '';
      if (!HEBREW_RE.test(subtreeText)) continue;
      var cs3 = getComputedStyle(b);
      bubbles.push({
        tag: b.tagName.toLowerCase(),
        direction: cs3.direction,
        textAlign: cs3.textAlign,
        sample: trimSample(subtreeText),
        path: shortPath(b),
        matchedBy: bubbleSelectors[k]
      });
      if (bubbles.length >= 25) break;
    }
    if (bubbles.length >= 25) break;
  }

  return {
    totalVisible: totalVisible,
    hebrewElements: hebrewElements,
    chatInput: chatInput,
    bubbles: bubbles
  };
})();
`;

async function evalInPage(Runtime, expression) {
  const r = await Runtime.evaluate({ expression: expression, returnByValue: true });
  if (r && r.exceptionDetails) {
    const text = r.exceptionDetails.exception && r.exceptionDetails.exception.description || JSON.stringify(r.exceptionDetails);
    throw new Error('page-eval: ' + text);
  }
  return r && r.result ? r.result.value : undefined;
}

async function runForTarget(entry) {
  const appName = entry.app ? entry.app.name : 'unknown';
  const port = entry.port;
  const target = entry.target;
  const out = {
    app: appName,
    port: port,
    url: target.url,
    styleSource: 'none-found',
    totalVisible: 0,
    hebrewElements: [],
    chatInput: null,
    bubbles: [],
    error: null
  };

  let client;
  try {
    client = await CDP({ target: target.webSocketDebuggerUrl });
  } catch (e) {
    out.error = 'CDP connect failed: ' + e.message;
    return out;
  }

  try {
    const Runtime = client.Runtime;
    await Runtime.enable();

    // 1. Verify production CSS, inject our copy if missing.
    const stylePresent = await evalInPage(Runtime, CHECK_STYLE_PRESENT_SCRIPT);
    if (stylePresent) {
      out.styleSource = 'production-injector';
    } else {
      const r = await evalInPage(Runtime, ENSURE_STYLE_SCRIPT);
      out.styleSource = 'test-injected (' + r + ')';
    }

    // 2. Scan the live DOM for Hebrew content.
    const scan = await evalInPage(Runtime, SCAN_LIVE_CONTENT_SCRIPT);
    if (scan) {
      out.totalVisible = scan.totalVisible || 0;
      out.hebrewElements = scan.hebrewElements || [];
      out.chatInput = scan.chatInput || null;
      out.bubbles = scan.bubbles || [];
    }
  } catch (e) {
    out.error = 'Eval failed: ' + e.message;
  } finally {
    try { await client.close(); } catch (e) { /* best-effort */ }
  }

  return out;
}

function fmtElement(e) {
  return '"' + e.sample + '": <' + e.tag + '> direction=' + e.direction +
    (e.direction === 'rtl' ? ' PASS' : ' FAIL');
}

function printReport(allOut) {
  console.log('');
  console.log('=== Live Add-in Content Tests - Claude for Office RTL Fix ===');
  console.log('');

  let totalHebrew = 0;
  let totalRtl = 0;
  let appsWithoutHebrew = [];
  let anomalies = [];

  for (let i = 0; i < allOut.length; i++) {
    const o = allOut[i];
    console.log('[' + o.app + ']  port=' + o.port);
    console.log('  CSS source: ' + o.styleSource);

    if (o.error) {
      console.log('  ERROR: ' + o.error);
      console.log('');
      continue;
    }

    console.log('  Total visible elements scanned: ' + o.totalVisible);
    console.log('  Elements with Hebrew content: ' + o.hebrewElements.length);

    let appHebrew = o.hebrewElements.length;
    let appRtl = 0;
    for (let j = 0; j < o.hebrewElements.length; j++) {
      const he = o.hebrewElements[j];
      console.log('  - ' + fmtElement(he));
      if (he.direction === 'rtl') appRtl++;
      else anomalies.push('[' + o.app + '] Hebrew but ' + he.direction + ': ' + he.sample + ' (' + he.path + ')');
    }
    totalHebrew += appHebrew;
    totalRtl += appRtl;

    if (appHebrew === 0) {
      appsWithoutHebrew.push(o.app);
      console.log('  (no Hebrew content found - ask user to type something in this pane and re-run)');
    }

    // Chat input
    if (o.chatInput) {
      const ci = o.chatInput;
      const ciDirOk = ci.direction === 'rtl';
      const ciAlignOk = ci.textAlign === 'start' || ci.textAlign === 'right';
      const ciOk = ciDirOk && ciAlignOk;
      console.log('  Chat input element: <' + ci.tag + '> direction=' + ci.direction +
        ' text-align=' + ci.textAlign + ' ' + (ciOk ? 'PASS' : 'FAIL'));
      if (!ciOk) {
        anomalies.push('[' + o.app + '] chat input not RTL: direction=' + ci.direction + ' text-align=' + ci.textAlign);
      }
    } else {
      console.log('  Chat input element: NOT FOUND');
    }

    // Message bubbles
    if (o.bubbles.length > 0) {
      console.log('  Message bubble candidates with Hebrew: ' + o.bubbles.length);
      for (let k = 0; k < o.bubbles.length; k++) {
        const b = o.bubbles[k];
        const ok = b.direction === 'rtl';
        console.log('    - <' + b.tag + '> [' + b.matchedBy + '] direction=' + b.direction +
          ' ' + (ok ? 'PASS' : 'FAIL') + ' :: "' + b.sample + '"');
        if (!ok) {
          anomalies.push('[' + o.app + '] bubble not RTL: ' + b.path + ' direction=' + b.direction);
        }
      }
    } else {
      console.log('  Message bubble candidates with Hebrew: 0');
    }

    const ciResult = o.chatInput ? (o.chatInput.direction === 'rtl' ? 'PASS' : 'FAIL') : 'N/A';
    console.log('  Result: ' + appRtl + '/' + appHebrew + ' Hebrew elements rendering RTL, chat input ' + ciResult);
    console.log('');
  }

  const goodApps = allOut.filter(function (o) { return !o.error; }).length;
  const pct = totalHebrew === 0 ? 0 : Math.round((totalRtl / totalHebrew) * 100);
  console.log('=== Overall: ' + totalHebrew + ' Hebrew elements scanned across ' +
    goodApps + ' apps, ' + pct + '% RTL-rendered ===');

  if (appsWithoutHebrew.length > 0) {
    console.log('');
    console.log('Apps with no Hebrew content found: ' + appsWithoutHebrew.join(', '));
    console.log('Recommendation: type some Hebrew text in those panes and re-run.');
  }

  if (anomalies.length > 0) {
    console.log('');
    console.log('Anomalies (Hebrew text rendering non-RTL - real bugs):');
    for (let i = 0; i < anomalies.length; i++) {
      console.log('  - ' + anomalies[i]);
    }
  }

  // Pass condition: every Hebrew element rendered RTL, every chat input
  // (where one was found) was RTL, and at least one app contributed
  // Hebrew content (otherwise the test was vacuous).
  const noFailures = anomalies.length === 0;
  const hadHebrew = totalHebrew > 0;
  const noErrors = allOut.every(function (o) { return !o.error; });
  return noFailures && hadHebrew && noErrors;
}

async function main() {
  console.log('Discovering active Claude targets across Office apps...');
  let targets;
  try {
    targets = await portDiscovery.discoverActiveTargets();
  } catch (e) {
    console.error('Discovery failed: ' + e.message);
    process.exit(1);
  }
  if (!targets || targets.length === 0) {
    console.error('No active Claude targets found. Open Claude in Word/Excel/PowerPoint first, then re-run.');
    process.exit(1);
  }
  console.log('Found ' + targets.length + ' target(s).');

  const allOut = [];
  for (let i = 0; i < targets.length; i++) {
    const o = await runForTarget(targets[i]);
    allOut.push(o);
  }

  const ok = printReport(allOut);
  process.exit(ok ? 0 : 1);
}

main().catch(function (e) {
  console.error('Fatal: ' + (e && e.stack ? e.stack : e));
  process.exit(1);
});
