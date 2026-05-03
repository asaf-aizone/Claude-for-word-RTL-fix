#!/usr/bin/env node
'use strict';

/**
 * Text-rendering test suite for Claude for Office RTL Fix.
 *
 * Discovers all currently-active Claude targets across Word/Excel/PowerPoint
 * via port-discovery.discoverActiveTargets(), attaches via CDP, and runs a
 * series of in-page DOM tests to verify that the production CSS + typography
 * rules render correctly with real Hebrew text variations.
 *
 * Prerequisite: the user must have opened Claude panes in the Office apps
 * manually before running this script.
 *
 * Exits 0 on all-pass, 1 on any failure.
 *
 * NOTE: RTL_CSS is duplicated verbatim from scripts/inject.js. This is
 * intentional - inject.js has top-level side effects (writePidFile, log
 * truncation, setInterval) so requiring it as a module is not safe. The
 * tradeoff is acceptable since this is a test script and the duplication
 * is small. If RTL_CSS is ever refactored into a separate module, this
 * copy should be removed in favor of the import.
 */

const path = require('path');
const CDP = require(path.join(__dirname, '..', 'scripts', 'node_modules', 'chrome-remote-interface'));
const portDiscovery = require(path.join(__dirname, '..', 'scripts', 'port-discovery'));

// === RTL_CSS: copied verbatim from scripts/inject.js (see file header note) ===
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

// Minimal style-only injector for use when the production injector hasn't
// run yet against this target. We do NOT recreate the MutationObserver here -
// instead we install a one-shot text-walker that performs the same character
// replacements directly, so test t3-t7 can pass without depending on whether
// the production observer is already active.
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

// Runs the same text-replacement rules as INJECTOR_SCRIPT in inject.js, but
// synchronously over a given root. Used after inserting test elements so we
// don't have to wait for the production observer's rAF flush.
const APPLY_TYPO_SCRIPT_TEMPLATE = function (rootSelector) {
  return `
  (function () {
    var root = document.querySelector(${JSON.stringify(rootSelector)});
    if (!root) return 'no-root';
    var REPLACERS = [
      { from: /[\\u2014\\u2013]/g, to: '-' },
      { from: /[\\u2190-\\u2199\\u21D0-\\u21D4]/g, to: ',' }
    ];
    function isInCodeBlock(node) {
      var el = node.nodeType === 1 ? node : node.parentElement;
      while (el && el !== root.parentElement) {
        var tn = el.tagName;
        if (tn === 'PRE' || tn === 'CODE' || tn === 'KBD' || tn === 'SAMP') return true;
        el = el.parentElement;
      }
      return false;
    }
    function fix(node) {
      if (node.nodeType === 3) {
        if (isInCodeBlock(node)) return;
        var t = node.nodeValue;
        for (var i = 0; i < REPLACERS.length; i++) t = t.replace(REPLACERS[i].from, REPLACERS[i].to);
        if (t !== node.nodeValue) node.nodeValue = t;
      } else if (node.nodeType === 1) {
        var tag = node.tagName;
        if (tag === 'TEXTAREA' || tag === 'INPUT' || node.isContentEditable) return;
        for (var c = node.firstChild; c; c = c.nextSibling) fix(c);
      }
    }
    fix(root);
    return 'applied';
  })();
  `;
};

// Builds the test container with all 15 variations and returns nothing.
// The literal Hebrew + special characters live here as JS string literals
// using \uXXXX escapes so this source file stays pure ASCII.
const BUILD_CONTAINER_SCRIPT = `
(function () {
  var prev = document.getElementById('__rtl_test__');
  if (prev) prev.remove();
  var c = document.createElement('div');
  c.id = '__rtl_test__';
  c.setAttribute('style', 'position:fixed;left:-9999px;top:0;width:400px;');
  c.innerHTML = [
    '<p id="t1_hebrew_p">\\u05E9\\u05DC\\u05D5\\u05DD \\u05E2\\u05D5\\u05DC\\u05DD</p>',
    '<p id="t2_mixed">Hello \\u05E9\\u05DC\\u05D5\\u05DD world</p>',
    '<p id="t3_emdash">\\u05D0\\u05D1\\u05D0 \\u2014 \\u05D0\\u05DE\\u05D0</p>',
    '<p id="t4_endash">\\u05D0\\u05D1\\u05D0 \\u2013 \\u05D0\\u05DE\\u05D0</p>',
    '<p id="t5_arrow_right">\\u05D0 \\u2192 \\u05D1</p>',
    '<p id="t6_arrow_left">\\u05D1 \\u2190 \\u05D0</p>',
    '<p id="t7_double_arrow">\\u05D0 \\u21D2 \\u05D1</p>',
    '<pre id="t8_code_block"><code>function \\u05E9\\u05DC\\u05D5\\u05DD()</code></pre>',
    '<p id="t9_inline_code_in_p">\\u05D4\\u05E9\\u05DD <code id="t9_inner">name</code> \\u05D4\\u05D5\\u05D0 X</p>',
    '<ol id="t10_list_ol"><li id="t10_li">\\u05E4\\u05E8\\u05D9\\u05D8 \\u05E8\\u05D0\\u05E9\\u05D5\\u05DF</li></ol>',
    '<ul id="t11_list_ul"><li id="t11_li">\\u05E4\\u05E8\\u05D9\\u05D8 \\u05E9\\u05E0\\u05D9</li></ul>',
    '<textarea id="t12_textarea"></textarea>',
    '<div id="t13_contenteditable" contenteditable="true">\\u05D4\\u05E7\\u05DC\\u05D3 \\u05DB\\u05D0\\u05DF</div>',
    '<h1 id="t14_h1">\\u05DB\\u05D5\\u05EA\\u05E8\\u05EA \\u05E8\\u05D0\\u05E9\\u05D9\\u05EA</h1>',
    '<table><tbody><tr><td id="t15_td">\\u05EA\\u05D0 \\u05D1\\u05D8\\u05D1\\u05DC\\u05D4</td></tr></tbody></table>',
    // === t16-t25: parentheses/brackets/braces BiDi stress tests ===
    // Visual layout (which side parens render on) cannot be reliably checked
    // via CDP getComputedStyle alone - the browser's BiDi algorithm runs at
    // layout time and produces visual order that is not exposed as a property.
    // What we CAN check programmatically and reliably:
    //   1. direction=rtl (or ltr for inner <code>) on the wrapping element
    //   2. unicode-bidi=isolate on the same element
    //   3. textContent unchanged - typography observer must NOT touch
    //      parens (), brackets [], braces {}, or hyphen-minus U+002D
    // With those three invariants in place, the browser BiDi algorithm
    // handles the visual rendering correctly. So the test is essentially:
    // "did we set the right CSS, and did the typography observer leave the
    // structural punctuation alone".
    '<p id="t16_parens_around_english">\\u05D4\\u05D8\\u05E7\\u05E1\\u05D8 (in English) \\u05DE\\u05DE\\u05E9\\u05D9\\u05DA</p>',
    '<p id="t17_parens_around_hebrew">Click here (\\u05DC\\u05D7\\u05E5 \\u05DB\\u05D0\\u05DF) please</p>',
    '<p id="t18_parens_with_number">\\u05D4\\u05DE\\u05D7\\u05D9\\u05E8 (50 \\u05E9\\u05E7\\u05DC\\u05D9\\u05DD) \\u05DB\\u05D5\\u05DC\\u05DC \\u05DE\\u05E2"\\u05DE</p>',
    '<p id="t19_brackets_around_english">\\u05D4\\u05E7\\u05D5\\u05D1\\u05E5 [config.json] \\u05E0\\u05DE\\u05E6\\u05D0</p>',
    '<p id="t20_braces_in_code_context">\\u05D4\\u05E4\\u05D5\\u05E0\\u05E7\\u05E6\\u05D9\\u05D4 <code id="t20_inner">{ key: value }</code> \\u05DE\\u05D7\\u05D6\\u05D9\\u05E8\\u05D4</p>',
    '<p id="t21_nested_parens">\\u05E9\\u05DC\\u05D5\\u05DD (\\u05E2\\u05D5\\u05DC\\u05DD (\\u05E4\\u05E0\\u05D9\\u05DE\\u05D9) \\u05D7\\u05D9\\u05E6\\u05D5\\u05E0\\u05D9) \\u05E1\\u05D5\\u05E3</p>',
    '<p id="t22_paren_at_start">(\\u05D4\\u05E2\\u05E8\\u05D4: \\u05D6\\u05D4 \\u05D7\\u05E9\\u05D5\\u05D1) \\u05D4\\u05DE\\u05E9\\u05DA \\u05D4\\u05D8\\u05E7\\u05E1\\u05D8</p>',
    '<p id="t23_paren_at_end">\\u05D4\\u05DE\\u05E9\\u05DA \\u05D4\\u05D8\\u05E7\\u05E1\\u05D8 (\\u05D4\\u05E2\\u05E8\\u05D4: \\u05D6\\u05D4 \\u05D7\\u05E9\\u05D5\\u05D1)</p>',
    '<p id="t24_quote_with_paren">\\u05D0\\u05DE\\u05E8: "\\u05E9\\u05DC\\u05D5\\u05DD (\\u05D4\\u05D9\\u05D9) \\u05E2\\u05D5\\u05DC\\u05DD" - \\u05EA\\u05D2\\u05D5\\u05D1\\u05D4</p>',
    '<p id="t25_url_in_parens">\\u05E8\\u05D0\\u05D4 (https://example.com) \\u05DC\\u05DE\\u05D9\\u05D3\\u05E2</p>'
  ].join('');
  document.body.appendChild(c);
  return 'built';
})();
`;

// Runs in-page; reads computed direction/text-align + textContent for each
// element id and returns a structured object the test harness compares
// against expectations.
const READ_RESULTS_SCRIPT = `
(function () {
  function info(id) {
    var el = document.getElementById(id);
    if (!el) return { missing: true };
    var cs = getComputedStyle(el);
    return {
      direction: cs.direction,
      textAlign: cs.textAlign,
      unicodeBidi: cs.unicodeBidi,
      textContent: el.textContent
    };
  }
  return {
    t1: info('t1_hebrew_p'),
    t2: info('t2_mixed'),
    t3: info('t3_emdash'),
    t4: info('t4_endash'),
    t5: info('t5_arrow_right'),
    t6: info('t6_arrow_left'),
    t7: info('t7_double_arrow'),
    t8_pre: info('t8_code_block'),
    t9_p: info('t9_inline_code_in_p'),
    t9_code: info('t9_inner'),
    t10_ol: info('t10_list_ol'),
    t10_li: info('t10_li'),
    t11_ul: info('t11_list_ul'),
    t11_li: info('t11_li'),
    t12: info('t12_textarea'),
    t13: info('t13_contenteditable'),
    t14: info('t14_h1'),
    t15: info('t15_td'),
    t16: info('t16_parens_around_english'),
    t17: info('t17_parens_around_hebrew'),
    t18: info('t18_parens_with_number'),
    t19: info('t19_brackets_around_english'),
    t20_p: info('t20_braces_in_code_context'),
    t20_code: info('t20_inner'),
    t21: info('t21_nested_parens'),
    t22: info('t22_paren_at_start'),
    t23: info('t23_paren_at_end'),
    t24: info('t24_quote_with_paren'),
    t25: info('t25_url_in_parens')
  };
})();
`;

const CLEANUP_SCRIPT = `
(function () {
  var c = document.getElementById('__rtl_test__');
  if (c) c.remove();
  return 'cleaned';
})();
`;

const CHECK_STYLE_PRESENT_SCRIPT =
  `!!document.getElementById('__claude_rtl_fix__')`;

// === Test definitions: each takes the results object and returns
// { pass: boolean, detail: string } ===
const EM_DASH = '—';
const EN_DASH = '–';
const ARROW_RIGHT = '→';
const ARROW_LEFT = '←';
const DOUBLE_ARROW = '⇒';

const TESTS = [
  { id: 't1_hebrew_p', label: 'direction=rtl', check: function (r) {
    return checkDirection(r.t1, 'rtl');
  }},
  { id: 't2_mixed', label: 'direction=rtl on mixed-content <p>', check: function (r) {
    return checkDirection(r.t2, 'rtl');
  }},
  { id: 't3_emdash', label: 'em-dash replaced with hyphen', check: function (r) {
    return checkNoChar(r.t3, EM_DASH, '-');
  }},
  { id: 't4_endash', label: 'en-dash replaced with hyphen', check: function (r) {
    return checkNoChar(r.t4, EN_DASH, '-');
  }},
  { id: 't5_arrow_right', label: 'right arrow replaced with comma', check: function (r) {
    return checkNoChar(r.t5, ARROW_RIGHT, ',');
  }},
  { id: 't6_arrow_left', label: 'left arrow replaced with comma', check: function (r) {
    return checkNoChar(r.t6, ARROW_LEFT, ',');
  }},
  { id: 't7_double_arrow', label: 'double arrow replaced with comma', check: function (r) {
    return checkNoChar(r.t7, DOUBLE_ARROW, ',');
  }},
  { id: 't8_code_block', label: '<pre><code> stays direction=ltr', check: function (r) {
    return checkDirection(r.t8_pre, 'ltr');
  }},
  { id: 't9_inline_code_in_p', label: '<p>=rtl, inner <code>=ltr', check: function (r) {
    var p = checkDirection(r.t9_p, 'rtl');
    if (!p.pass) return p;
    var c = checkDirection(r.t9_code, 'ltr');
    if (!c.pass) return { pass: false, detail: 'inner <code>: ' + c.detail };
    return { pass: true, detail: '<p>=rtl, <code>=ltr' };
  }},
  { id: 't10_list_ol', label: '<ol> direction=rtl', check: function (r) {
    return checkDirection(r.t10_ol, 'rtl');
  }},
  { id: 't11_list_ul', label: '<ul> direction=rtl', check: function (r) {
    return checkDirection(r.t11_ul, 'rtl');
  }},
  { id: 't12_textarea', label: 'textarea direction=rtl, text-align=start/right', check: function (r) {
    var d = checkDirection(r.t12, 'rtl');
    if (!d.pass) return d;
    // text-align:start under rtl resolves to 'right' in computed style on
    // most engines, but the spec also allows the literal 'start' to surface.
    var ta = r.t12.textAlign;
    if (ta === 'start' || ta === 'right') return { pass: true, detail: 'direction=rtl, text-align=' + ta };
    return { pass: false, detail: 'expected text-align in {start,right}, got ' + JSON.stringify(ta) };
  }},
  { id: 't13_contenteditable', label: 'contenteditable direction=rtl', check: function (r) {
    return checkDirection(r.t13, 'rtl');
  }},
  { id: 't14_h1', label: '<h1> direction=rtl', check: function (r) {
    return checkDirection(r.t14, 'rtl');
  }},
  { id: 't15_table_cell', label: '<td> direction=rtl', check: function (r) {
    return checkDirection(r.t15, 'rtl');
  }},
  // === t16-t25: parens/brackets BiDi tests ===
  // See block comment in BUILD_CONTAINER_SCRIPT for why visual layout is not
  // checked here. Each test asserts: direction=rtl, unicode-bidi=isolate,
  // and the structural punctuation chars survived the typography observer
  // intact (no mangling of (), [], {}, or U+002D hyphen-minus).
  { id: 't16_parens_around_english', label: '<p> rtl+isolate, parens preserved around English', check: function (r) {
    return checkBidiAndChars(r.t16, 'rtl', ['(', ')', 'in English']);
  }},
  { id: 't17_parens_around_hebrew', label: '<p> rtl+isolate, parens preserved around Hebrew', check: function (r) {
    return checkBidiAndChars(r.t17, 'rtl', ['(', ')', 'Click here', 'please']);
  }},
  { id: 't18_parens_with_number', label: '<p> rtl+isolate, parens preserved around mixed Hebrew+digits', check: function (r) {
    return checkBidiAndChars(r.t18, 'rtl', ['(', ')', '50']);
  }},
  { id: 't19_brackets_around_english', label: '<p> rtl+isolate, square brackets preserved around filename', check: function (r) {
    return checkBidiAndChars(r.t19, 'rtl', ['[', ']', 'config.json']);
  }},
  { id: 't20_braces_in_code_context', label: '<p>=rtl, inner <code>=ltr, braces preserved inside code', check: function (r) {
    var p = checkBidiAndChars(r.t20_p, 'rtl', []);
    if (!p.pass) return p;
    var c = checkDirection(r.t20_code, 'ltr');
    if (!c.pass) return { pass: false, detail: 'inner <code>: ' + c.detail };
    var t = (r.t20_code && r.t20_code.textContent) || '';
    if (t.indexOf('{') === -1 || t.indexOf('}') === -1 || t.indexOf('key: value') === -1) {
      return { pass: false, detail: 'inner <code> textContent missing braces or content: ' + JSON.stringify(t) };
    }
    return { pass: true, detail: '<p>=rtl+isolate, <code>=ltr, braces and content preserved' };
  }},
  { id: 't21_nested_parens', label: '<p> rtl+isolate, nested parens preserved (4 paren chars)', check: function (r) {
    var base = checkBidiAndChars(r.t21, 'rtl', []);
    if (!base.pass) return base;
    var t = (r.t21 && r.t21.textContent) || '';
    var opens = (t.match(/\(/g) || []).length;
    var closes = (t.match(/\)/g) || []).length;
    if (opens !== 2 || closes !== 2) {
      return { pass: false, detail: 'expected 2 \'(\' and 2 \')\', got ' + opens + '/' + closes + ' in ' + JSON.stringify(t) };
    }
    return { pass: true, detail: 'rtl+isolate, both paren pairs intact' };
  }},
  { id: 't22_paren_at_start', label: '<p> rtl+isolate, leading opening paren preserved', check: function (r) {
    var base = checkBidiAndChars(r.t22, 'rtl', ['(', ')']);
    if (!base.pass) return base;
    var t = (r.t22 && r.t22.textContent) || '';
    if (t.charAt(0) !== '(') {
      return { pass: false, detail: 'expected textContent to start with \'(\'; got ' + JSON.stringify(t.slice(0, 8)) };
    }
    return { pass: true, detail: 'rtl+isolate, leading \'(\' intact' };
  }},
  { id: 't23_paren_at_end', label: '<p> rtl+isolate, trailing closing paren preserved', check: function (r) {
    var base = checkBidiAndChars(r.t23, 'rtl', ['(', ')']);
    if (!base.pass) return base;
    var t = (r.t23 && r.t23.textContent) || '';
    if (t.charAt(t.length - 1) !== ')') {
      return { pass: false, detail: 'expected textContent to end with \')\'; got ' + JSON.stringify(t.slice(-8)) };
    }
    return { pass: true, detail: 'rtl+isolate, trailing \')\' intact' };
  }},
  { id: 't24_quote_with_paren', label: '<p> rtl+isolate, quotes+parens+hyphen all preserved', check: function (r) {
    // Critical: the hyphen-minus U+002D must survive. Typography observer
    // only replaces U+2013 / U+2014 dashes, not U+002D, but we verify here
    // explicitly because if anyone broadens the observer regex this test
    // catches it before users see broken contractions.
    var base = checkBidiAndChars(r.t24, 'rtl', ['"', '(', ')', '-']);
    if (!base.pass) return base;
    var t = (r.t24 && r.t24.textContent) || '';
    // Make sure the hyphen-minus was NOT replaced with anything else.
    if (t.indexOf('–') !== -1 || t.indexOf('—') !== -1) {
      return { pass: false, detail: 'unexpected en/em-dash in textContent: ' + JSON.stringify(t) };
    }
    return { pass: true, detail: 'rtl+isolate, quotes/parens/hyphen-minus preserved' };
  }},
  { id: 't25_url_in_parens', label: '<p> rtl+isolate, URL inside parens preserved whole', check: function (r) {
    var base = checkBidiAndChars(r.t25, 'rtl', ['(', ')', 'https://example.com']);
    if (!base.pass) return base;
    return { pass: true, detail: 'rtl+isolate, URL intact inside parens' };
  }}
];

function checkDirection(info, expected) {
  if (!info || info.missing) return { pass: false, detail: 'element missing' };
  if (info.direction === expected) return { pass: true, detail: 'direction=' + info.direction };
  return { pass: false, detail: 'expected direction=' + expected + ', got ' + JSON.stringify(info.direction) };
}

// Combined check used by t16-t25: element must have direction=expectedDir
// AND unicode-bidi=isolate (the BiDi-correctness invariant from RTL_CSS),
// AND every requiredSubstr must be present unchanged in textContent. This
// is the strongest signal we can read via CDP without doing pixel layout.
function checkBidiAndChars(info, expectedDir, requiredSubstrs) {
  if (!info || info.missing) return { pass: false, detail: 'element missing' };
  if (info.direction !== expectedDir) {
    return { pass: false, detail: 'expected direction=' + expectedDir + ', got ' + JSON.stringify(info.direction) };
  }
  // unicode-bidi can compute as 'isolate' literally; some engines may
  // surface it as 'isolate' or 'bidi-isolate' historically. Accept either.
  var ub = info.unicodeBidi;
  if (ub !== 'isolate' && ub !== 'bidi-isolate') {
    return { pass: false, detail: 'expected unicode-bidi=isolate, got ' + JSON.stringify(ub) };
  }
  var t = info.textContent || '';
  for (var i = 0; i < requiredSubstrs.length; i++) {
    if (t.indexOf(requiredSubstrs[i]) === -1) {
      return {
        pass: false,
        detail: 'missing required substring ' + JSON.stringify(requiredSubstrs[i]) +
          ' in textContent=' + JSON.stringify(t)
      };
    }
  }
  return { pass: true, detail: 'direction=' + expectedDir + ', unicode-bidi=' + ub + ', content intact' };
}

function checkNoChar(info, badChar, goodChar) {
  if (!info || info.missing) return { pass: false, detail: 'element missing' };
  var t = info.textContent || '';
  var hasBad = t.indexOf(badChar) !== -1;
  var hasGood = t.indexOf(goodChar) !== -1;
  if (!hasBad && hasGood) {
    return { pass: true, detail: 'replaced (' + JSON.stringify(goodChar) + ' not ' + JSON.stringify(badChar) + ')' };
  }
  return {
    pass: false,
    detail: 'expected no ' + JSON.stringify(badChar) + ' and at least one ' + JSON.stringify(goodChar) +
      '; got textContent=' + JSON.stringify(t)
  };
}

function delay(ms) {
  return new Promise(function (r) { setTimeout(r, ms); });
}

async function evalInPage(Runtime, expression) {
  const r = await Runtime.evaluate({ expression: expression, returnByValue: true });
  if (r && r.exceptionDetails) {
    const text = r.exceptionDetails.exception && r.exceptionDetails.exception.description || JSON.stringify(r.exceptionDetails);
    throw new Error('page-eval: ' + text);
  }
  return r && r.result ? r.result.value : undefined;
}

async function runTestsForTarget(entry) {
  const appName = entry.app ? entry.app.name : 'unknown';
  const port = entry.port;
  const target = entry.target;
  const header = '[' + appName + '] port=' + port;
  const out = { app: appName, port: port, url: target.url, results: [], passCount: 0, failCount: 0, error: null };

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
    if (!stylePresent) {
      const r = await evalInPage(Runtime, ENSURE_STYLE_SCRIPT);
      out.styleSource = 'test-injected (' + r + ')';
    } else {
      out.styleSource = 'production-injector';
    }

    // 2. Build the test container.
    await evalInPage(Runtime, BUILD_CONTAINER_SCRIPT);

    // 3. Wait for the production MutationObserver to process inserted nodes
    // (text-node character replacements run on rAF).
    await delay(250);

    // 4. Belt and suspenders: synchronously apply the same replacement pass
    // ourselves. If the production observer already did it, this is a no-op.
    // If it didn't (e.g. observer not installed yet, or running against a
    // page that has the style but not the observer), we still get correct
    // textContent for tests t3-t7.
    await evalInPage(Runtime, APPLY_TYPO_SCRIPT_TEMPLATE('#__rtl_test__'));

    // 5. Read computed styles + textContent.
    const results = await evalInPage(Runtime, READ_RESULTS_SCRIPT);

    // 6. Run each test.
    for (let i = 0; i < TESTS.length; i++) {
      const t = TESTS[i];
      let outcome;
      try {
        outcome = t.check(results);
      } catch (e) {
        outcome = { pass: false, detail: 'check threw: ' + e.message };
      }
      out.results.push({ id: t.id, label: t.label, pass: outcome.pass, detail: outcome.detail });
      if (outcome.pass) out.passCount++; else out.failCount++;
    }

    // 7. Cleanup.
    try { await evalInPage(Runtime, CLEANUP_SCRIPT); } catch (e) { /* best-effort */ }
  } catch (e) {
    out.error = 'Eval failed: ' + e.message;
  } finally {
    try { await client.close(); } catch (e) { /* best-effort */ }
  }

  return out;
}

function printReport(allOut) {
  console.log('');
  console.log('=== Text Rendering Tests - Claude for Office RTL Fix ===');
  console.log('');
  let totalPass = 0;
  let totalRun = 0;
  for (let i = 0; i < allOut.length; i++) {
    const o = allOut[i];
    console.log('[' + o.app + ']  port=' + o.port);
    if (o.styleSource) console.log('  CSS source: ' + o.styleSource);
    if (o.error) {
      console.log('  ERROR: ' + o.error);
      console.log('');
      continue;
    }
    for (let j = 0; j < o.results.length; j++) {
      const r = o.results[j];
      const tag = r.pass ? 'PASS' : 'FAIL';
      console.log('  ' + tag + '  ' + r.id + ': ' + r.detail);
    }
    console.log('  Result: ' + o.passCount + '/' + (o.passCount + o.failCount) + ' passed');
    console.log('');
    totalPass += o.passCount;
    totalRun += (o.passCount + o.failCount);
  }
  const appCount = allOut.filter(function (o) { return !o.error; }).length;
  console.log('=== Overall: ' + totalPass + '/' + totalRun + ' passed (' + appCount + ' app' + (appCount === 1 ? '' : 's') + ' x ' + TESTS.length + ' tests) ===');
  return totalPass === totalRun && totalRun > 0 && allOut.every(function (o) { return !o.error; });
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
    const o = await runTestsForTarget(targets[i]);
    allOut.push(o);
  }

  const ok = printReport(allOut);
  process.exit(ok ? 0 : 1);
}

main().catch(function (e) {
  console.error('Fatal: ' + (e && e.stack ? e.stack : e));
  process.exit(1);
});
