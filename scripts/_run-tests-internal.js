#!/usr/bin/env node
// Automated test runner - sends messages to Claude panel, screenshots responses.
const CDP = require('chrome-remote-interface');
const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = 9222;
const OUT_DIR = path.join(__dirname, '..', 'test-results');
if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR, { recursive: true });

const TESTS = process.argv.slice(2).length > 0
  ? JSON.parse(fs.readFileSync(process.argv[2], 'utf8'))
  : [];

function listTargets() {
  return new Promise((resolve, reject) => {
    http.get(`http://localhost:${PORT}/json/list`, (res) => {
      let body = '';
      res.on('data', c => body += c);
      res.on('end', () => { try { resolve(JSON.parse(body)); } catch(e) { reject(e); } });
    }).on('error', reject);
  });
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

async function sendMessage(client, text) {
  const { Runtime, Input } = client;

  // Focus the editor
  await Runtime.evaluate({ expression: `document.querySelector('.tiptap.ProseMirror').focus()` });
  await sleep(200);

  // Clear existing content
  await Input.dispatchKeyEvent({ type: 'keyDown', modifiers: 2, key: 'a', code: 'KeyA', windowsVirtualKeyCode: 65, nativeVirtualKeyCode: 65 });
  await Input.dispatchKeyEvent({ type: 'keyUp', modifiers: 2, key: 'a', code: 'KeyA', windowsVirtualKeyCode: 65, nativeVirtualKeyCode: 65 });
  await Input.dispatchKeyEvent({ type: 'keyDown', key: 'Delete', code: 'Delete', windowsVirtualKeyCode: 46, nativeVirtualKeyCode: 46 });
  await Input.dispatchKeyEvent({ type: 'keyUp', key: 'Delete', code: 'Delete', windowsVirtualKeyCode: 46, nativeVirtualKeyCode: 46 });
  await sleep(200);

  // Insert text
  await Input.insertText({ text });
  await sleep(500);

  // Press Enter to send
  await Input.dispatchKeyEvent({ type: 'keyDown', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13, nativeVirtualKeyCode: 13 });
  await Input.dispatchKeyEvent({ type: 'keyUp', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13, nativeVirtualKeyCode: 13 });
}

async function waitForResponse(client, maxWaitMs = 60000) {
  const { Runtime } = client;
  const start = Date.now();
  let lastLen = -1;
  let stableCount = 0;

  while (Date.now() - start < maxWaitMs) {
    await sleep(1500);
    const r = await Runtime.evaluate({
      expression: `(function(){
        const stopBtn = document.querySelector('button[aria-label="Stop response"], button[aria-label*="Stop"]');
        const body = document.body.innerText.length;
        return JSON.stringify({stopping: !!stopBtn, bodyLen: body});
      })()`,
      returnByValue: true
    });
    const info = JSON.parse(r.result.value);
    if (!info.stopping) {
      if (info.bodyLen === lastLen) {
        stableCount++;
        if (stableCount >= 2) return true;
      } else {
        stableCount = 0;
        lastLen = info.bodyLen;
      }
    } else {
      stableCount = 0;
      lastLen = info.bodyLen;
    }
  }
  return false;
}

async function screenshot(client, filepath) {
  const { Page } = client;
  await Page.enable();
  const { data } = await Page.captureScreenshot({ format: 'png', captureBeyondViewport: false });
  fs.writeFileSync(filepath, Buffer.from(data, 'base64'));
}

async function scrollToBottom(client) {
  await client.Runtime.evaluate({
    expression: `(function(){
      const scrollers = document.querySelectorAll('div');
      for (const s of scrollers) {
        if (s.scrollHeight > s.clientHeight && s.clientHeight > 100) {
          s.scrollTop = s.scrollHeight;
        }
      }
      window.scrollTo(0, document.body.scrollHeight);
    })()`
  });
}

(async () => {
  const targets = await listTargets();
  const t = targets.find(x => /claude\.ai|anthropic/i.test(x.url || ''));
  if (!t) { console.error('No Claude target'); process.exit(1); }

  const client = await CDP({ target: t.webSocketDebuggerUrl });
  await client.Page.enable();
  await client.Runtime.enable();

  for (let i = 0; i < TESTS.length; i++) {
    const test = TESTS[i];
    console.log(`\n[Test ${i + 1}/${TESTS.length}] ${test.name}`);
    console.log(`  Prompt: ${test.prompt.slice(0, 80)}...`);

    await sendMessage(client, test.prompt);
    console.log('  Sent. Waiting for response...');
    const ok = await waitForResponse(client);
    if (!ok) console.log('  WARN: response did not stabilize within timeout');

    await sleep(1000);
    await scrollToBottom(client);
    await sleep(500);

    const file = path.join(OUT_DIR, `test-${String(i + 1).padStart(2, '0')}-${test.name}.png`);
    await screenshot(client, file);
    console.log(`  Screenshot: ${file}`);
    await sleep(2000);
  }

  await client.close();
  console.log('\nAll tests complete.');
})().catch(e => { console.error(e); process.exit(1); });
