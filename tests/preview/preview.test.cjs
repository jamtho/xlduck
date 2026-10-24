// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ channel: 'msedge', headless: true });
  let passed = 0, failed = 0, knownFailures = 0;
  try {
    const page = await browser.newPage();
    await page.route('**/*', route => route.abort());
    let html = fs.readFileSync(path.join(__dirname, '../../XlDuck/Preview/preview.html'), 'utf8');
    html = html.replace(/<script src="[^"]+"><\/script>/g, '')
      .replace(/<link[^>]+href="https:[^>]+>/g, '')
      .replace('<head>', '<head><script>window.chrome = {webview: {addEventListener(){}}};</script>');
    await page.setContent(html);
    const observations = await page.evaluate(() => {
      const value = 'value" onmouseover="window.previewProbe=1';
      renderState({kind: 'table', title: 'Test', table: {
        columns: [{name: 'value', type: 'VARCHAR'}], rows: [[value]], rowCount: 1, previewRowCount: 1
      }});
      const cell = document.querySelector('.data-grid td');
      cell.dispatchEvent(new MouseEvent('mouseover', {bubbles: true}));
      return {
        text: cell.textContent,
        expectedText: value,
        handler: cell.getAttribute('onmouseover'),
        executed: window.previewProbe === 1,
        fraction: parseDurationToSeconds('00:00:00.5000000'),
        seconds: parseDurationToSeconds('00:01:02')
      };
    });
    function check(name, test, known = false) {
      try { test(); passed++; console.log(`PASS ${name}`); }
      catch (error) {
        if (known) knownFailures++; else failed++;
        console.log(`${known ? 'KNOWN FAILURE' : 'FAIL'} ${name}: ${error.message}`);
      }
    }
    check('table text is preserved', () => assert.equal(observations.text, observations.expectedText));
    check('whole-second intervals', () => assert.equal(observations.seconds, 62));
    check('table values cannot add HTML attributes or execute code', () => {
      assert.equal(observations.handler, null);
      assert.equal(observations.executed, false);
    }, true);
    check('fractional interval seconds', () => assert.equal(observations.fraction, 0.5), true);
    console.log(`${passed} passed, ${failed} unexpected failures, ${knownFailures} known defects reproduced.`);
    process.exitCode = failed || (knownFailures && !process.argv.includes('--allow-known-defects')) ? 1 : 0;
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
