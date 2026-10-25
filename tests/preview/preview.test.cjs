// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ channel: process.env.XLDUCK_TEST_BROWSER === 'chromium' ? undefined : 'msedge', headless: true });
  let passed = 0, failed = 0, knownFailures = 0;
  try {
    const page = await browser.newPage();
    const external = [], errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.addInitScript(() => { window.chrome = {webview: {addEventListener(){}}}; });
    await page.route('**/*', async route => {
      const url = new URL(route.request().url());
      if (url.origin !== 'https://xlduck.invalid') { external.push(url.href); await route.abort(); return; }
      const name = path.posix.basename(url.pathname);
      const folder = name.startsWith('preview.') ? '../../XlDuck/Preview' : '../../artifacts/preview-assets';
      const file = path.join(__dirname, folder, name);
      if (!fs.existsSync(file)) throw new Error(`Missing bundled asset: ${file}. Run scripts/Get-PreviewAssets.ps1.`);
      await route.fulfill({path: file, contentType: name.endsWith('.js') ? 'text/javascript' : name.endsWith('.css') ? 'text/css' : 'text/html'});
    });
    await page.goto('https://xlduck.invalid/preview.html');
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
        seconds: parseDurationToSeconds('00:01:02'),
        duckDays: parseDurationToSeconds('2 days 00:00:00.25'),
        boolean: chartValue('true', 'Boolean', 'BOOLEAN'),
        textNumber: chartValue('00123', 'Text', 'VARCHAR'),
        nested: chartValue('{"x":9007199254740993}', 'Json', 'JSON'),
        nestedQuantitative: isQuantitative('INTEGER[]'),
        unsafeIntegerRejected: (() => { try { chartValue('9007199254740993', 'Integer', 'HUGEINT'); return false; } catch { return true; } })(),
        decimalRejected: (() => { try { chartValue('1.234567890123456789', 'Decimal', 'DECIMAL(20,18)'); return false; } catch { return true; } })(),
        literalQuestion: formatSql("SELECT '?' AS question")
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
    });
    check('fractional interval seconds', () => assert.equal(observations.fraction, 0.5));
    check('DuckDB day intervals', () => assert.equal(observations.duckDays, 172800.25));
    check('chart values preserve text, booleans and nested exact JSON', () => {
      assert.equal(observations.boolean, true); assert.equal(observations.textNumber, '00123');
      assert.equal(observations.nested, '{"x":9007199254740993}'); assert.equal(observations.nestedQuantitative, false);
    });
    check('chart precision loss requires an explicit SQL conversion', () => {
      assert.equal(observations.unsafeIntegerRejected, true); assert.equal(observations.decimalRejected, true);
    });
    check('SQL preview does not change literal question marks', () => assert.ok(observations.literalQuestion.includes("'?'")));
    const status = await page.evaluate(() => {
      const before = document.getElementById('content').innerHTML;
      renderState({kind: 'status', message: 'Computing — previous result', retainPrevious: true});
      const stale = document.getElementById('content').style.opacity;
      const retained = before === document.getElementById('content').innerHTML;
      renderState({kind: 'status', message: 'New selection', retainPrevious: false});
      return {stale, retained, cleared: document.getElementById('content').childElementCount === 0};
    });
    check('pending previews retain labelled old data only for the same selection', () => {
      assert.equal(status.stale, '0.45'); assert.equal(status.retained, true); assert.equal(status.cleared, true);
    });
    for (const template of ['bar','line','point','area','histogram','heatmap','boxplot']) {
      await page.evaluate(template => renderState({kind:'plot', title:'Chart', plot: {
        template, overrides: {x:'x',y:'y',value:'y'}, columns:['x','y'], types:['DOUBLE','DOUBLE'],
        rows:[[1,2],[2,3],[3,4]], kinds:[['Number','Number'],['Number','Number'],['Number','Number']], rowCount:'3'
      }}), template);
      await page.waitForFunction(() => activeView !== null || document.querySelector('.plot-error'));
      const rendered = await page.evaluate(() => ({canvas:!!document.querySelector('#plot-target canvas'), error:document.querySelector('.plot-error')?.textContent}));
      check(`bundled ${template} chart renders under CSP without network or eval`, () => { assert.equal(rendered.error, undefined); assert.equal(rendered.canvas,true); });
    }
    check('offline chart rendering makes no external requests', () => assert.deepEqual(external, []));
    await page.setViewportSize({width:240,height:650});
    await page.evaluate(() => renderState({kind:'plot',title:'Chart',handle:'xlduck://p/1::abcdefghijk',plot:{
      template:'line',overrides:{x:'month',y:'records',title:'Monthly mail records in the California research selection (UTC)'},
      columns:['month','records'],types:['VARCHAR','DOUBLE'],rows:[['Jan',1354],['Feb',932]],
      kinds:[['Text','Number'],['Text','Number']],rowCount:2
    }}));
    await page.waitForFunction(() => activeView !== null);
    const narrow = await page.evaluate(() => {
      const target = document.getElementById('plot-target');
      const info = document.querySelector('.plot-info');
      const titleWidths = [];
      function visit(item) {
        if (typeof item.text === 'string' && item.text.startsWith('Monthly mail records') && item.bounds)
          titleWidths.push(item.bounds.x2 - item.bounds.x1);
        for (const child of item.items || []) visit(child);
      }
      visit(activeView.scenegraph().root);
      return {width:target.clientWidth,canvas:target.querySelector('canvas').getBoundingClientRect().width,
        infoWidth:info.clientWidth,infoContent:info.scrollWidth,fullTitle:info.textContent.includes('California research selection (UTC)'),titleWidths};
    });
    check('narrow preview fits its chart and preserves the full title in wrapped details', () => {
      assert.ok(narrow.canvas <= narrow.width + 1, JSON.stringify(narrow));
      assert.ok(narrow.infoContent <= narrow.infoWidth + 1, JSON.stringify(narrow));
      assert.equal(narrow.fullTitle,true);
      assert.ok(narrow.titleWidths.length > 0 && narrow.titleWidths.every(width => width <= narrow.width), JSON.stringify(narrow));
    });
    await page.setViewportSize({width:600,height:650});
    await page.waitForFunction(width => document.querySelector('#plot-target canvas').getBoundingClientRect().width > width + 100, narrow.canvas);
    check('selected chart grows with the preview pane without reselection', () => assert.ok(true));
    await page.setViewportSize({width:1280,height:720});
    const fields = await page.evaluate(() => {
      const spec = buildVegaSpec('point',{x:'a.b',y:'c[0]'},['a.b','c[0]'],['DOUBLE','DOUBLE'],[{'a.b':1,'c[0]':2}]);
      let rejected = false;
      try { buildVegaSpec('histogram',{x:'label'},['label'],['VARCHAR'],[{label:'00123'}]); } catch { rejected = true; }
      return {x:spec.encoding.x.field,y:spec.encoding.y.field,original:spec.data.values[0]['a.b'], rejected,
        helper:durationField(['duration','__xlduck_duration_0'],0)};
    });
    check('literal SQL chart fields and duration helpers cannot collide or become nested paths', () => {
      assert.equal(fields.x,'a\\.b'); assert.equal(fields.y,'c\\[0\\]'); assert.equal(fields.original,1);
      assert.equal(fields.helper,'__xlduck_duration_0_'); assert.equal(fields.rejected,true);
    });
    const mapValidation = await page.evaluate(async () => {
      const plot = {overrides:{lat:'lat',lon:'lon'}, columns:['lat','lon'], types:['DOUBLE','DOUBLE']};
      const cases = [
        [{...plot,overrides:{lat:'missing',lon:'lon'}},[{lat:51.5,lon:0}]],
        [{...plot,types:['VARCHAR','DOUBLE']},[{lat:'51.5junk',lon:0}]],
        [plot,[{lat:91,lon:0}]], [plot,[{lat:0,lon:181}]], [plot,[{lat:null,lon:0}]],
        [{...plot,overrides:{...plot.overrides,basemap:'typo'}},[{lat:0,lon:0}]]
      ];
      const rejected = [];
      for (const [definition, values] of cases) {
        const rows = values.map(row => [row.lat,row.lon]);
        renderState({kind:'plot',title:'Invalid map',plot:{...definition,template:'map',rows,rowCount:'1',
          kinds:rows.map(row => row.map(value => value === null ? 'Null' : typeof value === 'string' ? 'Text' : 'Number'))}});
        await new Promise(requestAnimationFrame);
        rejected.push(!!document.querySelector('#plot-target .plot-error') && activeMap === null);
      }
      return rejected;
    });
    check('maps reject unknown fields, text coercion and invalid coordinates before drawing', () => assert.ok(mapValidation.every(Boolean)));
    await page.evaluate(() => renderState({kind:'plot', title:'Map', plot: {
      template:'map', overrides:{lat:'lat',lon:'lon'}, columns:['lat','lon'], types:['DOUBLE','DOUBLE'],
      rows:[[51.5,-0.1],[52,0.1]], kinds:[['Number','Number'],['Number','Number']], rowCount:'2'
    }}));
    await page.waitForFunction(() => !!document.querySelector('.leaflet-container'));
    check('bundled map renders with unavailable tiles', () => assert.ok(true));
    check('preview has no uncaught JavaScript errors', () => assert.deepEqual(errors, []));
    console.log(`${passed} passed, ${failed} unexpected failures, ${knownFailures} known defects reproduced.`);
    process.exitCode = failed || (knownFailures && !process.argv.includes('--allow-known-defects')) ? 1 : 0;
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
