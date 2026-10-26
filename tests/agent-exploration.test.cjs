// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
const {readFileSync} = require('node:fs');
const {join} = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const sandbox = {setTimeout};
vm.createContext(sandbox);
vm.runInContext(readFileSync(join(__dirname, '../docs/agent-exploration.js'), 'utf8'), sandbox);
const withoutTimers = {};
vm.createContext(withoutTimers);
vm.runInContext(readFileSync(join(__dirname, '../docs/agent-exploration.js'), 'utf8'), withoutTimers);

function fixture(states, wrap = value => value) {
    const calls = [];
    let pending = [];
    const ctx = {
        workbook: {worksheets: {getItem(name) {
            assert.equal(name, 'Research');
            return {evaluate(formula) {
                assert.ok(formula.length <= 255, 'Exceeded native Evaluate expression limit');
                calls.push(formula); const response = {}; pending.push({formula, response}); return response;
            }};
        }}},
        async sync() {
            for (const {formula, response} of pending) {
                if (typeof states === 'function') response.value = wrap(JSON.stringify(states(formula)));
                else { assert.ok(states.length, 'Unexpected extra evaluation'); response.value = wrap(JSON.stringify(states.shift())); }
            }
            pending = [];
        }
    };
    return {ctx, calls};
}
(async () => {
    const request = {op: 'query', id: 'test-1', sql: 'SELECT "quoted" FROM ?'};
    let f = fixture([{state: 'queued', session: 's1'}, {state: 'running', session: 's1'}, {state: 'succeeded', session: 's1'}]);
    assert.equal((await sandbox.xlduckExplore(f.ctx, 'Research', request)).state, 'succeeded');
    assert.equal(f.calls.length, 3);
    assert.ok(f.calls[0].includes('SELECT'));
    assert.ok(f.calls.slice(1).every(call => call.includes('poll') && !call.includes('SELECT')));
    f = fixture([{state: 'initializing'}, {state: 'queued', session: 's1'}, {state: 'failed', session: 's1'}]);
    assert.equal((await sandbox.xlduckExplore(f.ctx, 'Research', request)).state, 'failed');
    assert.equal(f.calls[0], f.calls[1]);
    f = fixture([{state: 'queued', session: 's1'}]);
    assert.equal((await sandbox.xlduckExplore(f.ctx, 'Research', request, 0)).state, 'queued');
    f = fixture([{state: 'running', session: 's1'}, {state: 'failed', session: 's2'}]);
    await assert.rejects(sandbox.xlduckExplore(f.ctx, 'Research', request), /session changed/);
    f = fixture([{state: 'awaiting_calculation', session: 's1', resultReady: true}]);
    assert.equal((await sandbox.xlduckWaitCell(f.ctx, 'Research', 'A1')).state, 'awaiting_calculation');
    assert.equal(f.calls.length, 1);
    await assert.rejects(sandbox.xlduckWaitCell(f.ctx, 'Research', 'A1),Other()'), /cell address/);
    // Live desktop Office.js can wrap an XLL scalar result in a one-cell array.
    for (const wrap of [value => [value], value => [[value]]]) {
        f = fixture([{state: 'queued', session: 's1'}, {state: 'succeeded', session: 's1'}], wrap);
        assert.equal((await sandbox.xlduckExplore(f.ctx, 'Research', request)).state, 'succeeded');
        assert.equal(f.calls.length, 2);
        f = fixture([{state: 'succeeded', session: 's1', resultReady: true}], wrap);
        assert.equal((await sandbox.xlduckWaitCell(f.ctx, 'Research', 'A1')).resultReady, true);
    }
    for (const wrap of [() => null, () => [], value => [value, value], value => [[value, value]], value => [[[value]]]]) {
        f = fixture([{state: 'succeeded'}], wrap);
        await assert.rejects(sandbox.xlduckExplore(f.ctx, 'Research', request), /did not return one JSON value/);
        f = fixture([{state: 'succeeded'}], wrap);
        await assert.rejects(sandbox.xlduckWaitCell(f.ctx, 'Research', 'A1'), /did not return one JSON value/);
    }
    // The connected agent's Office.js sandbox does not provide browser timers.
    f = fixture([{state: 'queued', session: 's1'}], value => [value]);
    assert.equal((await withoutTimers.xlduckExplore(f.ctx, 'Research', request)).state, 'queued');
    assert.equal(f.calls.length, 1);
    f = fixture([{state: 'succeeded', session: 's1'}], value => [value]);
    assert.equal((await withoutTimers.xlduckExplore(f.ctx, 'Research', {op:'poll', id:request.id})).state, 'succeeded');
    assert.ok(f.calls[0].includes('poll') && !f.calls[0].includes('SELECT'));
    f = fixture([{state:'running',session:'s1'}], value => [value]);
    assert.equal((await withoutTimers.xlduckWaitCell(f.ctx, 'Research', 'A1')).state, 'running');
    assert.equal(f.calls.length, 1);
    const long = {op: 'query', id: 'a'.repeat(64), sql: 'SELECT ? /*' + '"\\\n🙂'.repeat(500) + '*/', arguments: [42]};
    const chunks = [];
    let submits = 0;
    f = fixture(formula => {
        const command = JSON.parse(formula.slice('DuckExplore(AZ1000,"'.length, -2).replaceAll('""', '"'));
        if (command.op === 'stage') { chunks[command.index] = command.text; return {state: 'staged', session: 's1'}; }
        assert.equal(command.op, 'submit'); submits++;
        assert.deepEqual(JSON.parse(chunks.join('')), long);
        return {state: 'succeeded', session: 's1'};
    }, value => [value]);
    assert.equal((await sandbox.xlduckExplore(f.ctx, 'Research', long)).state, 'succeeded');
    assert.equal(submits, 1); assert.ok(chunks.length > 1);
    console.log('PASS agent helpers poll one request, bound waiting, preserve session identity and stop for manual calculation');
    console.log('PASS long requests transfer in bounded Evaluate expressions without changing Unicode, escapes or arguments');
    console.log('PASS live Office.js scalar/one-cell response shapes work; missing or multi-cell responses fail explicitly');
    console.log('PASS timer-free agent hosts return pending status for a later poll without spinning or resubmitting');
})().catch(error => { console.error(error); process.exitCode = 1; });
