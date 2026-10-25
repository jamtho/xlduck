// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
// Paste these helpers into a permitted Office.js script. They never write cells.
// Worksheet.evaluate requires ExcelApiDesktop 1.1; use worksheet formulas if unavailable.
async function xlduckExplore(ctx, sheetName, request, waitMs = 4000) {
    if (!Number.isFinite(waitMs) || waitMs < 0 || waitMs > 10000) throw new Error('waitMs must be between 0 and 10000.');
    if (typeof request.id !== 'string' || !/^[A-Za-z0-9-]{1,64}$/.test(request.id)) throw new Error('Supply a request id of up to 64 letters, digits or hyphens.');
    const sheet = ctx.workbook.worksheets.getItem(sheetName);
    const deadline = Date.now() + waitMs;
    let next = request;
    let session = null;
    function formula(input) {
        return `DuckExplore(AZ1000,"${JSON.stringify(input).replaceAll('"', '""')}")`;
    }
    const transferred = formula(request).length > 255;
    function read(response) {
        if (typeof response.value !== 'string' || !response.value.startsWith('{')) throw new Error('DuckExplore did not return JSON. Check this XlDuck build and formula-evaluation support.');
        const status = JSON.parse(response.value);
        if (session && status.session && status.session !== session) throw new Error('The workbook session changed; review and start a new investigation explicitly.');
        session = status.session || session;
        return status;
    }
    if (transferred) {
        const text = JSON.stringify(request);
        if (text.length > 30000) throw new Error('Exploration request exceeds 30000 characters.');
        const parts = [];
        let part = '';
        // Measure encoded formula size, including escaped quotes and JSON backslashes.
        for (const character of text) {
            const probe = {op: 'stage', id: request.id, index: 2047, count: 2048, text: part + character};
            if (formula(probe).length > 255) {
                if (!part) throw new Error('Request id leaves no room for a transfer part.');
                parts.push(part); part = character;
            } else part += character;
        }
        if (part) parts.push(part);
        if (parts.length > 2048) throw new Error('Exploration transfer exceeds 2048 parts.');
        // Every part is acknowledged; no computation starts until the explicit submit.
        for (let offset = 0; offset < parts.length; offset += 32) {
            const pending = parts.slice(offset, offset + 32).map((text, index) => sheet.evaluate(formula({op: 'stage', id: request.id, index: offset + index, count: parts.length, text})));
            await ctx.sync();
            for (const response of pending) {
                const status = read(response);
                if (status.state !== 'staged') return status;
            }
        }
        next = {op: 'submit', id: request.id};
    }
    for (;;) {
        const response = sheet.evaluate(formula(next));
        await ctx.sync();
        const status = read(response);
        if (transferred && status.state === 'initializing') return status;
        if (['succeeded', 'failed', 'cancelled'].includes(status.state) || Date.now() >= deadline) return status;
        // Initialization admits no operation. Once admitted, every wait is a poll of the same id.
        next = status.state === 'initializing' ? request : {op: 'poll', id: request.id};
        await new Promise(resolve => setTimeout(resolve, 100));
    }
}

async function xlduckWaitCell(ctx, sheetName, address, waitMs = 4000) {
    if (!/^[A-Za-z]{1,3}[1-9][0-9]{0,6}$/.test(address)) throw new Error('Supply one A1-style cell address.');
    if (!Number.isFinite(waitMs) || waitMs < 0 || waitMs > 10000) throw new Error('waitMs must be between 0 and 10000.');
    const sheet = ctx.workbook.worksheets.getItem(sheetName);
    const deadline = Date.now() + waitMs;
    let session = null;
    for (;;) {
        const response = sheet.evaluate(`DuckCellStatus(${address})`);
        await ctx.sync();
        if (typeof response.value !== 'string' || !response.value.startsWith('{')) throw new Error('DuckCellStatus did not return JSON.');
        const status = JSON.parse(response.value);
        if (session && status.session && status.session !== session) throw new Error('The workbook session changed while waiting.');
        session = status.session || session;
        if (['succeeded', 'failed', 'cancelled', 'awaiting_calculation', 'not_evaluated', 'paused'].includes(status.state) || Date.now() >= deadline) return status;
        await new Promise(resolve => setTimeout(resolve, 100));
    }
}

// In run_officejs, after the helper declarations, for example:
// return await xlduckExplore(ctx, 'Research', {
//     op: 'schema', id: 'schema-001', sql: 'SELECT * FROM read_parquet(?)',
//     arguments: ['s3://your-bucket/your-file.parquet']
// });
