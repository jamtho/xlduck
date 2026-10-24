# Agent Guidelines

## Commit messages and code

- Use concise, descriptive imperative commit messages; conventional prefixes where useful. No emoji or generated/co-author attribution.
- Target .NET 8+ and C# 12+; use file-scoped namespaces and nullable reference types.
- Keep UDF containers static and functions focused. Catch failures at Excel boundaries and return actionable `#ERROR: message` values.
- Keep DuckDB and native calls in `XlDuck.Engine`/`XlDuck.Worker`, outside Excel. Do not add an in-process native fast path.

## Structure and invariants

- `XlDuck/Host`: production UDFs, commands, workbook COM lifecycle and ribbon.
- `XlDuck/Preview`: task-pane UI, worker-backed preview, bundled HTML/JS/CSS.
- `XlDuck.Host`: testable approval, date, formula and navigation policies.
- `XlDuck.Core`: immutable typed requests, SQL-aware binding, leases, queues and protocol.
- `XlDuck.Engine`, `XlDuck.Client`, `XlDuck.Worker`: one owned DuckDB worker per enabled workbook.

Excel owns the formula DAG. Unchanged SQL and typed inputs reuse retained results. F9, Shift+F9 and unchanged F2/Enter are not source refreshes. Restart affects one workbook. Opening an approved saved workbook starts a fresh session. Never change global RTD latency or calculation mode to improve performance.

Worksheet SQL is read-and-compute only. Privileged internal temporary-table and connection operations must remain separate from user SQL. No unrestricted config, writes, exports or extension installation. Workers are fault isolation, not OS sandboxes. Preserve exact types and make approximation explicit.

Workbook state and COM access stay on Excel's main thread. A formula owns leases independently of RTD subscriptions. Do not dispose a workbook on BeforeClose alone: closure may be cancelled. Do not release borrowed Excel RCWs with ReleaseComObject. Retire stale completions by identity and current formula checks.

## Development and verification

Acquire exact build dependencies with `scripts/Get-ReaderExtensions.ps1` and `scripts/Get-PreviewAssets.ps1`. Build the 64-bit folder distribution; include worker, readers and preview assets. Do not distribute an XLL alone or stale 32-bit output.

After code changes, run the appropriate core/browser checks and the isolated production integration suite:

```powershell
powershell -ExecutionPolicy Bypass -File tests/Run-CoreTests.ps1 -WithReaders
powershell -ExecutionPolicy Bypass -File tests/Run-Tests.ps1 -WithPreview
```

Use `-NoBuild` for an existing build, `-NoRestore` on the core script after restore, and `-AddInPath` on production tests for an isolated package. Add a reproducing regression for a bug unless there is a sound reason not to. Tests should verify externally meaningful behaviour.

`scripts/Reload-AddIn.ps1` opens a new visible development Excel instance using a private build copy. It leaves existing instances alone. Close that instance manually. Ignored copies live in `artifacts/development`; remove a copy only after its owning Excel process has exited.

## Excel automation safety

1. Use `scripts/ExcelSession.ps1` to create and verify an owned process/workbook and to wait boundedly for values. Never select a test instance through ActiveWorkbook or the Running Object Table.
2. Write `.Formula2`, not `.Formula`, so dynamic arrays spill correctly. Assert `.Value2` independently of column width or formatting.
3. Do not change Trust Center settings or dismiss unrelated/security prompts automatically.
4. Never close unrelated workbooks, suppress their save prompts or kill all Excel processes. Failed tests must release their own resources in `finally`.
5. Keep original process identity/handles for cleanup; avoid PID-reuse and process-exit races.

Read `docs/FORMULAS.md` for the user/agent contract and `PLAN-modernization.md` for remaining release gates. Update these when behaviour changes. The user-provided `skill/xlduck.zip` is an original archive; preserve it when updating the unpacked agent guidance.
