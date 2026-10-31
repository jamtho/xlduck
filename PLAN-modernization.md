# XlDuck modernization plan

XlDuck is a public Excel add-in for researchers composing SQL through cells and retained table handles. AI working inside Excel is expected to write most of the SQL. The rewrite shipped as 1.0.0; the current patch is **1.0.1**, putting the workbook tag first in handles. Earlier design experiments and candidate-by-candidate evidence remain in Git history. The maintained [release guide](docs/RELEASE.md) distinguishes tested coverage from accepted follow-up work.

## Agreed product scope

- Excel owns the formula dependency graph. One owned DuckDB worker per ordinary workbook provides execution and fault isolation; native DuckDB never runs inside Excel.
- Most worksheet computations should finish within a few seconds. Extensive expensive preparation belongs outside the workbook or behind deliberate retained tables.
- SQL is read-and-compute only. No worksheet writes, exports, arbitrary extension installation or unrestricted configuration. Privileged internal temporary-table operations remain separate.
- Ordinary workbooks start automatically. Preserve Excel Protected View and reject sandboxed workbooks; do not add a second approval gate or change Trust Center. Broad file/network reads are supported, so workers are not an OS security sandbox.
- Keep SQL as ordinary cell text when useful, and execute it from formulas referencing those cells. Named SQL and positional arguments remain supported.
- Keep the chart UX: select a plot handle and view it in the preview pane. Selecting inside DuckOut/DuckQueryOut spills also previews their retained table. In-sheet charts and report distribution are deferred.
- Prioritize automatic calculation, while preserving supported manual calculation behaviour. Never change Excel's global calculation mode or RTD latency.
- Connection definitions and resource budgets travel with workbook metadata. Keys stay in Windows Credential Manager with a local connection-ID-and-destination association. Save after configuration or legacy migration.
- Agent guidance is a small bootstrap skill that reads current `DuckHelp` content. Manual provider setup is acceptable for v1; discovery improvements follow real research experience.

## Computation contract

| Event | Behaviour |
| --- | --- |
| Excel reevaluates unchanged SQL and typed inputs | Reuse the retained result |
| F9, Shift+F9 or unchanged F2/Enter | Preserve Excel's calculation semantics; no implicit source refresh |
| SQL, scalar/range input or upstream handle changes | Execute the new computation when Excel evaluates it |
| External files, S3 objects, time or random values change | No automatic invalidation or polling |
| An ordinary saved workbook opens | Start a fresh workbook session and rebuild according to calculation mode |
| Restart XlDuck for this workbook | Cancel old work, discard that session and rebuild, preserving formulas and settings |

Handles identify immutable results or recipes and resolve only through the caller workbook's live registry. Cell aliases use `xlduck://ABCD/t/123` (also `f` and `p`); four-letter tags are never reused within an Excel process, but can coincide across processes or later launches. Reference cells rather than copying tokens. Worker identities remain session-qualified. Stale completions cannot revive results after an edit or restart. Formula leases are independent of RTD subscriptions.

A fragment is a query recipe, not a source snapshot. Materialize a table to reuse a particular computed result. Separate reads of changing sources do not form a transactional snapshot. Reuse applies to retained results, not an unlimited history. Failed or cancelled work has explicit retry rather than an automatic retry loop.

Preserve exact types and expose approximation explicitly. Keep full results in DuckDB, label previews and bound worksheet output. Excel spill errors, pending work, policy failures, source errors and resource exhaustion must remain distinguishable.

## Implemented milestones

1. **Safe automation and regression baseline:** owned Excel processes/workbooks, bounded waits, process-identity cleanup and exit-code checks; no unrelated workbook closure.
2. **Execution model:** measured separate workers, queued execution, immediate retained results and asynchronous completion without global Excel setting changes.
3. **Typed core:** SQL-aware binding, immutable requests, typed values, leases, cancellation, bounded protocol and workbook-local aliases.
4. **Read-and-compute boundary:** recursive SQL policy, curated readers, separate privileged operations, native Excel trust integration and explicit connection settings.
5. **Workbook and preview lifecycle:** restart, save/reopen, Save As, cancelled closure, two-workbook isolation, manual calculation, exact dates and selected-chart updates.
6. **Research and agent interface:** embedded help, connection discovery, schema inspection, per-cell status, structured errors, bounded exploration, Office.js helper and live OpenAI Excel trials.
7. **Distribution:** complete x64 package, pinned dependencies/notices, per-user install/rollback/removal, manifest and provenance checks, synthetic demo, screenshots and CI release publishing.

## Release gates

- Keep public docs, embedded help and versioned downloads aligned with the workbook-first handle format. Existing formulas rebuild their handles on open; copied token text is not a persistence format.
- Build from committed clean source. Run core/readers, browser, helper, package and installer checks, then native Excel and demo tests against the actual package.
- Update and merge the PR after CI passes. Tag the merged source; verify published asset hashes and package contents.
- Preserve private research workbook edits and the user's original skill archive outside the release changes.

Clean-machine/VM trials, signing, broader compatibility, additional provider/Claude trials, more display transitions, skill triggering improvements and in-sheet charts are accepted follow-up work. They do not block v1; see [tested scope](docs/RELEASE.md#tested-scope-and-follow-up-work). No cloud spend is needed for this release.
