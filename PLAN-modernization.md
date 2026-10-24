# XlDuck modernization plan

XlDuck will be a polished public Excel add-in for researchers who compose SQL computations through cells and table handles. AI working inside Excel is expected to write most of the SQL. The product should make experimentation fast, results dependable, and intermediate computations easy to inspect.

The design decisions below reflect the discussion through 24 October 2026. The user authorized execution on 24 October 2026. Work starts with the safe test harness and execution prototype, followed by an architecture checkpoint. Existing workbooks and APIs need not remain backward compatible.

Milestone 1 is complete: the isolated Excel harness has 95 passing legacy assertions, explicit isolation and cleanup tests, and maintained core and browser regression baselines. The baselines reproduce 13 core defects and two browser defects; those product defects remain unfixed.

Milestone 2 has established the execution choice: one worker per workbook, serial execution within each worker, and targeted result publication in automatic calculation mode. Real Excel tests cover short chains, branching, reuse, cancellation, shared requests, restart, closure, worker failure, and a 38-second trunk query. Manual calculation needed three explicit passes for a simple asynchronous chain and remains limited. See [the architecture checkpoint](prototypes/Execution/DECISION.md) and its saved measurements. The public runtime/deployment compatibility gate remains open; the experiment uses installed .NET 8 and does not yet validate a .NET 10 public build. Production code and chart UX have not been migrated.

## Agreed product scope

- Keep Excel, Excel-DNA, DuckDB, and the formula DAG as the foundation. Excel owns worksheet dependencies; XlDuck manages computation and result lifetimes.
- Design for sophisticated users doing research. Most queries should finish quickly, generally within a few seconds. Occasional 30 to 60 second trunk queries can work, but extensive expensive preparation belongs outside the workbook or behind better materialization choices.
- Support read-and-compute SQL only. Worksheet SQL cannot modify external data, issue general DDL or DML, export files, install arbitrary extensions, or change security settings. Internal temporary tables and engine spill files remain necessary implementation details.
- Use broad read access for enabled workbooks, with a coarse enablement decision for unfamiliar workbooks. Do not build permissions for individual queries, folders, buckets, or tables. Previously trusted workbooks can run automatically on open.
- Preserve the current chart experience: select a plot cell and view its chart in the preview pane. Improve correctness, responsiveness, status, and cleanup within that experience.
- Prioritize automatic calculation. Support manual calculation if the prototype establishes acceptable behaviour without compromising the main experience. Never silently change Excel's application-wide calculation mode.
- Defer in-sheet charts, dashboard construction, report distribution, and snapshots intended for recipients without the add-in. Keep chart definitions independent of the preview host so another display surface can be added later.

## Computation contract

Within a workbook session, a query with unchanged SQL and unchanged typed inputs reuses its existing successful result. Inputs include scalar values, captured range contents, and referenced handles. Referencing a different upstream result is a changed input. Handles identify immutable results or immutable query definitions; their contents cannot change invisibly.

| Event | Behaviour |
| --- | --- |
| Excel reevaluates unchanged SQL and inputs | Return the existing result without executing the SQL again |
| F2 and Enter on an unchanged formula | Reuse the result; this is not a refresh gesture |
| F9, Shift F9, or a full calculation | Preserve Excel's calculation scope and reuse unchanged XlDuck results |
| SQL or an input changes | Execute the new computation when Excel evaluates it; downstream cells observe the resulting new handle or value |
| An external file or S3 object changes | Do nothing automatically; no polling, expiry timer, or source freshness inference |
| A trusted workbook opens | Start a fresh workbook session and rebuild its computations in the supported automatic calculation mode |
| Restart XlDuck for this workbook | Discard that workbook's computation session, create a fresh one, and rebuild its results |

The same contract applies to table, scalar, and spilled query results, and to executing a fragment for output or preview. Plot display and selection must reuse the relevant computed data rather than independently rereading a source on every selection.

A fragment remains a query recipe, not a snapshot of external data. A changed downstream query that uses a fragment can execute its external reads again. Materializing a table is how the researcher retains a particular computed result for reuse. No automatic source classification or separate refresh graph is required.

Queries containing time-dependent or random expressions follow the same reuse rule. External database contents and time are not implicit cache invalidators. This is not a guarantee that separate reads of changing external sources form a transactional snapshot.

Reuse is a guarantee for currently retained results, not a requirement to remember every historical formula or input combination indefinitely. Retain results still needed by live formulas, fragments, plots, or running work. Bound any additional cache. Failed and cancelled work needs a deliberate retry path rather than indefinite reuse or an automatic retry loop on every recalculation.

Restart is the only explicit refresh operation required in v1. It cancels old work, clears workbook result state, and starts a new session while preserving formulas, connection settings, and the preview UI. It does not restart other workbooks. Every completion carries its session and request identity so an old result cannot reappear after restart or an edit. If the selected architecture uses one worker per workbook, restarting that worker may implement the operation directly.

## Access and correctness boundaries

Read-only SQL is a product capability boundary, not a claim that arbitrary hostile SQL is safely sandboxed. Broad file and network reads still carry confidentiality implications. A separate worker process supplies fault isolation; it is not automatically an operating-system security sandbox.

Enforce the supported read-and-compute surface using the engine's statement information, controlled connection setup, curated extensions and reader functions, and read-only source connections or credentials where available. A SELECT prefix test is insufficient: read-looking statements and extension functions can have side effects. Reject unsupported capabilities before execution and test actual mutation paths. User SQL must never be promoted into the privileged internal channel used to manage temporary results.

Replace unrestricted worksheet configuration with explicit connection and application settings. Do not retain DuckExecute, unrestricted DuckConfig, or their startup ordering as compatibility features. Settings that affect query meaning must either be immutable for a session or restart the affected workbook session.

Data conversions must be explicit. Preserve text, booleans, numbers, empty values, nulls, errors, date/time meanings, and nested data according to a documented conversion policy. Never silently convert an unsupported value to a blank or a precise integer to a different number. Keep full results in DuckDB; label sampled previews and reject oversized sheet output with an actionable message instead of silently truncating it.

## Implementation sequence

### Milestone 1 Safe development and regression baseline

Make the test and reload workflow safe before using it for the rewrite. Automation must create and identify its own Excel process and temporary workbook, and close only resources it owns. Remove closing unrelated workbooks without saving, killing all Excel processes, and clearing an arbitrary active worksheet. Update the development instructions alongside the scripts.

Establish a small core test project, browser checks for preview rendering, and isolated Excel integration tests. Replace fixed sleeps with bounded waits for observable state. Keep useful existing assertions, but remove tests that require capabilities deliberately dropped by the new contract. Add failing regression cases before fixing the corresponding defects.

Acceptance: the harness runs without modifying an unrelated open workbook; useful baseline behaviours and the confirmed correctness failures are reproducible. A diagnostic build or test run must not be described as a completed product fix.

### Milestone 2 Execution prototype and architecture decision

Build the smallest real Excel path that can evaluate a source query, dependent analyses, scalar and spilled outputs, and a preview. Compare background execution inside Excel with a worker process outside Excel. Use actual Excel integration to measure publication and dependency propagation, rather than timing DuckDB alone.

The prototype must establish:

- End-to-end latency for short linear and branching DAGs, including repeated edits and cache hits. Report engine time separately from scheduling, communication, and Excel notification time.
- Behaviour with an occasional slow trunk query, cancellation, formula deletion or replacement, workbook closure, and session restart during execution.
- Isolation between two workbooks, and the memory and startup cost of separate sessions or workers.
- Transfer cost for captures, previews, and bounded sheet output, while intermediate tables remain in the engine.
- Behaviour in manual calculation mode, including whether completed asynchronous work and its dependents become visible predictably.
- A supported runtime and dependency baseline, including loading alongside other Excel add-ins and distributing native DuckDB. Evaluate a supported .NET LTS runtime and current compatible Excel-DNA, DuckDB.NET, and WebView2 versions; do not upgrade by version number alone.

Use the measurements to choose the execution location, concurrency model, and manual calculation support level. A worker is preferred only if its reliability benefits justify its latency, packaging, and transfer costs. Avoid a thread per formula or blocking Excel while waiting for query work. Do not change a global Excel RTD setting silently to improve a benchmark.

Acceptance: a short architecture decision records the measured comparison, chosen design, supported calculation modes, and performance regression workloads. Keep any viable prototype code only where it fits the chosen design. The user need only revisit this decision if the findings require a product tradeoff outside the agreed scope.

### Milestone 3 Typed core and workbook sessions

Introduce a UI-independent core for typed arguments, SQL preparation, query requests, results, and session lifetimes. Make all execution entry points use it. Bind scalar values as parameters; resolve table and fragment references through a structured or SQL-aware mechanism that respects literals and comments. Escape identifiers correctly. Do not retain regular-expression replacement of every question mark.

Implement typed capture hashing, deterministic unique column names, and efficient range ingestion. Define tested policies for null and empty values, mixed columns, large integers and decimals, nested types, dates, timestamps, and intervals. Use culture-independent serialization.

Implement per-workbook ownership, session-qualified handles, bounded scheduling, in-flight work reuse, retained successful results, cancellation, and deterministic disposal. Stored fragments and plots must own the dependencies they need. Acquire and release result ownership atomically, including exceptions before execution begins. Roll back partial internal results on failure.

Acceptance: identical live requests do not execute twice; distinct typed inputs cannot collide; obsolete completions are ignored; all output paths agree on values and types; repeated edit, cancel, restart, and close cycles do not accumulate orphan tables or unbounded dictionaries. Referenced results survive temporary Excel subscription changes without forcing unchanged queries to reread external data.

### Milestone 4 Read and compute enforcement and connections

Implement the agreed statement and capability policy before exposing the new engine as the normal worksheet path. Include every entry point, including fragment validation and execution, previews, and direct outputs. Separate privileged internal operations from worksheet SQL.

Provide a practical replacement for the existing S3 and local-file setup, then add only other explicitly supported readers. Keep connection secrets out of formulas where practical and redact credentials, signed URLs, raw query arguments, and sensitive SQL from routine logs. Use coarse workbook enablement and preserve the existing decision across ordinary restarts; define safe behaviour for unfamiliar workbook copies.

Acceptance: supported local and S3 reads work, while tested DDL, DML, file exports, extension installation, settings changes, multiple-statement bypasses, and supported-extension side-effect paths fail before their effects occur. Both core and Excel tests cover the policy. Documentation accurately describes broad read access and its limitations.

### Milestone 5 Excel integration and existing preview experience

Replace the legacy execution path with the chosen implementation and explicitly register only intended Excel functions and commands. Wire workbook open, close, edits, restart, and add-in shutdown to the correct session lifecycle. Keep F9 and Shift F9 unchanged. Preserve useful pause, resume, and cancel controls with consistent behaviour across execution paths; resuming does not itself request fresh external data.

Keep selecting a handle cell as the main preview interaction. Update the pane when that selected computation completes or changes, not only when selection changes. Preserve the last successful chart while pending or failed, clearly labelled as old; never pass its old data into new calculations as though current. Track the selected computation's state without building a second worksheet dependency engine.

Fix HTML attribute injection, typed JSON conversion, locale-sensitive numbers, interval conversion, and fragment/plot inconsistencies. Consolidate chart definitions where duplication can cause drift. Bundle or pin required renderer assets, constrain WebView2 navigation and content capabilities, and preserve intentional map/network behaviour explicitly. Dispose panes and event handlers when their owning window closes.

Acceptance: the current table and chart workflows work with correct data and clear sampling limits; pending previews settle without selecting another cell; HTML-like source text cannot execute code; closing or restarting a workbook does not leave active panes, callbacks, or queries behind. Existing chart UX remains recognizable, with no in-sheet chart feature added.

### Milestone 6 Research workflow and agent documentation

Document the final function contract and examples for quick exploratory DAGs, lazy composition, useful materialization points, captured ranges, S3 reads, and precise output handling. Explain reuse, unchanged F2 and Enter behaviour, workbook restart, retry, calculation-mode support, and rebuilding on open.

Provide inspectable schema, row counts, bounded samples, errors, and computation state through a supported surface available to the Excel agent as well as the preview UI. Specify how that agent invokes supported commands; do not assume it can access a ribbon, VBA, or a particular tool without verification. This is an integration contract, not a request to add a new AI chat product or model provider.

Update README.md, ARCHITECTURE.md, AGENTS.md, and the XlDuck agent guidance to reflect the implementation. The existing untracked skill directory contains user work and must be inspected and preserved when updating it. Remove obsolete claims and unsafe development instructions rather than accumulating parallel contradictory documentation.

Acceptance: a researcher and an Excel agent can create and inspect an analysis, change an assumption, reuse a shared trunk, recover from a failed query, and restart the workbook using the documented surface. The documented examples are exercised by integration tests where feasible.

### Milestone 7 Public release preparation

Complete a reproducible Windows 64-bit build and release process with explicit versions, native dependency packaging, clean-machine installation checks, and clear WebView2/runtime requirements. Do not distribute a 32-bit build without its required native engine. Establish signing, installation, upgrade, and removal behaviour appropriate to a public add-in.

Run core and browser checks in CI and keep Excel-dependent checks in an environment with Excel installed. Include dependency and licence notices, bounded diagnostics, and a useful support report that avoids secrets. Verify loading with representative other add-ins before finalizing the runtime choice for distribution.

Acceptance: the packaged add-in installs and runs representative research workbooks on a clean supported machine; upgrade and removal work; release artefacts match the tested source. Public publication and any paid services or signing purchases remain separate user decisions.

## Regression evidence to preserve

The review found a successful build and 95 passing existing assertions, alongside failures outside those assertions. The new suite should preserve useful baseline coverage and explicitly address these findings:

| Area | Cases |
| --- | --- |
| Parameter handling | Numeric-looking text, leading apostrophes, empty arguments followed by values, question marks inside strings and comments |
| Capture identity | Number versus text, null versus empty string, unambiguous dimensions and value boundaries, colliding deduplicated headers |
| Value conversion | Quoted column names, nested values containing nulls, large integers beyond exact Excel numeric representation, decimal commas, fractional intervals |
| Lifetimes | References leaked during preparation failure, query completion after its cell is deleted, fragment and plot dependency retention, cleanup after partial table creation |
| Recalculation | Consistent reuse across table/scalar/array paths, changed typed inputs, fresh sessions on reopen, restart during work, two workbook isolation |
| Preview | Attribute injection, updates after completion without reselection, fragment plotting, typed data and explicit limits |
| Public surface | Accidental helper UDF registration, mutation functions and unrestricted configuration, consistent pause/cancel behaviour |
| Tooling | No clearing user sheets, closing unrelated workbooks, or terminating unrelated Excel processes |

The earlier diagnostic probes are under the ignored XlDuck/obj/review-20261024 directory. Port relevant scenarios into maintained tests; do not depend on that directory remaining present after a clean build.

## Execution and review process

Implement one milestone at a time in reviewable changes. Run checks appropriate to each change, including a regression case for each bug fix. Keep the working add-in usable at integration checkpoints rather than replacing every subsystem in one untested change.

Use the prototype decision as the first architectural checkpoint. Then demonstrate one complete path from source through dependent handles to scalar, spilled, and chart output before broadening coverage. Review the integrated research workflow before release packaging.

Routine implementation decisions can be resolved within the agreed principles. Return to the user when evidence would change product behaviour, access boundaries, or scope. Do not add per-source refresh controls, fine-grained permissions, in-sheet charts, full manual-calculation support, or report distribution merely because the rewrite makes them possible.
