# Typed core and workbook sessions

This is the next implementation of XlDuck's computational core. The existing `XlDuck/` add-in still uses the legacy implementation. The new core is exercised through a separate Excel test add-in; it is not yet the public worksheet path.

## Project boundaries

| Project | Responsibility |
| --- | --- |
| `XlDuck.Core` | Immutable typed requests and handles, SQL argument preparation, ownership, scheduling, transport contracts |
| `XlDuck.Engine` | One DuckDB connection, native cancellation and appender ingestion, temporary result storage, typed readback |
| `XlDuck.Worker` | Private child process hosting one engine session over bounded framed standard input/output |
| `XlDuck.Client` | Worker startup, request correlation, cancellation, failure detection and owned spill-directory cleanup; no DuckDB dependency |
| `tests/XlDuck.TestAddIn` | Real Excel verification of the core and targeted publication; not a production host |

These projects currently target installed .NET 8 with Excel-DNA 1.9.0 and DuckDB.NET 1.5.6.1. The public runtime, coexistence and distribution gate in the [architecture checkpoint](../prototypes/Execution/DECISION.md) remains open.

## Request and ownership contract

`WorkbookSession.Acquire` returns a `ResultLease`. A live formula owns that lease independently of transient RTD subscriptions. A request's identity includes its kind, SQL, typed scalar arguments, immutable upstream handles, capture dimensions/content and plot options. Identical retained requests share work and results. Releasing one owner does not cancel another owner's computation; releasing the final owner of pending work requests cancellation.

Handles include a session ID, result kind and result ID. Old-session and foreign-workbook handles are rejected. A changed request receives a new handle. Restart creates a new session; it does not reuse results from the previous session.

Materialized tables own their stored result. During execution they retain their inputs; afterwards the snapshot can survive without those inputs. Fragments retain their referenced recipes and tables. A plot retains one materialized data result, including when its source is a fragment. Output and preview readers acquire temporary data leases while reading. The host must retain the materialization for as long as its live output or preview needs reuse.

Only unowned materialized results and failed table-request records enter the bounded historical cache (32 by default). Unowned fragments and plots retire immediately, avoiding hidden retention of entire dependency graphs. Live results are never evicted by history pressure. Dropping a stored result is tracked work; cleanup failure faults the session and requires restart.

Failed requests remain stable while retained. `Acquire(request, retryFailed: true)` creates a new attempt with a new handle; existing owners of the failure keep their original result. The public retry command is still to be wired. A cancelled request is not a successful cache entry. Session disposal cancels work, waits for reads and cleanup, disposes the backend, and releases bookkeeping.

## SQL arguments

Use positional `?` arguments. Scalar values are bound parameters, including text that looks numeric, empty strings and leading apostrophes. Table handles become escaped internal identifiers. Fragment handles become parenthesized subqueries with recursively renumbered parameters. Plot handles cannot be used as SQL table arguments; use their data through the output/preview API.

The scanner respects single-quoted strings, escaped `E'...'` strings, quoted identifiers, dollar-quoted strings, line comments and nested block comments. A final semicolon is accepted; multiple statements are rejected. Named `$parameters` are reserved for the binder. Identifiers containing `$` must be quoted. Placeholder count must match argument count. Nesting is capped at 64 and cumulative expansion work is bounded; materialize a shared intermediate result if fragment expansion becomes excessive.

DuckDB's own SQL parser also validates the single SELECT query before internal materialization. This is only the statement check. The full reader/function/extension capability policy is not implemented, and **external access is disabled** in this stage. Do not enable external access or replace the normal worksheet path until that policy and workbook enablement are implemented. Worker isolation is not an OS security sandbox.

## Values

Transport preserves type alongside invariant text. It does not guess types from text contents.

| Value | Current policy |
| --- | --- |
| Text | Preserve exactly, including numeric-looking text, apostrophes and empty strings |
| Empty cell / SQL null | Distinct request identities; both bind as SQL NULL and display as blank |
| Boolean | Typed boolean throughout |
| Excel number | Finite double; no parsing numeric-looking text into a number |
| Integer | Exact invariant integer; Excel receives a number only within the safe integer range ±(2^53−1), otherwise text |
| Decimal | Exact text through transport and sheet output; scalar arguments use an explicit DECIMAL cast with up to 38 digits; SQL `CAST AS DOUBLE` opts into approximation |
| Date / time / timestamp | Typed invariant values; host-neutral display helper provides OA date serials and fractional days where representable, otherwise text; final host must account for workbook date system and formatting |
| Zoned or nanosecond timestamp | Preserve exact textual representation; do not silently remove zone or nanoseconds |
| Interval | Preserve DuckDB's invariant representation, including fractional seconds |
| Nested data | Preserve JSON, including nested nulls and exact numeric text; chart rendering must not silently round large JSON numbers |
| Nonfinite number / Excel error / unsupported input | Fail explicitly instead of replacing with blank or an unrelated value |

Captures require a homogeneous nonblank type per column. Mixed columns fail with an actionable error; no implicit text/number coercion. All-blank columns become nullable VARCHAR. Decimal capture is currently unsupported; exact decimal scalar binding and SQL results are supported. Headers are unique without stealing explicitly supplied names, and identifier quotes are escaped.

Excel `Value2` supplies dates as numbers: formatting alone is not a typed date argument. Final host date-system handling and chart conversions remain integration work.

## Bounds and worker failure

Defaults are 64 pending computations/cleanup jobs, 64 active output reads, one million capture cells and eight million input characters. Worker admission is separately capped at 64 requests. SQL expansion has an eight-million-character cumulative construction budget, which includes intermediate fragment construction. Transport frames are capped at 32 MiB. Output is limited by requested rows/cells and eight million characters. These are development defaults, not a public tuning API.

Sheet output rejects oversized results without truncation. Preview output returns a bounded sample with total row count and `IsSample`. Scalar output requires exactly one row and one column. Intermediate tables remain in DuckDB. The worker configures two execution threads, a 512 MB engine memory limit and at most 1 GiB of spill; the engine limit does not bound all process memory or live results.

Cancellation waits for worker acknowledgement before releasing source tables. If the worker cannot acknowledge within five seconds, the client terminates only its owned child. The deadline includes blocked pipe writes. Disposal allows two seconds for a graceful exit, then terminates that child. Worker exit fails both outstanding and subsequent calls. A fresh workbook session is the recovery path.

The client owns a newly created temporary spill directory and removes it after shutdown or worker failure. It never takes ownership of a user directory. An abrupt Excel/host crash can still leave this directory behind; orphan recovery is remaining packaging/lifecycle work.

## Evidence and remaining host work

On 24 October 2026: 17 core specifications, 14 direct-engine cases, 17 worker cases and nine real Excel integration checks pass. Tests cover exact values, placeholder parsing, shared work, fragments/plots, bounded history/admission, native cancellation, physical table cleanup, worker failure, blocked transport and excessive fragment expansion. New regression cases were observed failing before the transport and expansion fixes.

The Excel path exercises a source plus eight dependent queries, scalar and spilled outputs, captured ranges and fragment plot data. Input edits settled in about 30–41 ms across the recorded checks on this development machine. Full calculation and unchanged formula reentry reused live results with historical caching disabled. Clearing formulas released all result nodes, and restart preserved another workbook's results.

Publication queues a macro on Excel's main thread, verifies the current session and binding, then dirties and calculates the requesting cell. The UDF returns its completed cached value. Excel's global RTD throttle is unchanged. Automatic calculation is the primary mode; the harness does not force publication in manual mode.

The test add-in deliberately has a narrow surface: three SQL arguments, one core invocation per cell, workbook-name identity, explicit initialization and preview JSON rather than a chart pane. It does not implement public enablement, Save As/rename, cancelled workbook closure, all formula/edit patterns, date-system handling, production diagnostics or the final function registry. Those must be implemented and tested when replacing the production host; copying the harness unchanged is not sufficient.

Run [the new-core suite](../tests/README.md) separately from the still-legacy regression suite. Passing new-core tests does not fix legacy entry points or the known browser defects.
