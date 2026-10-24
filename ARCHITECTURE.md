# XlDuck architecture

XlDuck uses Excel's dependency graph to connect immutable query results. Excel stays responsive while each enabled workbook computes through its own DuckDB worker. The production worksheet path is `XlDuck/Host/Functions.cs`; no legacy execution path remains.

## Boundaries

| Project | Responsibility |
| --- | --- |
| `XlDuck` | Excel-DNA functions/commands, COM workbook lifecycle, ribbon, selected-cell preview |
| `XlDuck.Host` | Testable approval storage, Excel date conversion and formula/navigation policies |
| `XlDuck.Core` | Typed requests, SQL-aware parameter preparation, leases, bounded scheduling and transport contracts |
| `XlDuck.Client` | Owned worker process, framed stdio, cancellation, failure detection and spill cleanup |
| `XlDuck.Worker` | One workbook's private engine process |
| `XlDuck.Engine` | DuckDB connection, parsed capability enforcement, curated readers, temporary tables and typed readback |

The Excel process does not load native DuckDB. Each worker owns one in-memory connection and executes serially; DuckDB can use two threads internally. Default limits include 512 MB engine memory, 1 GiB spill and bounded request/response sizes. These do not bound every allocation in the worker or Excel.

## Computation and publication

A request identity includes kind, exact SQL, typed scalar values, upstream handles, capture contents/dimensions and plot options. A formula holds result leases independently of temporary RTD subscriptions. Multiple/nested XlDuck calls in one formula retain separate bindings. Identical retained requests share execution; a changed request receives a new immutable handle.

Handles have the form `xlduck://<session>/<kind>/<result>`. They are opaque and reject foreign or expired sessions. Materialized results keep a snapshot; fragments keep recipes and their inputs. A plot owns one materialization of its source, including fragment sources. Unowned historical table results are bounded; live results survive cache pressure. See [the core contract](XlDuck.Core/README.md).

Pending functions return Excel `#N/A` immediately. Work runs outside the Excel main thread. Completion queues a macro, checks workbook/session/binding identity and the current formula, then dirties and calculates the requesting cell in automatic mode. The UDF returns its cached value or array, and Excel propagates dependencies. RTD provides a fallback completion notification; it does not carry result arrays. No blocking grace period or global RTD interval change is required.

Manual mode leaves calculation to the user. The next explicit calculation sees completed cached values and can start subsequent work. Pause holds new computations without cancelling the running one. Cancel interrupts outstanding work and reads while retaining successful results. Retry deliberately resubmits failures. Restart retires all leases and the worker for one workbook and rebuilds its XlDuck formula anchors without calculating unrelated formula anchors explicitly.

## Workbook lifecycle

The main-thread host identifies live workbooks by COM identity and sheets by Excel's stable sheet ID. File names and captions are not session identities. Save As and sheet renaming preserve the open session. Edits and sheet deletion release obsolete bindings; publication ignores old formulas. `BeforeClose` merely schedules reconciliation because Excel may still cancel closure. A session is retired only after its workbook is absent.

Local approval records contain the canonical path, SHA-256 fingerprint of the saved workbook and non-secret connection settings. Writes use a temporary file and atomic replacement; reads fail closed on corruption or mismatch. Open workbooks are checked asynchronously. A matching saved revision enables a fresh session; a copy or external modification does not inherit approval. Enabled saves update the approval. Settings changes restart the affected session.

Excel-DNA's `IExcelAddIn.AutoClose` is insufficient for application exit; the ribbon's COM shutdown/disconnection callbacks also retire the host. Worker disposal first requests graceful exit, then terminates only its owned child if necessary. Panes and callbacks are disposed as windows/workbooks close. A host and its worker hold file leases on their private spill directory. The next host reclaims marked crash leftovers only when neither lease remains; active, unmarked and linked directories are left alone.

## Read and compute boundary

SQL is parsed before execution and checked throughout its tree, including nested expressions and macros. Exactly one SELECT and reviewed built-in functions/readers are admitted. Engine-generated temporary-table operations use a separate privileged path. Worksheet SQL cannot export, mutate, install extensions or configure credentials. Native and extension versions/hashes are pinned, autoload/install and persistent secrets disabled, and session settings locked.

Enabled workbooks have broad access to readable files and networks under the user's account. This is a capability boundary and fault-isolation design, not an operating-system sandbox. [The engine policy](XlDuck.Engine/README.md) records the exact supported surface and limitations.

## Preview

The existing interaction remains: select a table, fragment or plot handle and view the task pane. Selection and host completion events drive a cancellable preview fetch; there is no second dependency graph. A pending or failed selected computation can retain the last successful chart with a visible stale label. Another selection clears it. Preview data cannot become a worksheet input.

Tables show at most 200 rows; charts at most 5,000. Samples carry total row counts and an explicit label. Typed JSON preserves text and exact values. Decimal charting requires an explicit approximate SQL cast; unsafe integers are rejected. HTML escaping covers text and attribute contexts.

WebView2 serves the bundled HTML/JS/CSS from a fixed virtual origin. CSP prohibits inline scripts, eval, connections and frames. Vega uses its expression interpreter. Navigation, downloads, new windows, host objects and permissions are blocked; fixed OpenStreetMap/Esri image tile paths are the only external requests allowed by the host. The renderer libraries and their licences come from integrity-pinned build dependencies.

## Evidence and remaining release work

The production suite exercises trust, typed DAGs, reuse, nested calls, date systems, workbook isolation, cancellation/retry, save/reopen, cancelled closure and formula movement. Native preview smoke tests verify WebView startup and updates without reselection; browser tests render all eight templates and check injection, exact conversions and stale states. See [tests](tests/README.md).

Visual layout at different DPI settings, live AWS/SSO and enterprise networks, clean-machine installation, signing, runtime coexistence and upgrade/removal remain explicit release gates. Earlier prototype measurements and the execution decision are preserved under [prototypes/Execution](prototypes/Execution/DECISION.md).
