# XlDuck architecture

XlDuck uses Excel's dependency graph to connect immutable query results. Excel stays responsive while each ordinary workbook computes through its own DuckDB worker. The production worksheet path is `XlDuck/Host/Functions.cs`; no legacy execution path remains.

## Boundaries

| Project | Responsibility |
| --- | --- |
| `XlDuck` | Excel-DNA functions/commands, COM workbook lifecycle, ribbon, selected-cell preview |
| `XlDuck.Host` | Connection settings and Windows credential storage, Excel date conversion and formula/navigation policies |
| `XlDuck.Core` | Typed requests, SQL-aware parameter preparation, leases, bounded scheduling and transport contracts |
| `XlDuck.Client` | Owned worker process, framed stdio, cancellation, failure detection and spill cleanup |
| `XlDuck.Worker` | One workbook's private engine process |
| `XlDuck.Engine` | DuckDB connection, parsed capability enforcement, curated readers, temporary tables and typed readback |

The Excel process does not load native DuckDB. Each worker owns one in-memory connection and executes serially; DuckDB can use two threads internally. Default limits include 512 MB engine memory, 1 GiB spill and bounded request/response sizes. These do not bound every allocation in the worker or Excel.

## Computation and publication

A request identity includes kind, exact SQL, typed scalar values, upstream handles, capture contents/dimensions and plot options. A formula holds result leases independently of temporary RTD subscriptions. Multiple/nested XlDuck calls in one formula retain separate bindings. Identical retained requests share execution; a changed request receives a new immutable handle.

Cell handles have the form `xlduck://<kind>/<22-character result ID>`. The token preserves a full 128-bit result ID without printing the session UUID. The caller workbook resolves it through its live handle registry; foreign, retired and previous-session tokens fail. Internal requests and worker frames still carry the complete session-qualified identity. Materialized results keep a snapshot; fragments keep recipes and their inputs. A plot owns one materialization of its source, including fragment sources. Unowned historical table results are bounded; live results survive cache pressure. See [the core contract](XlDuck.Core/README.md).

Pending functions return Excel `#N/A` immediately. Work runs outside the Excel main thread. Completion queues a macro, checks workbook/session/binding identity and the current formula, then dirties and calculates the requesting cell in automatic mode. The UDF returns its cached value or array, and Excel propagates dependencies. RTD provides a fallback completion notification; it does not carry result arrays. No blocking grace period or global RTD interval change is required.

Manual mode leaves calculation to the user. The next explicit calculation sees completed cached values and can start subsequent work. Pause holds new computations without cancelling the running one. Cancel interrupts outstanding work and reads while retaining successful results. Retry deliberately resubmits failures. Restart retires all leases and the worker for one workbook and rebuilds its XlDuck formula anchors without calculating unrelated formula anchors explicitly.

## Workbook lifecycle

The main-thread host identifies live workbooks by COM identity and sheets by Excel's stable sheet ID. File names and captions are not session identities. Save As and sheet renaming preserve the open session. Edits and sheet deletion release obsolete bindings; publication ignores old formulas. `BeforeClose` merely schedules reconciliation because Excel may still cancel closure. A session is retired only after its workbook is absent.

Only ordinary members of Excel's `Workbooks` collection are registered; Protected View workbooks are in a separate collection. The host also checks `Workbook.Application.IsSandboxed` before session creation, UDF access, preview selection and publication. Leaving Protected View through Excel's Enable Editing creates a new ordinary session. No custom enable/disable or saved-revision approval gate remains. Workers still start lazily at the first computation.

Local connection records contain the canonical workbook path and non-secret settings. Atomic writes survive ordinary saves and external edits; Save As records the settings for the new path without restarting. An unfamiliar path uses defaults. Saved access keys are referenced by random IDs in Windows Credential Manager, resolved only when starting a worker, and passed over its private stdio channel into temporary scoped DuckDB secrets. Removing a connection does not delete a key that another workbook might reference; users can remove unused XlDuck/S3 entries in Windows Credential Manager. Settings changes restart only the affected session.

Excel-DNA's `IExcelAddIn.AutoClose` is insufficient for application exit; the ribbon's COM shutdown/disconnection callbacks also retire the host. Worker disposal first requests graceful exit, then terminates only its owned child if necessary. Panes and callbacks are disposed as windows/workbooks close. A host and its worker hold file leases on their private spill directory. The next host reclaims marked crash leftovers only when neither lease remains; active, unmarked and linked directories are left alone.

## Read and compute boundary

SQL is parsed before execution and checked throughout its tree, including nested expressions and macros. Exactly one SELECT and reviewed built-in functions/readers are admitted. Engine-generated temporary-table operations use a separate privileged path. Worksheet SQL cannot export, mutate, install extensions or configure credentials. Native and extension versions/hashes are pinned, autoload/install and persistent secrets disabled, and session settings locked.

Ordinary workbooks have broad access to readable files and networks under the user's account. This is a capability boundary and fault-isolation design, not an operating-system sandbox. [The engine policy](XlDuck.Engine/README.md) records the exact supported surface and limitations.

## Preview

The existing interaction remains: select a table, fragment or plot handle and view the task pane. Selection and host completion events drive a cancellable preview fetch; there is no second dependency graph. A pending or failed selected computation can retain the last successful chart with a visible stale label. Another selection clears it. Preview data cannot become a worksheet input.

Tables show at most 200 rows; charts at most 5,000. Samples carry total row counts and an explicit label. Typed JSON preserves text and exact values. Decimal charting requires an explicit approximate SQL cast; unsafe integers are rejected. HTML escaping covers text and attribute contexts.

WebView2 serves the bundled HTML/JS/CSS from a fixed virtual origin. CSP prohibits inline scripts, eval, connections and frames. Vega uses its expression interpreter. Navigation, downloads, new windows, host objects and permissions are blocked; fixed OpenStreetMap/Esri image tile paths are the only external requests allowed by the host. The renderer libraries and their licences come from integrity-pinned build dependencies.

## Evidence and remaining release work

The production suite exercises trust, typed DAGs, reuse, nested calls, date systems, workbook isolation, cancellation/retry, save/reopen, cancelled closure and formula movement. Native preview smoke tests verify WebView startup and updates without reselection; browser tests render all eight templates and check injection, exact conversions and stale states. See [tests](tests/README.md).

Visual layout at different DPI settings, live AWS/SSO and enterprise networks, clean-machine installation, signing, runtime coexistence and upgrade/removal remain explicit release gates. Earlier prototype measurements and the execution decision are preserved under [prototypes/Execution](prototypes/Execution/DECISION.md).
