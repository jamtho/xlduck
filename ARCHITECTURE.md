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

The Excel process does not load native DuckDB. Each worker owns one in-memory connection and executes serially. Host-controlled resource defaults are two query threads, 1024 MiB DuckDB memory and 4 GiB spill. The Resources dialog stores validated per-workbook settings and restarts that workbook to apply them. These budgets do not bound every allocation in the worker or Excel, and several workers can consume their budgets at once.

Core entries retain timing, state and dependency aliases independently of released input leases. Protocol v3 carries structured public failure codes across the worker boundary. `DuckCellStatus` aggregates the formula's live bindings, synchronous failures and Excel display state; it accepts a reference even before a result handle exists. It describes last evaluated inputs, and reports ready-but-unpublished work separately in manual mode.

`DuckSchema` and fragment schema inspection use a checked zero-row query; source binding may still read metadata or infer types. `DuckExplore` is a bounded read-only interface anchored explicitly to a workbook. It uses the same session and worker, with eight tracked requests, small typed results, 30-second deadlines and idempotent polling. Exploration leases never replace worksheet ownership. The Office.js bridge uses desktop Worksheet.evaluate; native integration tests and a live Windows Excel agent trial establish the complete roundtrip in the observed host.

## Computation and publication

A request identity includes kind, exact SQL, typed scalar values, upstream handles, capture contents/dimensions and plot options. A formula holds result leases independently of temporary RTD subscriptions. Multiple/nested XlDuck calls in one formula retain separate bindings. Identical retained requests share execution; a changed request receives a new immutable handle.

Cell handles have the form `xlduck://<session tag>/<kind>/<integer>`, for example `xlduck://ABCD/t/42`. The caller workbook resolves the compact alias through its live registry; unregistered and retired tokens fail. Tags are four random uppercase letters, never reused within the Excel process. They are display discriminators, not authentication: unrelated Excel processes can coincidentally issue the same alias, so always reference producing cells rather than copying handle text. Numbers are never reused within a session. Internal requests and worker frames still carry the complete session-qualified UUID identity. Materialized results keep a snapshot; fragments keep recipes and their inputs. A plot owns one materialization of its source, including fragment sources. Unowned historical table results are bounded; live results survive cache pressure. See [the core contract](XlDuck.Core/README.md).

Pending functions return Excel `#N/A` immediately. Work runs outside the Excel main thread. Completion queues a macro, checks workbook/session/binding identity and the current formula, then dirties and calculates the requesting cell in automatic mode. The UDF returns its cached value or array, and Excel propagates dependencies. RTD provides a fallback completion notification; it does not carry result arrays. No blocking grace period or global RTD interval change is required.

Manual mode leaves calculation to the user. The next explicit calculation sees completed cached values and can start subsequent work. Pause holds new computations without cancelling the running one. Cancel interrupts outstanding work and reads while retaining successful results. Retry deliberately resubmits failures. Restart retires all leases and the worker for one workbook and rebuilds its XlDuck formula anchors without calculating unrelated formula anchors explicitly.

## Workbook lifecycle

The main-thread host identifies live workbooks by COM identity and sheets by Excel's stable sheet ID. File names and captions are not session identities. Save As and sheet renaming preserve the open session. Edits and sheet deletion release obsolete bindings; publication ignores old formulas. `BeforeClose` merely schedules reconciliation because Excel may still cancel closure. A session is retired only after its workbook is absent.

Only ordinary members of Excel's `Workbooks` collection are registered; Protected View workbooks are in a separate collection. The host also checks `Workbook.Application.IsSandboxed` before session creation, UDF access, preview selection and publication. Leaving Protected View through Excel's Enable Editing creates a new ordinary session. No custom enable/disable or saved-revision approval gate remains. Workers still start lazily at the first computation.

Connection definitions and resource budgets live in a versioned custom XML part inside each workbook. Connections have stable IDs and readable names; Save As, moves and copies retain them. Explicit serialization excludes vault references and secrets. Local credential associations bind a connection ID and destination (endpoint, HTTPS, addressing style and signing region) to a Windows Credential Manager entry. Copies on the same computer can reuse matching associations; another computer needs its own credentials. A name change preserves the association, while a destination change requires one for the new server. Keys resolve only when starting a worker and pass over its private stdio channel into temporary scoped DuckDB secrets.

Legacy path-based records migrate on open, including OneDrive URL records written by earlier builds; the workbook must be saved to retain migration. Conflicting records, invalid metadata and duplicate parts fail explicitly instead of falling back to defaults. Applying valid settings repairs the metadata. Removing a connection does not delete a key that another workbook might reference; users can remove unused XlDuck/S3 entries in Windows Credential Manager. Settings changes restart only the affected session. CSV and legacy .xls cannot retain this metadata.

Excel-DNA's `IExcelAddIn.AutoClose` is insufficient for application exit; the ribbon's COM shutdown/disconnection callbacks also retire the host. Worker disposal first requests graceful exit, then terminates only its owned child if necessary. Panes and callbacks are disposed as windows/workbooks close. A host and its worker hold file leases on their private spill directory. The next host reclaims marked crash leftovers only when neither lease remains; active, unmarked and linked directories are left alone.

## Read and compute boundary

SQL is parsed before execution and checked throughout its tree, including nested expressions and macros. Exactly one SELECT and reviewed built-in functions/readers are admitted. Engine-generated temporary-table operations use a separate privileged path. Worksheet SQL cannot export, mutate, install extensions or configure credentials. Native and extension versions/hashes are pinned, autoload/install and persistent secrets disabled, and session settings locked.

Ordinary workbooks have broad access to readable files and networks under the user's account. This is a capability boundary and fault-isolation design, not an operating-system sandbox. [The engine policy](XlDuck.Engine/README.md) records the exact supported surface and limitations.

## Preview

The existing interaction remains: select a table, fragment or plot handle and view the task pane. Selection and host completion events drive a cancellable preview fetch; there is no second dependency graph. A pending or failed selected computation can retain the last successful chart with a visible stale label. Another selection clears it. Preview data cannot become a worksheet input.

Tables show at most 200 rows; charts at most 5,000. Samples carry total row counts and an explicit label. Typed JSON preserves text and exact values. Decimal charting requires an explicit approximate SQL cast; unsafe integers are rejected. HTML escaping covers text and attribute contexts.

WebView2 serves the bundled HTML/JS/CSS from a fixed virtual origin. CSP prohibits inline scripts, eval, connections and frames. Vega uses its expression interpreter. Navigation, downloads, new windows, host objects and permissions are blocked; fixed OpenStreetMap/Esri image tile paths are the only external requests allowed by the host. The renderer libraries and their licences come from integrity-pinned build dependencies.

## Evidence and remaining release work

`DuckHelp` is a pure text function backed by documentation embedded in the host assembly at build time. It needs no workbook session or worker. A tiny separately installed agent skill bootstraps into this help; the detailed workflow is embedded from `docs/AGENT-WORKFLOW.md`, not from the bootstrap itself. `DuckCellStatus` and `DuckExplore` instead use explicit workbook references and the existing host/session policy. The optional Office.js helper bridges an agent's formula-evaluation tool to those functions, normalizing single-cell results and returning pending status when the host has no timers. Normal worksheet calculation does not use this helper. See [agent integration](docs/AGENT-INTEGRATION.md) for the live trial and the distinction between formula, macro and agent-skill capabilities.

The production suite exercises trust, typed DAGs, reuse, nested calls, date systems, workbook isolation, cancellation/retry, save/reopen, cancelled closure and formula movement. Native preview smoke tests verify WebView startup and updates without reselection; browser tests render all eight templates and check injection, exact conversions and stale states. See [tests](tests/README.md).

The rewrite has passed the development-machine installation lifecycle, live OpenAI Excel exploration, and AIS workbook migration and copied-workbook S3 trials. Clean-machine installation, improved skill discovery, Claude execution, more display configurations and broader runtime/provider coexistence are [explicit follow-up trials](docs/RELEASE.md#tested-scope-and-follow-up-work). Signing is deferred. Earlier prototype measurements and the execution decision are preserved under [prototypes/Execution](prototypes/Execution/DECISION.md).
