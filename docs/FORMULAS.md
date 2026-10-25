# XlDuck formula and agent reference

XlDuck 0.2 runs read-and-compute SQL through a separate worker for each ordinary workbook. An Excel coding agent should build a small, inspectable formula DAG, materialize shared sources and use ordinary Excel dependencies for changes. Most computations should finish within a few seconds; substantial preparation belongs outside the sheet.

## Build and inspect an analysis

```excel
A1: =DuckQuery("SELECT category, value FROM read_parquet(?)", "C:\research\observations.parquet")
A2: =DuckQuery("SELECT category, avg(value)::DOUBLE mean_value FROM ? GROUP BY category", A1)
A3: =DuckPlot(A2, "bar", "x", "category", "y", "mean_value")
D1: =DuckOut(A2)
G1: =DuckInspect(A1, "schema")
J1: =DuckInspect(A1, "rows")
```

Discover the actual schema before choosing columns: `=DuckSchema("SELECT * FROM read_parquet(?)", sourceCell)` spills column names and types without retaining data rows. It can read source metadata; CSV/JSON schema inference can scan/sample data. Inspect a few real examples next, then retain the columns the investigation needs. Use the real path or source supplied by the user. Local, HTTP and S3 CSV/JSON/Parquet readers are supported. Use explicit reader functions, rather than SQL replacement scans such as `FROM 'file.parquet'`. `DuckCapture(A1:C20)` captures the header row and the values below it; its result is another table handle.

Handles are short tokens such as `xlduck://t/42::Ab3dE6gH9jk`, meaningful only inside the producing workbook session. `t`, `f` and `p` identify tables, fragments and plots. The number identifies a result; the shared suffix distinguishes its session. Numbers are never reused within a session. Treat the complete token as opaque: do not extract dimensions, construct tokens, copy them between workbooks or invent internal table names. Reference the producing cell. `DuckInspect(handle)` spills column names and types; `DuckInspect(handle,"rows")` returns the row count; `DuckInspect(handle,"definition")` returns its request as JSON, including SQL and arguments. A fragment's schema binds its query with zero result rows; its row count requires materialization. Inspecting its definition does not read its source. Definitions can contain private query values, so do not include them in routine diagnostics.

## SQL parameters and composition

`DuckQuery`, `DuckFrag`, `DuckQueryOut`, `DuckQueryOutScalar` and `DuckSchema` accept SQL followed by up to eight positional arguments. Each unquoted `?` binds to the next argument. Question marks inside SQL strings, quoted identifiers and comments remain literal. Values are bound parameters; table handles become table references; fragments become subqueries. Use aliases explicitly:

```excel
=DuckQuery("SELECT l.id, r.value FROM ? l JOIN ? r ON l.id=r.id WHERE r.value > ?", A1, B1, C1)
=DuckQueryOutScalar("SELECT CAST(? AS DATE)", DuckDate(D1))
```

Do not concatenate user values or credentials into SQL. Column/table identifiers are not scalar parameters: choose identifiers in the SQL and quote them with SQL double quotes (doubled again inside an Excel string). An empty argument does not remove later arguments. Numeric-looking text remains text, including leading zeroes and apostrophes. Plot handles are not SQL table arguments; refer to the plot's source table instead.

Fragments allow composition without an intermediate snapshot:

```excel
A1: =DuckFrag("SELECT range AS n FROM range(100)")
A2: =DuckFrag("SELECT n FROM ? WHERE n >= ?", A1, 10)
A3: =DuckQuery("SELECT sum(n) AS total FROM ?", A2)
A4: =DuckScalar(A3)
```

`DuckFrag` validates the supported capability surface without reading external data. A downstream query can read a fragment's source again when that downstream request changes. Use `DuckQuery` at a shared trunk when several analyses need the same retained snapshot. A fragment consumed by output or a plot gets a retained materialization while that consumer remains live.

## Supported SQL at a glance

| Need | Supported route |
| --- | --- |
| Filtering, joins, aggregates, windows, CTEs, recursive CTEs | One SELECT query, including WITH |
| Lists/structs and flattening | `SELECT unnest([1,2,3]) AS n`, recursive `unnest(..., recursive := true)`, or FROM-form UNNEST |
| Pivot | Explicit pivot values: `PIVOT (...) ON category IN ('a','b') USING sum(amount)` |
| Source schema or binding check | `DuckSchema(sql, arg1, …, arg8)`; alternatively a SELECT wrapped with `LIMIT 0` |
| Retained table schema | `DuckInspect(handle)` |
| Syntax/capability validation without source access | `DuckFrag` or exploration `validate`; these do **not** check source availability, columns or runtime types |
| DESCRIBE/SHOW | Not materializable worksheet SELECTs; use DuckSchema/DuckInspect |
| Writes, exports, settings, extension installation, dynamic SQL | Unavailable |

The pinned engine is DuckDB 1.5.6. Its reviewed built-ins and explicit reader allowlist are the capability contract; arbitrary extensions cannot add functions. Unsupported SQL fails before execution. `DuckSchema` also binds column names and types, and therefore can fail on inaccessible sources or nonexistent columns. It does not guarantee that every row will convert successfully during execution.

## Status and errors

`=DuckCellStatus(A1)` accepts a **reference** to the producing cell, even while A1 is `#N/A` and has no handle. It returns JSON without submitting or refreshing the query. With an API supporting formula evaluation, evaluate that same formula without placing it in a cell. Poll again to obtain current state; a saved status-cell value is just its most recent evaluation.

| `state` | Meaning / next action |
| --- | --- |
| `initializing`, `queued`, `running`, `reading` | Wait boundedly, then inspect again |
| `waiting_inputs` | An upstream value is unavailable; inspect its producing cell |
| `paused` | Resume through the workbook command/ribbon |
| `awaiting_calculation` | Result is ready, but Excel has not consumed it; another calculation pass may be needed in manual mode |
| `succeeded` | The last evaluated inputs produced a successful value |
| `failed`, `cancelled` | Inspect `failure.code`, `message` and `advice` |
| `not_evaluated` | No tracked XlDuck invocation for this formula; calculate or check the reference |

Status includes `session`, `evaluationId`, `resultReady`, `excelState`, `calculationMode`, and `computations`. Each computation has its handle, kind, direct dependency handles, timestamps, elapsed execution milliseconds and result row count where known. `reused` means its lease joined an already existing computation when acquired. Stable handles and timestamps across calculations demonstrate reuse. Dependencies describe the request at execution time, even when its input leases have since been released; they are not a source-freshness or cost estimate. A fragment's elapsed time covers validation only.

Status explicitly reports `scope: "last_evaluated_inputs"`. It cannot tell whether uncalculated edits in manual mode should have produced a different result. Do not treat it as a promise of external freshness. A changed `session` invalidates previous handles and polling assumptions. `evaluationId` identifies the currently tracked formula/bindings, not a workbook-wide calculation barrier.

`#N/A` is the asynchronous marker. `#VALUE!` is **not** a normal running state. A completed query with a blocked Excel spill reports `failure.code: "excel_spill"` and `resultReady: true`; changing the SQL is usually the wrong fix. Error categories include `invalid_input`, `input_error`, `sql_validation`, `sql_binding`, `data_conversion`, `source_access`, `memory_limit`, `spill_limit`, `output_shape`, `output_limit`, `worker_unavailable`, `timeout`, `cancelled`, `excel_spill`, `excel_error`, and a fallback `execution`. Pending work is represented by state, not an error code. Worksheet errors retain the readable `#ERROR: …` form; use JSON status rather than parsing their wording.

## Reuse and recovery

| Event | Result |
| --- | --- |
| Unchanged SQL, typed inputs and handles | Reuse the retained result |
| F9, Shift+F9, full calculation, unchanged F2/Enter | Excel's usual calculation scope; no implicit source refresh |
| Changed SQL, scalar/capture contents or upstream handle | New computation |
| Changed external file, clock or random seed | No automatic invalidation |
| Reopen an ordinary saved workbook | New session and source reads in automatic calculation mode |
| Restart workbook | Cancel and discard that workbook's session, then rebuild |
| Retry failed results | Resubmit failures deliberately; successful retained work stays |

`#N/A` can mean initialization, asynchronous work or an unavailable input. Wait for completion before treating it as an SQL failure. `#ERROR: …` is an actionable computation failure. Protected View prevents query execution; ordinary workbooks initialize automatically. Use `DuckCellStatus` for a particular query; `DuckStatus()` reports workbook `ready`, connection count, `connectionsSaved`, `session`, `paused`, `runtime`, `resources`, aggregate statistics and a status message. Do not retry by adding random tokens to SQL. A lost worker requires Restart workbook; simply retrying a query cannot revive it.

Automatic calculation is primary. Manual mode may require repeated explicit calculation passes as asynchronous dependencies complete. Do not silently change Excel's application-wide calculation mode, RTD throttle or the meaning of F9/Shift+F9.

Pause holds new computations; an already running computation may finish. Resume starts the held requests without refreshing successful results. Cancel interrupts outstanding work and output reads; successful tables remain available. Retry or a changed formula starts a cancelled computation again.

## Memory-conscious exploration

Start with schema and a small metadata projection. For a message archive, retain IDs, dates, senders and subjects first; aggregate/filter those, then fetch `left(body, 500)` for a handful of chosen IDs from the original source. Verify the actual column names before writing that SQL. Avoid retaining every body or aggregating huge lists simply to inspect a sample. LIMIT bounds result rows, not all scanning, sorting, joining or aggregation work; even an excerpt query can read large source blocks. Expensive preparation belongs outside Excel.

**XlDuck > Resources** controls each workbook's worker: defaults are **1024 MiB memory, 4 GiB temporary disk and 2 query threads**. Allowed ranges are 128–65536 MiB, 1–1024 GiB and 1–16 threads. These are explicit budgets, not recommendations to maximize them. Several workers can use their budgets simultaneously. DuckDB's memory setting does not cap total process memory, and not every operation spills. The dialog saves settings locally by workbook path; Apply restarts only that workbook. Worksheet SQL cannot change budgets. Memory/spill failures suggest query changes and these host controls.

## Values and output bounds

`DuckOut(handle,[maxRows])` and `DuckQueryOut` spill column headers followed by rows. The default is 10,000 data rows; `DuckOut` accepts 1 through 1,000,000. Output also has a one-million-cell limit, an eight-million-character limit and a 32 MiB transport frame limit. Oversized output fails instead of silently truncating. Query a deliberate `LIMIT`/aggregation when a sample is wanted. `DuckScalar` and `DuckQueryOutScalar` require exactly one row and one column; they never quietly choose the first cell.

| Value | Behaviour |
| --- | --- |
| Number, text, Boolean | Type preserved |
| Empty worksheet cell | SQL NULL; distinct from empty text |
| Excel errors in capture/arguments | Rejected or propagated; never silently blanked |
| Exact DECIMAL | Exact text on output; explicit `CAST AS DOUBLE` chooses approximation |
| Integer outside ±9,007,199,254,740,991 | Exact text; never rounded silently |
| Nested data/JSON | JSON text, preserving nested nulls and numbers |
| Captured mixed text and numeric column | Error requiring an explicit source/type choice |
| DATE/TIMESTAMP output | Excel serial where representable, using the workbook's 1900/1904 date system; otherwise text |
| TIME | Fraction of a day on sheet output |
| Time zone timestamps and INTERVAL | Invariant text on sheet output |

Worksheet date formatting is not a SQL type. Use `DuckDate`/`DuckDateTime` plus `CAST(? AS DATE/TIMESTAMP)` for serial date arguments. Helpers return ISO text without SQL quotes and reject Excel's fictional 1900-02-29. Format date output cells explicitly if needed.

Captures support up to one million cells. Headers are made deterministically unique without stealing explicitly supplied names. Numeric/text types participate in request identity, so a text `"1"` never collides with numeric `1`.

## Charts and preview

Select a plot handle and open Preview Pane. Tables show up to 200 rows; charts up to 5,000, labelled when sampled. Aggregate or explicitly sample in SQL for a deliberate display. A pending/failed chart may show its previous successful state with a stale label; that state is never an input to new computations.

`DuckPlot(source,template,name1,value1,…,name6,value6)` supports:

| Template | Required options |
| --- | --- |
| `bar`, `line`, `point`, `area`, `boxplot` | `x`, `y` column names |
| `histogram` | `x` |
| `heatmap` | `x`, `y`, `value` |
| `map` | `lat`, `lon` |

Options also include `color`, `label`, `tooltip`, `title`, `xmin`, `xmax`, `ymin`, `ymax`, and map `basemap` (`osm` or `satellite`). Axis-domain overrides must be numeric. Lines follow data order: use SQL `ORDER BY` when ordering matters. Charts require an explicit `CAST AS DOUBLE` for exact decimals. Integers beyond JavaScript's exact range require an explicit VARCHAR label or approximate DOUBLE cast. Calendar intervals cannot become fixed seconds automatically; convert them in SQL. Prefer simple, unique column aliases for chart fields.

Map coordinates must be numeric, finite, non-null and within latitude -90 to 90 and longitude -180 to 180. Cast text coordinates and filter missing/invalid locations explicitly in SQL. A map rejects invalid coordinates instead of quietly omitting or coercing points.

The libraries are bundled. Maps still request tiles from OpenStreetMap/Esri, exposing the requested viewport to the provider. Missing WebView2 produces a repair message, while worksheet computations remain available.

## Workbook access and S3-compatible storage

Ordinary workbooks start automatically. Protected View workbooks cannot start queries; use Excel's Enable Editing only for workbooks you trust. A workbook that bypasses Protected View can use broad local-file and network reads under your account and configured credentials. XlDuck does not add a second approval dialog or change Trust Center. SQL remains read-and-compute only; DDL/DML, COPY/export, ATTACH, SET/PRAGMA, extension installation and dynamic SQL are unavailable. This is not an OS sandbox for hostile workbooks.

Use **Connections** to add a bucket/path such as `s3://research/`, an endpoint such as `storage.example.com:9000`, and credentials. The endpoint has no scheme or path. Saved access keys go to Windows Credential Manager; environment and anonymous access are also supported. Environment authentication uses `AWS_ACCESS_KEY_ID`, `AWS_SECRET_ACCESS_KEY` and optional `AWS_SESSION_TOKEN`, including for non-AWS servers. Advanced settings contain HTTPS (on by default), path-style URLs and an optional session token. Named profiles and SSO remain optional provider choices.

The S3 signing protocol includes a region string even for compatible services hosted elsewhere. XlDuck supplies `us-east-1` internally, so the dialog does not ask for it. This does not route traffic to AWS: the configured endpoint determines the server. Providers requiring a different value can use the optional `Region` field in `DuckSetConnections`; editing the connection in the dialog preserves that override.

Connection settings are saved on this computer by workbook path, not workbook fingerprint. External edits preserve settings; Save As carries them to the new path; copies at a different path use defaults. Unsaved workbook settings persist when first saved. Applying settings restarts this workbook. Secrets stay outside cells and settings files. Removing a connection leaves its saved key available to other workbooks; unused `XlDuck/S3/…` entries can be removed in Windows Credential Manager. `Scope` selects credentials; it is not an access restriction. Live Hetzdev S3 reads and recovery from a simulated S3 outage are tested; broader provider/proxy compatibility remains a release validation item.

## Commands available to automation

### Read-only agent inspection prototype

`DuckExplore(anchorCell, requestJson)` starts or polls a bounded inspection using the anchor's existing workbook session, worker, connections and SQL policy. Use a blank single-cell anchor in the intended workbook. It does not write cells, create worksheet formulas or change settings. With desktop Office.js, [`Worksheet.evaluate`](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheet#excel-excel-worksheet-evaluate-member(1)) (ExcelApiDesktop 1.1) provides a possible bridge from `run_officejs`. The actual tool schema must permit Office.js execution; formula access alone is insufficient. Native Excel Evaluate is integration-tested. The connected agent's Office.js roundtrip remains to be verified in a later trial.

```js
const sheet = ctx.workbook.worksheets.getItem("Research");
const request = {op: "query", id: "inspect-001", sql: "SELECT unnest([1,2,3]) AS n", maxRows: 2};
const escaped = JSON.stringify(request).replaceAll('"', '""');
const response = sheet.evaluate(`DuckExplore(AZ1000,"${escaped}")`);
await ctx.sync();
return JSON.parse(response.value);
```

Use a unique `id` per exploration, then poll `{"op":"poll","id":"inspect-001"}`. Reusing the same id and inputs is idempotent; different inputs with that id fail. Do not start a new query on each poll. Terminal states are `succeeded`, `failed`, `cancelled`. Polls return the session ID: do not silently restart an investigation if that changes. Initialization can require retrying the original start request because no operation has been admitted yet.

| `op` | Additional fields and result |
| --- | --- |
| `schema` | `sql`, optional `arguments`; zero-row schema/binding check |
| `query` | `sql`, optional `arguments`, `maxRows` (default 20, maximum 200); bounded typed rows and `hasMore` |
| `validate` | `sql`, optional `arguments`; syntax/capabilities only, explicitly labelled in `result.validation` |
| `inspect` | `handle`; retained metadata and definition; a fragment returns its recipe without materialization |
| `poll` | `id`; reads the existing operation's state |

Arguments may be JSON scalars or exact `{kind,text}` values, e.g. `{"kind":"Integer","text":"9007199254740993"}`. Handle strings resolve only in the anchor's session. Returned cells are typed `{kind,text}` objects; do not coerce large integers or decimals into JavaScript numbers. `hasMore` reports truncation of the bounded query result, not the source's full row count. Result JSON is capped at 30000 characters, output at 4000 cells, SQL at 24000 characters, arguments at eight. Eight operations are tracked per workbook; completed operations can be evicted to admit new ones and expire after two minutes without polling. Each operation has a 30-second deadline including queue time. Timeout releases its own lease and cancels unshared work, never another formula's computation. An expired/unknown poll fails; restarting exploration requires a deliberate new id.

The helper [agent-exploration.js](agent-exploration.js) performs short bounded waits and returns pending status when its wait budget expires. It never edits the workbook or restarts anything. Native Evaluate has a **255-character expression limit**, so the helper stages longer JSON requests in bounded parts before submitting them. Each expression stays within that limit, including JSON/Excel escaping. No SQL executes during staging. Transfer is in memory, scoped to the workbook session, with eight uploads, at most 2048 parts/30000 characters each, and two-minute inactivity expiry. `stage` uses `id`, `index`, `count`, `text`; `submit` uses `id`. Duplicate parts must match; incomplete or altered transfers fail. A repeated submit polls the same operation. Prefer the helper over constructing this transport manually. Transfer precedes the query's execution deadline and can take longer than the helper's polling budget.

Use exploration for investigation, then leave reproducible worksheet formulas as the final analysis. The usual unchanged-input reuse rules still apply; a new exploration id is not a source refresh.

### Workbook commands

Commands are Excel macros, not worksheet functions. A host with Excel COM/VBA access can invoke them with `Application.Run`; always supply the workbook name, or an explicit empty string for the active workbook:

```vb
Application.Run "DuckRestartWorkbook", "Research.xlsx"
Application.Run "DuckRetryFailed", "Research.xlsx"
Application.Run "DuckTogglePreview"
```

The same workbook-name argument applies to `DuckPauseQueries`, `DuckResumeQueries` and `DuckInterrupt`. `DuckSetConnections(settingsJson,workbookName)` applies non-secret settings and restarts the named ordinary workbook. JSON uses property names `S3`, `Scope`, `Credentials`, `Region`, `Endpoint`, `Profile`, `UseSsl`, `PathStyle`, `CredentialId`; `Credentials` is currently numeric: 0 anonymous, 1 environment, 2 named profile, 3 SSO, 4 a saved access key. `CredentialId` references a key already saved through the dialog, never the secret itself. Example using environment credentials:

```json
{"S3":[{"Scope":"s3://research/","Credentials":1,"Endpoint":"storage.example.com:9000","PathStyle":true,"UseSsl":true}]}
```

`DuckSetResources(resourcesJson,workbookName)` accepts `{"MemoryMiB":2048,"SpillGiB":8,"Threads":2}` and preserves the workbook's connections. It restarts only that workbook. `DuckSetConnections` replaces the supplied workbook settings; its optional `Resources` object uses the same fields and defaults if omitted.

Check the actual Excel agent's tool schema before claiming it can run a command. Formula read/write access alone does not imply COM/VBA or macro execution. If command execution is unavailable, use the corresponding ribbon action as the handoff; never imitate a command with a side-effecting formula. The observed connected agent offers `run_officejs`, which does not itself establish macro-command access. Worksheet inspection, Evaluate and the COM commands are covered by isolated Excel tests; no commands were executed in the user's live research workbook during this implementation.

When writing formulas through COM, use `.Formula2`, not `.Formula`, to preserve dynamic arrays. Read `.Value2` for assertions rather than formatted display text. Never attach integration tests to a user's active workbook or change unrelated workbooks.
