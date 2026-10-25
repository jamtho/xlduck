# XlDuck formula and agent reference

XlDuck 0.2 runs read-and-compute SQL through a separate worker for each ordinary workbook. An Excel coding agent should build a small, inspectable formula DAG, materialize shared sources and use ordinary Excel dependencies for changes. Most computations should finish within a few seconds; substantial preparation belongs outside the sheet.

## Build and inspect an analysis

```excel
A1: =DuckQuery("SELECT * FROM read_parquet(?)", "C:\research\observations.parquet")
A2: =DuckQuery("SELECT category, avg(value)::DOUBLE mean_value FROM ? GROUP BY category", A1)
A3: =DuckPlot(A2, "bar", "x", "category", "y", "mean_value")
D1: =DuckOut(A2)
G1: =DuckInspect(A1, "schema")
J1: =DuckInspect(A1, "rows")
```

Use the real path or source supplied by the user. Local, HTTP and S3 CSV/JSON/Parquet readers are supported. Use explicit reader functions, rather than SQL replacement scans such as `FROM 'file.parquet'`. `DuckCapture(A1:C20)` captures the header row and the values below it; its result is another table handle.

Handles are short tokens such as `T42~Ab3dE6gH9jk`, meaningful only inside the producing workbook session. `T`, `F` and `P` identify tables, fragments and plots. The number identifies a result; the shared suffix distinguishes its session. Numbers are never reused within a session. Treat the complete token as opaque: do not extract dimensions, construct tokens, copy them between workbooks or invent internal table names. Reference the producing cell. `DuckInspect(handle)` spills column names and types; `DuckInspect(handle,"rows")` returns the row count; `DuckInspect(handle,"definition")` returns its request as JSON, including SQL and arguments. Inspecting a fragment's schema/rows executes a retained materialization; inspecting its definition does not read its source. Definitions can contain private query values, so do not include them in routine diagnostics.

## SQL parameters and composition

`DuckQuery`, `DuckFrag`, `DuckQueryOut` and `DuckQueryOutScalar` accept SQL followed by up to eight positional arguments. Each unquoted `?` binds to the next argument. Question marks inside SQL strings, quoted identifiers and comments remain literal. Values are bound parameters; table handles become table references; fragments become subqueries. Use aliases explicitly:

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

`#N/A` can mean initialization, asynchronous work or an unavailable input. Wait for completion before treating it as an SQL failure. `#ERROR: …` is an actionable computation failure. Protected View prevents query execution; ordinary workbooks initialize automatically. Poll boundedly and inspect `DuckStatus()` when diagnosing progress. It reports `ready`, connection count, `connectionsSaved`, `session`, `paused`, `runtime`, computation statistics and a status message. Do not retry by adding random tokens to SQL. A lost worker requires Restart workbook; simply retrying a query cannot revive it.

Automatic calculation is primary. Manual mode may require repeated explicit calculation passes as asynchronous dependencies complete. Do not silently change Excel's application-wide calculation mode, RTD throttle or the meaning of F9/Shift+F9.

Pause holds new computations; an already running computation may finish. Resume starts the held requests without refreshing successful results. Cancel interrupts outstanding work and output reads; successful tables remain available. Retry or a changed formula starts a cancelled computation again.

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

Check the actual Excel agent's tool schema before claiming it can run a command. Formula read/write access alone does not imply COM/VBA or macro execution. If command execution is unavailable, use the corresponding ribbon action as the handoff; never imitate a command with a side-effecting formula. No connected Excel agent session was available during implementation to validate that additional integration. Worksheet inspection and the COM commands are covered by the isolated Excel tests.

When writing formulas through COM, use `.Formula2`, not `.Formula`, to preserve dynamic arrays. Read `.Value2` for assertions rather than formatted display text. Never attach integration tests to a user's active workbook or change unrelated workbooks.
