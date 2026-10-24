# XlDuck formula and agent reference

XlDuck 0.2 runs read-and-compute SQL through a separate worker for each enabled workbook. An Excel coding agent should build a small, inspectable formula DAG, materialize shared sources and use ordinary Excel dependencies for changes. Most computations should finish within a few seconds; substantial preparation belongs outside the sheet.

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

Handles are opaque `xlduck://…` strings, meaningful only inside the producing workbook session. Do not parse dimensions from them, copy them between workbooks or invent internal table names. Reference the producing cell. `DuckInspect(handle)` spills column names and types; `DuckInspect(handle,"rows")` returns the row count; `DuckInspect(handle,"definition")` returns its request as JSON, including SQL and arguments. Inspecting a fragment's schema/rows executes a retained materialization; inspecting its definition does not read its source. Definitions can contain private query values, so do not include them in routine diagnostics.

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
| Reopen a trusted saved workbook | New session and source reads in automatic calculation mode |
| Restart workbook | Cancel and discard that workbook's session, then rebuild |
| Retry failed results | Resubmit failures deliberately; successful retained work stays |

`#N/A` can mean asynchronous work or an unavailable input. Wait for completion before treating it as an SQL failure. `#ERROR: …` is an actionable computation failure; `#BLOCKED: …` requires workbook enablement. Poll boundedly and inspect `DuckStatus()` when diagnosing progress. Do not retry by adding random tokens to SQL. A lost worker requires Restart workbook; simply retrying a query cannot revive it.

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

## Enablement and S3

Enable / Connections gives the workbook broad read access. The local approval is bound to the saved path and saved file fingerprint. Copies and externally modified workbooks require a new decision. There are no per-query or per-bucket permission prompts. SQL remains read-and-compute only; DDL/DML, COPY/export, ATTACH, SET/PRAGMA, extension installation and dynamic SQL are unavailable. This is not a sandbox for hostile workbooks.

Use the connections dialog for S3 routing. Choose environment credentials, a named AWS profile, SSO or anonymous access. Region, optional endpoint and path-style addressing are non-secret settings. Authenticate outside Excel and restart after renewing credentials. `Scope` chooses credentials; it is not an access restriction. Do not put access keys or secret-creation SQL in cells. Live AWS/SSO and enterprise-network compatibility remain release validation items.

## Commands available to automation

Commands are Excel macros, not worksheet functions. A host with Excel COM/VBA access can invoke them with `Application.Run`; always supply the workbook name, or an explicit empty string for the active workbook:

```vb
Application.Run "DuckEnableWorkbook", "Research.xlsx"
Application.Run "DuckRestartWorkbook", "Research.xlsx"
Application.Run "DuckRetryFailed", "Research.xlsx"
Application.Run "DuckTogglePreview"
```

The same workbook-name argument applies to `DuckDisableWorkbook`, `DuckPauseQueries`, `DuckResumeQueries` and `DuckInterrupt`. `DuckSetConnections(settingsJson,workbookName)` requires an already enabled workbook and restarts it. JSON uses property names `S3`, `Scope`, `Credentials`, `Region`, `Endpoint`, `Profile`, `UseSsl`, `PathStyle`; `Credentials` is currently numeric: 0 anonymous, 1 environment, 2 named profile, 3 SSO. Example:

```json
{"S3":[{"Scope":"s3://research/","Credentials":2,"Profile":"research","Region":"eu-west-2"}]}
```

Check the actual Excel agent's tool schema before claiming it can run a command. Formula read/write access alone does not imply COM/VBA or macro execution. If command execution is unavailable, use the corresponding ribbon action as the handoff; never imitate a command with a side-effecting formula. No connected Excel agent session was available during implementation to validate that additional integration. Worksheet inspection and the COM commands are covered by the isolated Excel tests.

When writing formulas through COM, use `.Formula2`, not `.Formula`, to preserve dynamic arrays. Read `.Value2` for assertions rather than formatted display text. Never attach integration tests to a user's active workbook or change unrelated workbooks.
