---
name: xlduck
description: Build and inspect research analyses with XlDuck 0.2, including schema-first exploration, DuckCellStatus, bounded DuckExplore queries, typed SQL formula DAGs, supported readers and recovery in ordinary workbooks.
---

# XlDuck research workflow

XlDuck computes read-and-compute SQL through a worker for each ordinary workbook. Excel owns the dependency graph. Most queries should be quick; materialize shared sources and keep expensive preparation outside the sheet.

1. Check `DuckStatus()` and the actual Excel tool capabilities. Ordinary workbooks start automatically; Protected View must be left through Excel by the user. Connections and restart are ribbon/host actions, not formulas. If the agent cannot invoke supported macros, hand those actions to the user; never invent COM access or a configuration UDF.
2. Inspect the actual schema first with `DuckSchema(sql,args...)`, then a small metadata sample. If the tool permits Office.js and desktop `Worksheet.evaluate`, the `DuckExplore` prototype can inspect without writing formulas; see [reference.md](reference.md) and [agent-exploration.js](agent-exploration.js). Native Evaluate is tested; the actual agent roundtrip still needs validation. Never claim macro access from Office.js alone.
3. Retain selected columns with `DuckQuery`, or use `DuckCapture` with a header row. Bind values and handles with up to eight positional `?` arguments; do not interpolate values. Keep SQL in ordinary or named cells when that makes the analysis easier to inspect. Use `DuckInspect` for retained schema, rows or definitions. Handles such as `xlduck://t/42::Ab3dE6gH9jk` are opaque; reference producing cells and never construct or copy tokens between workbooks.
4. Compose analyses with table/fragment references. `DuckQuery` retains a snapshot. `DuckFrag` is a recipe: a changed downstream computation can read its sources again. Materialize a shared trunk when several branches need the same snapshot.
5. Use `DuckOut` for a bounded spill with headers, or `DuckScalar` for exactly one row and column. Exact decimals and large integers stay text. Use explicit SQL casts to choose approximation; decimal charts require `CAST AS DOUBLE`.
6. Use `DuckPlot` for a chart viewed by selecting its cell in Preview Pane. No in-sheet chart placement is required. Samples are labelled; an old chart can remain visibly stale during computation and must not be treated as current data.
7. Poll `DuckCellStatus(producingCell)` with a bounded wait. Its JSON distinguishes running, failure, cancellation, waiting inputs and awaiting manual calculation; `DuckStatus()` is only aggregate workbook status. `#N/A` can be pending; `#VALUE!` is not a running marker. Use `failure.code` to distinguish SQL, source, memory and Excel spill failures. Status describes last evaluated inputs, not uncalculated manual-mode edits or source freshness. Unchanged F2/Enter and F9/Shift+F9 reuse successful results. Use Retry deliberately and Restart for fresh sources or a failed worker.

For large archives, retain IDs/dates/senders/subjects, then fetch a few short body excerpts by ID. LIMIT alone does not bound scan/sort/join memory. Resources controls per-workbook memory, spill and thread budgets; defaults are 1 GiB, 4 GiB and two threads. SQL cannot change them. DESCRIBE/SHOW are unavailable; use DuckSchema. SELECT-list UNNEST, recursive UNNEST and FROM UNNEST are supported. DuckFrag/exploration validation checks syntax and capabilities, not source columns; DuckSchema binds columns without retaining data rows. Keep final analyses as reproducible worksheet formulas, even when exploration used the direct prototype.

Worksheet SQL cannot write/export data, change settings, install extensions or run dynamic SQL. Ordinary workbooks have broad read access under the user's account and configured credentials; this is not an OS sandbox. Keep credentials out of cells. Configure S3-compatible endpoints and credentials through Connections; saved keys live in Windows Credential Manager.

Date cells are serial numbers. Use `DuckDate`/`DuckDateTime` and explicit SQL DATE/TIMESTAMP casts. Preserve numeric-looking text, blanks, empty strings and booleans. Do not silently coerce mixed capture columns. Manual calculation may need several explicit passes; never change global calculation/RTD settings silently.

When using COM, write `.Formula2` and read `.Value2`. Do not attach tests to a user's active workbook. The full signatures, limits, charts, commands and connection schema are in [reference.md](reference.md).
