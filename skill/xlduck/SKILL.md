---
name: xlduck
description: Build and inspect research analyses with XlDuck 0.2 Excel functions such as DuckQuery, DuckFrag, DuckCapture, DuckOut, DuckScalar, DuckInspect and DuckPlot. Use for SQL formula DAGs, typed outputs, supported file/S3 readers and recovery in enabled workbooks.
---

# XlDuck research workflow

XlDuck computes read-and-compute SQL through a worker for each enabled workbook. Excel owns the dependency graph. Most queries should be quick; materialize shared sources and keep expensive preparation outside the sheet.

1. Check `DuckStatus()` and the actual Excel tool capabilities. Enablement and restart are ribbon/host commands, not formulas. If the agent cannot invoke supported macros, hand those actions to the user; never invent COM access or a configuration UDF.
2. Create a source with `DuckQuery` or `DuckCapture`. A capture includes its header row. Bind values and upstream handles with positional `?` arguments (up to eight); do not interpolate values into SQL.
3. Inspect `DuckInspect(handle,"schema")`, `DuckInspect(handle,"rows")` and a deliberate bounded output. Handles are opaque session references, not numeric data. Never invent internal table names or copy a handle into another workbook.
4. Compose analyses with table/fragment references. `DuckQuery` retains a snapshot. `DuckFrag` is a recipe: a changed downstream computation can read its sources again. Materialize a shared trunk when several branches need the same snapshot.
5. Use `DuckOut` for a bounded spill with headers, or `DuckScalar` for exactly one row and column. Exact decimals and large integers stay text. Use explicit SQL casts to choose approximation; decimal charts require `CAST AS DOUBLE`.
6. Use `DuckPlot` for a chart viewed by selecting its cell in Preview Pane. No in-sheet chart placement is required. Samples are labelled; an old chart can remain visibly stale during computation and must not be treated as current data.
7. Wait boundedly while cells return `#N/A`. Inspect actionable `#ERROR:` messages. Unchanged F2/Enter and F9/Shift+F9 reuse successful results; they do not refresh sources. Use Retry for deliberate failure retries and Restart workbook for fresh sources or a failed worker.

Worksheet SQL cannot write/export data, change settings, install extensions or run dynamic SQL. Enabled workbooks have broad read access under the user's account and configured credentials; this is not an OS sandbox. Keep credentials out of cells. Configure S3 through host settings and authenticate outside Excel.

Date cells are serial numbers. Use `DuckDate`/`DuckDateTime` and explicit SQL DATE/TIMESTAMP casts. Preserve numeric-looking text, blanks, empty strings and booleans. Do not silently coerce mixed capture columns. Manual calculation may need several explicit passes; never change global calculation/RTD settings silently.

When using COM, write `.Formula2` and read `.Value2`. Do not attach tests to a user's active workbook. The full signatures, limits, charts, commands and connection schema are in [reference.md](reference.md).
