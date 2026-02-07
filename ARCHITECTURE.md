# Architecture

## Overview

XLDuck is an Excel add-in that exposes DuckDB's SQL engine to spreadsheet users. The core idea is to enable **dataflow-style computation** where intermediate query results can be stored as handles and referenced by downstream queries, creating a DAG of computations across the sheet.

## Core Concepts

### Handles

A handle is a string reference to stored data or deferred SQL, formatted as:
```
duck://table/1234|10x4    (table handle - materialized data with dimensions)
duck://frag/1234         (fragment handle - deferred SQL)
```

Where:
- `duck://` - protocol prefix
- `table` or `frag` - type identifier
- `1234` - auto-generated numeric ID
- `|10x4` - (table handles only) row x column dimensions

Handles are displayed in cells and can be passed to other functions as table references.

### Status URLs

Special cell values use a `#duck://` URL format:
```
#duck://blocked/config|Waiting for DuckConfigReady()   (waiting for config)
#duck://error/syntax|SQL syntax error near...          (syntax error)
#duck://error/notfound|Table 'foo' does not exist      (not found error)
#duck://error/http|HTTP 403 on S3 bucket               (HTTP error)
#duck://error/query|...                                (general query error)
#duck://error/internal|...                             (internal error)
```

The `#` prefix follows Excel's convention for special status values.

### Result Storage

Query results from `DuckQuery` are stored as DuckDB temp tables. The .NET layer keeps metadata (table name, column names, row count) for each handle:

```csharp
class StoredResult {
    string DuckTableName;   // e.g. "_xlduck_res_abc123..."
    string[] ColumnNames;
    long RowCount;
}
```

Metadata is kept in a `Dictionary<string, StoredResult>` keyed by handle. The actual data stays in DuckDB, avoiding memory copies between .NET and DuckDB.

### Fragment Storage

SQL fragments from `DuckFrag` are stored as deferred SQL text, not executed results:

```csharp
class StoredFragment {
    string Sql;
    object[] Args;  // Bound parameters for recursive resolution
}
```

Fragments enable lazy evaluation - the SQL is validated (via EXPLAIN) at creation time but not executed until used.

### Query Parameter Binding

When a query references a stored result, users specify `?` placeholders with positional arguments:

```excel
=DuckQuery("SELECT * FROM ? WHERE region = ?", A1, "EU")
```

Where A1 contains a handle like `duck://table/1234`.

Arguments are passed positionally after the SQL string, replacing `?` placeholders left-to-right.

## Data Flow

### Query Execution

```
DuckQuery("SELECT ...")
    → CREATE TEMP TABLE _xlduck_res_xxx AS SELECT ...
    → Get schema via PRAGMA table_info
    → Get row count via SELECT COUNT(*)
    → Store metadata (table name, columns, count)
    → Return handle to cell
```

### Query with References

```
DuckQuery("SELECT * FROM ?", "duck://table/1")
    → Resolve positional ? arguments left-to-right
    → For each argument:
        → If table handle: substitute DuckDB table name directly
        → If fragment handle: recursively resolve and inline as subquery
        → If string: escape and quote as SQL literal
    → Increment refcount on referenced tables (prevents drop during query)
    → CREATE TEMP TABLE _xlduck_res_xxx AS [resolved SQL]
    → Decrement refcounts (may trigger table drops if count reaches zero)
    → Store metadata, return new handle
```

### Fragment Creation

```
DuckFrag("SELECT * FROM ? WHERE x > 5", A1)
    → Resolve positional arguments (for validation only)
    → Run EXPLAIN to validate SQL
    → Decrement refcounts on any referenced tables
    → Store original SQL + args
    → Return fragment handle
```

### Fragment Resolution

When a fragment is used as a parameter, it's resolved recursively:

```
DuckQuery("SELECT * FROM ?", "duck://frag/1")
    → Look up fragment f/1
    → Resolve fragment's own arguments recursively
    → Inline resolved SQL as: (SELECT ...)
    → Continue with outer query resolution
```

Circular references (fragment A → B → A) are detected and raise an error.

### Materialization

```
DuckOut("duck://table/1")
    → Look up handle metadata
    → SELECT * FROM temp_table LIMIT 200001
    → Convert to Excel array with headers
    → Add truncation footer if >200K rows
    → Return as spilled array
```

**Output Limit**: DuckOut caps output at 200,000 rows to prevent Excel from becoming unresponsive. A footer row indicates when truncation occurs.

### Why Temp Tables?

Query results are stored as DuckDB temp tables rather than in .NET memory. This approach:

1. **Avoids memory copies**: Data stays in DuckDB; no copying to .NET and back
2. **Supports large datasets**: Can handle millions of rows efficiently
3. **Enables zero-copy references**: When a query references a table handle, it uses the existing temp table directly

The trade-off is that all intermediate results consume DuckDB memory until their handles are no longer referenced. Reference counting ensures tables are dropped when no longer needed.

## Excel Functions

| Function | Purpose |
|----------|---------|
| `DuckQuery(sql, [arg1, arg2, ...])` | Execute SQL, return table handle. Up to 8 positional `?` arguments. |
| `DuckQueryAfterConfig(sql, [arg1, arg2, ...])` | Same as DuckQuery, but waits for `DuckConfigReady()` before executing. |
| `DuckFrag(sql, [arg1, arg2, ...])` | Create SQL fragment for lazy evaluation. Validated but not executed. |
| `DuckFragAfterConfig(sql, [arg1, arg2, ...])` | Same as DuckFrag, but waits for `DuckConfigReady()` before executing. |
| `DuckOut(handle)` | Output handle (table or frag) as spilled array with headers. |
| `DuckQueryOut(sql, [arg1, arg2, ...])` | Execute SQL and output directly as spilled array. Combo of DuckQuery + DuckOut. |
| `DuckExecute(sql)` | Execute DDL/DML (CREATE, INSERT, etc.) |
| `DuckConfigReady()` | Signal that configuration is complete. `AfterConfig` functions wait for this. |
| `DuckVersion()` | Return add-in version (0.1) |
| `DuckLibraryVersion()` | Return DuckDB library version |

**When to use which:**
- `DuckQuery` - Materialize and cache results (good for expensive queries used multiple times)
- `DuckQueryAfterConfig` - Same as DuckQuery, for queries that need runtime config (S3 endpoints, etc.)
- `DuckFrag` - Defer execution, allow query optimization across composed fragments
- `DuckFragAfterConfig` - Same as DuckFrag, for fragments that need runtime config
- `DuckOut` - Display results from either handle type
- `DuckQueryOut` - One-off queries where you just want the output

## Known Issues and Workarounds

### HUGEINT/BigInteger Conversion

DuckDB's aggregate functions (SUM, etc.) return HUGEINT/INT128 types that .NET and Excel don't handle natively. The add-in automatically converts these to `double` for Excel compatibility. This may lose precision for very large integers.

### Parameter Limit

Excel-DNA doesn't support `params` arrays in UDFs. Instead, we use explicit optional parameters, limiting queries to 8 positional arguments. This covers most use cases; complex joins needing more can use subqueries or intermediate handles.

## RTD and Lifecycle Management

### RTD-Based Functions

`DuckQuery` and `DuckFrag` use Excel's RTD (Real-Time Data) mechanism for lifecycle tracking. This enables:

1. **Reference counting**: Handles are automatically cleaned up when no longer referenced by any cell
2. **Cell lifecycle awareness**: When a cell is deleted or its formula changes, the handle's reference count decrements
3. **Automatic cleanup**: Handles with zero references are evicted; their DuckDB temp tables are dropped

### Timeout Budget

To avoid RTD's 2-second throttle delay, queries use a timeout budget:

- Queries completing within **1 second** return results directly (synchronous)
- Slower queries return "Loading..." immediately, then update asynchronously

This provides responsive UX for fast queries while supporting long-running operations.

### Configuration Gate (AfterConfig functions)

Queries needing runtime configuration (e.g., S3 endpoints) use the `AfterConfig` variants which wait for setup:

```excel
=DuckFragAfterConfig("SELECT * FROM read_parquet(?)", A1)
```

`DuckQueryAfterConfig` and `DuckFragAfterConfig` wait until `DuckConfigReady()` is called, typically from VBA `Auto_Open`:

```vba
Sub Auto_Open()
    Application.Run "DuckExecute", "SET s3_endpoint = '127.0.0.1:9000'"
    Application.Run "DuckConfigReady"
End Sub
```

Downstream queries that depend on a blocked query (input starts with `#duck://blocked/`) also wait automatically.

## Session Lifecycle

- Results (DuckDB temp tables) persist for the Excel session
- Closing Excel clears all handles and temp tables
- No persistence to disk (yet)
- DuckDB runs in-memory mode
- Reference counting automatically cleans up unused handles and drops their temp tables

## Concurrency Model

The add-in uses a single shared DuckDB connection protected by locks. RTD callbacks and Excel function calls may occur on different threads.

Current approach: simple locking around connection access. This works but has limitations.

**Future consideration**: A dedicated worker thread (actor pattern) that owns the DuckDB connection would be cleaner. All database operations would be queued to this worker, eliminating lock contention and ensuring operations like temp table drops don't race with queries. Deferred for now as current locking is sufficient, but may be needed as complexity grows.

## Preview Pane

The XLDuck ribbon tab includes a "Preview Pane" toggle button that opens a task pane for inspecting handles.

### Architecture

```
Excel Selection Change
    → PreviewPaneManager (singleton)
    → Debounce (500ms)
    → PreviewController (serial queue)
    → PreviewDataProvider (data access)
    → PreviewPane (WebView2 host)
    → JavaScript UI
```

**Key components:**

| File | Purpose |
|------|---------|
| `PreviewPane.cs` | WinForms UserControl hosting WebView2, exposed via COM for CustomTaskPane |
| `PreviewPaneManager.cs` | Singleton managing panes per Excel window, handles selection events |
| `PreviewController.cs` | Debounce and serial queue to avoid UI thrashing |
| `PreviewDataProvider.cs` | Extracts preview data from ResultStore/FragmentStore |
| `PreviewModels.cs` | JSON-serializable models for WebView2 communication |
| `preview.html` | Embedded HTML/CSS/JS UI for rendering previews |

### COM Interop for .NET 8

CustomTaskPane requires a COM-visible control. .NET 6+ requires the `[ComDefaultInterface]` pattern:

```csharp
[ComVisible(true)]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface IPreviewPane { }

[ComVisible(true)]
[ComDefaultInterface(typeof(IPreviewPane))]
[ClassInterface(ClassInterfaceType.None)]
public class PreviewPane : UserControl, IPreviewPane
```

Without this pattern, Excel throws "Unable to create specified ActiveX control".

### Preview Types

**Table handles** show:
- SQL text and positional arguments (if any)
- Schema table: column names and DuckDB types
- Data grid: first 200 rows of data
- Row/column counts in the title

**Fragment handles** show:
- SQL text
- Bound positional arguments

**Plot handles** show:
- Interactive Vega-Lite chart
- Template name and row count
- Field bindings (x, y, color, title)

**Error handles** show:
- Error category and message

### WebView2 Integration

The pane uses Microsoft Edge WebView2 for rich HTML rendering:

- User data folder: `%LOCALAPPDATA%\XlDuck\WebView2` (avoids permission issues)
- Communication: JSON via `PostWebMessageAsString` / `window.chrome.webview.addEventListener`
- Graceful fallback to Label control if WebView2 runtime not installed

### JSON Serialization

Uses source-generated `JsonSerializerContext` for AOT compatibility. All row data is converted to `string?[]` because source-gen JSON cannot serialize boxed primitives in `object[]`.

### Debounce and Serial Queue

Selection changes fire rapidly. The controller:
1. Debounces for 500ms before processing
2. Queues requests serially to avoid race conditions
3. Cancels pending work when new selection arrives

## Plotting

XLDuck supports interactive charts via `DuckPlot`, rendered in the preview pane using Vega-Lite.

### Design

Plotting uses a **template-based** approach rather than requiring users to write Vega-Lite JSON:

```excel
=DuckPlot(data, "bar", "x", "region", "y", "sales", "color", "product")
```

Templates are hardcoded Vega-Lite specs. Users specify field bindings via overrides.

### Templates

| Template | Mark | Use Case |
|----------|------|----------|
| `bar` | bar | Aggregated values per category |
| `line` | line + points | Time series, trends |
| `point` | point | Scatter plots, correlations |
| `area` | area | Cumulative/stacked time series |
| `histogram` | bar (binned) | Distribution of a single column |
| `heatmap` | rect | Two categories with color intensity |
| `boxplot` | boxplot | Distribution comparison across categories |

### PlotStore

Plot configurations are stored similarly to fragments:

```csharp
class StoredPlot {
    string DataHandle;      // duck://table/... or duck://frag/...
    string Template;        // "bar", "line", etc.
    Dictionary<string, string> Overrides;  // x, y, color, title
}
```

Uses RTD lifecycle for automatic cleanup when cells no longer reference the plot.

### Data Caps

To prevent browser crashes with large datasets, plots enforce hard limits:
- Max rows: 50,000
- Max cells: 500,000

Exceeding limits shows an error prompting the user to aggregate or filter in SQL.

### Type Inference

Field types are inferred from DuckDB column types:
- VARCHAR, TEXT → `nominal`
- INTEGER, DOUBLE, etc. → `quantitative`
- DATE, TIMESTAMP → `temporal`

### Vega-Lite Integration

The preview pane loads Vega-Lite from CDN and renders charts via `vegaEmbed()`. Data is sent as column arrays and converted to Vega-Lite's `values` format in JavaScript.

## Future Considerations

- **Handle comments**: Allow user annotations on handles for readability
- **Persistence**: Save/load handle stores to disk
