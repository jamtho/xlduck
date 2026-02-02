# Architecture

## Overview

XlDuck is an Excel add-in that exposes DuckDB's SQL engine to spreadsheet users. The core idea is to enable **dataflow-style computation** where intermediate query results can be stored as handles and referenced by downstream queries, creating a DAG of computations across the sheet.

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

Query results from `DuckQuery` are stored in .NET memory, not as DuckDB tables. This allows users to create many intermediate results without polluting DuckDB's catalog.

Storage structure:
```csharp
class StoredResult {
    string[] ColumnNames;
    Type[] ColumnTypes;
    List<object[]> Rows;
}
```

Results are kept in a `Dictionary<string, StoredResult>` keyed by handle.

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

When a query references a stored result, users specify placeholders with `:name` syntax:

```excel
=DuckQuery("SELECT * FROM :sales WHERE region = 'EU'", "sales", A1)
```

Where A1 contains a handle like `duck://table/1234`.

Parameters are passed as name/value pairs after the SQL string.

## Data Flow

### Query Execution

```
DuckQuery("SELECT ...")
    → Execute in DuckDB
    → Read results into .NET memory
    → Generate handle
    → Store in dictionary
    → Return handle to cell
```

### Query with References

```
DuckQuery("SELECT * FROM :src", "src", "duck://table/1")
    → Parse SQL for :placeholders
    → For each placeholder:
        → If table handle (t): create temp table from stored rows
        → If fragment handle (f): recursively resolve and inline as subquery
        → Replace :name with temp table name or (subquery SQL)
    → Execute query in DuckDB
    → Drop temp tables
    → Store new result, return new handle
```

### Fragment Creation

```
DuckFrag("SELECT * FROM :src WHERE x > 5", "src", A1)
    → Resolve parameters (for validation only)
    → Run EXPLAIN to validate SQL
    → Drop any temp tables created for validation
    → Store original SQL + args
    → Return fragment handle
```

### Fragment Resolution

When a fragment is used as a parameter, it's resolved recursively:

```
DuckQuery("SELECT * FROM :data", "data", "duck://frag/1")
    → Look up fragment f/1
    → Resolve fragment's own parameters recursively
    → Inline resolved SQL as: (SELECT ...)
    → Continue with outer query resolution
```

Circular references (fragment A → B → A) are detected and raise an error.

### Materialization

```
DuckOut("duck://table/1")
    → Look up handle in storage
    → Convert to Excel array with headers
    → Return as spilled array
```

### Why Temp Tables?

DuckDB.NET doesn't expose a direct way to query Arrow memory or register external data sources from .NET. The workaround is:

1. Store results in .NET memory
2. When referenced, hydrate into a DuckDB temp table (INSERT rows)
3. Query the temp table
4. Drop it after

This involves copying data twice (DuckDB → .NET → DuckDB), which is inefficient for large datasets. Future optimization could use Arrow format with DuckDB's `arrow_scan()` if DuckDB.NET exposes the necessary bindings.

The goal is to support millions of rows efficiently. The current temp table approach works but will need optimization (likely via Arrow) for large-scale use.

## Excel Functions

| Function | Purpose |
|----------|---------|
| `DuckQuery(sql, [n1, v1, ...])` | Execute SQL, return table handle. Up to 4 `:name` placeholders. Add `"@config"` to wait for config. |
| `DuckFrag(sql, [n1, v1, ...])` | Create SQL fragment for lazy evaluation. Validated but not executed. Add `"@config"` to wait for config. |
| `DuckOut(handle)` | Output handle (table or frag) as spilled array with headers. |
| `DuckQueryOut(sql, [n1, v1, ...])` | Execute SQL and output directly as spilled array. Combo of DuckQuery + DuckOut. |
| `DuckExecute(sql)` | Execute DDL/DML (CREATE, INSERT, etc.) |
| `DuckConfigReady()` | Signal that configuration is complete. Queries with `@config` wait for this. |
| `DuckVersion()` | Return add-in version (0.1) |
| `DuckLibraryVersion()` | Return DuckDB library version |

**When to use which:**
- `DuckQuery` - Materialize and cache results (good for expensive queries used multiple times)
- `DuckFrag` - Defer execution, allow query optimization across composed fragments
- `DuckOut` - Display results from either handle type
- `DuckQueryOut` - One-off queries where you just want the output

## Known Issues and Workarounds

### HUGEINT/BigInteger Conversion

DuckDB's aggregate functions (SUM, etc.) return HUGEINT/INT128 types that .NET and Excel don't handle natively. The add-in automatically converts these to `double` for Excel compatibility. This may lose precision for very large integers.

### Parameter Limit

Excel-DNA doesn't support `params` arrays in UDFs. Instead, we use explicit optional parameters, limiting queries to 4 name/value pairs (8 parameters). This covers most use cases; complex joins needing more can use subqueries or intermediate handles.

## RTD and Lifecycle Management

### RTD-Based Functions

`DuckQuery` and `DuckFrag` use Excel's RTD (Real-Time Data) mechanism for lifecycle tracking. This enables:

1. **Reference counting**: Handles are automatically cleaned up when no longer referenced by any cell
2. **Cell lifecycle awareness**: When a cell is deleted or its formula changes, the handle's reference count decrements
3. **Automatic cleanup**: Handles with zero references are evicted from storage

### Timeout Budget

To avoid RTD's 2-second throttle delay, queries use a timeout budget:

- Queries completing within **200ms** return results directly (synchronous)
- Slower queries return "Loading..." immediately, then update asynchronously

This provides responsive UX for fast queries while supporting long-running operations.

### Configuration Gate (@config)

Queries needing runtime configuration (e.g., S3 endpoints) can wait for setup:

```excel
=DuckFrag("SELECT * FROM read_parquet(:url)", "url", A1, "@config")
```

The `@config` sentinel causes the query to wait until `DuckConfigReady()` is called, typically from VBA `Auto_Open`:

```vba
Sub Auto_Open()
    Application.Run "DuckExecute", "SET s3_endpoint = '127.0.0.1:9000'"
    Application.Run "DuckConfigReady"
End Sub
```

Downstream queries that depend on a blocked query (input starts with `#duck://blocked/`) also wait automatically.

### Bulk Insert Optimization

When materializing table handles into temp tables, the add-in uses DuckDB's Appender API for ~300x faster bulk inserts compared to row-by-row INSERT statements.

## Session Lifecycle

- All stored results persist for the Excel session
- Closing Excel clears all handles
- No persistence to disk (yet)
- DuckDB runs in-memory mode
- Reference counting automatically cleans up unused handles

## Future Considerations

- **Arrow optimization**: Avoid temp table round-trip by using Arrow memory directly
- **Handle comments**: Allow user annotations on handles for readability
- **Persistence**: Save/load handle stores to disk
