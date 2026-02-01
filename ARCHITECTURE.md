# Architecture

## Overview

XlDuck is an Excel add-in that exposes DuckDB's SQL engine to spreadsheet users. The core idea is to enable **dataflow-style computation** where intermediate query results can be stored as handles and referenced by downstream queries, creating a DAG of computations across the sheet.

## Core Concepts

### Handles

A handle is a string reference to a stored query result, formatted as:
```
duck://t/1234
```

Where:
- `duck://` - protocol prefix
- `t` - type identifier (currently `t` for table/result set)
- `1234` - auto-generated numeric ID

Handles are displayed in cells and can be passed to other functions as table references.

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

### Query Parameter Binding

When a query references a stored result, users specify placeholders with `:name` syntax:

```excel
=DuckQuery("SELECT * FROM :sales WHERE region = 'EU'", "sales", A1)
```

Where A1 contains a handle like `duck://t/1234`.

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
DuckQuery("SELECT * FROM :src", "src", "duck://t/1")
    → Parse SQL for :placeholders
    → For each placeholder:
        → Look up handle in storage
        → Create temp DuckDB table from stored rows
        → Replace :name with temp table name
    → Execute query in DuckDB
    → Drop temp tables
    → Store new result, return new handle
```

### Materialization

```
DuckQueryOut("duck://t/1")
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

For typical spreadsheet use cases (thousands of rows, not millions), this overhead is acceptable.

## Excel Functions

| Function | Purpose |
|----------|---------|
| `DuckQuery(sql, [n1, v1, n2, v2, n3, v3, n4, v4])` | Execute SQL, return handle. Up to 4 `:name` placeholders. |
| `DuckQueryOut(handle)` | Output stored result as spilled array with headers. |
| `DuckExecute(sql)` | Execute DDL/DML (CREATE, INSERT, etc.) |
| `DuckVersion()` | Return DuckDB version |

All queries return handles. Use `DuckQueryOut` to materialize results to the sheet.

## Known Issues and Workarounds

### HUGEINT/BigInteger Conversion

DuckDB's aggregate functions (SUM, etc.) return HUGEINT/INT128 types that .NET and Excel don't handle natively. The add-in automatically converts these to `double` for Excel compatibility. This may lose precision for very large integers.

### Parameter Limit

Excel-DNA doesn't support `params` arrays in UDFs. Instead, we use explicit optional parameters, limiting queries to 4 name/value pairs (8 parameters). This covers most use cases; complex joins needing more can use subqueries or intermediate handles.

## Session Lifecycle

- All stored results persist for the Excel session
- Closing Excel clears all handles
- No persistence to disk (yet)
- DuckDB runs in-memory mode

## Future Considerations

- **Arrow optimization**: Avoid temp table round-trip by using Arrow memory directly
- **Handle comments**: Allow user annotations on handles for readability
- **Persistence**: Save/load handle stores to disk
- **Handle cleanup**: Manual or automatic garbage collection of unused handles
