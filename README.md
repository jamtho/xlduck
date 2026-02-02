# XlDuck

Excel add-in wrapping DuckDB for in-cell SQL queries.

## Prerequisites

- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)
- [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0) (required for Excel add-ins)
- Microsoft Excel (64-bit)

## Build

```
cd XlDuck
dotnet build
```

## Run

Open the add-in directly to launch Excel with it loaded:

```
XlDuck\bin\Debug\net8.0-windows\XlDuck-AddIn64.xll
```

## Excel Functions

| Function | Description |
|----------|-------------|
| `=DuckQuery(sql, ...)` | Execute SQL, return a table handle (`duck://table/1\|10x4` = 10 rows, 4 cols) |
| `=DuckFrag(sql, ...)` | Create SQL fragment for lazy evaluation (`duck://frag/...`) |
| `=DuckOut(handle)` | Output a handle as a spilled array |
| `=DuckQueryOut(sql, ...)` | Execute SQL and output directly as array |
| `=DuckExecute(sql)` | Execute DDL/DML statements |
| `=DuckVersion()` | XlDuck add-in version (0.1) |
| `=DuckLibraryVersion()` | DuckDB library version |

## Examples

### Basic Usage

```excel
A1: =DuckQueryOut("SELECT * FROM range(5)")
→ | range |
  | 0     |
  | 1     |
  | 2     |
  | 3     |
  | 4     |
```

### Using Handles for Chaining

```excel
A1: =DuckQuery("SELECT * FROM range(10)")
→ duck://table/1|10x1

B1: =DuckQuery("SELECT * FROM :src WHERE range > 5", "src", A1)
→ duck://table/2|4x1

C1: =DuckQuery("SELECT SUM(range) AS total FROM :data", "data", B1)
→ duck://table/3|1x1

D1: =DuckOut(C1)
→ | total |
  | 30    |
```

### Parameter Binding

Use `:name` placeholders with name/value pairs (up to 4 pairs):

```excel
=DuckQuery("SELECT * FROM :t1 JOIN :t2 ON t1.id = t2.id", "t1", A1, "t2", B1)
```

### Lazy Evaluation with Fragments

Fragments (`duck://frag/...`) defer SQL execution - the SQL is inlined as a subquery when used:

```excel
A1: =DuckFrag("SELECT * FROM range(10)")
→ duck://frag/1

B1: =DuckFrag("SELECT * FROM :src WHERE range >= 5", "src", A1)
→ duck://frag/2

C1: =DuckOut(B1)
→ | range |
  | 5     |
  | 6     |
  | 7     |
  | 8     |
  | 9     |
```

When `DuckOut(B1)` executes, it builds and runs:
```sql
SELECT * FROM (SELECT * FROM (SELECT * FROM range(10)) WHERE range >= 5)
```

Fragments are validated at creation time (EXPLAIN), so SQL errors appear early.

Use fragments for:
- Building query pipelines without materializing intermediate results
- Allowing DuckDB to optimize the entire composed query
- Reducing memory usage for complex transformations

### Reading Files

DuckDB can read CSV, Parquet, JSON, and other file formats directly:

```excel
=DuckQueryOut("SELECT * FROM read_csv_auto('C:/data/sales.csv')")

=DuckQueryOut("SELECT * FROM read_parquet('C:/data/events.parquet') WHERE date > '2024-01-01'")

=DuckQueryOut("SELECT * FROM read_json_auto('C:/data/config.json')")
```

Combine with fragments for reusable data sources:

```excel
A1: =DuckFrag("SELECT * FROM read_csv_auto('C:/data/sales.csv')")
B1: =DuckQueryOut("SELECT region, SUM(amount) FROM :sales GROUP BY region", "sales", A1)
```

DuckDB can also read from URLs and S3 - see [DuckDB documentation](https://duckdb.org/docs/data/overview) for details.

### Pivot Tables

DuckDB has built-in PIVOT support for reshaping data:

```excel
A1: =DuckFrag("SELECT * FROM (VALUES ('Q1','North',100), ('Q1','South',150), ('Q2','North',200), ('Q2','South',250)) AS sales(quarter, region, amount)")

B1: =DuckQueryOut("PIVOT :data ON region USING SUM(amount)", "data", A1)
→ | quarter | North | South |
  | Q1      | 100   | 150   |
  | Q2      | 200   | 250   |
```

See [DuckDB PIVOT documentation](https://duckdb.org/docs/sql/statements/pivot) for more examples.

## Credits

Several design ideas take inspiration from the superb [PyXLL add-in](https://www.pyxll.com), which you should check out immediately if you've ever considered integrating Python code with your sheets.
