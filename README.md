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
| `=DuckVersion()` | Returns DuckDB version |
| `=DuckQuery(sql, ...)` | Execute SQL, return single value |
| `=DuckQueryArray(sql, ...)` | Execute SQL, return array with headers |
| `=DuckQueryLazy(sql)` | Execute SQL, store result, return handle |
| `=DuckExecute(sql)` | Execute DDL/DML statements |

## Examples

### Basic Queries

```excel
=DuckQuery("SELECT 42 * 2")
→ 84

=DuckQueryArray("SELECT * FROM range(3)")
→ | range |
  | 0     |
  | 1     |
  | 2     |

=DuckExecute("CREATE TABLE test(id INT, name VARCHAR)")
→ OK (0 rows affected)
```

### Lazy Evaluation with Handles

Store intermediate results and reference them in downstream queries:

```excel
A1: =DuckQueryLazy("SELECT * FROM range(5)")
→ duck://t/1

A2: =DuckQuery("SELECT SUM(range) FROM :src", "src", A1)
→ 10

A3: =DuckQueryArray("SELECT * FROM :data WHERE range > 2", "data", A1)
→ | range |
  | 3     |
  | 4     |
```

Parameters use `:name` placeholders with name/value pairs (up to 4 pairs supported):

```excel
=DuckQuery("SELECT * FROM :t1 JOIN :t2 ON ...", "t1", A1, "t2", B1)
```
