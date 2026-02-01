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
| `=DuckQuery("SELECT 1+1")` | Execute SQL, return single value |
| `=DuckQueryArray("SELECT * FROM range(5)")` | Execute SQL, return array with headers |
| `=DuckExecute("CREATE TABLE t(x INT)")` | Execute DDL/DML statements |

## Examples

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
