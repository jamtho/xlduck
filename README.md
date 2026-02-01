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
| `=DuckQuery(sql, ...)` | Execute SQL, return a handle |
| `=DuckQueryOut(handle)` | Output a handle as a spilled array |
| `=DuckExecute(sql)` | Execute DDL/DML statements |
| `=DuckVersion()` | Returns DuckDB version |

## Examples

### Basic Usage

```excel
A1: =DuckQuery("SELECT * FROM range(5)")
→ duck://t/1

A2: =DuckQueryOut(A1)
→ | range |
  | 0     |
  | 1     |
  | 2     |
  | 3     |
  | 4     |
```

### Chaining Queries with Handles

Store intermediate results and reference them in downstream queries:

```excel
A1: =DuckQuery("SELECT * FROM range(10)")
→ duck://t/1

B1: =DuckQuery("SELECT * FROM :src WHERE range > 5", "src", A1)
→ duck://t/2

C1: =DuckQuery("SELECT SUM(range) AS total FROM :data", "data", B1)
→ duck://t/3

D1: =DuckQueryOut(C1)
→ | total |
  | 30    |
```

### Parameter Binding

Use `:name` placeholders with name/value pairs (up to 4 pairs):

```excel
=DuckQuery("SELECT * FROM :t1 JOIN :t2 ON t1.id = t2.id", "t1", A1, "t2", B1)
```
