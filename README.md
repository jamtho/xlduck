# XLDuck

Excel add-in wrapping DuckDB for in-cell SQL queries.

## Prerequisites

- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)
- [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0) (required for Excel add-ins)
- Microsoft Excel (64-bit)
- [WebView2 Runtime](https://developer.microsoft.com/microsoft-edge/webview2/) (optional, for preview pane)

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
| `=DuckQueryAfterConfig(sql, ...)` | Same as DuckQuery, but waits for `DuckConfigReady()` first |
| `=DuckFrag(sql, ...)` | Create SQL fragment for lazy evaluation (`duck://frag/...`) |
| `=DuckFragAfterConfig(sql, ...)` | Same as DuckFrag, but waits for `DuckConfigReady()` first |
| `=DuckCapture(range)` | Capture a sheet range as a DuckDB table (first row = headers) |
| `=DuckDate(cell)` | Convert Excel date to SQL date string (`2023-01-01`) |
| `=DuckDateTime(cell)` | Convert Excel date/time to SQL datetime string (`2023-01-01 14:30:00`) |
| `=DuckOut(handle)` | Output a handle as a spilled array |
| `=DuckQueryOut(sql, ...)` | Execute SQL and output directly as array |
| `=DuckQueryOutScalar(sql, ...)` | Execute SQL and return a single value (first column, first row) |
| `=DuckPlot(data, template, ...)` | Create a chart from data (`duck://plot/...`) |
| `=DuckExecute(sql)` | Execute DDL/DML statements |
| `=DuckConfigReady()` | Signal that configuration is complete |
| `=DuckVersion()` | XLDuck add-in version (0.1) |
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

B1: =DuckQuery("SELECT * FROM ? WHERE range > 5", A1)
→ duck://table/2|4x1

C1: =DuckQuery("SELECT SUM(range) AS total FROM ?", B1)
→ duck://table/3|1x1

D1: =DuckOut(C1)
→ | total |
  | 30    |
```

### Parameter Binding

Use `?` placeholders for positional arguments (up to 8):

```excel
=DuckQuery("SELECT * FROM ? JOIN ? ON t1.id = t2.id", A1, B1)
```

### Lazy Evaluation with Fragments

Fragments (`duck://frag/...`) defer SQL execution - the SQL is inlined as a subquery when used:

```excel
A1: =DuckFrag("SELECT * FROM range(10)")
→ duck://frag/1

B1: =DuckFrag("SELECT * FROM ? WHERE range >= 5", A1)
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
B1: =DuckQueryOut("SELECT region, SUM(amount) FROM ? GROUP BY region", A1)
```

DuckDB can also read from URLs and S3 - see [DuckDB documentation](https://duckdb.org/docs/data/overview) for details.

### Capturing Sheet Data

`DuckCapture` brings Excel range data into DuckDB for querying. The first row is treated as headers, the rest as data:

```excel
A1:C4 contains:
  | name    | age | city    |
  | alice   | 30  | NYC     |
  | bob     | 25  | LA      |
  | charlie | 35  | Chicago |

D1: =DuckCapture(A1:C4)
→ duck://table/1|3x3

E1: =DuckQueryOut("SELECT * FROM ? WHERE age > 28", D1)
→ | name    | age | city    |
  | alice   | 30  | NYC     |
  | charlie | 35  | Chicago |
```

Column types are inferred automatically: all-numeric columns become `DOUBLE`, all-boolean become `BOOLEAN`, everything else becomes `VARCHAR`. Empty cells are treated as `NULL`.

Combine with other functions for analysis:

```excel
A1: =DuckCapture(Sheet2!A1:D100)
B1: =DuckQueryOut("SELECT department, AVG(salary) FROM ? GROUP BY department", A1)
```

### Date Parameters

Excel stores dates as serial numbers, so passing a date cell directly as a `?` parameter won't work — DuckDB receives a number like `44927` instead of a date. Use `DuckDate` or `DuckDateTime` to convert:

```excel
C1: 1/1/2023                    (Excel date)
D1: 12/31/2023                  (Excel date)

E1: =DuckQuery("SELECT * FROM ? WHERE date BETWEEN ? AND ?", A1, DuckDate(C1), DuckDate(D1))

F1: =DuckQuery("SELECT * FROM ? WHERE timestamp > ?", A1, DuckDateTime(C1))
```

### Pivot Tables

DuckDB has built-in PIVOT support for reshaping data:

```excel
A1: =DuckFrag("SELECT * FROM (VALUES ('Q1','North',100), ('Q1','South',150), ('Q2','North',200), ('Q2','South',250)) AS sales(quarter, region, amount)")

B1: =DuckQueryOut("PIVOT ? ON region USING SUM(amount)", A1)
→ | quarter | North | South |
  | Q1      | 100   | 150   |
  | Q2      | 200   | 250   |
```

See [DuckDB PIVOT documentation](https://duckdb.org/docs/sql/statements/pivot) for more examples.

### Plotting

Create interactive charts with `DuckPlot`. Select a plot handle cell and open the Preview Pane to view.

```excel
A1: =DuckQuery("SELECT region, SUM(sales) as total FROM (VALUES ('North', 100), ('South', 150), ('East', 80), ('West', 120)) AS t(region, sales) GROUP BY region")

B1: =DuckPlot(A1, "bar", "x", "region", "y", "total", "title", "Sales by Region")
→ duck://plot/1
```

**Templates:**

| Template | Use Case |
|----------|----------|
| `bar` | Aggregated values per category |
| `line` | Time series, trends |
| `point` | Scatter plots, correlations |
| `area` | Cumulative/stacked time series |
| `histogram` | Distribution of values (only needs `x`) |
| `heatmap` | Two categories with color intensity (needs `x`, `y`, `value`) |
| `boxplot` | Distribution comparison across categories |

**Overrides:**
- `x` - field for x-axis (required)
- `y` - field for y-axis (required, except histogram)
- `color` - field for color/series (optional)
- `value` - field for color intensity (heatmap only)
- `title` - chart title (optional)

**Examples:**

Bar chart:
```excel
A1: =DuckQuery("SELECT region, SUM(sales) as total FROM (VALUES ('North', 100), ('South', 150), ('East', 80), ('West', 120)) AS t(region, sales) GROUP BY region")
B1: =DuckPlot(A1, "bar", "x", "region", "y", "total", "title", "Sales by Region")
```

Line chart with multiple series:
```excel
A1: =DuckQuery("SELECT x as day, 'A' as product, x*10 as sales FROM range(20) UNION ALL SELECT x, 'B', x*7+20 FROM range(20)")
B1: =DuckPlot(A1, "line", "x", "day", "y", "sales", "color", "product")
```

Scatter plot:
```excel
A1: =DuckQuery("SELECT random()*100 as x, random()*100 as y FROM range(200)")
B1: =DuckPlot(A1, "point", "x", "x", "y", "y")
```

Area chart with stacked series:
```excel
A1: =DuckQuery("SELECT x as month, 'Product A' as product, x*5+10 as revenue FROM range(12) UNION ALL SELECT x, 'Product B', x*3+20 FROM range(12)")
B1: =DuckPlot(A1, "area", "x", "month", "y", "revenue", "color", "product")
```

Histogram (distribution):
```excel
A1: =DuckQuery("SELECT random()*100 as value FROM range(1000)")
B1: =DuckPlot(A1, "histogram", "x", "value", "title", "Value Distribution")
```

Boxplot (compare distributions):
```excel
A1: =DuckQuery("SELECT category, value FROM (SELECT 'A' as category, random()*50 as value FROM range(100) UNION ALL SELECT 'B', random()*50+25 FROM range(100))")
B1: =DuckPlot(A1, "boxplot", "x", "category", "y", "value")
```

Heatmap:
```excel
A1: =DuckQuery("SELECT day, hour, temp FROM (SELECT d.d as day, h.h as hour, (15 + d.d + h.h*0.5 + random()*5)::INT as temp FROM range(7) AS d(d), range(24) AS h(h))")
B1: =DuckPlot(A1, "heatmap", "x", "hour", "y", "day", "value", "temp")
```

## Pause Queries

The XLDuck ribbon tab includes a **Pause Queries** toggle button. When paused, all query execution is deferred — cells show `#duck://blocked/paused|Queries paused` instead of running. Toggle off to resume: all deferred queries execute automatically and results flow to cells.

Use this when building complex formulas to avoid triggering expensive queries during editing. For example, if you have a chain of queries (A1 feeds B1 feeds C1) and need to restructure them, pausing prevents each intermediate edit from triggering a cascade of expensive executions.

```excel
1. Click "Pause Queries" in the XLDuck ribbon tab
2. Edit formulas freely — cells show "Queries paused" instead of executing
3. Click "Pause Queries" again to resume — all queries execute and results appear
```

Also available programmatically via VBA:
```vba
Application.Run "DuckPauseQueries"   ' pause
Application.Run "DuckResumeQueries"  ' resume
```

**Pause vs Cancel — two different behaviors:**
- **Pause**: "I'm editing, don't run anything yet" — queries are deferred and resume automatically when toggled off
- **Cancel**: "Kill this query, I don't want it" — queries error out permanently, new queries after cancellation run normally

Pausing while a query is running will cancel it and defer it for re-execution on resume. Queries that depend on other paused queries resolve naturally through Excel's recalculation chain — root queries execute first, then dependents recalculate with the correct handles.

## Query Engine Busy

Synchronous functions (`DuckOut`, `DuckQueryOut`, `DuckQueryOutScalar`, `DuckExecute`) may show a "Query engine busy" error if a background query is running. This prevents Excel from freezing during long-running queries. Press F9 to recalculate once the background query completes.

## Cancel Query

The XLDuck ribbon tab includes a **Cancel Query** button that interrupts the running query and cancels all pending queued queries. The connection remains valid after cancellation — subsequent queries work normally.

Also available programmatically via VBA:
```vba
Application.Run "DuckInterrupt"
```

## Preview Pane

The XLDuck ribbon tab includes a toggle to open a preview pane on the right side of the window. When you select a cell containing a handle:

- **Table handles**: Shows SQL with positional arguments, column schema, and the first 200 rows of data
- **Fragment handles**: Shows the SQL text and positional arguments
- **Plot handles**: Shows an interactive Vega-Lite chart
- **Error handles**: Shows the error category and message

Requires WebView2 Runtime (falls back to plain text if not installed).

## Logging

Log files are written to `%LOCALAPPDATA%\XlDuck\` (typically `C:\Users\<you>\AppData\Local\XlDuck\`). Each Excel session creates a new log file named `xlduck-{timestamp}.log`. Files older than 7 days are automatically deleted on startup.

Logs include query timing, RTD lifecycle events, handle reference counting, and error details. To tail the current session's log:

```powershell
Get-ChildItem "$env:LOCALAPPDATA\XlDuck\xlduck-*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1 | Get-Content -Wait
```

## Credits

Several design ideas take inspiration from the superb [PyXLL add-in](https://www.pyxll.com), which you should check out immediately if you've ever considered integrating Python code with your sheets.
