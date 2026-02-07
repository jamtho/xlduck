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
| `=DuckOut(handle)` | Output a handle as a spilled array |
| `=DuckQueryOut(sql, ...)` | Execute SQL and output directly as array |
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
A1: =DuckQuery("SELECT weekday, hour, avg_temp FROM temperature_data")
B1: =DuckPlot(A1, "heatmap", "x", "hour", "y", "weekday", "value", "avg_temp")
```

## Preview Pane

The XLDuck ribbon tab includes a toggle to open a preview pane on the right side of the window. When you select a cell containing a handle:

- **Table handles**: Shows SQL with positional arguments, column schema, and the first 200 rows of data
- **Fragment handles**: Shows the SQL text and positional arguments
- **Plot handles**: Shows an interactive Vega-Lite chart
- **Error handles**: Shows the error category and message

Requires WebView2 Runtime (falls back to plain text if not installed).

## Credits

Several design ideas take inspiration from the superb [PyXLL add-in](https://www.pyxll.com), which you should check out immediately if you've ever considered integrating Python code with your sheets.
