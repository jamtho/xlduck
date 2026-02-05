# Vega-Lite Plotting V1 - Implementation Plan

## Goal

Add charting to XlDuck via a template-based `DuckPlot` function that renders interactive Vega-Lite charts in the preview pane.

## Design Principles

- **Template-first**: Users pick from hardcoded chart templates by name, not write Vega-Lite JSON
- **Override-based customization**: Field bindings and options via name/value pairs
- **Reuse existing patterns**: PlotStore follows ResultStore/FragmentStore patterns; preview uses existing WebView2 infrastructure
- **Fail clearly**: Hard data caps with explicit errors, not silent truncation

## Function Signature

```excel
=DuckPlot(data, template, [n1], [v1], [n2], [v2], [n3], [v3], [n4], [v4])
```

**Parameters:**
- `data` - Table or fragment handle (`duck://table/...` or `duck://frag/...`)
- `template` - Template name: `"bar"`, `"line"`, `"point"`, `"area"`
- `n1, v1, ...` - Override name/value pairs (up to 4 pairs)

**Returns:** `duck://plot/<id>`

**Examples:**
```excel
=DuckPlot(A1, "bar", "x", "region", "y", "sales")
=DuckPlot(A1, "line", "x", "date", "y", "price", "color", "symbol")
=DuckPlot(A1, "point", "x", "height", "y", "weight", "title", "Height vs Weight")
```

## Templates

### V1 Templates

| Name | Mark | Use Case |
|------|------|----------|
| `bar` | `bar` | Category comparison |
| `line` | `line` | Time series |
| `point` | `point` | Scatter / correlation |
| `area` | `area` | Cumulative time series |

### Template Structure

Each template is a Vega-Lite JSON object with placeholders for field bindings:

```json
{
  "$schema": "https://vega.github.io/schema/vega-lite/v5.json",
  "mark": "bar",
  "encoding": {
    "x": {"field": null, "type": "auto"},
    "y": {"field": null, "type": "auto"}
  }
}
```

Field types are inferred from DuckDB column types at render time.

## Overrides

### Required
| Name | Purpose |
|------|---------|
| `x` | Field name for x-axis encoding |
| `y` | Field name for y-axis encoding |

### Optional
| Name | Purpose |
|------|---------|
| `color` | Field name for color encoding (creates series) |
| `title` | Chart title |

### Type Inference

Map DuckDB types to Vega-Lite types:

| DuckDB Type | Vega-Lite Type |
|-------------|----------------|
| VARCHAR, TEXT, ENUM | `nominal` |
| INTEGER, BIGINT, DOUBLE, DECIMAL | `quantitative` |
| DATE, TIMESTAMP, TIMESTAMPTZ | `temporal` |
| BOOLEAN | `nominal` |

## PlotStore

### Data Model

```csharp
public sealed record StoredPlot(
    string DataHandle,      // duck://table/... or duck://frag/...
    string Template,        // "bar", "line", etc.
    Dictionary<string, string> Overrides,  // x, y, color, title, etc.
    DateTime CreatedUtc
);
```

### Handle Format

`duck://plot/<id>` where id is auto-incrementing integer.

### Lifecycle

Use RTD refcount pattern (same as ResultStore/FragmentStore). Plot is cleaned up when no cells reference it.

## Preview Pane Integration

### Message Contract

Extend preview message types with a new `plot` kind:

```typescript
interface PlotPreviewModel {
  kind: "plot";
  title: string;
  handle: string;
  plot: {
    spec: object;      // Complete Vega-Lite spec with data
    error?: string;    // Render error if any
  };
}
```

### Data Flow

```
Selection Change (duck://plot/...)
  → PreviewController (debounce)
  → PreviewDataProvider.GetPlotPreview()
      → Resolve data handle → get rows/columns
      → Check data cap → fail if exceeded
      → Look up template
      → Apply overrides + type inference
      → Build complete Vega-Lite spec with inline data
  → PostMessage to WebView2
  → JavaScript renders via vegaEmbed()
```

### Data Packaging

Send data as:
```json
{
  "columns": ["region", "sales"],
  "types": ["nominal", "quantitative"],
  "rows": [["North", 100], ["South", 150], ...]
}
```

JavaScript builds `data.values` array of objects for Vega-Lite.

## Data Caps

### Limits
- Max rows: 50,000
- Max cells (rows × columns): 500,000
- Max JSON payload: 10 MB

### Behavior When Exceeded

Return error preview:
```
Error: Dataset too large for plotting (125,000 rows)
Maximum: 50,000 rows

Use DuckQuery to aggregate or filter your data:
  =DuckQuery("SELECT region, SUM(sales) FROM :data GROUP BY region", "data", A1)
```

## Vega-Lite Bundle

### Libraries
- `vega` 5.x
- `vega-lite` 5.x
- `vega-embed` 6.x

### Embedding
Bundle minified JS in `preview.html` as inline scripts (no CDN dependency).

### Render Options
```javascript
vegaEmbed('#chart', spec, {
  actions: false,      // Hide export menu
  renderer: 'canvas'   // Better performance for large data
});
```

## Implementation Checklist

### Phase 1: PlotStore + DuckPlot Function
- [ ] Add `PlotStore.cs` (following FragmentStore pattern)
- [ ] Add `DuckPlot` UDF in `DuckFunctions.cs`
- [ ] Add templates as embedded resource or static dictionary
- [ ] Validate template name, required overrides

### Phase 2: Preview Pane Rendering
- [ ] Add `PlotPreviewModel` to `PreviewModels.cs`
- [ ] Add `GetPlotPreview()` to `PreviewDataProvider.cs`
- [ ] Implement data cap checking
- [ ] Implement type inference from DuckDB schema
- [ ] Build complete Vega-Lite spec with data

### Phase 3: WebView2 Integration
- [ ] Bundle vega/vega-lite/vega-embed in `preview.html`
- [ ] Add plot rendering JavaScript
- [ ] Handle render errors gracefully
- [ ] Test with each template type

### Phase 4: Polish
- [ ] Error messages for invalid template, missing overrides
- [ ] Test with edge cases (empty data, single row, null values)
- [ ] Update documentation

## Files to Create/Modify

| File | Action |
|------|--------|
| `XlDuck/PlotStore.cs` | Create |
| `XlDuck/PlotTemplates.cs` | Create (template definitions) |
| `XlDuck/DuckFunctions.cs` | Add DuckPlot |
| `XlDuck/Preview/PreviewModels.cs` | Add PlotPreviewModel |
| `XlDuck/Preview/PreviewDataProvider.cs` | Add GetPlotPreview |
| `XlDuck/Preview/preview.html` | Add Vega-Lite + render logic |

## Out of Scope (V2)

- `DuckVegaLiteJson` / `DuckVegaLiteFile` (custom specs)
- SpecStore
- File-based specs with reload
- Sampling/downsampling params
- Snapshot export to worksheet
- Additional templates (histogram, heatmap, etc.)
- Template-specific options (stacked, interpolate, etc.)

## V2 Compatibility

Design choices that preserve V2 extensibility:
- `duck://plot/<id>` namespace (not template-specific)
- PlotStore can later store spec handle reference
- Override mechanism can grow without API change
- Preview rendering path can handle custom specs later
