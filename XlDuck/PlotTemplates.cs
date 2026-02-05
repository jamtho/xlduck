// Copyright (c) 2026 James Thompson
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

using System.Text.Json;
using System.Text.Json.Nodes;

namespace XlDuck;

/// <summary>
/// Hardcoded Vega-Lite chart templates.
/// </summary>
public static class PlotTemplates
{
    /// <summary>
    /// Available template names.
    /// </summary>
    public static readonly string[] TemplateNames = { "bar", "line", "point", "area", "histogram", "heatmap", "boxplot" };

    // Templates that only require x (y is auto-generated)
    private static readonly HashSet<string> _xOnlyTemplates = new() { "histogram" };

    // Templates that require a 'value' field for color intensity
    private static readonly HashSet<string> _valueTemplates = new() { "heatmap" };

    private static readonly Dictionary<string, JsonObject> _templates = new()
    {
        ["bar"] = CreateTemplate("bar"),
        ["line"] = CreateTemplate("line"),
        ["point"] = CreateTemplate("point"),
        ["area"] = CreateTemplate("area"),
        ["histogram"] = CreateHistogramTemplate(),
        ["heatmap"] = CreateHeatmapTemplate(),
        ["boxplot"] = CreateBoxplotTemplate(),
    };

    /// <summary>
    /// Check if a template name is valid.
    /// </summary>
    public static bool IsValidTemplate(string name)
    {
        return _templates.ContainsKey(name);
    }

    /// <summary>
    /// Get a template by name. Returns null if not found.
    /// </summary>
    public static JsonObject? GetTemplate(string name)
    {
        return _templates.TryGetValue(name, out var template) ? template.DeepClone().AsObject() : null;
    }

    /// <summary>
    /// Build a complete Vega-Lite spec from a template and overrides.
    /// </summary>
    public static JsonObject BuildSpec(
        string templateName,
        Dictionary<string, string> overrides,
        string[] columnNames,
        string[] columnTypes,
        JsonArray dataValues)
    {
        var spec = GetTemplate(templateName)
            ?? throw new ArgumentException($"Unknown template: {templateName}");

        // Apply field bindings to encoding
        var encoding = spec["encoding"]?.AsObject()
            ?? throw new InvalidOperationException("Template missing encoding");

        // Required: x
        if (!overrides.TryGetValue("x", out var xField))
            throw new ArgumentException("Missing required override: x");

        var xType = GetVegaType(xField, columnNames, columnTypes);

        // Handle special templates
        if (templateName == "histogram")
        {
            // Histogram: x is binned, y is count
            encoding["x"] = new JsonObject
            {
                ["field"] = xField,
                ["type"] = "quantitative",
                ["bin"] = true
            };
            encoding["y"] = new JsonObject
            {
                ["aggregate"] = "count"
            };
        }
        else if (templateName == "heatmap")
        {
            // Heatmap: x and y are categories, color is the value
            if (!overrides.TryGetValue("y", out var yFieldHeat))
                throw new ArgumentException("Missing required override: y");
            if (!overrides.TryGetValue("value", out var valueField))
                throw new ArgumentException("Missing required override: value (for color intensity)");

            var yTypeHeat = GetVegaType(yFieldHeat, columnNames, columnTypes);

            encoding["x"] = new JsonObject
            {
                ["field"] = xField,
                ["type"] = xType == "quantitative" ? "ordinal" : xType
            };
            encoding["y"] = new JsonObject
            {
                ["field"] = yFieldHeat,
                ["type"] = yTypeHeat == "quantitative" ? "ordinal" : yTypeHeat
            };
            encoding["color"] = new JsonObject
            {
                ["field"] = valueField,
                ["type"] = "quantitative",
                ["aggregate"] = "mean"
            };
        }
        else if (templateName == "boxplot")
        {
            // Boxplot: x is category, y is values to summarize
            if (!overrides.TryGetValue("y", out var yFieldBox))
                throw new ArgumentException("Missing required override: y");

            encoding["x"] = new JsonObject
            {
                ["field"] = xField,
                ["type"] = xType == "quantitative" ? "nominal" : xType
            };
            encoding["y"] = new JsonObject
            {
                ["field"] = yFieldBox,
                ["type"] = "quantitative"
            };
        }
        else
        {
            // Standard templates: x and y required
            if (!overrides.TryGetValue("y", out var yField))
                throw new ArgumentException("Missing required override: y");

            var yType = GetVegaType(yField, columnNames, columnTypes);

            encoding["x"] = new JsonObject
            {
                ["field"] = xField,
                ["type"] = xType
            };
            encoding["y"] = new JsonObject
            {
                ["field"] = yField,
                ["type"] = yType
            };

            // Optional: color (not for heatmap which handles it specially)
            if (overrides.TryGetValue("color", out var colorField))
            {
                var colorType = GetVegaType(colorField, columnNames, columnTypes);
                encoding["color"] = new JsonObject
                {
                    ["field"] = colorField,
                    ["type"] = colorType
                };
            }
        }

        // Optional: title (all templates)
        if (overrides.TryGetValue("title", out var title))
        {
            spec["title"] = title;
        }

        // Add data
        spec["data"] = new JsonObject
        {
            ["values"] = dataValues
        };

        return spec;
    }

    /// <summary>
    /// Create a base template for a given mark type.
    /// </summary>
    private static JsonObject CreateTemplate(string mark)
    {
        var template = new JsonObject
        {
            ["$schema"] = "https://vega.github.io/schema/vega-lite/v5.json",
            ["width"] = "container",
            ["height"] = 300,
            ["mark"] = new JsonObject
            {
                ["type"] = mark,
                ["tooltip"] = true
            },
            ["encoding"] = new JsonObject()
        };

        // Add line-specific defaults
        if (mark == "line")
        {
            template["mark"]!["point"] = true;
        }

        return template;
    }

    /// <summary>
    /// Create histogram template (binned x, count y).
    /// </summary>
    private static JsonObject CreateHistogramTemplate()
    {
        return new JsonObject
        {
            ["$schema"] = "https://vega.github.io/schema/vega-lite/v5.json",
            ["width"] = "container",
            ["height"] = 300,
            ["mark"] = new JsonObject
            {
                ["type"] = "bar",
                ["tooltip"] = true
            },
            ["encoding"] = new JsonObject()
        };
    }

    /// <summary>
    /// Create heatmap template (rect marks with color encoding).
    /// </summary>
    private static JsonObject CreateHeatmapTemplate()
    {
        return new JsonObject
        {
            ["$schema"] = "https://vega.github.io/schema/vega-lite/v5.json",
            ["width"] = "container",
            ["height"] = 300,
            ["mark"] = new JsonObject
            {
                ["type"] = "rect",
                ["tooltip"] = true
            },
            ["encoding"] = new JsonObject()
        };
    }

    /// <summary>
    /// Create boxplot template.
    /// </summary>
    private static JsonObject CreateBoxplotTemplate()
    {
        return new JsonObject
        {
            ["$schema"] = "https://vega.github.io/schema/vega-lite/v5.json",
            ["width"] = "container",
            ["height"] = 300,
            ["mark"] = new JsonObject
            {
                ["type"] = "boxplot",
                ["extent"] = "min-max"
            },
            ["encoding"] = new JsonObject()
        };
    }

    /// <summary>
    /// Map a column's DuckDB type to a Vega-Lite type.
    /// </summary>
    private static string GetVegaType(string fieldName, string[] columnNames, string[] columnTypes)
    {
        var index = Array.IndexOf(columnNames, fieldName);
        if (index < 0)
        {
            // Field not found - default to nominal
            return "nominal";
        }

        var duckType = columnTypes[index].ToUpperInvariant();

        // Temporal types
        if (duckType.Contains("DATE") || duckType.Contains("TIME") || duckType.Contains("TIMESTAMP"))
            return "temporal";

        // Quantitative types
        if (duckType.Contains("INT") || duckType.Contains("FLOAT") || duckType.Contains("DOUBLE") ||
            duckType.Contains("DECIMAL") || duckType.Contains("NUMERIC") || duckType.Contains("REAL") ||
            duckType.Contains("BIGINT") || duckType.Contains("SMALLINT") || duckType.Contains("TINYINT") ||
            duckType.Contains("HUGEINT"))
            return "quantitative";

        // Everything else is nominal
        return "nominal";
    }
}
