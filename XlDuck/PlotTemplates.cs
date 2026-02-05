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
    public static readonly string[] TemplateNames = { "bar", "line", "point", "area" };

    private static readonly Dictionary<string, JsonObject> _templates = new()
    {
        ["bar"] = CreateTemplate("bar"),
        ["line"] = CreateTemplate("line"),
        ["point"] = CreateTemplate("point"),
        ["area"] = CreateTemplate("area"),
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

        // Required: x and y
        if (!overrides.TryGetValue("x", out var xField))
            throw new ArgumentException("Missing required override: x");
        if (!overrides.TryGetValue("y", out var yField))
            throw new ArgumentException("Missing required override: y");

        // Find column types for x and y
        var xType = GetVegaType(xField, columnNames, columnTypes);
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

        // Optional: color
        if (overrides.TryGetValue("color", out var colorField))
        {
            var colorType = GetVegaType(colorField, columnNames, columnTypes);
            encoding["color"] = new JsonObject
            {
                ["field"] = colorField,
                ["type"] = colorType
            };
        }

        // Optional: title
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
