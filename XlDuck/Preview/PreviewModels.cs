// Copyright (c) 2026 James Thompson
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

using System.Text.Json;
using System.Text.Json.Serialization;

namespace XlDuck.Preview;

/// <summary>
/// Base class for preview models sent to WebView2.
/// </summary>
public abstract class PreviewModel
{
    [JsonPropertyName("kind")]
    public abstract string Kind { get; }

    [JsonPropertyName("title")]
    public string Title { get; set; } = "";

    [JsonPropertyName("handle")]
    public string? Handle { get; set; }

    [JsonPropertyName("message")]
    public string? Message { get; set; }

    public string ToJson()
    {
        return JsonSerializer.Serialize(this, GetType(), PreviewJsonContext.Default);
    }
}

/// <summary>
/// Preview model for empty/placeholder state.
/// </summary>
public class EmptyPreviewModel : PreviewModel
{
    public override string Kind => "empty";
}

/// <summary>
/// Preview model for error state.
/// </summary>
public class ErrorPreviewModel : PreviewModel
{
    public override string Kind => "error";
}

/// <summary>
/// Column schema information.
/// </summary>
public class ColumnInfo
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    [JsonPropertyName("type")]
    public string Type { get; set; } = "";
}

/// <summary>
/// Table preview data.
/// </summary>
public class TablePreviewData
{
    [JsonPropertyName("rowCount")]
    public long RowCount { get; set; }

    [JsonPropertyName("colCount")]
    public int ColCount { get; set; }

    [JsonPropertyName("duckTableName")]
    public string DuckTableName { get; set; } = "";

    [JsonPropertyName("columns")]
    public List<ColumnInfo> Columns { get; set; } = new();

    [JsonPropertyName("rows")]
    public List<string?[]> Rows { get; set; } = new();

    [JsonPropertyName("previewRowCount")]
    public int PreviewRowCount { get; set; }

    [JsonPropertyName("sql")]
    public string? Sql { get; set; }

    [JsonPropertyName("args")]
    public List<FragmentArg>? Args { get; set; }
}

/// <summary>
/// Preview model for table handles.
/// </summary>
public class TablePreviewModel : PreviewModel
{
    public override string Kind => "table";

    [JsonPropertyName("table")]
    public TablePreviewData Table { get; set; } = new();
}

/// <summary>
/// Fragment argument (position label and value).
/// </summary>
public class FragmentArg
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    [JsonPropertyName("value")]
    public string Value { get; set; } = "";
}

/// <summary>
/// Fragment preview data.
/// </summary>
public class FragPreviewData
{
    [JsonPropertyName("sql")]
    public string Sql { get; set; } = "";

    [JsonPropertyName("args")]
    public List<FragmentArg> Args { get; set; } = new();
}

/// <summary>
/// Preview model for fragment handles.
/// </summary>
public class FragPreviewModel : PreviewModel
{
    public override string Kind => "frag";

    [JsonPropertyName("frag")]
    public FragPreviewData Frag { get; set; } = new();
}

/// <summary>
/// Plot preview data.
/// </summary>
public class PlotPreviewData
{
    [JsonPropertyName("template")]
    public string Template { get; set; } = "";

    [JsonPropertyName("columns")]
    public List<string> Columns { get; set; } = new();

    [JsonPropertyName("types")]
    public List<string> Types { get; set; } = new();

    [JsonPropertyName("rows")]
    public List<string?[]> Rows { get; set; } = new();

    [JsonPropertyName("overrides")]
    public Dictionary<string, string> Overrides { get; set; } = new();

    [JsonPropertyName("rowCount")]
    public long RowCount { get; set; }

    [JsonPropertyName("error")]
    public string? Error { get; set; }
}

/// <summary>
/// Preview model for plot handles.
/// </summary>
public class PlotPreviewModel : PreviewModel
{
    public override string Kind => "plot";

    [JsonPropertyName("plot")]
    public PlotPreviewData Plot { get; set; } = new();
}

/// <summary>
/// JSON serialization context for preview models (AOT-friendly).
/// </summary>
[JsonSourceGenerationOptions(
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    WriteIndented = false)]
[JsonSerializable(typeof(EmptyPreviewModel))]
[JsonSerializable(typeof(ErrorPreviewModel))]
[JsonSerializable(typeof(TablePreviewModel))]
[JsonSerializable(typeof(FragPreviewModel))]
[JsonSerializable(typeof(PlotPreviewModel))]
internal partial class PreviewJsonContext : JsonSerializerContext
{
}
