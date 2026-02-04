// Copyright (c) 2026 James Thompson
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

namespace XlDuck.Preview;

/// <summary>
/// Provides preview data for handles.
/// </summary>
public static class PreviewDataProvider
{
    private const int PreviewRowLimit = 200;

    /// <summary>
    /// Get a preview model for the given cell value.
    /// </summary>
    public static PreviewModel GetPreview(string? cellValue)
    {
        if (string.IsNullOrEmpty(cellValue))
        {
            return new EmptyPreviewModel
            {
                Title = "No Selection",
                Message = "Select a cell containing a handle to preview"
            };
        }

        // Check for error handles
        if (cellValue.StartsWith(DuckFunctions.ErrorPrefix))
        {
            return GetErrorPreview(cellValue);
        }

        // Check for blocked handles
        if (cellValue.StartsWith(DuckFunctions.BlockedPrefix))
        {
            return GetBlockedPreview(cellValue);
        }

        // Check for table handles
        if (ResultStore.IsHandle(cellValue))
        {
            return GetTablePreview(cellValue);
        }

        // Check for fragment handles
        if (FragmentStore.IsHandle(cellValue))
        {
            return GetFragmentPreview(cellValue);
        }

        // Not a handle
        return new EmptyPreviewModel
        {
            Title = "Not a Handle",
            Message = "Select a cell containing a handle to preview"
        };
    }

    private static PreviewModel GetErrorPreview(string handle)
    {
        // Format: #duck://error/category|message
        var pipeIndex = handle.IndexOf('|');
        string category, message;

        if (pipeIndex >= 0)
        {
            var prefix = handle[..pipeIndex];
            category = prefix.Replace(DuckFunctions.ErrorPrefix, "");
            message = handle[(pipeIndex + 1)..];
        }
        else
        {
            category = "error";
            message = handle.Replace(DuckFunctions.ErrorPrefix, "");
        }

        return new ErrorPreviewModel
        {
            Title = $"Error ({category})",
            Handle = handle,
            Message = message
        };
    }

    private static PreviewModel GetBlockedPreview(string handle)
    {
        // Format: #duck://blocked/reason|message
        var pipeIndex = handle.IndexOf('|');
        string message = pipeIndex >= 0
            ? handle[(pipeIndex + 1)..]
            : "Waiting...";

        return new ErrorPreviewModel
        {
            Title = "Blocked",
            Handle = handle,
            Message = message
        };
    }

    private static PreviewModel GetTablePreview(string handle)
    {
        var stored = ResultStore.Get(handle);
        if (stored == null)
        {
            return new ErrorPreviewModel
            {
                Title = "Handle Not Found",
                Handle = handle,
                Message = "This handle may have been released"
            };
        }

        try
        {
            var conn = DuckFunctions.GetConnection();
            var tableData = new TablePreviewData
            {
                RowCount = stored.RowCount,
                ColCount = stored.ColumnNames.Length,
                DuckTableName = stored.DuckTableName
            };

            // Get column schema via PRAGMA table_info
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = $"PRAGMA table_info('{stored.DuckTableName}')";
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    tableData.Columns.Add(new ColumnInfo
                    {
                        Name = reader.GetString(reader.GetOrdinal("name")),
                        Type = reader.GetString(reader.GetOrdinal("type"))
                    });
                }
            }

            // Get preview rows
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = $"SELECT * FROM \"{stored.DuckTableName}\" LIMIT {PreviewRowLimit}";
                using var reader = cmd.ExecuteReader();

                var fieldCount = reader.FieldCount;
                while (reader.Read())
                {
                    var row = new object?[fieldCount];
                    for (int i = 0; i < fieldCount; i++)
                    {
                        row[i] = reader.IsDBNull(i) ? null : ConvertForJson(reader.GetValue(i));
                    }
                    tableData.Rows.Add(row);
                }
                tableData.PreviewRowCount = tableData.Rows.Count;
            }

            return new TablePreviewModel
            {
                Title = $"Table ({stored.RowCount:N0} rows, {stored.ColumnNames.Length} cols)",
                Handle = handle,
                Table = tableData
            };
        }
        catch (Exception ex)
        {
            return new ErrorPreviewModel
            {
                Title = "Query Error",
                Handle = handle,
                Message = ex.Message
            };
        }
    }

    private static PreviewModel GetFragmentPreview(string handle)
    {
        var fragment = FragmentStore.Get(handle);
        if (fragment == null)
        {
            return new ErrorPreviewModel
            {
                Title = "Fragment Not Found",
                Handle = handle,
                Message = "This fragment may have been released"
            };
        }

        var fragData = new FragPreviewData
        {
            Sql = fragment.Sql
        };

        // Parse args as name/value pairs
        for (int i = 0; i + 1 < fragment.Args.Length; i += 2)
        {
            var name = fragment.Args[i]?.ToString() ?? "";
            var value = fragment.Args[i + 1]?.ToString() ?? "";

            // Skip @config sentinel
            if (name == DuckFunctions.ConfigSentinel)
            {
                i--; // Sentinel is standalone, adjust index
                continue;
            }

            fragData.Args.Add(new FragmentArg
            {
                Name = name,
                Value = value
            });
        }

        return new FragPreviewModel
        {
            Title = "Fragment",
            Handle = handle,
            Frag = fragData
        };
    }

    /// <summary>
    /// Convert a DuckDB value to a JSON-safe representation.
    /// </summary>
    private static object? ConvertForJson(object value)
    {
        if (value == null || value == DBNull.Value)
            return null;

        // Handle BigInteger (HUGEINT)
        if (value is System.Numerics.BigInteger bigInt)
            return bigInt.ToString();

        // Handle DateTime
        if (value is DateTime dt)
            return dt.ToString("O"); // ISO 8601

        // Handle DateTimeOffset
        if (value is DateTimeOffset dto)
            return dto.ToString("O");

        // Handle TimeSpan
        if (value is TimeSpan ts)
            return ts.ToString();

        // Handle decimal (preserve precision as string)
        if (value is decimal dec)
            return dec.ToString();

        // Handle byte arrays (blobs)
        if (value is byte[] bytes)
            return $"(blob, {bytes.Length} bytes)";

        // Handle Guid
        if (value is Guid guid)
            return guid.ToString();

        // Primitive types pass through
        if (value is string || value is bool ||
            value is int || value is long || value is short || value is byte ||
            value is uint || value is ulong || value is ushort || value is sbyte ||
            value is float || value is double)
            return value;

        // Fallback: convert to string
        return value.ToString();
    }
}
