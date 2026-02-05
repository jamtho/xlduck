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

        // Check for plot handles
        if (PlotStore.IsHandle(cellValue))
        {
            return GetPlotPreview(cellValue);
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
                    var row = new string?[fieldCount];
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

    private const int PlotRowLimit = 50_000;
    private const int PlotCellLimit = 500_000;

    private static PreviewModel GetPlotPreview(string handle)
    {
        var plot = PlotStore.Get(handle);
        if (plot == null)
        {
            return new ErrorPreviewModel
            {
                Title = "Plot Not Found",
                Handle = handle,
                Message = "This plot may have been released"
            };
        }

        var plotData = new PlotPreviewData
        {
            Template = plot.Template,
            Overrides = plot.Overrides
        };

        try
        {
            // Resolve data handle
            string duckTableName;
            long rowCount;
            string[] columnNames;

            if (ResultStore.IsHandle(plot.DataHandle))
            {
                var stored = ResultStore.Get(plot.DataHandle);
                if (stored == null)
                {
                    plotData.Error = "Data handle not found - it may have been released";
                    return new PlotPreviewModel
                    {
                        Title = $"Plot ({plot.Template})",
                        Handle = handle,
                        Plot = plotData
                    };
                }
                duckTableName = stored.DuckTableName;
                rowCount = stored.RowCount;
                columnNames = stored.ColumnNames;
            }
            else if (FragmentStore.IsHandle(plot.DataHandle))
            {
                // For fragments, we need to materialize to get data
                // This is more complex - for now, return an error suggesting to use table handles
                plotData.Error = "Fragment handles for plots not yet supported. Use DuckQuery to materialize first.";
                return new PlotPreviewModel
                {
                    Title = $"Plot ({plot.Template})",
                    Handle = handle,
                    Plot = plotData
                };
            }
            else
            {
                plotData.Error = $"Invalid data handle: {plot.DataHandle}";
                return new PlotPreviewModel
                {
                    Title = $"Plot ({plot.Template})",
                    Handle = handle,
                    Plot = plotData
                };
            }

            plotData.RowCount = rowCount;

            // Check data caps
            if (rowCount > PlotRowLimit)
            {
                plotData.Error = $"Dataset too large for plotting ({rowCount:N0} rows). Maximum: {PlotRowLimit:N0} rows. Use DuckQuery to aggregate or filter your data.";
                return new PlotPreviewModel
                {
                    Title = $"Plot ({plot.Template})",
                    Handle = handle,
                    Plot = plotData
                };
            }

            var totalCells = rowCount * columnNames.Length;
            if (totalCells > PlotCellLimit)
            {
                plotData.Error = $"Dataset too large for plotting ({totalCells:N0} cells). Maximum: {PlotCellLimit:N0} cells.";
                return new PlotPreviewModel
                {
                    Title = $"Plot ({plot.Template})",
                    Handle = handle,
                    Plot = plotData
                };
            }

            var conn = DuckFunctions.GetConnection();

            // Get column types via PRAGMA table_info
            var columnTypes = new List<string>();
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = $"PRAGMA table_info('{duckTableName}')";
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    columnTypes.Add(reader.GetString(reader.GetOrdinal("type")));
                }
            }

            plotData.Columns = columnNames.ToList();
            plotData.Types = columnTypes;

            // Get all rows (within limit)
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = $"SELECT * FROM \"{duckTableName}\" LIMIT {PlotRowLimit}";
                using var reader = cmd.ExecuteReader();

                var fieldCount = reader.FieldCount;
                while (reader.Read())
                {
                    var row = new string?[fieldCount];
                    for (int i = 0; i < fieldCount; i++)
                    {
                        row[i] = reader.IsDBNull(i) ? null : ConvertForJson(reader.GetValue(i));
                    }
                    plotData.Rows.Add(row);
                }
            }

            return new PlotPreviewModel
            {
                Title = $"Plot ({plot.Template}, {rowCount:N0} rows)",
                Handle = handle,
                Plot = plotData
            };
        }
        catch (Exception ex)
        {
            plotData.Error = ex.Message;
            return new PlotPreviewModel
            {
                Title = $"Plot ({plot.Template})",
                Handle = handle,
                Plot = plotData
            };
        }
    }

    /// <summary>
    /// Convert a DuckDB value to a string for JSON serialization.
    /// </summary>
    private static string ConvertForJson(object value)
    {
        // Special formatting for certain types
        return value switch
        {
            bool b => b ? "true" : "false",
            DateTime dt => dt.ToString("O"),
            DateTimeOffset dto => dto.ToString("O"),
            byte[] bytes => $"(blob, {bytes.Length} bytes)",
            _ => value.ToString() ?? ""
        };
    }
}
