// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Data;
using System.Text.Json;
using DuckDB.NET.Data;
using XlDuck.Core;

namespace XlDuck.Engine;

/// <summary>Capability policy for the pinned engine, applied before binding or executing worksheet SQL.</summary>
internal sealed class ReadComputePolicy
{
    internal const string EngineVersion = "v1.5.6";
    private readonly DuckDBConnection _connection;
    private readonly HashSet<string> _expressions = new(StringComparer.OrdinalIgnoreCase);
    private static readonly HashSet<string> SafeVolatile = new(StringComparer.OrdinalIgnoreCase)
        { "random", "uuid", "uuidv4", "uuidv7", "gen_random_uuid", "error" };
    private static readonly HashSet<string> ForbiddenExpressions = new(StringComparer.OrdinalIgnoreCase)
        { "json_serialize_plan", "json_serialize_sql", "json_deserialize_sql", "current_query", "getvariable", "finalize", "combine" };
    private static readonly HashSet<string> Tables = new(StringComparer.OrdinalIgnoreCase)
    {
        "range", "generate_series", "unnest", "repeat", "repeat_row", "json_each", "json_tree",
        "read_csv", "read_csv_auto", "read_parquet", "parquet_scan", "read_json", "read_json_auto",
        "read_json_objects", "read_json_objects_auto", "read_ndjson", "read_ndjson_auto", "read_ndjson_objects",
        "read_text", "read_blob", "glob", "sniff_csv", "parquet_metadata", "parquet_schema",
        "parquet_file_metadata", "parquet_kv_metadata", "parquet_full_metadata", "parquet_bloom_probe",
        "duckdb_tables", "duckdb_columns", "duckdb_types", "duckdb_functions", "duckdb_keywords",
        "duckdb_settings", "duckdb_extensions", "pragma_version", "pragma_platform", "pragma_table_info",
        "pragma_collations", "pg_timezone_names", "icu_calendar_names"
    };

    internal ReadComputePolicy(DuckDBConnection connection)
    {
        _connection = connection;
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT version()";
        if (!string.Equals((string)command.ExecuteScalar()!, EngineVersion, StringComparison.Ordinal))
            throw new InvalidOperationException("This DuckDB version has not been validated for XlDuck's read-and-compute policy.");
        // Snapshot only the pinned built-ins, before loading the explicitly supported reader extensions.
        // Function metadata alone is not the policy: table functions, dynamic binding and macros need separate checks.
        command.CommandText = "SELECT function_name, function_type, has_side_effects, macro_definition FROM duckdb_functions() WHERE internal AND database_name='system' AND schema_name='main'";
        var macros = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        using (var reader = command.ExecuteReader())
        {
            while (reader.Read())
            {
                var name = reader.GetString(0);
                var kind = reader.GetString(1);
                if (ForbiddenExpressions.Contains(name) || name.StartsWith("__", StringComparison.Ordinal)) continue;
                if (kind is "scalar" or "aggregate" && (!reader.GetBoolean(2) || SafeVolatile.Contains(name))) _expressions.Add(name);
                if (kind == "macro")
                {
                    if (!macros.TryGetValue(name, out var definitions)) macros[name] = definitions = [];
                    definitions.Add(reader.GetString(3));
                }
            }
        }
        // Admit scalar macros only when every overload expands entirely into already admitted capabilities.
        bool changed;
        do
        {
            changed = false;
            foreach (var (name, definitions) in macros)
            {
                if (_expressions.Contains(name)) continue;
                try
                {
                    foreach (var definition in definitions) Validate("SELECT " + definition, []);
                    changed |= _expressions.Add(name);
                }
                catch (ArgumentException) { /* Unsupported macro remains unavailable. */ }
            }
        } while (changed);
    }

    internal void Validate(string sql, IEnumerable<ResultHandle> inputs)
    {
        using var command = _connection.CreateCommand();
        command.CommandText = "SELECT json_serialize_sql(CAST($sql AS VARCHAR))";
        command.Parameters.Add(new DuckDBParameter("sql", DbType.String, sql));
        using var parsed = JsonDocument.Parse((string)command.ExecuteScalar()!, new JsonDocumentOptions { MaxDepth = 256 });
        var root = parsed.RootElement;
        if (root.GetProperty("error").GetBoolean() || root.GetProperty("statements").GetArrayLength() != 1)
            throw new ArgumentException("Only one read-and-compute SELECT query is supported. DDL, DML, settings and multiple statements are not allowed. Use explicit IN values for PIVOT.");
        var ownedTables = inputs.Select(input => input.TableName).ToHashSet(StringComparer.OrdinalIgnoreCase);
        Walk(root.GetProperty("statements")[0], new(StringComparer.OrdinalIgnoreCase), ownedTables);
    }

    private void Walk(JsonElement node, HashSet<string> ctes, HashSet<string> ownedTables, bool tableFunction = false)
    {
        if (node.ValueKind == JsonValueKind.Array)
        {
            foreach (var child in node.EnumerateArray()) Walk(child, ctes, ownedTables);
            return;
        }
        if (node.ValueKind != JsonValueKind.Object) return;
        if (node.TryGetProperty("cte_map", out var map))
        {
            ctes = new(ctes, StringComparer.OrdinalIgnoreCase);
            // A nonrecursive definition cannot see itself or later definitions. Treating every
            // name as in scope would let a self-reference fall back to an unowned physical table.
            foreach (var entry in map.GetProperty("map").EnumerateArray())
            {
                Walk(entry.GetProperty("value"), ctes, ownedTables);
                ctes.Add(entry.GetProperty("key").GetString()!);
            }
        }
        var type = Text(node, "type");
        if (type == "BASE_TABLE")
        {
            var name = Text(node, "table_name");
            var schema = Text(node, "schema_name").ToLowerInvariant(); var catalog = Text(node, "catalog_name").ToLowerInvariant();
            if (!(schema == "" && catalog == "" && ctes.Contains(name)) &&
                !(schema is "" or "main" && catalog is "" or "temp" or "memory" && ownedTables.Contains(name)))
                throw new ArgumentException("Reference a table handle or a CTE, or use an explicit supported reader such as read_csv, read_parquet or read_json.");
        }
        if (node.TryGetProperty("function_name", out var functionName))
        {
            var name = functionName.GetString()!;
            var schema = Text(node, "schema").ToLowerInvariant(); var catalog = Text(node, "catalog").ToLowerInvariant();
            if (schema is not ("" or "main") || catalog is not ("" or "system") ||
                !(tableFunction ? Tables.Contains(name) : _expressions.Contains(name)))
                throw new ArgumentException($"Function '{name}' is not available in XlDuck's read-and-compute SQL.");
            if (node.TryGetProperty("export_state", out var export) && export.GetBoolean())
                throw new ArgumentException("Exported aggregate state is not supported.");
            if (tableFunction && name.ToLowerInvariant() is "read_csv" or "read_csv_auto")
                RejectCaptureSideEffects(node);
        }
        foreach (var property in node.EnumerateObject())
        {
            if (property.Name == "cte_map") continue;
            var scope = ctes;
            // Only the recursive term sees the working table; the seed still needs real inputs.
            if (type == "RECURSIVE_CTE_NODE" && property.Name == "right")
                scope = new(ctes, StringComparer.OrdinalIgnoreCase) { Text(node, "cte_name") };
            Walk(property.Value, scope, ownedTables, type == "TABLE_FUNCTION" && property.Name == "function");
        }
    }

    private static void RejectCaptureSideEffects(JsonElement function)
    {
        // CSV reject tables are mutable auxiliary state outside result ownership; no worksheet path creates them.
        foreach (var child in function.GetProperty("children").EnumerateArray())
        {
            var alias = Text(child, "alias").ToLowerInvariant();
            if (alias is "store_rejects" or "rejects_table" or "rejects_scan")
                throw new ArgumentException("CSV reject-table options are unavailable. Clean or inspect the input outside the worksheet query.");
            if (Text(child, "type") == "COMPARE_EQUAL" && child.TryGetProperty("left", out var left) && left.TryGetProperty("column_names", out var names) &&
                names.EnumerateArray().Any(name => name.GetString()?.ToLowerInvariant() is "store_rejects" or "rejects_table" or "rejects_scan"))
                throw new ArgumentException("CSV reject-table options are unavailable. Clean or inspect the input outside the worksheet query.");
        }
    }
    private static string Text(JsonElement node, string name) => node.TryGetProperty(name, out var property) && property.ValueKind == JsonValueKind.String ? property.GetString()! : "";
}
