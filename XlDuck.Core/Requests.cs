// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Collections.Immutable;
using System.Security.Cryptography;
using System.Text.Json;

namespace XlDuck.Core;

public enum RequestKind { Query, Capture, Fragment, Plot }
public sealed record CaptureData(ImmutableArray<string> Headers, ImmutableArray<ImmutableArray<CellValue>> Rows)
{
    public void Validate()
    {
        if (Headers.IsDefaultOrEmpty) throw new ArgumentException("A capture needs at least one column.");
        if (Rows.IsDefault) throw new ArgumentException("Capture rows are missing.");
        foreach (var row in Rows)
        {
            if (row.IsDefault || row.Length != Headers.Length) throw new ArgumentException("Capture dimensions do not match the headers.");
            foreach (var value in row) value.Validate();
        }
    }
    public static ImmutableArray<string> UniqueHeaders(IEnumerable<string?> headers)
    {
        var original = headers.Select((name, i) => string.IsNullOrWhiteSpace(name) ? $"Column{i + 1}" : name).ToArray();
        var reserved = original.ToHashSet(StringComparer.OrdinalIgnoreCase);
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var result = ImmutableArray.CreateBuilder<string>();
        foreach (var name in original)
        {
            var candidate = name;
            var suffix = 2;
            while (used.Contains(candidate))
            {
                do { candidate = $"{name}_{suffix++}"; } while (reserved.Contains(candidate));
            }
            used.Add(candidate);
            result.Add(candidate);
        }
        return result.ToImmutable();
    }
}

public sealed record ComputationRequest(RequestKind Kind, string Sql, ImmutableArray<QueryArgument> Arguments,
    CaptureData? Capture = null, string? PlotTemplate = null, ImmutableSortedDictionary<string, string>? PlotOptions = null)
{
    public static ComputationRequest Query(string sql, params QueryArgument[] arguments) => new(RequestKind.Query, sql, [.. arguments]);
    public static ComputationRequest Fragment(string sql, params QueryArgument[] arguments) => new(RequestKind.Fragment, sql, [.. arguments]);
    public static ComputationRequest CaptureRange(CaptureData data) => new(RequestKind.Capture, "", [], data);
    public static ComputationRequest Plot(ResultHandle source, string template, IEnumerable<KeyValuePair<string, string>>? options = null) =>
        new(RequestKind.Plot, "", [QueryArgument.Reference(source)], PlotTemplate: template,
            PlotOptions: (options ?? []).ToImmutableSortedDictionary(StringComparer.Ordinal));
    public string Identity()
    {
        Validate();
        return Convert.ToHexString(SHA256.HashData(JsonSerializer.SerializeToUtf8Bytes(this)));
    }
    public void Validate()
    {
        if (Arguments.IsDefault) throw new ArgumentException("Arguments are missing.");
        foreach (var argument in Arguments) argument.Validate();
        if (Kind is RequestKind.Query or RequestKind.Fragment && string.IsNullOrWhiteSpace(Sql)) throw new ArgumentException("SQL is empty.");
        if (Kind == RequestKind.Capture) (Capture ?? throw new ArgumentException("Capture is missing.")).Validate();
        if (Kind == RequestKind.Plot && (Arguments.Length != 1 || Arguments[0].Handle is null || string.IsNullOrWhiteSpace(PlotTemplate)))
            throw new ArgumentException("A plot needs one source handle and a template.");
        if (!Enum.IsDefined(Kind)) throw new ArgumentException("Unknown computation kind.");
    }
}

public sealed record BoundParameter(string Name, CellValue Value);
public sealed record PreparedQuery(string Sql, ImmutableArray<BoundParameter> Parameters, ImmutableArray<ResultHandle> Tables);
public sealed record ColumnSchema(string Name, string Type);
public sealed record ComputationResult(ResultHandle Handle, ImmutableArray<ColumnSchema> Columns, long RowCount);
public enum ReadPurpose { Output, Preview, Scalar }
public sealed record ReadRequest(ReadPurpose Purpose, int MaxRows = 10000, int MaxCells = 1000000);
public sealed record TablePage(ImmutableArray<ColumnSchema> Columns, ImmutableArray<ImmutableArray<CellValue>> Rows, long TotalRows, bool IsSample);
public sealed record BackendRequest(ResultHandle Output, PreparedQuery? Query = null, CaptureData? Capture = null);

public interface IExecutionBackend : IAsyncDisposable
{
    Task ValidateAsync(PreparedQuery query, CancellationToken cancellationToken);
    Task<ComputationResult> ExecuteAsync(BackendRequest request, CancellationToken cancellationToken);
    Task<TablePage> ReadAsync(ResultHandle handle, ReadRequest request, CancellationToken cancellationToken);
    Task DropAsync(ResultHandle handle);
}
