// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Globalization;

namespace XlDuck.ExecutionPrototype;

// Deliberately narrow transport for measuring execution, not the final public query API.
public sealed record Atom(string Kind, string Text)
{
    public static Atom From(object? value) => value switch
    {
        null or DBNull => new("null", ""),
        string text => new("text", text),
        bool flag => new("boolean", flag ? "true" : "false"),
        double or float or int or short or byte => new("number", Convert.ToString(value, CultureInfo.InvariantCulture)!),
        long integer when Math.Abs((decimal)integer) <= 9007199254740991m => new("number", integer.ToString(CultureInfo.InvariantCulture)),
        _ => throw new NotSupportedException($"Prototype does not support {value.GetType().Name}.")
    };
    public object ToObject() => Kind switch
    {
        "null" => DBNull.Value,
        "text" => Text,
        "boolean" => bool.Parse(Text),
        "number" => double.Parse(Text, CultureInfo.InvariantCulture),
        _ => throw new InvalidOperationException($"Unknown atom kind {Kind}.")
    };
}

public sealed record Request(Guid Id, string Operation, string Sql = "", string Input = "",
    Atom? Value = null, int Limit = 100, Atom[][]? Capture = null);
public sealed record Result(Guid Id, string Handle = "", string[]? Columns = null,
    Atom[][]? Rows = null, long RowCount = 0, double EngineMs = 0, string? Error = null);

public interface IBackend : IAsyncDisposable
{
    Task<Result> ExecuteAsync(Request request, CancellationToken cancellationToken);
}
