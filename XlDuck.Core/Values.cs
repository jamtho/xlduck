// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Globalization;
using System.Numerics;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace XlDuck.Core;

[JsonConverter(typeof(JsonStringEnumConverter<ValueKind>))]
public enum ValueKind { Null, Empty, Text, Boolean, Number, Integer, Decimal, Date, Time, Timestamp, TimestampOffset, Interval, Json, Error }

/// <summary>Lossless transport value. Text is canonical and culture independent; it is never inferred from cell text.</summary>
public sealed record CellValue(ValueKind Kind, string Text)
{
    public static readonly CellValue Null = new(ValueKind.Null, "");
    public static readonly CellValue Empty = new(ValueKind.Empty, "");
    public static CellValue String(string value) => new(ValueKind.Text, value);
    public static CellValue Number(double value) => double.IsFinite(value)
        ? new(ValueKind.Number, value.ToString("R", CultureInfo.InvariantCulture))
        : throw new ArgumentException("Non-finite numbers cannot be represented in Excel. Cast to VARCHAR explicitly.");
    public static CellValue From(object? value) => value switch
    {
        null or DBNull => Null,
        string text => String(text),
        bool flag => new(ValueKind.Boolean, flag ? "true" : "false"),
        double number => Number(number),
        float number => Number(number),
        sbyte or byte or short or ushort or int or uint or long or ulong or BigInteger =>
            new(ValueKind.Integer, Convert.ToString(value, CultureInfo.InvariantCulture)!),
        decimal number => new(ValueKind.Decimal, number.ToString(CultureInfo.InvariantCulture)),
        DateOnly date => new(ValueKind.Date, date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)),
        TimeOnly time => new(ValueKind.Time, time.ToString("HH:mm:ss.fffffff", CultureInfo.InvariantCulture)),
        DateTime timestamp => new(ValueKind.Timestamp, timestamp.ToString("yyyy-MM-ddTHH:mm:ss.fffffff", CultureInfo.InvariantCulture)),
        DateTimeOffset timestamp => new(ValueKind.TimestampOffset, timestamp.ToString("O", CultureInfo.InvariantCulture)),
        _ => throw new NotSupportedException($"Unsupported value type {value.GetType().Name}. Convert it explicitly in SQL.")
    };

    /// <summary>Excel-facing value, excluding host-specific blank/error tokens.</summary>
    public object? ToDisplayValue() => Kind switch
    {
        ValueKind.Null or ValueKind.Empty => null,
        ValueKind.Boolean => bool.Parse(Text),
        ValueKind.Number => double.Parse(Text, CultureInfo.InvariantCulture),
        ValueKind.Integer when BigInteger.Abs(BigInteger.Parse(Text, CultureInfo.InvariantCulture)) <= 9007199254740991L =>
            (double)BigInteger.Parse(Text, CultureInfo.InvariantCulture),
        // Exact decimals and large integers stay text. CAST AS DOUBLE explicitly opts into approximation.
        ValueKind.Date or ValueKind.Timestamp => DisplayDate(Text),
        ValueKind.Time => TimeOnly.Parse(Text, CultureInfo.InvariantCulture).ToTimeSpan().TotalDays,
        _ => Text
    };

    public void Validate()
    {
        ArgumentNullException.ThrowIfNull(Text);
        switch (Kind)
        {
            case ValueKind.Null or ValueKind.Empty when Text.Length != 0: throw new ArgumentException("Blank values cannot carry text.");
            case ValueKind.Boolean when Text is not ("true" or "false"): throw new ArgumentException("Invalid boolean.");
            case ValueKind.Number: _ = Number(double.Parse(Text, CultureInfo.InvariantCulture)); break;
            case ValueKind.Integer: _ = BigInteger.Parse(Text, CultureInfo.InvariantCulture); break;
            case ValueKind.Decimal: _ = DecimalScale(); break;
            case ValueKind.Json: using (JsonDocument.Parse(Text)) { } break;
            case ValueKind.Error: throw new ArgumentException($"Input contains an Excel error: {Text}");
        }
        if (!Enum.IsDefined(Kind)) throw new ArgumentException("Unknown value kind.");
    }
    public int DecimalScale()
    {
        var digits = Text.StartsWith('-') || Text.StartsWith('+') ? Text[1..] : Text;
        var parts = digits.Split('.');
        if (parts.Length > 2 || parts[0].Length == 0 || parts.Any(part => part.Length == 0 || part.Any(c => !char.IsAsciiDigit(c))))
            throw new ArgumentException("Invalid exact decimal value.");
        var scale = parts.Length == 2 ? parts[1].Length : 0;
        if (parts[0].TrimStart('0').Length + scale > 38 || scale > 38)
            throw new ArgumentException("Decimal exceeds DuckDB's 38-digit precision. Keep it as text or convert explicitly.");
        return scale;
    }
    private static object DisplayDate(string text)
    {
        if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
            try { return date.ToOADate(); } catch (OverflowException) { }
        return text;
    }
}

public enum ResultKind { Table, Fragment, Plot }
public readonly record struct ResultHandle(Guid Session, Guid Id, ResultKind Kind)
{
    public override string ToString() => $"xlduck://{Session:N}/{Kind.ToString().ToLowerInvariant()}/{Id:N}";
    public static ResultHandle Parse(string text)
    {
        if (!text.StartsWith("xlduck://", StringComparison.Ordinal)) throw new ArgumentException("Invalid XlDuck handle.");
        var parts = text[9..].Split('/');
        if (parts.Length != 3 || !Guid.TryParseExact(parts[0], "N", out var session) ||
            !Guid.TryParseExact(parts[2], "N", out var id) || !Enum.TryParse<ResultKind>(parts[1], true, out var kind) || !Enum.IsDefined(kind))
            throw new ArgumentException("Invalid XlDuck handle.");
        return new(session, id, kind);
    }
    public string TableName => "_xlduck_" + Id.ToString("N");
}

public sealed record QueryArgument(CellValue? Value = null, ResultHandle? Handle = null)
{
    public static QueryArgument Scalar(object? value) => new(CellValue.From(value));
    public static QueryArgument Reference(ResultHandle handle) => new(Handle: handle);
    public void Validate()
    {
        if ((Value is null) == (Handle is null)) throw new ArgumentException("An argument must contain one typed value or one handle.");
        Value?.Validate();
    }
}
