// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Security.Cryptography;

namespace XlDuck.Core;

/// <summary>Compact worksheet aliases. Full result and session identities stay in the engine.</summary>
public static class CellHandle
{
    private const string Prefix = "xlduck://";

    internal static string NewSessionTag() => Convert.ToBase64String(RandomNumberGenerator.GetBytes(8))
        .TrimEnd('=').Replace('+', '-').Replace('/', '_');

    internal static string Format(ResultKind kind, long number, string tag) =>
        Prefix + (kind switch { ResultKind.Table => "t", ResultKind.Fragment => "f", ResultKind.Plot => "p", _ => throw new ArgumentOutOfRangeException(nameof(kind)) })
        + "/" + number.ToString(CultureInfo.InvariantCulture) + "::" + tag;

    public static bool IsToken([NotNullWhen(true)] string? text)
    {
        if (text is null || text.Length is < 25 or > 43 || !text.StartsWith(Prefix, StringComparison.Ordinal)) return false;
        var body = text.AsSpan(Prefix.Length);
        if (body[0] is not ('t' or 'f' or 'p') || body[1] != '/') return false;
        var separator = body.Length - 13;
        if (body[separator] != ':' || body[separator + 1] != ':' || body[2] is < '1' or > '9') return false;
        if (!long.TryParse(body.Slice(2, separator - 2), NumberStyles.None, CultureInfo.InvariantCulture, out var number) || number < 1) return false;
        foreach (var character in body[(separator + 2)..])
            if (!char.IsAsciiLetterOrDigit(character) && character is not ('-' or '_')) return false;
        return true;
    }
}
