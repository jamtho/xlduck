// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Security.Cryptography;

namespace XlDuck.Core;

/// <summary>Compact worksheet aliases. Full result and session identities stay in the engine.</summary>
public static class CellHandle
{
    internal static string NewSessionTag() => Convert.ToBase64String(RandomNumberGenerator.GetBytes(8))
        .TrimEnd('=').Replace('+', '-').Replace('/', '_');

    internal static string Format(ResultKind kind, long number, string tag) =>
        (kind switch { ResultKind.Table => "T", ResultKind.Fragment => "F", ResultKind.Plot => "P", _ => throw new ArgumentOutOfRangeException(nameof(kind)) })
        + number.ToString(CultureInfo.InvariantCulture) + "~" + tag;

    public static bool IsToken([NotNullWhen(true)] string? text)
    {
        if (text is null || text.Length is < 14 or > 32 || text[0] is not ('T' or 'F' or 'P')) return false;
        var separator = text.Length - 12;
        if (text[separator] != '~' || text[1] is < '1' or > '9') return false;
        if (!long.TryParse(text.AsSpan(1, separator - 1), NumberStyles.None, CultureInfo.InvariantCulture, out var number) || number < 1) return false;
        foreach (var character in text.AsSpan(separator + 1))
            if (!char.IsAsciiLetterOrDigit(character) && character is not ('-' or '_')) return false;
        return true;
    }
}
