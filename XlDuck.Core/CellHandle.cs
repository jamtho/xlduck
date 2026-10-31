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

    private const int TagSpace = 26 * 26 * 26 * 26;
    private static readonly HashSet<int> IssuedTags = [];

    // A short display discriminator, not an authentication token. Never recycle a
    // tag in this process, including after a workbook closes or restarts.
    internal static string NewSessionTag()
    {
        int number;
        lock (IssuedTags)
        {
            if (IssuedTags.Count == TagSpace)
                throw new InvalidOperationException("XlDuck session tags are exhausted. Save your work and restart Excel.");
            number = RandomNumberGenerator.GetInt32(TagSpace);
            while (!IssuedTags.Add(number)) number = (number + 1) % TagSpace;
        }
        Span<char> tag = stackalloc char[4];
        for (var i = tag.Length - 1; i >= 0; i--) { tag[i] = (char)('A' + number % 26); number /= 26; }
        return new string(tag);
    }

    internal static string Format(ResultKind kind, long number, string tag) =>
        Prefix + tag + "/" + (kind switch { ResultKind.Table => "t", ResultKind.Fragment => "f", ResultKind.Plot => "p", _ => throw new ArgumentOutOfRangeException(nameof(kind)) })
        + "/" + number.ToString(CultureInfo.InvariantCulture);

    public static bool IsToken([NotNullWhen(true)] string? text)
    {
        if (text is null || text.Length is < 17 or > 35 || !text.StartsWith(Prefix, StringComparison.Ordinal)) return false;
        var body = text.AsSpan(Prefix.Length);
        if (body[4] != '/' || body[5] is not ('t' or 'f' or 'p')) return false;
        if (body[6] != '/' || body[7] is < '1' or > '9') return false;
        foreach (var character in body[..4])
            if (character is < 'A' or > 'Z') return false;
        if (!long.TryParse(body[7..], NumberStyles.None, CultureInfo.InvariantCulture, out var number) || number < 1) return false;
        return true;
    }
}
