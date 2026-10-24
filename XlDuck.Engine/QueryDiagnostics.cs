// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Text.RegularExpressions;

namespace XlDuck.Engine;

internal static class QueryDiagnostics
{
    // Remove common engine excerpts, quoted values and signed URLs. This is diagnostic hygiene,
    // not a confidentiality boundary: enabled SQL can read data and deliberately return it.
    internal static string PublicMessage(string message)
    {
        var excerpt = Regex.Match(message, @"(?m)^LINE \d+:", RegexOptions.None, TimeSpan.FromMilliseconds(100));
        if (excerpt.Success) message = message[..excerpt.Index];
        var installationHint = message.IndexOf("Please try installing", StringComparison.OrdinalIgnoreCase);
        if (installationHint >= 0) message = message[..installationHint] + "This capability is not available in XlDuck's supported reader set.";
        message = Regex.Replace(message, @"(?i)\b(?:https?|s3)://[^\s<>""']+", "[remote source]", RegexOptions.None, TimeSpan.FromMilliseconds(100));
        message = Regex.Replace(message, @"'(?:''|[^'])*'", "'[value]'", RegexOptions.None, TimeSpan.FromMilliseconds(100));
        return message.Length > 2000 ? message[..2000] + "…" : message.Trim();
    }
}
