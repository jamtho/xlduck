// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Text.RegularExpressions;
using XlDuck.Core;

namespace XlDuck.Engine;

internal static class QueryDiagnostics
{
    internal static ComputationFailure Failure(string message)
    {
        // Classify inside the engine. Callers must not parse engine prose to decide what to do.
        var text = message.ToLowerInvariant();
        if (text.Contains("max_temp_directory_size") || text.Contains("no space left") || text.Contains("disk full"))
            return new("spill_limit", "The query exhausted its temporary disk budget or available disk space.",
                "Project fewer columns and reduce large sorts or joins. Review this workbook's Resources settings if more disk is appropriate.");
        if (text.Contains("out of memory") || text.Contains("failed to allocate") || text.Contains("failed to pin block"))
            return new("memory_limit", "The query exceeded its memory budget.",
                "Inspect the schema first; select metadata columns, filter or aggregate early, then fetch a few text excerpts by ID. LIMIT alone does not bound scan or sort memory. Review this workbook's Resources settings if more memory is appropriate.");
        var code = text.StartsWith("parser error") ? "sql_validation"
            : text.StartsWith("binder error") || text.StartsWith("catalog error") ? "sql_binding"
            : text.StartsWith("conversion error") || text.StartsWith("invalid input error") ? "data_conversion"
            : text.StartsWith("http error") || text.StartsWith("io error") || text.Contains("http exception") ? "source_access"
            : "execution";
        return new(code, PublicMessage(message));
    }
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
