// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Text.Json;
using ExcelDna.Integration;
using XlDuck.Core;
using XlDuck.Host;

namespace XlDuck;

public static partial class DuckFunctions
{
    [ExcelFunction(IsMacroType = true, IsVolatile = true, Description = "List this workbook's configured S3 scopes, endpoints and authentication methods as JSON. No credentials or source reads. Supply an anchor cell when using Evaluate.")]
    public static string DuckConnections(
        [ExcelArgument(AllowReference = true, Description = "Optional in a worksheet formula; required for Evaluate. One cell in the intended workbook.")] object anchor = null!,
        [ExcelArgument(Description = "Optional zero-based connection offset; use nextOffset if the JSON result has another page.")] int offset = 0)
    {
        try
        {
            var book = anchor switch
            {
                null or ExcelMissing => WorkbookHost.Caller().Book,
                ExcelReference reference => WorkbookHost.Context(reference).Book,
                _ => throw new ArgumentException("Supply one cell reference in the intended workbook, for example DuckConnections(A1).")
            };
            if (book.Session is null)
            {
                if (book.Status is not null) throw new InvalidOperationException(book.Status);
                return JsonSerializer.Serialize(new { version = 1, state = "initializing" }, DiagnosticJson);
            }
            return ConnectionDiscovery.Describe(book.Settings, offset);
        }
        catch (Exception error)
        {
            return JsonSerializer.Serialize(new { version = 1, state = "failed", failure = ComputationFailure.From(error) }, DiagnosticJson);
        }
    }
}
