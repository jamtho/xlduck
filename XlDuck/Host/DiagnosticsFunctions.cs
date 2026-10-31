// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Collections.Immutable;
using ExcelDna.Integration;
using XlDuck.Core;

namespace XlDuck;

public static partial class DuckFunctions
{
    private static readonly JsonSerializerOptions DiagnosticJson = new(JsonSerializerDefaults.Web);

    private sealed record ExploreInput(string Op, string Id, string? Sql = null, JsonElement[]? Arguments = null, string? Handle = null, int MaxRows = 20,
        int Index = 0, int Count = 0, string? Text = null);
    [ExcelFunction(IsMacroType = true, IsVolatile = true, Description = "Read-only agent exploration: start or poll a bounded query using a workbook anchor and JSON request. Returns JSON; does not edit worksheet cells.")]
    public static string DuckExplore([ExcelArgument(AllowReference = true)] object anchor, string requestJson)
    {
        var id = ""; var session = Guid.Empty;
        try
        {
            if (requestJson.Length > 30000) throw new ArgumentException("Exploration request exceeds 30000 characters.");
            if (anchor is not ExcelReference reference) throw new ArgumentException("Supply one cell reference in the intended workbook as the anchor.");
            var book = WorkbookHost.Context(reference).Book;
            session = book.Session?.Id ?? Guid.Empty;
            var input = JsonSerializer.Deserialize<ExploreInput>(requestJson, DiagnosticJson) ?? throw new ArgumentException("Supply an exploration request.");
            id = input.Id;
            if (book.Exploration is null)
            {
                if (book.Status is not null) throw new InvalidOperationException(book.Status);
                return JsonSerializer.Serialize(new { version = 1, id, state = "initializing" }, DiagnosticJson);
            }
            if (input.Op == "stage")
            {
                book.ExplorationTransfer!.Add(input.Id, input.Index, input.Count, input.Text ?? throw new ArgumentException("Missing transfer text."));
                return JsonSerializer.Serialize(new { version = 1, id, session, state = "staged" }, DiagnosticJson);
            }
            if (input.Op == "submit")
            {
                var assembled = book.ExplorationTransfer!.Complete(input.Id);
                if (assembled is null)
                {
                    // A repeated submit with no staged text polls the existing operation.
                    if (book.Exploration.TryPoll(input.Id) is { } existing) return ExplorationSession.Serialize(existing);
                    throw new ArgumentException("Exploration transfer expired or was not staged.");
                }
                input = JsonSerializer.Deserialize<ExploreInput>(assembled, DiagnosticJson) ?? throw new ArgumentException("Invalid exploration transfer.");
                if (input.Id != id || input.Op is not ("schema" or "query" or "validate" or "inspect")) throw new ArgumentException("Transfer id or operation does not match.");
            }
            CellValue Value(JsonElement element) => element.ValueKind switch
            {
                JsonValueKind.Null => CellValue.Null,
                JsonValueKind.String => CellValue.From(element.GetString()),
                JsonValueKind.True => CellValue.From(true),
                JsonValueKind.False => CellValue.From(false),
                JsonValueKind.Number when element.TryGetInt64(out var integer) => CellValue.From(integer),
                JsonValueKind.Number => CellValue.From(element.GetDouble()),
                JsonValueKind.Object => element.Deserialize<CellValue>(DiagnosticJson) ?? throw new ArgumentException("Invalid typed argument."),
                _ => throw new ArgumentException("Arguments must be scalars, handle strings or typed {kind,text} values.")
            };
            return ExplorationSession.Serialize(book.Exploration.Invoke(new(input.Op, input.Id, input.Sql,
                input.Arguments?.Select(Value).ToImmutableArray() ?? [], input.Handle, input.MaxRows)));
        }
        catch (Exception error)
        {
            return ExplorationSession.Serialize(new(1, id, session, "failed", Failure: ComputationFailure.From(error)));
        }
    }

    [ExcelFunction(IsMacroType = true, IsVolatile = true, Description = "Inspect one formula cell's execution and Excel display state as JSON. Does not run or refresh its query.")]
    public static string DuckCellStatus([ExcelArgument(AllowReference = true)] object target)
    {
        try
        {
            if (target is not ExcelReference reference) throw new ArgumentException("Pass a cell reference, for example DuckCellStatus(A1).");
            var (book, address, formula) = WorkbookHost.Context(reference);
            var cell = book.Cells.GetValueOrDefault(address);
            if (cell?.Formula != formula) cell = null;
            var bindings = cell?.Bindings.Values.Where(binding => !cell.Touched || binding.Seen).ToArray() ?? [];
            var computations = bindings.Select(binding => new
            {
                computation = book.Session!.GetStatus(binding.Lease.Handle),
                reused = binding.Lease.Reused,
                resultReady = binding.Value.IsCompletedSuccessfully,
                published = binding.Published
            }).ToArray();
            var failure = cell?.ImmediateFailure ?? bindings.Where(binding => binding.Value.IsFaulted)
                .Select(binding => ComputationFailure.From(binding.Value.Exception!)).FirstOrDefault();
            if (book.Session is null && book.Status is not null)
                failure = new("connection_settings", book.Status, "Open XlDuck > Connections to repair the workbook configuration.");
            var cancelled = bindings.Any(binding => binding.Value.IsCanceled);
            var ready = bindings.Length > 0 && bindings.All(binding => binding.Value.IsCompletedSuccessfully);
            var pending = computations.Select(item => item.computation.State).FirstOrDefault(state => state is not ("succeeded" or "failed" or "cancelled"));
            var state = failure is not null ? "failed" : book.Session is null ? "initializing" : cancelled ? "cancelled"
                : pending ?? (bindings.Any(binding => !binding.Value.IsCompleted) ? "reading"
                    : cell?.WaitingInputs == true ? "waiting_inputs"
                    : bindings.Any(binding => !binding.Published) ? "awaiting_calculation"
                    : ready ? "succeeded" : "not_evaluated");
            var range = (Microsoft.Office.Interop.Excel.Range)book.Sheets[address.Sheet].Cells[address.Row + 1, address.Column + 1];
            object value = range.Value2;
            var errorCode = value is int excelError && excelError < 0 ? excelError & 0xffff : 0;
            string? excelState = errorCode switch { 2045 => "spill_error", 2015 => "value_error", 2042 => "na", 0 => null, _ => "excel_error" };
            if (errorCode != 0 && errorCode != 2042)
            {
                state = "failed";
                failure = new(errorCode == 2045 ? "excel_spill" : "excel_error",
                    errorCode == 2045 ? "Excel cannot spill the completed result." : "The formula contains an Excel error.",
                    errorCode == 2045 ? "Clear the spill area or move the output formula." : "Inspect the full Excel formula and its inputs. This is not a running-query marker.");
            }
            else if (errorCode == 2042 && state == "succeeded")
            {
                state = "failed"; failure = new("excel_error", "The formula returned #N/A after its XlDuck computations completed.");
            }
            if (cancelled) failure ??= ComputationFailure.From(new OperationCanceledException());
            var identity = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(
                book.Session?.Id + ":" + formula + ":" + string.Join('|', bindings.Select(binding => binding.Key + binding.Lease.CellText).Order())
                + ":" + cell?.WaitingInputs + ":" + cell?.ImmediateFailure?.Code)));
            var json = JsonSerializer.Serialize(new
            {
                version = 1, state, session = book.Session?.Id, evaluationId = identity,
                resultReady = ready, excelState, failure, computations,
                calculationMode = WorkbookHost.Application.Calculation == Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual ? "manual" : "automatic",
                scope = "last_evaluated_inputs"
            }, DiagnosticJson);
            return json.Length <= 30000 ? json : JsonSerializer.Serialize(new { version = 1, state = "failed",
                failure = new ComputationFailure("output_limit", "This formula has too many computations for one status cell. Inspect smaller parts of its calculation.") }, DiagnosticJson);
        }
        catch (Exception error)
        {
            return JsonSerializer.Serialize(new { version = 1, state = "failed", failure = ComputationFailure.From(error) }, DiagnosticJson);
        }
    }

    [ExcelFunction(IsMacroType = true, Description = "Bind a SELECT and spill its column names and types without retaining its data rows. Reader schema discovery may access the source.")]
    public static object DuckSchema(object sql, object arg1 = null!, object arg2 = null!, object arg3 = null!, object arg4 = null!, object arg5 = null!, object arg6 = null!, object arg7 = null!, object arg8 = null!)
    {
        try
        {
            var args = new List<object> { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8 };
            while (args.Count > 0 && args[^1] is ExcelMissing or null) args.RemoveAt(args.Count - 1);
            foreach (var input in args.Prepend(sql))
                if (input is ExcelError || input is string error && error.StartsWith("#ERROR:", StringComparison.Ordinal)) return Propagate(input);
            if (sql is not string statement) throw new ArgumentException("SQL must be text.");
            var request = ComputationRequest.Query(statement, [.. args.Select(Argument)]) with { RowLimit = 0 };
            var key = "schema:" + request.Identity();
            return Evaluate(key, book =>
            {
                var lease = book.Session!.Acquire(request);
                async Task<object> Schema()
                {
                    var result = await lease.Completion.ConfigureAwait(false);
                    var schema = new object[result.Columns.Length + 1, 2]; schema[0, 0] = "Column"; schema[0, 1] = "Type";
                    for (var i = 0; i < result.Columns.Length; i++) { schema[i + 1, 0] = result.Columns[i].Name; schema[i + 1, 1] = result.Columns[i].Type; }
                    return schema;
                }
                return new(key, lease, Schema());
            });
        }
        catch (Exception error) { return Fail(error); }
    }
}
