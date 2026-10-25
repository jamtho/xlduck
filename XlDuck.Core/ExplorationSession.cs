// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Collections.Immutable;
using System.Text.Json;

namespace XlDuck.Core;

public sealed record ExplorationRequest(string Op, string Id, string? Sql = null, ImmutableArray<CellValue> Arguments = default,
    string? Handle = null, int MaxRows = 20)
{
    public ImmutableArray<CellValue> Arguments { get; init; } = Arguments.IsDefault ? [] : Arguments;
}
public sealed record ExplorationResult(ImmutableArray<ColumnSchema> Columns, ImmutableArray<ImmutableArray<CellValue>> Rows,
    bool HasMore = false, string? Validation = null, ComputationRequest? Definition = null);
public sealed record ExplorationStatus(int Version, string Id, Guid Session, string State, ExplorationResult? Result = null,
    ComputationFailure? Failure = null, ComputationStatus? Computation = null);

/// <summary>Bounded, read-only exploration in the existing session. No workbook/COM access.</summary>
public sealed class ExplorationSession(WorkbookSession session, TimeSpan? timeout = null) : IDisposable
{
    private sealed class Operation(ExplorationRequest request)
    {
        public readonly ExplorationRequest Request = request;
        public readonly string Identity = JsonSerializer.Serialize(request);
        public readonly CancellationTokenSource Cancellation = new();
        public DateTimeOffset LastAccess = DateTimeOffset.UtcNow;
        public ResultHandle? Handle;
        public Task<ExplorationStatus> Work = null!;
    }
    private readonly object _sync = new();
    private readonly Dictionary<string, Operation> _operations = new(StringComparer.Ordinal);
    private bool _disposed;
    private readonly TimeSpan _timeout = timeout ?? TimeSpan.FromSeconds(30);
    private static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web);

    public ExplorationStatus? TryPoll(string id)
    {
        lock (_sync)
        {
            ObjectDisposedException.ThrowIf(_disposed, this); Prune();
            if (!_operations.TryGetValue(id, out var operation)) return null;
            operation.LastAccess = DateTimeOffset.UtcNow;
            return Status(operation);
        }
    }

    public ExplorationStatus Invoke(ExplorationRequest request)
    {
        lock (_sync)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (string.IsNullOrWhiteSpace(request.Id) || request.Id.Length > 64 || request.Id.Any(c => !char.IsAsciiLetterOrDigit(c) && c != '-'))
                throw new ArgumentException("Supply a unique request id of up to 64 letters, digits or hyphens.");
            Prune();
            if (_operations.TryGetValue(request.Id, out var existing))
            {
                if (request.Op != "poll" && existing.Identity != JsonSerializer.Serialize(request))
                    throw new ArgumentException("This request id is already in use with different inputs. Use a new id.");
                existing.LastAccess = DateTimeOffset.UtcNow;
                return Status(existing);
            }
            if (request.Op == "poll") throw new ArgumentException("The exploration request expired or belongs to an earlier session. Start it again with a new id.");
            if (request.Op is not ("schema" or "query" or "validate" or "inspect")) throw new ArgumentException("Choose schema, query, validate, inspect or poll.");
            if (request.MaxRows is < 1 or > 200) throw new ArgumentException("Exploration maxRows must be between 1 and 200.");
            if (request.Sql?.Length > 24000 || !request.Arguments.IsDefault && request.Arguments.Length > 8) throw new ArgumentException("Exploration accepts at most 24000 SQL characters and eight arguments.");
            if (_operations.Count >= 8)
            {
                var oldest = _operations.Values.Where(operation => operation.Work.IsCompleted).MinBy(operation => operation.LastAccess);
                if (oldest is null) throw new ComputationException(new("queue_full", "Eight exploration requests are already running.", "Poll those requests before starting more."));
                Remove(oldest);
            }
            var operation = new Operation(request);
            _operations.Add(request.Id, operation);
            operation.Work = RunAsync(operation);
            return Status(operation);
        }
    }
    public static string Serialize(ExplorationStatus status)
    {
        var json = JsonSerializer.Serialize(status, Json);
        if (json.Length <= 30000) return json;
        return JsonSerializer.Serialize(new ExplorationStatus(1, status.Id, status.Session, "failed", Failure:
            new("output_limit", "Exploration output exceeds 30000 JSON characters.", "Select fewer columns or rows, and shorten text excerpts explicitly with left(body, 500).")), Json);
    }
    private ExplorationStatus Status(Operation operation)
    {
        if (operation.Work.IsCompletedSuccessfully) return operation.Work.Result;
        var computation = GetStatus(operation.Handle);
        var state = computation?.State is "succeeded" ? "reading" : computation?.State ?? "queued";
        return new(1, operation.Request.Id, session.Id, state, Computation: computation);
    }
    private ComputationStatus? GetStatus(ResultHandle? handle)
    {
        try { return handle is { } value ? session.GetStatus(value) : null; }
        catch (ArgumentException) { return null; } // Completed exploration may no longer be retained.
    }
    private async Task<ExplorationStatus> RunAsync(Operation operation)
    {
        var request = operation.Request;
        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(operation.Cancellation.Token);
        deadline.CancelAfter(_timeout);
        try
        {
            ResultLease lease;
            if (request.Op == "inspect")
                lease = session.AcquireExisting(session.ResolveCellHandle(request.Handle ?? throw new ArgumentException("Supply a retained handle to inspect.")));
            else
            {
                var arguments = request.Arguments.IsDefault ? [] : request.Arguments.Select(value =>
                {
                    value.Validate();
                    return value.Kind == ValueKind.Text && CellHandle.IsToken(value.Text)
                        ? QueryArgument.Reference(session.ResolveCellHandle(value.Text)) : new QueryArgument(value);
                }).ToArray();
                var computation = request.Op == "validate" ? ComputationRequest.Fragment(request.Sql ?? "", arguments)
                    : ComputationRequest.Query(request.Sql ?? "", arguments) with { RowLimit = request.Op == "schema" ? 0 : request.MaxRows + 1 };
                lease = session.Acquire(computation);
            }
            using (lease)
            {
                operation.Handle = lease.Handle;
                var result = await lease.WaitAsync(deadline.Token).ConfigureAwait(false);
                ExplorationResult data;
                if (request.Op == "query")
                {
                    var page = await session.ReadAsync(lease.Handle, new(ReadPurpose.Preview, request.MaxRows, 4000), deadline.Token).ConfigureAwait(false);
                    data = new(page.Columns, page.Rows, page.IsSample);
                }
                else data = new(result.Columns, [], Validation: request.Op == "validate" ? "syntax_and_capabilities_only" : null,
                    Definition: request.Op == "inspect" ? session.GetDefinition(lease.Handle) : null);
                // A fragment inspection returns its recipe only; it never silently materializes the source.
                var status = new ExplorationStatus(1, request.Id, session.Id, "succeeded", data, Computation: session.GetStatus(lease.Handle));
                if (JsonSerializer.Serialize(status, Json).Length > 30000)
                    throw new ComputationException(new("output_limit", "Exploration output exceeds 30000 JSON characters.", "Select fewer columns or rows, and shorten text excerpts explicitly with left(body, 500)."));
                return status;
            }
        }
        catch (OperationCanceledException)
        {
            var failure = operation.Cancellation.IsCancellationRequested || session.IsClosed || !deadline.IsCancellationRequested
                ? ComputationFailure.From(new OperationCanceledException())
                : new ComputationFailure("timeout", $"Exploration exceeded its {_timeout.TotalSeconds:g}-second deadline.", "Reduce the query or prepare the expensive intermediate result outside Excel.");
            return new(1, request.Id, session.Id, failure.Code == "timeout" ? "failed" : "cancelled", Failure: failure);
        }
        catch (Exception error) { return new(1, request.Id, session.Id, "failed", Failure: ComputationFailure.From(error), Computation: GetStatus(operation.Handle)); }
    }
    private void Prune()
    {
        foreach (var operation in _operations.Values.Where(operation => DateTimeOffset.UtcNow - operation.LastAccess > TimeSpan.FromMinutes(2)).ToArray()) Remove(operation);
    }
    private void Remove(Operation operation)
    {
        _operations.Remove(operation.Request.Id); operation.Cancellation.Cancel();
        _ = operation.Work.ContinueWith(_ => operation.Cancellation.Dispose(), TaskScheduler.Default);
    }
    public void Dispose()
    {
        lock (_sync)
        {
            if (_disposed) return;
            _disposed = true;
            foreach (var operation in _operations.Values.ToArray()) Remove(operation);
        }
    }
}
