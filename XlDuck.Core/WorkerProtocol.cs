// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Buffers.Binary;
using System.Text.Json;

namespace XlDuck.Core;

public enum WorkerOperation { Execute, Read, Drop, Cancel, Inspect, Validate }
public sealed record WorkerStartup(ConnectionSettings Settings, int Version = WorkerProtocol.Version);
public sealed record WorkerRequest(Guid Id, WorkerOperation Operation, BackendRequest? Execution = null, ResultHandle? Handle = null, ReadRequest? Read = null, int Version = WorkerProtocol.Version, PreparedQuery? Query = null);
public sealed record WorkerResponse(Guid Id, ComputationResult? Result = null, TablePage? Page = null, string? Error = null, bool Cancelled = false, bool Ready = false, int Version = WorkerProtocol.Version, int? StoredTables = null);

/// <summary>Length-prefixed UTF-8 frames keep malformed or oversized transfers from growing without bounds.</summary>
public static class WorkerProtocol
{
    public const int Version = 2;
    public const int MaxFrameBytes = 32 * 1024 * 1024;
    public static async Task WriteAsync<T>(Stream stream, T message, CancellationToken token = default)
    {
        var payload = JsonSerializer.SerializeToUtf8Bytes(message);
        if (payload.Length > MaxFrameBytes) throw new InvalidOperationException("Worker transfer exceeds 32 MiB. Reduce the capture or output size.");
        var header = new byte[4];
        BinaryPrimitives.WriteInt32LittleEndian(header, payload.Length);
        await stream.WriteAsync(header, token).ConfigureAwait(false);
        await stream.WriteAsync(payload, token).ConfigureAwait(false);
        await stream.FlushAsync(token).ConfigureAwait(false);
    }
    public static async Task<T?> ReadAsync<T>(Stream stream, CancellationToken token = default) where T : class
    {
        var header = new byte[4];
        var count = await stream.ReadAsync(header.AsMemory(0, 1), token).ConfigureAwait(false);
        if (count == 0) return null;
        await stream.ReadExactlyAsync(header.AsMemory(1), token).ConfigureAwait(false);
        var length = BinaryPrimitives.ReadInt32LittleEndian(header);
        if (length < 1 || length > MaxFrameBytes) throw new IOException("Invalid worker frame size.");
        var payload = new byte[length];
        await stream.ReadExactlyAsync(payload, token).ConfigureAwait(false);
        return JsonSerializer.Deserialize<T>(payload) ?? throw new IOException("Invalid worker message.");
    }
}
