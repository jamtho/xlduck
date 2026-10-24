// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Collections.Immutable;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using XlDuck.Core;

namespace XlDuck.Host;

public sealed record WorkbookSettings(ImmutableArray<S3Connection> S3 = default)
{
    public ImmutableArray<S3Connection> S3 { get; init; } = S3.IsDefault ? [] : S3;
    public ConnectionSettings ForSession(string readerDirectory) => new(true, readerDirectory, S3);
    public void Validate() => ForSession(Path.GetFullPath(".")).Validate();
}

/// <summary>User-local approval for a saved workbook revision. Copies and external edits do not inherit approval.</summary>
public sealed class WorkbookApprovalStore(string directory)
{
    private sealed record Approval(int Version, string Path, string Fingerprint, WorkbookSettings Settings);
    private readonly string _directory = Path.GetFullPath(directory);
    private string RecordPath(string path) => System.IO.Path.Combine(_directory,
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(Canonical(path)))) + ".json");
    private static string Canonical(string path) => System.IO.Path.GetFullPath(path).ToUpperInvariant();

    public async Task<WorkbookSettings?> ReadAsync(string path, CancellationToken token = default)
    {
        try
        {
            var record = RecordPath(path);
            if (!File.Exists(record) || new FileInfo(record).Length > 64 * 1024) return null;
            var approval = JsonSerializer.Deserialize<Approval>(await File.ReadAllTextAsync(record, token).ConfigureAwait(false));
            if (approval is null || approval.Version != 1 || approval.Path != Canonical(path) || approval.Settings is null) return null;
            approval.Settings.Validate();
            return approval.Fingerprint == await FingerprintAsync(path, token).ConfigureAwait(false) ? approval.Settings : null;
        }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException or JsonException or ArgumentException) { return null; }
    }

    public async Task RememberAsync(string path, WorkbookSettings settings, CancellationToken token = default)
    {
        settings.Validate();
        var approval = new Approval(1, Canonical(path), await FingerprintAsync(path, token).ConfigureAwait(false), settings);
        Directory.CreateDirectory(_directory);
        var record = RecordPath(path);
        var temporary = record + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            await File.WriteAllTextAsync(temporary, JsonSerializer.Serialize(approval), token).ConfigureAwait(false);
            token.ThrowIfCancellationRequested();
            File.Move(temporary, record, overwrite: true);
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }

    public void Forget(string path) => File.Delete(RecordPath(path));

    private static async Task<string> FingerprintAsync(string path, CancellationToken token)
    {
        var before = new FileInfo(path);
        var length = before.Length; var modified = before.LastWriteTimeUtc;
        await using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete,
            128 * 1024, FileOptions.Asynchronous | FileOptions.SequentialScan);
        var hash = await SHA256.HashDataAsync(stream, token).ConfigureAwait(false);
        var after = new FileInfo(path);
        if (after.Length != length || after.LastWriteTimeUtc != modified) throw new IOException("The workbook changed while recording approval. Save it again.");
        return Convert.ToHexString(hash);
    }
}
