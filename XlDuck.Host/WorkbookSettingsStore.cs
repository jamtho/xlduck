// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Collections.Immutable;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using XlDuck.Core;

namespace XlDuck.Host;

public sealed record WorkbookSettings(ImmutableArray<S3Connection> S3 = default, ResourceSettings? Resources = null)
{
    public ResourceSettings Resources { get; init; } = Resources ?? new();
    public ImmutableArray<S3Connection> S3 { get; init; } = S3.IsDefault ? [] : S3;
    public ConnectionSettings ForSession(string readerDirectory, Func<string, S3AccessKey>? loadKey = null)
    {
        if (S3.Any(connection => connection.Credentials == S3CredentialSource.AccessKey && connection.CredentialId is null))
            throw new InvalidOperationException("A workbook connection has no saved key on this computer for its endpoint. Open XlDuck > Connections and enter its access key.");
        var keys = S3.Where(connection => connection.Credentials == S3CredentialSource.AccessKey)
            .Select(connection => connection.CredentialId!).Distinct(StringComparer.Ordinal)
            .ToImmutableDictionary(id => id, id => (loadKey ?? throw new InvalidOperationException("An access-key store is required."))(id));
        return new(true, readerDirectory, S3, keys, Resources);
    }
    public void Validate()
    {
        (Resources ?? throw new ArgumentException("Resource settings are missing.")).Validate();
        if (S3.IsDefault || S3.Length > 32) throw new ArgumentException("At most 32 S3 connections are supported per workbook.");
        foreach (var connection in S3)
        {
            if (connection is null) throw new ArgumentException("Connection is missing.");
            // Portable definitions can exist before this Windows account supplies a key.
            (connection.Credentials == S3CredentialSource.AccessKey && connection.CredentialId is null
                ? connection with { CredentialId = Guid.Empty.ToString("N") } : connection).Validate();
        }
        if (S3.Select(connection => connection.Scope).Distinct(StringComparer.Ordinal).Count() != S3.Length)
            throw new ArgumentException("S3 connection scopes must be unique.");
        var ids = S3.Where(connection => connection.ConnectionId is not null).Select(connection => connection.ConnectionId!);
        if (ids.Distinct(StringComparer.OrdinalIgnoreCase).Count() != ids.Count()) throw new ArgumentException("Connection IDs must be unique within a workbook.");
    }
    public WorkbookSettings WithConnectionIds() => this with { S3 = [.. S3.Select(connection => connection with
    {
        ConnectionId = connection.ConnectionId?.ToLowerInvariant() ?? Guid.NewGuid().ToString("N"),
        Name = string.IsNullOrWhiteSpace(connection.Name) ? connection.Scope[..Math.Min(120, connection.Scope.Length)] : connection.Name
    })] };
}

/// <summary>Legacy path-based settings, retained only for migration and recovery.</summary>
public sealed class WorkbookSettingsStore(string directory)
{
    private sealed record SavedSettings(int Version, string Path, WorkbookSettings Settings);
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
            var saved = JsonSerializer.Deserialize<SavedSettings>(await File.ReadAllTextAsync(record, token).ConfigureAwait(false));
            if (saved is null || saved.Version != 1 || saved.Path != Canonical(path) || saved.Settings is null) return null;
            saved.Settings.Validate();
            return saved.Settings;
        }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException or JsonException or ArgumentException) { return null; }
    }

    public async Task SaveAsync(string path, WorkbookSettings settings, CancellationToken token = default)
    {
        settings.Validate();
        var saved = new SavedSettings(1, Canonical(path), settings);
        Directory.CreateDirectory(_directory);
        var record = RecordPath(path);
        var temporary = record + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            await File.WriteAllTextAsync(temporary, JsonSerializer.Serialize(saved), token).ConfigureAwait(false);
            token.ThrowIfCancellationRequested();
            File.Move(temporary, record, overwrite: true);
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }

    public void Forget(string path) => File.Delete(RecordPath(path));

    public async Task<WorkbookSettings?> ReadForMigrationAsync(string path, CancellationToken token = default)
    {
        if (string.IsNullOrEmpty(path)) return null;
        if (!Uri.TryCreate(path, UriKind.Absolute, out var uri) || uri.Scheme is not ("https" or "http"))
            return await ReadAsync(path, token).ConfigureAwait(false);
        // Earlier versions passed OneDrive URLs to GetFullPath, prepending Excel's
        // working directory. Match the complete legacy URL suffix, never a filename.
        if (!Directory.Exists(_directory)) return null;
        var suffix = path.Replace("://", ":/").Replace('/', '\\').ToUpperInvariant();
        var matches = new List<WorkbookSettings>();
        foreach (var record in Directory.EnumerateFiles(_directory, "*.json"))
        {
            token.ThrowIfCancellationRequested();
            try
            {
                if (new FileInfo(record).Length > 64 * 1024) continue;
                var saved = JsonSerializer.Deserialize<SavedSettings>(await File.ReadAllTextAsync(record, token).ConfigureAwait(false));
                if (saved is not { Version: 1, Settings: not null, Path: not null } ||
                    !(saved.Path == suffix || saved.Path.EndsWith("\\" + suffix, StringComparison.Ordinal))) continue;
                saved.Settings.Validate(); matches.Add(saved.Settings);
            }
            catch (Exception error) when (error is IOException or UnauthorizedAccessException or JsonException or ArgumentException) { }
        }
        if (matches.Select(settings => JsonSerializer.Serialize(settings)).Distinct(StringComparer.Ordinal).Count() > 1)
            throw new InvalidOperationException("Several different legacy connection records match this cloud workbook. Open Connections to choose its settings; the old records have been preserved.");
        return matches.FirstOrDefault();
    }

}
