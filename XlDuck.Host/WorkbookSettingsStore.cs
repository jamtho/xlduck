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
            connection.Validate();
        }
        if (S3.Select(connection => connection.Scope).Distinct(StringComparer.Ordinal).Count() != S3.Length)
            throw new ArgumentException("S3 connection scopes must be unique.");
    }
}

/// <summary>User-local connection settings by workbook path. This is not an execution permission store.</summary>
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

}
