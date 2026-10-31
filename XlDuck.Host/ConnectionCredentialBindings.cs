// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using XlDuck.Core;

namespace XlDuck.Host;

/// <summary>Local associations between portable connection identities and saved keys.</summary>
public sealed class ConnectionCredentialBindings(string directory)
{
    private sealed record Binding(int Version, string ConnectionId, string Destination, string CredentialId);
    private readonly string _directory = Path.GetFullPath(directory);
    public static string Destination(S3Connection connection)
    {
        // Include transport and endpoint-affecting options. Labels and scope paths
        // do not identify a server. AWS's default server also depends on its region.
        var endpoint = connection.Endpoint is null ? "amazon-s3:" + connection.Region.ToLowerInvariant()
            : new Uri((connection.UseSsl ? "https://" : "http://") + connection.Endpoint).GetLeftPart(UriPartial.Authority).ToLowerInvariant();
        return JsonSerializer.Serialize(new { endpoint, connection.UseSsl, connection.PathStyle, connection.Region });
    }
    private string RecordPath(S3Connection connection) => Path.Combine(_directory,
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(connection.ConnectionId!.ToLowerInvariant() + "\n" + Destination(connection)))) + ".json");
    public void Save(S3Connection connection)
    {
        connection.Validate();
        if (!Guid.TryParseExact(connection.ConnectionId, "N", out _) || !Guid.TryParseExact(connection.CredentialId, "N", out _) || connection.Credentials != S3CredentialSource.AccessKey)
            throw new ArgumentException("A saved-key binding needs valid connection and credential IDs.");
        Directory.CreateDirectory(_directory);
        var record = RecordPath(connection); var temporary = record + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            File.WriteAllText(temporary, JsonSerializer.Serialize(new Binding(1, connection.ConnectionId!.ToLowerInvariant(), Destination(connection), connection.CredentialId!)));
            File.Move(temporary, record, overwrite: true);
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }
    public WorkbookSettings Resolve(WorkbookSettings settings) => settings with { S3 = [.. settings.S3.Select(connection =>
        connection with { CredentialId = connection.Credentials == S3CredentialSource.AccessKey ? Read(connection) : null })] };
    private string? Read(S3Connection connection)
    {
        if (!Guid.TryParseExact(connection.ConnectionId, "N", out _)) return null;
        try
        {
            var path = RecordPath(connection);
            if (!File.Exists(path) || new FileInfo(path).Length > 8192) return null;
            var binding = JsonSerializer.Deserialize<Binding>(File.ReadAllText(path));
            return binding is { Version: 1 } && binding.ConnectionId == connection.ConnectionId!.ToLowerInvariant() &&
                binding.Destination == Destination(connection) && Guid.TryParseExact(binding.CredentialId, "N", out _)
                ? binding.CredentialId : null;
        }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException or JsonException) { return null; }
    }
}
