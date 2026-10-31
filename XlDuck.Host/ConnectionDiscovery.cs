// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Text.Json;
using XlDuck.Core;

namespace XlDuck.Host;

/// <summary>Public connection metadata only. Never resolves credentials or contacts a source.</summary>
public static class ConnectionDiscovery
{
    private static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web);
    private sealed record Connection(string Type, string? Name, string Scope, string? Endpoint, string Authentication, bool UseSsl, bool PathStyle, bool RequiresCredentialSetup);
    private sealed record Page(int Version, string State, IReadOnlyList<Connection> Connections, int Total, int? NextOffset);

    public static string Describe(WorkbookSettings settings, int offset = 0)
    {
        settings.Validate();
        if (offset < 0 || offset > settings.S3.Length) throw new ArgumentException("Connection offset must be between zero and the connection count. Start with zero.");
        var connections = new List<Connection>();
        string Serialize(int next) => JsonSerializer.Serialize(new Page(1, "succeeded", connections, settings.S3.Length,
            next < settings.S3.Length ? next : null), Json);
        var result = Serialize(offset);
        for (var index = offset; index < settings.S3.Length; index++)
        {
            var source = settings.S3[index];
            var authentication = source.Credentials switch
            {
                S3CredentialSource.Anonymous => "anonymous",
                S3CredentialSource.Environment => "environment",
                S3CredentialSource.AwsProfile => "named_profile",
                S3CredentialSource.AwsSso => "sso_profile",
                S3CredentialSource.AccessKey => "saved_access_key",
                _ => throw new ArgumentException("Unknown connection authentication method.")
            };
            // Deliberate projection: no credential IDs, profile names, keys or tokens.
            connections.Add(new("s3", source.Name, source.Scope, source.Endpoint, authentication, source.UseSsl, source.PathStyle,
                source.Credentials == S3CredentialSource.AccessKey && source.CredentialId is null));
            var candidate = Serialize(index + 1);
            if (candidate.Length > 30000)
            {
                if (connections.Count == 1) throw new ArgumentException("Connection metadata exceeds one cell's size limit.");
                break;
            }
            result = candidate;
        }
        return result;
    }
}
