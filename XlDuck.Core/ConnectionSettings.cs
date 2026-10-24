// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Collections.Immutable;

namespace XlDuck.Core;

public enum S3CredentialSource { Anonymous, Environment, AwsProfile, AwsSso }

/// <summary>Non-secret connection routing. Credentials remain in the user's environment or AWS profile.</summary>
public sealed record S3Connection(string Scope, S3CredentialSource Credentials = S3CredentialSource.Environment,
    string Region = "us-east-1", string? Endpoint = null, string? Profile = null, bool UseSsl = true, bool PathStyle = false)
{
    public void Validate()
    {
        if (!Enum.IsDefined(Credentials) || string.IsNullOrEmpty(Scope) || Scope.Length > 4096 || !Uri.TryCreate(Scope, UriKind.Absolute, out var scope) || scope.Scheme != "s3" || scope.Host.Length == 0 || scope.Query.Length != 0 || scope.Fragment.Length != 0 || scope.UserInfo.Length != 0)
            throw new ArgumentException("An S3 connection needs an s3://bucket/ scope without credentials or query parameters.");
        if (string.IsNullOrEmpty(Region) || Region.Length > 128 || Region.Any(c => !char.IsAsciiLetterOrDigit(c) && c != '-')) throw new ArgumentException("Invalid S3 region.");
        if (Credentials is S3CredentialSource.AwsProfile or S3CredentialSource.AwsSso && string.IsNullOrWhiteSpace(Profile)) throw new ArgumentException("Choose an AWS profile for this connection.");
        if (Profile is { Length: > 256 } || Profile?.Any(char.IsControl) == true) throw new ArgumentException("Invalid AWS profile name.");
        if (Endpoint is not null && (Endpoint.Length > 512 || !Uri.TryCreate("https://" + Endpoint, UriKind.Absolute, out var endpoint) || endpoint.UserInfo.Length != 0 || endpoint.AbsolutePath != "/" || endpoint.Query.Length != 0 || endpoint.Fragment.Length != 0))
            throw new ArgumentException("An S3 endpoint must be a hostname with an optional port, without a scheme, path or credentials.");
    }
}

/// <summary>Host-supplied settings fixed for a session. Worksheet SQL cannot construct or update these.</summary>
public sealed record ConnectionSettings(bool EnableExternalReads = false, string? ReaderExtensionDirectory = null,
    ImmutableArray<S3Connection> S3 = default)
{
    public ImmutableArray<S3Connection> S3 { get; init; } = S3.IsDefault ? [] : S3;
    public void Validate()
    {
        if (S3.IsDefault) throw new ArgumentException("Connection list is missing.");
        if (!S3.IsDefault && S3.Length > 32) throw new ArgumentException("At most 32 S3 connections are supported per session.");
        if (!S3.IsDefaultOrEmpty)
        {
            if (!EnableExternalReads || ReaderExtensionDirectory is null) throw new ArgumentException("S3 connections need external reads and the installed reader extensions.");
            foreach (var connection in S3)
            {
                if (connection is null) throw new ArgumentException("Connection list contains a missing connection.");
                connection.Validate();
            }
            if (S3.Select(connection => connection.Scope).Distinct(StringComparer.Ordinal).Count() != S3.Length) throw new ArgumentException("S3 connection scopes must be unique.");
        }
        if (ReaderExtensionDirectory is not null && (!EnableExternalReads || !Path.IsPathFullyQualified(ReaderExtensionDirectory)))
            throw new ArgumentException("Reader extensions need an absolute installation directory and enabled external reads.");
    }
}
