// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using DuckDB.NET.Data;
using System.Security.Cryptography;
using XlDuck.Core;

namespace XlDuck.Engine;

internal static class ConnectionSetup
{
    internal static DuckDBConnection Open(ConnectionSettings settings, string spillDirectory, out ReadComputePolicy policy)
    {
        settings.Validate();
        var builder = new DuckDBConnectionStringBuilder
        {
            ["DataSource"] = ":memory:", ["threads"] = "2", ["memory_limit"] = "512MB",
            ["temp_directory"] = spillDirectory, ["max_temp_directory_size"] = "1GiB",
            ["autoinstall_known_extensions"] = "false", ["autoload_known_extensions"] = "false",
            ["allow_community_extensions"] = "false", ["allow_unsigned_extensions"] = "false",
            ["enable_external_access"] = settings.EnableExternalReads ? "true" : "false",
            ["allow_persistent_secrets"] = "false", ["allow_unredacted_secrets"] = "false"
        };
        var connection = new DuckDBConnection(builder.ConnectionString);
        try
        {
            connection.Open();
            Execute(connection, "SET enable_logging=false; SET enable_http_logging=false;");
            policy = new(connection);
            if (settings.ReaderExtensionDirectory is { } extensionDirectory)
            {
                Load("httpfs");
                if (!settings.S3.IsDefaultOrEmpty && settings.S3.Any(s3 => s3.Credentials != S3CredentialSource.Anonymous)) Load("aws");
                void Load(string name)
                {
                    var path = Path.Combine(extensionDirectory, name + ".duckdb_extension");
                    if (!File.Exists(path)) throw new InvalidOperationException("The installed HTTP/S3 reader extensions are missing. Repair the XlDuck installation.");
                    using var file = File.OpenRead(path);
                    var expected = name == "httpfs" ? "9B0B0202CFA34F94F12BD9B2698002735BBCD0C1BD5FDA2E85450583EFD18E94" : "008CB2809ADD3AFB7E5BCDF27655BE129099D1B35B52D3248969A5A055F4C627";
                    if (Convert.ToHexString(SHA256.HashData(file)) != expected)
                        throw new InvalidOperationException("The HTTP/S3 reader extension does not match the pinned XlDuck build. Repair the installation.");
                    Execute(connection, "LOAD " + Literal(path));
                }
                for (var i = 0; !settings.S3.IsDefault && i < settings.S3.Length; i++)
                {
                    var s3 = settings.S3[i];
                    var clauses = new List<string> { "TYPE s3", "SCOPE " + Literal(s3.Scope), "REGION " + Literal(s3.Region),
                        "USE_SSL " + (s3.UseSsl ? "true" : "false"), "URL_STYLE " + Literal(s3.PathStyle ? "path" : "vhost") };
                    if (s3.Endpoint is not null) clauses.Add("ENDPOINT " + Literal(s3.Endpoint));
                    if (s3.Credentials == S3CredentialSource.Anonymous) clauses.Add("PROVIDER config");
                    else
                    {
                        clauses.Add("PROVIDER credential_chain");
                        clauses.Add("CHAIN " + Literal(s3.Credentials switch { S3CredentialSource.Environment => "env", S3CredentialSource.AwsSso => "sso", _ => "config" }));
                        if (s3.Profile is not null) clauses.Add("PROFILE " + Literal(s3.Profile));
                    }
                    try { Execute(connection, $"CREATE SECRET _xlduck_s3_{i} ({string.Join(',', clauses)})"); }
                    catch { throw new InvalidOperationException($"S3 connection {i + 1} could not load its credentials. Check its environment variables or authenticate the selected AWS profile, then restart XlDuck for this workbook."); }
                }
                // httpfs can populate legacy settings from the environment when first loaded.
                // Credentials belong in the scoped secrets above, not ordinary settings metadata.
                using var inspect = connection.CreateCommand();
                inspect.CommandText = "SELECT name FROM duckdb_settings() WHERE name IN ('s3_access_key_id','s3_secret_access_key','s3_session_token')";
                var legacySettings = new List<string>();
                using (var reader = inspect.ExecuteReader())
                    while (reader.Read()) legacySettings.Add(reader.GetString(0));
                foreach (var name in legacySettings) Execute(connection, $"SET {name}=''");
            }
            Execute(connection, "SET lock_configuration=true;");
            return connection;
        }
        catch { connection.Dispose(); throw; }
    }
    private static string Literal(string value) => "'" + value.Replace("'", "''", StringComparison.Ordinal) + "'";
    private static void Execute(DuckDBConnection connection, string sql)
    {
        using var command = connection.CreateCommand(); command.CommandText = sql; command.ExecuteNonQuery();
    }
}
