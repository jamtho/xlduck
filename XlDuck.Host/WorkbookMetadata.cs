// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Xml;
using System.Xml.Linq;
using XlDuck.Core;

namespace XlDuck.Host;

/// <summary>Portable definitions only. Credential Manager references never enter this format.</summary>
public static class WorkbookMetadata
{
    public const string Namespace = "urn:xlduck:workbook-settings:v1";
    public const int MaxLength = 1024 * 1024;
    private static readonly XNamespace Ns = Namespace;
    private static readonly JsonSerializerOptions Json = new() { UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow };
    private sealed record Definition(string ConnectionId, string Name, string Scope, S3CredentialSource Credentials,
        string Region, string? Endpoint, string? Profile, bool UseSsl, bool PathStyle);
    private sealed record Document(int Version, Definition[] Connections, ResourceSettings Resources);

    public static string Write(WorkbookSettings settings)
    {
        settings.Validate();
        if (settings.S3.Any(connection => connection.ConnectionId is null || connection.Name is null))
            throw new ArgumentException("Connections need persistent IDs and names before they can be saved.");
        var document = new Document(1, [.. settings.S3.Select(c => new Definition(c.ConnectionId!, c.Name!, c.Scope,
            c.Credentials, c.Region, c.Endpoint, c.Profile, c.UseSsl, c.PathStyle))], settings.Resources);
        var xml = new XElement(Ns + "settings", new XElement(Ns + "json", JsonSerializer.Serialize(document, Json))).ToString(SaveOptions.DisableFormatting);
        if (xml.Length > MaxLength) throw new ArgumentException("Workbook connection metadata is too large.");
        return xml;
    }

    public static WorkbookSettings Read(string xml)
    {
        if (xml.Length > MaxLength) throw new ArgumentException("Workbook connection metadata is too large.");
        using var reader = XmlReader.Create(new StringReader(xml), new XmlReaderSettings
        { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = MaxLength });
        var root = XElement.Load(reader);
        if (root.Name != Ns + "settings" || root.Elements().Count() != 1 || root.Element(Ns + "json") is not { } json)
            throw new ArgumentException("Invalid XlDuck workbook metadata.");
        var document = JsonSerializer.Deserialize<Document>(json.Value, Json);
        if (document is not { Version: 1, Connections: not null, Resources: not null } || document.Connections.Length > 32)
            throw new ArgumentException("Unsupported XlDuck workbook metadata. Update XlDuck or repair the connections.");
        if (document.Connections.Any(c => c is null || !Guid.TryParseExact(c.ConnectionId, "N", out _) || string.IsNullOrWhiteSpace(c.Name)))
            throw new ArgumentException("Workbook connections need valid IDs and names.");
        var settings = new WorkbookSettings([.. document.Connections.Select(c => new S3Connection(c.Scope, c.Credentials,
            c.Region, c.Endpoint, c.Profile, c.UseSsl, c.PathStyle, ConnectionId: c.ConnectionId.ToLowerInvariant(), Name: c.Name))], document.Resources);
        settings.Validate(); return settings;
    }
}
