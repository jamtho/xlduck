// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Text.RegularExpressions;

namespace XlDuck.Host;

public static partial class PreviewAccess
{
    public const string Host = "xlduck.invalid";
    public const string Document = "https://xlduck.invalid/preview.html";
    public static bool CanNavigate(string uri) => string.Equals(uri, Document, StringComparison.Ordinal);
    public static bool IsMapTile(string uri)
    {
        if (!Uri.TryCreate(uri, UriKind.Absolute, out var url) || url.Scheme != "https" || !url.IsDefaultPort ||
            url.UserInfo.Length != 0 || url.Query.Length != 0 || url.Fragment.Length != 0) return false;
        return url.Host == "tile.openstreetmap.org"
            ? OsmTile().IsMatch(url.AbsolutePath)
            : url.Host == "server.arcgisonline.com" && EsriTile().IsMatch(url.AbsolutePath);
    }
    [GeneratedRegex(@"^/[0-9]{1,2}/[0-9]{1,10}/[0-9]{1,10}\.png$")]
    private static partial Regex OsmTile();
    [GeneratedRegex(@"^/ArcGIS/rest/services/World_Imagery/MapServer/tile/[0-9]{1,2}/[0-9]{1,10}/[0-9]{1,10}$")]
    private static partial Regex EsriTile();
}
