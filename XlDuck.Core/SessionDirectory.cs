// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
namespace XlDuck.Core;

/// <summary>Owns a newly-created private temporary directory, never a caller-supplied directory.</summary>
public sealed class SessionDirectory : IDisposable
{
    private const string MarkerName = "owner.lock";
    private const string MarkerContent = "XlDuck spill ownership v1\n";
    private static readonly string Root = System.IO.Path.GetFullPath(System.IO.Path.Combine(System.IO.Path.GetTempPath(), "XlDuck", "spill"));
    private static readonly Lazy<int> StartupCleanup = new(ReclaimOrphans);
    private readonly string _parent;
    private readonly FileStream _ownership;
    private int _disposed;
    public string Path { get; }
    public SessionDirectory()
    {
        _ = StartupCleanup.Value;
        Directory.CreateDirectory(Root);
        if ((File.GetAttributes(Root) & FileAttributes.ReparsePoint) != 0)
            throw new IOException("The spill directory root must not be a symbolic link or junction.");
        var directory = Directory.CreateDirectory(System.IO.Path.Combine(Root, "xlduck-" + Guid.NewGuid().ToString("N")));
        Path = System.IO.Path.GetFullPath(directory.FullName);
        _parent = System.IO.Path.GetFullPath(directory.Parent!.FullName);
        _ownership = new FileStream(System.IO.Path.Combine(Path, MarkerName), FileMode.CreateNew, FileAccess.ReadWrite, FileShare.Read);
        _ownership.Write(System.Text.Encoding.UTF8.GetBytes(MarkerContent));
        _ownership.Flush();
    }
    /// <summary>The worker keeps the marker open until its native engine has shut down.</summary>
    public static FileStream HoldWorkerOwnership(string path) => new(System.IO.Path.Combine(path, MarkerName),
        FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

    /// <summary>Reclaim only marked directories under our dedicated root, never active hosts or workers.</summary>
    public static int ReclaimOrphans()
    {
        string[] candidates;
        try
        {
            if (!Directory.Exists(Root) || (File.GetAttributes(Root) & FileAttributes.ReparsePoint) != 0) return 0;
            candidates = Directory.EnumerateDirectories(Root, "xlduck-*").Take(256).ToArray();
        }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException) { return 0; }
        var reclaimed = 0;
        foreach (var path in candidates)
        {
            try
            {
                var name = System.IO.Path.GetFileName(path);
                if (name.Length != 39 || !Guid.TryParseExact(name[7..], "N", out _) ||
                    (File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0) continue;
                var marker = System.IO.Path.Combine(path, MarkerName);
                if (!File.Exists(marker) || (File.GetAttributes(marker) & FileAttributes.ReparsePoint) != 0) continue;
                // Both host and worker leases deny this access. Keep exclusive ownership
                // through deletion, permitting Windows to remove our open marker itself.
                using var lease = new FileStream(marker, FileMode.Open, FileAccess.ReadWrite, FileShare.Delete);
                if (lease.Length != MarkerContent.Length) continue;
                using var reader = new StreamReader(lease, leaveOpen: true);
                if (reader.ReadToEnd() != MarkerContent) continue;
                Directory.Delete(path, recursive: true); reclaimed++;
            }
            catch (Exception error) when (error is IOException or UnauthorizedAccessException) { /* Live owner or unavailable file: leave it alone. */ }
        }
        return reclaimed;
    }
    public void Dispose()
    {
        if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
        _ownership.Dispose();
        if (!Directory.Exists(Path)) return;
        var resolved = System.IO.Path.GetFullPath(Path);
        if (!string.Equals(System.IO.Path.GetDirectoryName(resolved), _parent, StringComparison.OrdinalIgnoreCase) ||
            !System.IO.Path.GetFileName(resolved).StartsWith("xlduck-", StringComparison.Ordinal) ||
            (File.GetAttributes(resolved) & FileAttributes.ReparsePoint) != 0)
            throw new IOException("Temporary directory ownership could not be verified; cleanup was refused.");
        Directory.Delete(resolved, recursive: true);
    }
}
