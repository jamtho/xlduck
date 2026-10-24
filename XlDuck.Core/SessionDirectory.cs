// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
namespace XlDuck.Core;

/// <summary>Owns a newly-created private temporary directory, never a caller-supplied directory.</summary>
public sealed class SessionDirectory : IDisposable
{
    private readonly string _parent;
    private int _disposed;
    public string Path { get; }
    public SessionDirectory()
    {
        var directory = Directory.CreateTempSubdirectory("xlduck-");
        Path = System.IO.Path.GetFullPath(directory.FullName);
        _parent = System.IO.Path.GetFullPath(directory.Parent!.FullName);
    }
    public void Dispose()
    {
        if (Interlocked.Exchange(ref _disposed, 1) != 0 || !Directory.Exists(Path)) return;
        var resolved = System.IO.Path.GetFullPath(Path);
        if (!string.Equals(System.IO.Path.GetDirectoryName(resolved), _parent, StringComparison.OrdinalIgnoreCase) ||
            !System.IO.Path.GetFileName(resolved).StartsWith("xlduck-", StringComparison.Ordinal) ||
            (File.GetAttributes(resolved) & FileAttributes.ReparsePoint) != 0)
            throw new IOException("Temporary directory ownership could not be verified; cleanup was refused.");
        Directory.Delete(resolved, recursive: true);
    }
}
