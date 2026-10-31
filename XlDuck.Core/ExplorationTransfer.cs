// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
namespace XlDuck.Core;

/// <summary>Bounded in-memory transfer for Excel Evaluate's 255-character expressions. No SQL executes here.</summary>
public sealed class ExplorationTransfer
{
    private sealed class Upload(int count)
    {
        public readonly string?[] Parts = new string?[count];
        public int Characters;
        public bool Submitted;
        public DateTimeOffset LastAccess = DateTimeOffset.UtcNow;
    }
    private readonly Dictionary<string, Upload> _uploads = new(StringComparer.Ordinal);
    // The host calls this only on Excel's main thread; it is scoped to one workbook session.
    public void Add(string id, int index, int count, string text)
    {
        ValidateId(id); Prune();
        if (count is < 1 or > 2048 || index < 0 || index >= count || text.Length > 1000)
            throw new ArgumentException("Invalid exploration transfer part.");
        if (!_uploads.TryGetValue(id, out var upload))
        {
            if (_uploads.Count >= 8)
            {
                var oldest = _uploads.Where(pair => pair.Value.Submitted).OrderBy(pair => pair.Value.LastAccess).FirstOrDefault();
                if (oldest.Value is null) throw new ArgumentException("Eight exploration transfers are already staged. Submit them or wait for expiry.");
                _uploads.Remove(oldest.Key);
            }
            _uploads.Add(id, upload = new(count));
        }
        if (upload.Parts.Length != count) throw new ArgumentException("Transfer part count changed. Use a new request id.");
        if (upload.Parts[index] is { } previous && previous != text) throw new ArgumentException("Transfer part changed. Use a new request id.");
        if (upload.Parts[index] is null)
        {
            if (upload.Characters + text.Length > 30000) throw new ArgumentException("Exploration request exceeds 30000 characters.");
            upload.Parts[index] = text; upload.Characters += text.Length;
        }
        upload.LastAccess = DateTimeOffset.UtcNow;
    }
    public string? Complete(string id)
    {
        ValidateId(id); Prune();
        if (!_uploads.TryGetValue(id, out var upload)) return null;
        if (upload.Parts.Any(part => part is null)) throw new ArgumentException("Exploration transfer is incomplete.");
        // Excel may evaluate a UDF more than once for one Evaluate call. Keep completed
        // bytes so repeat submission preserves even a validation failure, not an older success.
        upload.Submitted = true; upload.LastAccess = DateTimeOffset.UtcNow;
        return string.Concat(upload.Parts);
    }
    private void Prune()
    {
        foreach (var (id, upload) in _uploads.ToArray())
            if (DateTimeOffset.UtcNow - upload.LastAccess > TimeSpan.FromMinutes(2)) _uploads.Remove(id);
    }
    private static void ValidateId(string id)
    {
        if (string.IsNullOrWhiteSpace(id) || id.Length > 64 || id.Any(c => !char.IsAsciiLetterOrDigit(c) && c != '-'))
            throw new ArgumentException("Supply a request id of up to 64 letters, digits or hyphens.");
    }
}
