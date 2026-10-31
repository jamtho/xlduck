// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;
using XlDuck.Core;

namespace XlDuck.Host;

/// <summary>Generic credentials owned by this application and the current Windows user.</summary>
public static class CredentialStore
{
    private static string Target(string id)
    {
        if (!OperatingSystem.IsWindows()) throw new PlatformNotSupportedException("Saved access keys require Windows Credential Manager.");
        if (!Guid.TryParseExact(id, "N", out _)) throw new ArgumentException("Invalid credential reference.");
        return "XlDuck/S3/" + id;
    }
    public static void Save(string id, S3AccessKey key)
    {
        var target = Target(id); key.Validate();
        var bytes = JsonSerializer.SerializeToUtf8Bytes(key);
        nint buffer = 0;
        try
        {
            if (bytes.Length > 2560) throw new ArgumentException("These credentials exceed Windows Credential Manager's size limit. Use environment credentials for a longer session token.");
            buffer = Marshal.AllocHGlobal(bytes.Length); Marshal.Copy(bytes, 0, buffer, bytes.Length);
            var credential = new Credential { Type = 1, TargetName = target, CredentialBlobSize = (uint)bytes.Length, CredentialBlob = buffer, Persist = 2, UserName = key.KeyId };
            if (!CredWrite(ref credential, 0)) throw new Win32Exception(Marshal.GetLastWin32Error(), "Could not save the S3 access key in Windows Credential Manager.");
        }
        finally
        {
            CryptographicOperations.ZeroMemory(bytes);
            if (buffer != 0) { Marshal.Copy(bytes, 0, buffer, bytes.Length); Marshal.FreeHGlobal(buffer); }
        }
    }
    public static S3AccessKey Read(string id)
    {
        if (!CredRead(Target(id), 1, 0, out var pointer))
            throw new InvalidOperationException("The saved S3 access key is unavailable. Open Connections and enter it again.");
        byte[]? bytes = null;
        try
        {
            var credential = Marshal.PtrToStructure<Credential>(pointer);
            if (credential.CredentialBlobSize > 2560) throw new InvalidOperationException("The saved S3 access key is invalid.");
            bytes = new byte[credential.CredentialBlobSize]; Marshal.Copy(credential.CredentialBlob, bytes, 0, bytes.Length);
            var key = JsonSerializer.Deserialize<S3AccessKey>(bytes) ?? throw new InvalidOperationException("The saved S3 access key is invalid.");
            key.Validate(); return key;
        }
        finally { if (bytes is not null) CryptographicOperations.ZeroMemory(bytes); CredFree(pointer); }
    }
    public static void Delete(string id)
    {
        if (!CredDelete(Target(id), 1, 0) && Marshal.GetLastWin32Error() != 1168)
            throw new Win32Exception(Marshal.GetLastWin32Error(), "Could not remove the saved S3 access key.");
    }
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct Credential
    {
        public uint Flags, Type;
        public string? TargetName, Comment;
        public long LastWritten;
        public uint CredentialBlobSize;
        public nint CredentialBlob;
        public uint Persist, AttributeCount;
        public nint Attributes;
        public string? TargetAlias, UserName;
    }
    [DllImport("advapi32.dll", EntryPoint = "CredWriteW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)] private static extern bool CredWrite(ref Credential credential, uint flags);
    [DllImport("advapi32.dll", EntryPoint = "CredReadW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)] private static extern bool CredRead(string target, uint type, uint flags, out nint credential);
    [DllImport("advapi32.dll", EntryPoint = "CredDeleteW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)] private static extern bool CredDelete(string target, uint type, uint flags);
    [DllImport("advapi32.dll")] private static extern void CredFree(nint credential);
}
