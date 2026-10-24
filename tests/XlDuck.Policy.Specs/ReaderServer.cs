// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Collections.Concurrent;
using System.Net;
using System.Net.Sockets;
using System.Text;

// Local HTTP/S3-compatible read fixture. Never forwards requests or contacts a cloud service.
internal sealed class ReaderServer : IAsyncDisposable
{
    internal const string AccessKey = "xlduck-fixture-key";
    internal const string SessionToken = "xlduck-fixture-token";
    private readonly TcpListener _listener = new(IPAddress.Loopback, 0);
    private readonly CancellationTokenSource _stop = new();
    private readonly Task _loop;
    private readonly List<Task> _requests = [];
    private readonly byte[] _csv, _parquet;
    internal ConcurrentQueue<string> Requests { get; } = new();
    private int _authenticated;
    internal int AuthenticatedS3Reads => Volatile.Read(ref _authenticated);
    internal string Endpoint { get; }
    internal string BaseUrl => "http://" + Endpoint + "/";
    internal ReaderServer(byte[] csv, byte[] parquet)
    {
        _csv = csv; _parquet = parquet; _listener.Start();
        Endpoint = "127.0.0.1:" + ((IPEndPoint)_listener.LocalEndpoint).Port;
        _loop = Run();
    }
    private async Task Run()
    {
        try
        {
            while (!_stop.IsCancellationRequested)
            {
                var client = await _listener.AcceptTcpClientAsync(_stop.Token);
                _requests.Add(Respond(client));
            }
        }
        catch (OperationCanceledException) { }
    }
    private async Task Respond(TcpClient client)
    {
        using (client)
        {
            var stream = client.GetStream();
            using var reader = new StreamReader(stream, Encoding.ASCII, leaveOpen: true);
            var first = (await reader.ReadLineAsync(_stop.Token))!.Split(' ');
            var method = first[0]; var target = first[1];
            var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string? line;
            while (!string.IsNullOrEmpty(line = await reader.ReadLineAsync(_stop.Token)))
            {
                var colon = line.IndexOf(':'); if (colon > 0) headers[line[..colon]] = line[(colon + 1)..].Trim();
            }
            Requests.Enqueue(method);
            bool s3 = target.StartsWith("/bucket/", StringComparison.Ordinal);
            bool authenticated = headers.GetValueOrDefault("Authorization", "").Contains("Credential=" + AccessKey + "/", StringComparison.Ordinal) && headers.GetValueOrDefault("x-amz-security-token") == SessionToken;
            if (s3 && authenticated) Interlocked.Increment(ref _authenticated);
            var listing = target.Contains("list-type=2", StringComparison.Ordinal);
            var payload = listing ? Encoding.UTF8.GetBytes($"<?xml version=\"1.0\"?><ListBucketResult xmlns=\"http://s3.amazonaws.com/doc/2006-03-01/\"><Name>bucket</Name><Prefix></Prefix><KeyCount>1</KeyCount><MaxKeys>1000</MaxKeys><IsTruncated>false</IsTruncated><Contents><Key>sample.parquet</Key><LastModified>2026-10-24T00:00:00.000Z</LastModified><ETag>&quot;fixture&quot;</ETag><Size>{_parquet.Length}</Size><StorageClass>STANDARD</StorageClass></Contents></ListBucketResult>") :
                target.EndsWith("sample.csv", StringComparison.Ordinal) ? _csv : target.EndsWith("sample.parquet", StringComparison.Ordinal) ? _parquet : [];
            var status = payload.Length == 0 ? "404 Not Found" : s3 && !authenticated ? "403 Forbidden" : "200 OK";
            if (status != "200 OK") payload = [];
            var start = 0; var end = payload.Length - 1;
            var rangeHeader = "";
            if (payload.Length > 0 && headers.TryGetValue("Range", out var range) && range.StartsWith("bytes=", StringComparison.Ordinal))
            {
                var parts = range[6..].Split('-'); start = int.Parse(parts[0]);
                if (parts[1].Length > 0) end = Math.Min(end, int.Parse(parts[1]));
                rangeHeader = $"Content-Range: bytes {start}-{end}/{payload.Length}\r\n";
                status = "206 Partial Content";
            }
            var length = Math.Max(0, end - start + 1);
            var response = $"HTTP/1.1 {status}\r\nContent-Length: {length}\r\nAccept-Ranges: bytes\r\nETag: \"fixture\"\r\n{rangeHeader}Connection: close\r\n\r\n";
            await stream.WriteAsync(Encoding.ASCII.GetBytes(response), _stop.Token);
            if (method != "HEAD" && length > 0) await stream.WriteAsync(payload.AsMemory(start, length), _stop.Token);
        }
    }
    public async ValueTask DisposeAsync()
    {
        _stop.Cancel(); _listener.Stop();
        try { await _loop; await Task.WhenAll(_requests); } catch (OperationCanceledException) { }
        _stop.Dispose();
    }
}
