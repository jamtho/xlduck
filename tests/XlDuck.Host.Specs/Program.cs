// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using XlDuck.Core;
using XlDuck.Host;
using System.Text.Json;

using var files = new SessionDirectory();
var storePath = Path.Combine(files.Path, "connections");
var store = new WorkbookSettingsStore(storePath);
var workbook = Path.Combine(files.Path, "research.xlsx");
File.WriteAllText(workbook, "saved workbook revision");
int passed = 0, failed = 0;
async Task Check(string name, Func<Task> action)
{
    try { await action(); passed++; Console.WriteLine("PASS " + name); }
    catch (Exception error) { failed++; Console.WriteLine("FAIL " + name + ": " + error.Message); }
}
void Equal<T>(T expected, T actual) { if (!EqualityComparer<T>.Default.Equals(expected, actual)) throw new Exception($"Expected {expected}, got {actual}"); }
await Check("connection settings are optional and preserve routing without credentials", async () =>
{
    Equal(true, await store.ReadAsync(workbook) is null);
    var settings = new WorkbookSettings([new("s3://research/", S3CredentialSource.AwsProfile, Profile: "research")]);
    await store.SaveAsync(workbook, settings);
    Equal("research", (await new WorkbookSettingsStore(storePath).ReadAsync(workbook))!.S3[0].Profile);
});
await Check("resource settings persist with connections and old settings receive explicit defaults", async () =>
{
    var original = new WorkbookSettings([new("s3://research/", S3CredentialSource.AwsProfile, Profile: "research")], new(2048, 8, 3));
    await store.SaveAsync(workbook, original);
    var restored = (await store.ReadAsync(workbook))!;
    Equal(original.Resources, restored.Resources); Equal(original.Resources, restored.ForSession(Path.GetFullPath(".")).Resources);
    Equal(new ResourceSettings(), System.Text.Json.JsonSerializer.Deserialize<WorkbookSettings>("{}")!.Resources);
    foreach (var resources in new[] { new ResourceSettings(0), new ResourceSettings(SpillGiB: 0), new ResourceSettings(Threads: 17) })
    {
        try { (original with { Resources = resources }).Validate(); throw new Exception("Expected invalid budget rejection"); }
        catch (ArgumentException) { }
    }
});
await Check("connection settings survive external edits and remain scoped to the workbook path", async () =>
{
    var copy = Path.Combine(files.Path, "copy.xlsx"); File.Copy(workbook, copy);
    Equal(true, await store.ReadAsync(copy) is null);
    File.AppendAllText(workbook, "changed SQL");
    Equal("research", (await store.ReadAsync(workbook))!.S3[0].Profile);
    await store.SaveAsync(workbook, new());
    Equal(true, await store.ReadAsync(workbook) is not null);
    store.Forget(workbook);
    Equal(true, await store.ReadAsync(workbook) is null);
});
await Check("corrupt connection records are ignored and interrupted writes leave no temporary record", async () =>
{
    await store.SaveAsync(workbook, new());
    File.WriteAllText(Directory.GetFiles(storePath, "*.json").Single(), "not json");
    Equal(true, await store.ReadAsync(workbook) is null);
    using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
    try { await store.SaveAsync(workbook, new(), cancelled.Token); throw new Exception("Expected cancellation"); } catch (OperationCanceledException) { }
    Equal(0, Directory.GetFiles(storePath, "*.tmp").Length);
});
await Check("saved access keys stay out of settings and resolve only for worker startup", async () =>
{
    var id = Guid.NewGuid().ToString("N");
    var settings = new WorkbookSettings([new("s3://research/", S3CredentialSource.AccessKey, Endpoint: "storage.example.test:9000", PathStyle: true, CredentialId: id)]);
    await store.SaveAsync(workbook, settings);
    var key = new S3AccessKey("fixture-id", "fixture-secret", "fixture-token");
    var loaded = (await store.ReadAsync(workbook))!;
    var config = loaded.ForSession(Path.GetFullPath("."), requested => requested == id ? key : throw new Exception("Wrong key requested"));
    config.Validate(); Equal(key, config.AccessKeys![id]);
    Equal(false, File.ReadAllText(Directory.GetFiles(storePath, "*.json").Single()).Contains("fixture-secret"));
    Equal(false, key.ToString().Contains("fixture-secret"));
});
await Check("connection discovery exposes routing without resolving or disclosing credentials", () =>
{
    var credentialId = Guid.NewGuid().ToString("N"); // Deliberately absent from Credential Manager.
    var settings = new WorkbookSettings([.. Enum.GetValues<S3CredentialSource>().Select((source, index) =>
        new S3Connection($"s3://bucket-{index}/prefix/", source, Endpoint: "storage.example.test:9000",
            Profile: "private-profile", CredentialId: credentialId, UseSsl: false, PathStyle: true))]);
    var json = ConnectionDiscovery.Describe(settings);
    using var document = JsonDocument.Parse(json);
    var root = document.RootElement;
    Equal("succeeded", root.GetProperty("state").GetString());
    Equal(5, root.GetProperty("total").GetInt32());
    Equal(JsonValueKind.Null, root.GetProperty("nextOffset").ValueKind);
    var entries = root.GetProperty("connections");
    Equal("anonymous,environment,named_profile,sso_profile,saved_access_key",
        string.Join(',', entries.EnumerateArray().Select(item => item.GetProperty("authentication").GetString())));
    Equal("authentication,endpoint,name,pathStyle,requiresCredentialSetup,scope,type,useSsl", string.Join(',', entries[0].EnumerateObject().Select(item => item.Name).Order(StringComparer.Ordinal)));
    Equal("s3://bucket-0/prefix/", entries[0].GetProperty("scope").GetString());
    Equal("storage.example.test:9000", entries[0].GetProperty("endpoint").GetString());
    Equal(false, entries[0].GetProperty("useSsl").GetBoolean()); Equal(true, entries[0].GetProperty("pathStyle").GetBoolean());
    Equal(false, json.Contains(credentialId)); Equal(false, json.Contains("private-profile"));
    using var empty = JsonDocument.Parse(ConnectionDiscovery.Describe(new()));
    Equal(0, empty.RootElement.GetProperty("connections").GetArrayLength());
    Equal(0, empty.RootElement.GetProperty("total").GetInt32());
    foreach (var offset in new[] { -1, 6 })
    {
        try { ConnectionDiscovery.Describe(settings, offset); throw new Exception("Expected offset rejection"); } catch (ArgumentException) { }
    }
    return Task.CompletedTask;
});
await Check("large connection inventories page within one Excel cell without losing scopes or Unicode", () =>
{
    var settings = new WorkbookSettings([.. Enumerable.Range(0, 32).Select(index =>
        new S3Connection($"s3://bucket-{index}/" + new string('\u4e2d', 3900), S3CredentialSource.Anonymous))]);
    var found = new List<string>();
    var offset = 0; var pages = 0;
    do
    {
        var json = ConnectionDiscovery.Describe(settings, offset); Equal(true, json.Length <= 30000);
        using var document = JsonDocument.Parse(json); var root = document.RootElement;
        Equal(32, root.GetProperty("total").GetInt32());
        found.AddRange(root.GetProperty("connections").EnumerateArray().Select(item => item.GetProperty("scope").GetString()!));
        pages++;
        if (root.GetProperty("nextOffset").ValueKind == JsonValueKind.Null) break;
        var next = root.GetProperty("nextOffset").GetInt32(); Equal(true, next > offset); offset = next;
    } while (pages <= 32);
    Equal(true, pages > 1); Equal(true, found.SequenceEqual(settings.S3.Select(item => item.Scope)));
    using var end = JsonDocument.Parse(ConnectionDiscovery.Describe(settings, 32));
    Equal(0, end.RootElement.GetProperty("connections").GetArrayLength());
    return Task.CompletedTask;
});
await Check("workbook metadata preserves named connections and budgets without credential references", () =>
{
    var id = Guid.NewGuid().ToString("N");
    var original = new WorkbookSettings([new("s3://research/", S3CredentialSource.AccessKey, Endpoint: "storage.example.test:9000", CredentialId: id, Name: "Research & charts")], new(2048, 8, 3)).WithConnectionIds();
    var xml = WorkbookMetadata.Write(original);
    Equal(false, xml.Contains(id)); Equal(false, xml.Contains("CredentialId"));
    var restored = WorkbookMetadata.Read(xml);
    Equal(original.S3[0] with { CredentialId = null }, restored.S3[0]); Equal(original.Resources, restored.Resources);
    Equal(xml, WorkbookMetadata.Write(restored));
    Equal(original.S3[0].ConnectionId, restored.WithConnectionIds().S3[0].ConnectionId);
    using var discovery = JsonDocument.Parse(ConnectionDiscovery.Describe(restored));
    Equal(true, discovery.RootElement.GetProperty("connections")[0].GetProperty("requiresCredentialSetup").GetBoolean());
    try { restored.ForSession(files.Path, _ => throw new Exception("Unbound connection tried reading the vault")); throw new Exception("Missing local binding accepted"); }
    catch (InvalidOperationException error) { Equal(true, error.Message.Contains("Connections")); }
    return Task.CompletedTask;
});
await Check("credential bindings survive copies and label edits but reject changed destinations and foreign accounts", () =>
{
    var bindings = new ConnectionCredentialBindings(Path.Combine(files.Path, "bindings"));
    var source = new WorkbookSettings([new("s3://research/", S3CredentialSource.AccessKey, Endpoint: "STORAGE.example.test:443", CredentialId: Guid.NewGuid().ToString("N"))]).WithConnectionIds();
    bindings.Save(source.S3[0]);
    var portable = WorkbookMetadata.Read(WorkbookMetadata.Write(source));
    Equal(source.S3[0].CredentialId, bindings.Resolve(portable).S3[0].CredentialId);
    var renamed = portable with { S3 = [portable.S3[0] with { Name = "New label", Endpoint = "storage.example.test" }] };
    Equal(source.S3[0].CredentialId, bindings.Resolve(renamed).S3[0].CredentialId);
    foreach (var changed in new[] { portable.S3[0] with { Endpoint = "elsewhere.example.test" }, portable.S3[0] with { UseSsl = false },
        portable.S3[0] with { PathStyle = true }, portable.S3[0] with { Region = "local-2" }, portable.S3[0] with { ConnectionId = Guid.NewGuid().ToString("N") } })
        Equal<string?>(null, bindings.Resolve(new([changed])).S3[0].CredentialId);
    var secondAccount = new ConnectionCredentialBindings(Path.Combine(files.Path, "other-account"));
    Equal<string?>(null, secondAccount.Resolve(portable).S3[0].CredentialId);
    // Rebinding the new endpoint must not revoke the original workbook copy's key.
    var newEndpoint = source.S3[0] with { Endpoint = "elsewhere.example.test", CredentialId = Guid.NewGuid().ToString("N") };
    bindings.Save(newEndpoint);
    Equal(newEndpoint.CredentialId, bindings.Resolve(new([newEndpoint with { CredentialId = null }])).S3[0].CredentialId);
    Equal(source.S3[0].CredentialId, bindings.Resolve(portable).S3[0].CredentialId);
    return Task.CompletedTask;
});
await Check("invalid metadata cannot inject a credential ID or silently use defaults", () =>
{
    var xml = WorkbookMetadata.Write(new WorkbookSettings([new("s3://research/", S3CredentialSource.Anonymous)]).WithConnectionIds());
    var bad = new[] { xml.Replace("\"Version\":1", "\"Version\":99"), xml.Replace("\"Scope\":", "\"CredentialId\":\"" + Guid.NewGuid().ToString("N") + "\",\"Scope\":"),
        xml.Replace("\"ConnectionId\":\"", "\"ConnectionId\":\"bad"), "<!DOCTYPE settings [<!ENTITY x SYSTEM 'file:///missing'>]>" + xml,
        new string('x', WorkbookMetadata.MaxLength + 1), "<settings/>" };
    foreach (var input in bad)
    {
        var rejected = false;
        try { WorkbookMetadata.Read(input); } catch (Exception error) when (error is ArgumentException or JsonException or System.Xml.XmlException) { rejected = true; }
        Equal(true, rejected);
    }
    var duplicate = new WorkbookSettings([new("s3://one/", S3CredentialSource.Anonymous), new("s3://two/", S3CredentialSource.Anonymous)]).WithConnectionIds();
    try { (duplicate with { S3 = [duplicate.S3[0], duplicate.S3[1] with { ConnectionId = duplicate.S3[0].ConnectionId }] }).Validate(); throw new Exception("Duplicate IDs accepted"); }
    catch (ArgumentException) { }
    return Task.CompletedTask;
});
if (OperatingSystem.IsWindows()) await Check("legacy OneDrive settings migrate independently of current directory without guessing ambiguous records", async () =>
{
    var previous = Environment.CurrentDirectory;
    var migrationDirectory = Path.Combine(files.Path, "legacy-cloud");
    var legacy = new WorkbookSettingsStore(migrationDirectory);
    var firstCwd = Path.Combine(files.Path, "excel-cwd"); var secondCwd = Path.Combine(files.Path, "other-cwd");
    Directory.CreateDirectory(firstCwd); Directory.CreateDirectory(secondCwd);
    const string url = "https://d.docs.live.net/test-account/Documents/Research.xlsx";
    var settings = new WorkbookSettings([new("s3://cloud/", S3CredentialSource.AccessKey, CredentialId: Guid.NewGuid().ToString("N"))]);
    try
    {
        Environment.CurrentDirectory = firstCwd; await legacy.SaveAsync(url, settings);
        Environment.CurrentDirectory = secondCwd;
        Equal(settings.S3[0], (await legacy.ReadForMigrationAsync(url))!.S3[0]);
        Equal<WorkbookSettings?>(null, await legacy.ReadForMigrationAsync(url.Replace("test-account", "other-account")));
        await legacy.SaveAsync(url, new([new("s3://different/", S3CredentialSource.Anonymous)]));
        try { await legacy.ReadForMigrationAsync(url); throw new Exception("Ambiguous cloud records accepted"); }
        catch (InvalidOperationException) { }
        Equal(2, Directory.GetFiles(migrationDirectory, "*.json").Length);
    }
    finally { Environment.CurrentDirectory = previous; }
});
if (OperatingSystem.IsWindows()) await Check("Windows Credential Manager saves, replaces and removes only the test key", () =>
{
    var id = Guid.NewGuid().ToString("N");
    try
    {
        var first = new S3AccessKey("fixture-id", "fixture-secret", "fixture-token");
        CredentialStore.Save(id, first); Equal(first, CredentialStore.Read(id));
        var second = first with { Secret = "replacement-secret" };
        CredentialStore.Save(id, second); Equal(second, CredentialStore.Read(id));
        CredentialStore.Delete(id);
        try { CredentialStore.Read(id); throw new Exception("Missing key was accepted"); } catch (InvalidOperationException) { }
    }
    finally { CredentialStore.Delete(id); }
    return Task.CompletedTask;
});
await Check("date conversions respect both Excel date systems and reject the fictional leap day", () =>
{
    Equal("1900-01-01", ExcelValues.DateText(1, false, false));
    Equal("1900-03-01", ExcelValues.DateText(61, false, false));
    Equal("1904-01-01", ExcelValues.DateText(0, true, false));
    Equal(1d, ExcelValues.Display(new CellValue(ValueKind.Date, "1900-01-01"), false));
    Equal(0d, ExcelValues.Display(new CellValue(ValueKind.Date, "1904-01-01"), true));
    Equal(45292d - 1462, ExcelValues.Display(new CellValue(ValueKind.Date, "2024-01-01"), true));
    try { ExcelValues.DateText(60, false, false); throw new Exception("Expected fictional date rejection"); } catch (ArgumentException) { }
    return Task.CompletedTask;
});
await Check("restart selects XlDuck formula calls, not unrelated formulas or literal function names", () =>
{
    Equal(true, ExcelFormula.UsesXlDuck("=IF(A1,DuckScalar(DuckQuery(\"SELECT 1\")),0)"));
    Equal(true, ExcelFormula.UsesXlDuck("='C:\\Addins\\XlDuck.xll'!DuckQuery(\"SELECT 1\")"));
    Equal(true, ExcelFormula.UsesXlDuck("=DuckConnections()"));
    Equal(true, ExcelFormula.UsesXlDuck("=DuckHelp()"));
    Equal(true, ExcelFormula.UsesXlDuck("=DuckVersion()"));
    Equal(false, ExcelFormula.UsesXlDuck("=SUM(A1:A9)"));
    Equal(false, ExcelFormula.UsesXlDuck("=IF(A1,\"DuckQuery(\",\"\")"));
    Equal(false, ExcelFormula.UsesXlDuck("='DuckQuery('!A1"));
    return Task.CompletedTask;
});
await Check("preview navigation is fixed and only specific map tile URLs are permitted", () =>
{
    Equal(true, PreviewAccess.CanNavigate(PreviewAccess.Document));
    Equal(false, PreviewAccess.CanNavigate("https://xlduck.invalid/preview.html?x=1"));
    Equal(false, PreviewAccess.CanNavigate("https://example.org/"));
    Equal(true, PreviewAccess.IsMapTile("https://tile.openstreetmap.org/10/100/200.png"));
    Equal(true, PreviewAccess.IsMapTile("https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/4/1/2"));
    foreach (var uri in new[] { "http://tile.openstreetmap.org/1/2/3.png", "https://tile.openstreetmap.org/1/2/3.png?secret=x", "https://tile.openstreetmap.org.evil/1/2/3.png", "https://tile.openstreetmap.org:8443/1/2/3.png", "file:///C:/data.csv" })
        Equal(false, PreviewAccess.IsMapTile(uri));
    return Task.CompletedTask;
});
Console.WriteLine($"{passed} passed, {failed} failed.");
return failed == 0 ? 0 : 1;
