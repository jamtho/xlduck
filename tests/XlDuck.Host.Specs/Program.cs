// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using XlDuck.Core;
using XlDuck.Host;

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
    Equal(true, PreviewAccess.IsMapTile("https://a.tile.openstreetmap.org/10/100/200.png"));
    Equal(true, PreviewAccess.IsMapTile("https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/4/1/2"));
    foreach (var uri in new[] { "http://a.tile.openstreetmap.org/1/2/3.png", "https://a.tile.openstreetmap.org/1/2/3.png?secret=x", "https://a.tile.openstreetmap.org.evil/1/2/3.png", "https://a.tile.openstreetmap.org:8443/1/2/3.png", "file:///C:/data.csv" })
        Equal(false, PreviewAccess.IsMapTile(uri));
    return Task.CompletedTask;
});
Console.WriteLine($"{passed} passed, {failed} failed.");
return failed == 0 ? 0 : 1;
