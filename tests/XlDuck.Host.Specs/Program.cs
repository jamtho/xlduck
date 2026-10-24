// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using XlDuck.Core;
using XlDuck.Host;

using var files = new SessionDirectory();
var storePath = Path.Combine(files.Path, "approvals");
var store = new WorkbookApprovalStore(storePath);
var workbook = Path.Combine(files.Path, "research.xlsx");
File.WriteAllText(workbook, "saved workbook revision");
int passed = 0, failed = 0;
async Task Check(string name, Func<Task> action)
{
    try { await action(); passed++; Console.WriteLine("PASS " + name); }
    catch (Exception error) { failed++; Console.WriteLine("FAIL " + name + ": " + error.Message); }
}
void Equal<T>(T expected, T actual) { if (!EqualityComparer<T>.Default.Equals(expected, actual)) throw new Exception($"Expected {expected}, got {actual}"); }
await Check("unknown workbooks are disabled; approval preserves non-secret settings", async () =>
{
    Equal(true, await store.ReadAsync(workbook) is null);
    var settings = new WorkbookSettings([new("s3://research/", S3CredentialSource.AwsProfile, Profile: "research")]);
    await store.RememberAsync(workbook, settings);
    Equal("research", (await new WorkbookApprovalStore(storePath).ReadAsync(workbook))!.S3[0].Profile);
});
await Check("copied workbooks and externally modified revisions do not inherit approval", async () =>
{
    var copy = Path.Combine(files.Path, "copy.xlsx"); File.Copy(workbook, copy);
    Equal(true, await store.ReadAsync(copy) is null);
    File.AppendAllText(workbook, "changed SQL");
    Equal(true, await store.ReadAsync(workbook) is null);
    await store.RememberAsync(workbook, new());
    Equal(true, await store.ReadAsync(workbook) is not null);
    store.Forget(workbook);
    Equal(true, await store.ReadAsync(workbook) is null);
});
await Check("corrupt approval records fail closed and interrupted writes leave no temporary record", async () =>
{
    await store.RememberAsync(workbook, new());
    File.WriteAllText(Directory.GetFiles(storePath, "*.json").Single(), "not json");
    Equal(true, await store.ReadAsync(workbook) is null);
    using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
    try { await store.RememberAsync(workbook, new(), cancelled.Token); throw new Exception("Expected cancellation"); } catch (OperationCanceledException) { }
    Equal(0, Directory.GetFiles(storePath, "*.tmp").Length);
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
