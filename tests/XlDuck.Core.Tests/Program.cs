// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0

using System.Globalization;
using System.Reflection;
using XlDuck;
using XlDuck.Preview;

var allowKnownDefects = args.Contains("--allow-known-defects");
int passed = 0, failed = 0, knownFailures = 0;

void Check(string name, Action test, bool knownDefect = false)
{
    try
    {
        test();
        passed++;
        Console.WriteLine($"PASS {name}");
    }
    catch (Exception ex)
    {
        if (ex is TargetInvocationException { InnerException: not null }) ex = ex.InnerException;
        if (knownDefect) knownFailures++; else failed++;
        Console.WriteLine($"{(knownDefect ? "KNOWN FAILURE" : "FAIL")} {name}: {ex.Message}");
    }
}
static void Equal(object? expected, object? actual)
{
    if (!Equals(expected, actual))
        throw new Exception($"Expected {expected} ({expected?.GetType().Name}), got {actual} ({actual?.GetType().Name}).");
}
static void Number(double expected, object actual)
{
    if (actual is not (double or float or decimal or int or long or short or byte))
        throw new Exception($"Expected a number, got {actual} ({actual.GetType().Name}).");
    Equal(expected, Convert.ToDouble(actual, CultureInfo.InvariantCulture));
}
static object? Private(string name, params object?[] arguments) =>
    typeof(DuckFunctions).GetMethod(name, BindingFlags.Static | BindingFlags.NonPublic)!.Invoke(null, arguments);
static string RetainedTable(string sql)
{
    var handle = QueryExecutor.ExecuteQuery(sql, []);
    ResultStore.IncrementRefCount(handle);
    return handle;
}

Check("scalar numeric computation", () => Number(42d, DuckFunctions.DuckQueryOutScalar("SELECT 6 * 7")));
Check("scalar Unicode text", () => Equal("café", DuckFunctions.DuckQueryOutScalar("SELECT ?", "café")));
Check("materialized output", () => Number(3d, DuckFunctions.DuckOut(RetainedTable("SELECT 3 AS n"))[1, 0]));
Check("numeric-looking text stays text", () => Equal("VARCHAR:00123", DuckFunctions.DuckQueryOutScalar(
    "SELECT typeof(?) || ':' || CAST(? AS VARCHAR)", "00123", "00123")), true);
Check("leading apostrophe is preserved", () => Equal("'hello", DuckFunctions.DuckQueryOutScalar("SELECT ?", "'hello")), true);
Check("empty argument does not remove subsequent arguments", () =>
{
    var result = DuckFunctions.DuckQueryOut("SELECT ? AS a, ? AS b", "", 7d);
    Equal(2, result.GetLength(0));
    Equal("", result[1, 0]);
    Number(7d, result[1, 1]);
}, true);
Check("question mark in literal is not a parameter", () =>
{
    var result = DuckFunctions.DuckQueryOut("SELECT '?' AS literal, ? AS value", 7d);
    Equal(2, result.GetLength(0));
    Equal("?", result[1, 0]);
    Number(7d, result[1, 1]);
}, true);
Check("question mark in comment is not a parameter", () =>
    Number(7d, DuckFunctions.DuckQueryOutScalar("SELECT ? AS value -- why?", 7d)), true);
Check("capture identity distinguishes number and text", () => Equal(false, Equals(
    Private("ComputeRangeHash", new object[,] { { "v" }, { 1d } }),
    Private("ComputeRangeHash", new object[,] { { "v" }, { "1" } }))), true);
Check("capture identity distinguishes null and empty text", () => Equal(false, Equals(
    Private("ComputeRangeHash", new object[,] { { "v" }, { null! } }),
    Private("ComputeRangeHash", new object[,] { { "v" }, { "" } }))), true);
Check("capture generates unique headers", () =>
{
    var handle = DuckFunctions.CaptureRangeInternal(new object[,] { { "a", "a", "a_2" }, { 1d, 2d, 3d } });
    var columns = ResultStore.Get(handle)!.ColumnNames;
    Equal(3, columns.Distinct(StringComparer.OrdinalIgnoreCase).Count());
}, true);
Check("quoted column names survive output", () =>
{
    var result = DuckFunctions.DuckOut(RetainedTable("SELECT 1 AS \"a\"\"b\""));
    Equal("a\"b", result[0, 0]);
    Number(1d, result[1, 0]);
}, true);
Check("preparation failure releases references", () =>
{
    var handle = RetainedTable("SELECT 1 AS v");
    var before = ResultStore.GetRefCount(handle);
    _ = DuckFunctions.DuckQueryOut("SELECT * FROM ? WHERE v = ?", handle, "duck://table/missing");
    Equal(before, ResultStore.GetRefCount(handle));
}, true);
Check("large integers are not silently rounded", () =>
{
    var actual = DuckFunctions.DuckQueryOutScalar("SELECT 9007199254740993::HUGEINT");
    if (actual is double number && number == 9007199254740992d)
        throw new Exception("9007199254740993 was silently rounded to 9007199254740992.");
}, true);
Check("plot number representation is culture independent", () =>
{
    var plot = QueryExecutor.CreatePlot(RetainedTable("SELECT 1.25 AS value"), "bar", ["x", "value", "y", "value"]);
    var previous = CultureInfo.CurrentCulture;
    try
    {
        CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
        var value = ((PlotPreviewModel)PreviewDataProvider.GetPreview(plot)).Plot.Rows[0][0];
        Equal("1.25", Convert.ToString(value, CultureInfo.InvariantCulture));
    }
    finally { CultureInfo.CurrentCulture = previous; }
}, true);
Check("fragment plots accepted by creation can render", () =>
{
    var fragment = QueryExecutor.CreateFragment("SELECT 1 AS value", []);
    FragmentStore.IncrementRefCount(fragment);
    var plot = QueryExecutor.CreatePlot(fragment, "bar", ["x", "value", "y", "value"]);
    var data = ((PlotPreviewModel)PreviewDataProvider.GetPreview(plot)).Plot;
    if (!string.IsNullOrEmpty(data.Error)) throw new Exception(data.Error);
    Equal(1, data.Rows.Count);
}, true);

Console.WriteLine($"{passed} passed, {failed} unexpected failures, {knownFailures} known defects reproduced.");
if (knownFailures > 0 && allowKnownDefects)
    Console.WriteLine("Baseline mode allows known defects; this does not mean the product correctness tests pass.");
return failed > 0 || (knownFailures > 0 && !allowKnownDefects) ? 1 : 0;
