// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using ExcelDna.Integration;

namespace XlDuck;

public static partial class DuckFunctions
{
    private static readonly Lazy<string> WorkflowHelp = new(() => ReadHelp("Workflow.md"));
    private static readonly Lazy<string> ReferenceHelp = new(() => ReadHelp("Reference.md"));
    private static readonly Lazy<string> OfficeJsHelp = new(() => ReadHelp("OfficeJs.js"));

    [ExcelFunction(IsThreadSafe = true, Description = "Read guidance packaged with this XlDuck build. Omit topic for an index; topics: workflow, reference, officejs. Runs no queries.")]
    public static string DuckHelp([ExcelArgument(Description = "Optional: workflow, reference, officejs, or index.")] object topic)
    {
        try
        {
            var name = topic switch
            {
                null or ExcelMissing or ExcelEmpty => "index",
                string text => text.Trim().ToLowerInvariant(),
                _ => throw new ArgumentException("Help topic must be text. Use DuckHelp() for available topics.")
            };
            var content = name switch
            {
                "" or "index" => """
                    Guidance is embedded in the installed add-in and matches this build. No queries, external reads or workbook changes are performed by DuckHelp.

                    DuckHelp("workflow") - concise research workflow for humans and agents.
                    DuckHelp("reference") - function signatures, SQL capabilities, types, limits, connections and recovery.
                    DuckHelp("officejs") - JavaScript helpers for a host that supports Office.js Worksheet.evaluate.

                    Start with workflow, then consult reference as needed. Keep final analyses as worksheet formulas. Inspect schemas before large queries; retain metadata before body text. Poll DuckCellStatus(cell) to distinguish running work from errors.

                    An agent can read these functions through formula evaluation or an agreed help cell. Installing XlDuck does not automatically install an agent skill. Office.js access does not imply VBA/COM or ribbon-command access.
                    """,
                "workflow" => WorkflowHelp.Value,
                "reference" => ReferenceHelp.Value,
                "officejs" => OfficeJsHelp.Value,
                _ => throw new ArgumentException("Unknown help topic. Use DuckHelp() for available topics.")
            };
            var prefix = name == "officejs" ? "// " : "";
            var result = $"{prefix}XlDuck {DuckVersion()} help: {(name.Length == 0 ? "index" : name)}\n\n{content}";
            if (result.Length > 30000) throw new InvalidOperationException("Packaged help exceeds one cell. Use a smaller help topic.");
            return result;
        }
        catch (Exception error) { return Error(error); }
    }

    private static string ReadHelp(string resource)
    {
        using var stream = typeof(DuckFunctions).Assembly.GetManifestResourceStream("XlDuck.Help." + resource)
            ?? throw new InvalidOperationException("This build is missing packaged help.");
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}
