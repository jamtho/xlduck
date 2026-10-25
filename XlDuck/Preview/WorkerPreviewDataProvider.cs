// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using XlDuck.Core;

namespace XlDuck.Preview;

internal static class PreviewDataProvider
{
    internal static async Task<PreviewModel> GetPreviewAsync(WorkbookSession session, string value, CancellationToken token)
    {
        var handle = session.ResolveCellHandle(value);
        var definition = session.GetDefinition(handle);
        if (handle.Kind == ResultKind.Fragment)
            return new FragPreviewModel { Title = "SQL fragment", Handle = value, Frag = new()
            {
                Sql = definition.Sql,
                Args = definition.Arguments.Select((argument, i) => new FragmentArg { Name = "?" + (i + 1), Value = argument.Handle is { } dependency ? session.GetCellText(dependency) : argument.Value!.Text }).ToList()
            } };
        using var lease = session.AcquireData(handle);
        await lease.WaitAsync(token).ConfigureAwait(false);
        var page = await session.ReadAsync(lease.Handle, new(ReadPurpose.Preview, MaxRows: handle.Kind == ResultKind.Plot ? 5000 : 200), token).ConfigureAwait(false);
        var rows = page.Rows.Select(row => row.Select(cell => cell.Kind is ValueKind.Null or ValueKind.Empty ? null : cell.Text).ToArray()).ToList();
        if (handle.Kind == ResultKind.Plot)
            return new PlotPreviewModel { Title = "Chart", Handle = value, Plot = new()
            {
                Template = definition.PlotTemplate!, Overrides = definition.PlotOptions!.ToDictionary(),
                Columns = page.Columns.Select(column => column.Name).ToList(), Types = page.Columns.Select(column => column.Type).ToList(),
                Rows = rows, Kinds = page.Rows.Select(row => row.Select(cell => cell.Kind.ToString()).ToArray()).ToList(),
                RowCount = page.TotalRows, IsSample = page.IsSample
            } };
        return new TablePreviewModel { Title = "Table", Handle = value, Table = new()
        {
            Columns = page.Columns.Select(column => new ColumnInfo { Name = column.Name, Type = column.Type }).ToList(),
            Rows = rows, RowCount = page.TotalRows, ColCount = page.Columns.Length, PreviewRowCount = rows.Count,
            Sql = definition.Sql, IsSample = page.IsSample
        } };
    }
}
