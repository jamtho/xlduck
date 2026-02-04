// Copyright (c) 2026 James Thompson
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace XlDuck.Preview;

/// <summary>
/// Empty interface for COM default interface (required for .NET 6+ CTP support).
/// </summary>
[ComVisible(true)]
[Guid("F9A5E8C1-2B3D-4E6F-A7B8-C9D0E1F2A3B4")]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface IPreviewPane { }

/// <summary>
/// UserControl for preview display. Currently uses simple WinForms controls.
/// </summary>
[ComVisible(true)]
[Guid("E8F4D7B2-3A1C-4E9F-8B5D-2C7A6F0E1D3B")]
[ProgId("XlDuck.PreviewPane")]
[ComDefaultInterface(typeof(IPreviewPane))]
[ClassInterface(ClassInterfaceType.None)]
public class PreviewPane : UserControl, IPreviewPane
{
    private Label _titleLabel = null!;
    private Label _handleLabel = null!;
    private TextBox _contentBox = null!;

    public PreviewPane()
    {
        Log.Write("[PreviewPane] Constructor called");
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        Log.Write("[PreviewPane] InitializeComponent");
        SuspendLayout();

        BackColor = System.Drawing.Color.White;
        Dock = DockStyle.Fill;
        Padding = new Padding(8);

        // Title label
        _titleLabel = new Label
        {
            Text = "XlDuck Preview",
            Dock = DockStyle.Top,
            Font = new System.Drawing.Font("Segoe UI", 11, System.Drawing.FontStyle.Bold),
            Height = 30,
            ForeColor = System.Drawing.Color.FromArgb(30, 30, 30)
        };

        // Handle label
        _handleLabel = new Label
        {
            Text = "",
            Dock = DockStyle.Top,
            Font = new System.Drawing.Font("Consolas", 9),
            Height = 20,
            ForeColor = System.Drawing.Color.Gray
        };

        // Content text box
        _contentBox = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            Dock = DockStyle.Fill,
            Font = new System.Drawing.Font("Consolas", 9),
            ScrollBars = ScrollBars.Both,
            BackColor = System.Drawing.Color.FromArgb(250, 250, 250),
            Text = "Select a cell containing a handle to preview"
        };

        // Add controls in reverse order (bottom to top for Dock)
        Controls.Add(_contentBox);
        Controls.Add(_handleLabel);
        Controls.Add(_titleLabel);

        ResumeLayout(false);
        Log.Write("[PreviewPane] InitializeComponent done");
    }

    /// <summary>
    /// Set the preview state.
    /// </summary>
    public void SetState(PreviewModel model)
    {
        Log.Write($"[PreviewPane] SetState: {model.Kind}");

        _titleLabel.Text = model.Title ?? "Preview";
        _handleLabel.Text = model.Handle ?? "";

        switch (model)
        {
            case EmptyPreviewModel empty:
                _contentBox.Text = empty.Message ?? "Select a handle to preview";
                break;

            case ErrorPreviewModel error:
                _contentBox.Text = $"ERROR: {error.Message}";
                _contentBox.ForeColor = System.Drawing.Color.DarkRed;
                break;

            case TablePreviewModel table:
                _contentBox.ForeColor = System.Drawing.Color.Black;
                _contentBox.Text = FormatTablePreview(table.Table);
                break;

            case FragPreviewModel frag:
                _contentBox.ForeColor = System.Drawing.Color.Black;
                _contentBox.Text = FormatFragPreview(frag.Frag);
                break;

            default:
                _contentBox.Text = $"Unknown: {model.Kind}";
                break;
        }
    }

    private static string FormatTablePreview(TablePreviewData table)
    {
        var sb = new System.Text.StringBuilder();

        sb.AppendLine($"Rows: {table.RowCount:N0}  Columns: {table.ColCount}");
        sb.AppendLine();
        sb.AppendLine("=== Schema ===");
        foreach (var col in table.Columns)
        {
            sb.AppendLine($"  {col.Name}: {col.Type}");
        }

        sb.AppendLine();
        sb.AppendLine($"=== Data (first {table.PreviewRowCount} rows) ===");

        // Header
        sb.AppendLine(string.Join("\t", table.Columns.Select(c => c.Name)));
        sb.AppendLine(new string('-', 40));

        // Rows
        foreach (var row in table.Rows)
        {
            sb.AppendLine(string.Join("\t", row.Select(v => v?.ToString() ?? "NULL")));
        }

        return sb.ToString();
    }

    private static string FormatFragPreview(FragPreviewData frag)
    {
        var sb = new System.Text.StringBuilder();

        sb.AppendLine("=== SQL ===");
        sb.AppendLine(frag.Sql);

        if (frag.Args.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine("=== Parameters ===");
            foreach (var arg in frag.Args)
            {
                sb.AppendLine($"  :{arg.Name} = {arg.Value}");
            }
        }

        return sb.ToString();
    }
}
