// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Drawing;
using System.Windows.Forms;
using XlDuck.Core;

namespace XlDuck;

internal sealed class ResourcesDialog : Form
{
    internal ResourceSettings Settings => new((int)_memory.Value, (int)_spill.Value, (int)_threads.Value);
    private readonly NumericUpDown _memory = new() { Name = "MemoryMiB", Minimum = 128, Maximum = 65536, Increment = 128, ThousandsSeparator = true };
    private readonly NumericUpDown _spill = new() { Name = "SpillGiB", Minimum = 1, Maximum = 1024 };
    private readonly NumericUpDown _threads = new() { Name = "Threads", Minimum = 1, Maximum = 16 };

    internal ResourcesDialog(string workbookName, ResourceSettings current)
    {
        current.Validate();
        SuspendLayout();
        Text = "Resources — " + workbookName;
        AutoScaleDimensions = new(96, 96); AutoScaleMode = AutoScaleMode.Dpi;
        ClientSize = new(570, 340); MinimumSize = new(500, 340); StartPosition = FormStartPosition.CenterParent;
        _memory.Value = current.MemoryMiB; _spill.Value = current.SpillGiB; _threads.Value = current.Threads;
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new(16), ColumnCount = 2, RowCount = 6 };
        layout.ColumnStyles.Add(new(SizeType.Percent, 60)); layout.ColumnStyles.Add(new(SizeType.Percent, 40));
        var explanation = new Label { AutoSize = true, MaximumSize = new(520, 0), Margin = new(0, 0, 0, 12), Text =
            "Budgets apply to this workbook's worker. Several open workbooks can use their budgets at the same time. Changes restart XlDuck for this workbook. Save the workbook to keep these budgets with the file." };
        layout.Controls.Add(explanation, 0, 0); layout.SetColumnSpan(explanation, 2);
        void Row(int row, string label, NumericUpDown input)
        {
            layout.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new(0, 6, 12, 6) }, 0, row);
            input.Dock = DockStyle.Fill; input.Margin = new(0, 6, 0, 6); layout.Controls.Add(input, 1, row);
        }
        Row(1, "DuckDB memory budget (MiB)", _memory); Row(2, "Temporary disk budget (GiB)", _spill); Row(3, "Query threads", _threads);
        var note = new Label { AutoSize = true, MaximumSize = new(520, 0), Margin = new(0, 10, 0, 10), Text =
            "Memory is a DuckDB buffer budget, not a cap on the whole process. Some operations need extra memory or cannot spill to disk. Projecting fewer columns and filtering early still matter." };
        layout.Controls.Add(note, 0, 4); layout.SetColumnSpan(note, 2);
        var footer = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
        var cancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
        var apply = new Button { Text = "Apply and restart", AutoSize = true, DialogResult = DialogResult.OK };
        footer.Controls.Add(cancel); footer.Controls.Add(apply); layout.Controls.Add(footer, 0, 5); layout.SetColumnSpan(footer, 2);
        Controls.Add(layout); AcceptButton = apply; CancelButton = cancel;
        AutoScaleDimensions = new(96, 96); ResumeLayout(true);
    }
}
