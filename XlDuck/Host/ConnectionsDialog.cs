// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Windows.Forms;
using XlDuck.Core;
using XlDuck.Host;

namespace XlDuck;

internal sealed class ConnectionsDialog : Form
{
    private readonly DataGridView _connections = new() { Dock = DockStyle.Fill, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, AllowUserToAddRows = true, RowHeadersVisible = false };
    internal ConnectionsDialog(OpenWorkbook book)
    {
        Text = "XlDuck — " + book.Name; Width = 1000; Height = 460; StartPosition = FormStartPosition.CenterParent;
        var help = new Label { Dock = DockStyle.Top, Height = 104, Padding = new Padding(12), Text =
            "Enabled workbooks can read local files and supported network sources using your account. SQL can compute results but cannot write data or install extensions.\n\n" +
            "Add S3 connections below if needed. Credentials stay in environment variables or named AWS profiles; authenticate SSO outside Excel. Saving remembers this workbook revision on this computer. A copy at another path starts disabled." };
        _connections.Columns.Add("scope", "S3 scope");
        _connections.Columns.Add(new DataGridViewComboBoxColumn { Name = "credentials", HeaderText = "Credentials", DataSource = Enum.GetNames<S3CredentialSource>() });
        _connections.Columns.Add("profile", "AWS profile"); _connections.Columns.Add("region", "Region"); _connections.Columns.Add("endpoint", "Endpoint (optional)");
        _connections.Columns.Add(new DataGridViewCheckBoxColumn { Name = "pathStyle", HeaderText = "Path style" });
        foreach (var connection in book.Settings.S3)
            _connections.Rows.Add(connection.Scope, connection.Credentials.ToString(), connection.Profile, connection.Region, connection.Endpoint, connection.PathStyle);
        _connections.DefaultValuesNeeded += (_, args) => { args.Row.Cells[1].Value = nameof(S3CredentialSource.Environment); args.Row.Cells[3].Value = "us-east-1"; args.Row.Cells[5].Value = false; };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 48, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(6) };
        var apply = new Button { Text = book.Enabled ? "Apply and restart" : "Enable workbook", AutoSize = true };
        var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
        apply.Click += (_, _) =>
        {
            try
            {
                _connections.EndEdit();
                var connections = new List<S3Connection>();
                foreach (DataGridViewRow row in _connections.Rows)
                {
                    if (row.IsNewRow) continue;
                    string Value(int column) => Convert.ToString(row.Cells[column].Value)?.Trim() ?? "";
                    var scope = Value(0); if (scope.Length == 0) throw new ArgumentException("Each connection needs an S3 scope. Delete unused rows.");
                    if (!Enum.TryParse<S3CredentialSource>(Value(1), out var source)) throw new ArgumentException("Choose a credential source.");
                    connections.Add(new(scope, source, Value(3), Value(4).Length == 0 ? null : Value(4), Value(2).Length == 0 ? null : Value(2), PathStyle: row.Cells[5].Value is true));
                }
                WorkbookHost.Enable(book, new([.. connections])); DialogResult = DialogResult.OK; Close();
            }
            catch (Exception error) { MessageBox.Show(this, error.Message, "Connection settings", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        };
        buttons.Controls.Add(apply); buttons.Controls.Add(cancel);
        Controls.Add(_connections); Controls.Add(help); Controls.Add(buttons); AcceptButton = apply; CancelButton = cancel;
    }
}
