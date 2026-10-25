// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Drawing;
using System.Windows.Forms;
using XlDuck.Core;
using XlDuck.Host;

namespace XlDuck;

internal sealed class ConnectionsDialog : Form
{
    private sealed record Entry(S3Connection Connection, S3AccessKey? Key = null);
    private readonly List<Entry> _entries;
    private readonly ListView _list = new() { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, MultiSelect = false, HideSelection = false };
    internal ConnectionsDialog(OpenWorkbook book)
    {
        SuspendLayout();
        Text = "Connections — " + book.Name;
        AutoScaleDimensions = new(96, 96); AutoScaleMode = AutoScaleMode.Dpi;
        ClientSize = new(780, 380); MinimumSize = new(620, 320); StartPosition = FormStartPosition.CenterParent;
        _entries = book.Settings.S3.Select(connection => new Entry(connection)).ToList();
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new(12), ColumnCount = 1, RowCount = 4 };
        layout.RowStyles.Add(new(SizeType.AutoSize)); layout.RowStyles.Add(new(SizeType.Percent, 100));
        layout.RowStyles.Add(new(SizeType.AutoSize)); layout.RowStyles.Add(new(SizeType.AutoSize));
        layout.Controls.Add(new Label { AutoSize = true, MaximumSize = new(720, 0), Margin = new(0, 0, 0, 12), Text =
            "Connect to S3-compatible storage. Settings are kept on this computer for this workbook path; access keys are stored in Windows Credential Manager. Local files and public HTTP sources need no connection entry." }, 0, 0);
        _list.Columns.Add("Bucket / path", 260); _list.Columns.Add("Endpoint", 270); _list.Columns.Add("Credentials", 170);
        _list.SizeChanged += (_, _) =>
        {
            var width = Math.Max(300, _list.ClientSize.Width - SystemInformation.VerticalScrollBarWidth);
            _list.Columns[0].Width = width * 35 / 100; _list.Columns[1].Width = width * 40 / 100; _list.Columns[2].Width = width * 25 / 100;
        };
        layout.Controls.Add(_list, 0, 1);
        var actions = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill };
        Button Button(string text, Action action) { var button = new Button { Text = text, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink }; button.Click += (_, _) => Guard(action); return button; }
        actions.Controls.Add(Button("Add…", () => Edit(-1)));
        actions.Controls.Add(Button("Edit…", () => { if (_list.SelectedIndices.Count == 1) Edit(_list.SelectedIndices[0]); }));
        actions.Controls.Add(Button("Remove", () => { if (_list.SelectedIndices.Count == 1) { _entries.RemoveAt(_list.SelectedIndices[0]); RefreshList(); } }));
        _list.DoubleClick += (_, _) => Guard(() => { if (_list.SelectedIndices.Count == 1) Edit(_list.SelectedIndices[0]); });
        layout.Controls.Add(actions, 0, 2);
        var footer = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
        var cancel = new Button { Text = "Cancel", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, DialogResult = DialogResult.Cancel };
        var apply = Button("Apply and restart", () =>
        {
            var settings = book.Settings with { S3 = [.. _entries.Select(entry => entry.Connection)] }; settings.Validate();
            foreach (var entry in _entries.Where(entry => entry.Key is not null)) CredentialStore.Save(entry.Connection.CredentialId!, entry.Key!);
            WorkbookHost.ApplySettings(book, settings); DialogResult = DialogResult.OK; Close();
        });
        footer.Controls.Add(cancel); footer.Controls.Add(apply); layout.Controls.Add(footer, 0, 3);
        Controls.Add(layout); AcceptButton = apply; CancelButton = cancel; RefreshList();
        AutoScaleDimensions = new(96, 96); ResumeLayout(true);
    }
    private void Guard(Action action)
    {
        try { action(); } catch (Exception error) { MessageBox.Show(this, error.Message, "Connection settings", MessageBoxButtons.OK, MessageBoxIcon.Information); }
    }
    private void RefreshList()
    {
        _list.Items.Clear();
        foreach (var entry in _entries) _list.Items.Add(new ListViewItem([entry.Connection.Scope, entry.Connection.Endpoint ?? "Default S3 endpoint", ConnectionEditor.CredentialLabel(entry.Connection.Credentials)]));
    }
    private void Edit(int index)
    {
        var entry = index < 0 ? null : _entries[index];
        using var editor = new ConnectionEditor(entry?.Connection, entry?.Key);
        if (editor.ShowDialog(this) != DialogResult.OK) return;
        var edited = new Entry(editor.Connection!, editor.Key);
        if (index < 0) _entries.Add(edited); else _entries[index] = edited;
        RefreshList();
    }
}

internal sealed class ConnectionEditor : Form
{
    private static readonly S3CredentialSource[] Sources = [S3CredentialSource.AccessKey, S3CredentialSource.Environment, S3CredentialSource.Anonymous, S3CredentialSource.AwsProfile, S3CredentialSource.AwsSso];
    internal static string CredentialLabel(S3CredentialSource source) => source switch
    {
        S3CredentialSource.AccessKey => "Saved access key", S3CredentialSource.Environment => "Environment variables",
        S3CredentialSource.Anonymous => "Anonymous", S3CredentialSource.AwsProfile => "Named credential profile", _ => "SSO profile"
    };
    internal S3Connection? Connection { get; private set; }
    internal S3AccessKey? Key { get; private set; }
    internal ConnectionEditor(S3Connection? existing, S3AccessKey? editedKey)
    {
        SuspendLayout();
        Text = existing is null ? "Add S3-compatible connection" : "Edit S3-compatible connection";
        AutoScaleDimensions = new(96, 96); AutoScaleMode = AutoScaleMode.Dpi;
        ClientSize = new(580, 440); MinimumSize = new(530, 320); StartPosition = FormStartPosition.CenterParent;
        var outer = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, Padding = new(12) };
        outer.RowStyles.Add(new(SizeType.Percent, 100)); outer.RowStyles.Add(new(SizeType.AutoSize));
        var scroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
        var fields = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2 };
        fields.ColumnStyles.Add(new(SizeType.Absolute, 155)); fields.ColumnStyles.Add(new(SizeType.Percent, 100));
        void Row(string label, Control control)
        {
            var row = fields.RowCount++; fields.RowStyles.Add(new(SizeType.AutoSize));
            fields.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new(0, 6, 8, 6) }, 0, row);
            control.Dock = DockStyle.Fill; control.Margin = new(0, 5, 0, 5); fields.Controls.Add(control, 1, row);
        }
        var scope = new TextBox { Name = "Scope", Text = existing?.Scope ?? "s3://", PlaceholderText = "s3://bucket/ or s3://bucket/path/" };
        var endpoint = new TextBox { Name = "Endpoint", Text = existing?.Endpoint ?? "", PlaceholderText = "storage.example.com:9000" };
        var source = new ComboBox { Name = "Credentials", DropDownStyle = ComboBoxStyle.DropDownList };
        source.Items.AddRange(Sources.Select(CredentialLabel).ToArray()); source.SelectedIndex = Array.IndexOf(Sources, existing?.Credentials ?? S3CredentialSource.AccessKey);
        var keyId = new TextBox { Name = "KeyId" }; var secret = new TextBox { Name = "Secret", UseSystemPasswordChar = true }; var token = new TextBox { Name = "SessionToken", UseSystemPasswordChar = true };
        var profile = new TextBox { Name = "Profile", Text = existing?.Profile ?? "" };
        var tls = new CheckBox { Name = "Tls", Text = "Use HTTPS", Checked = existing?.UseSsl ?? true, AutoSize = true };
        var pathStyle = new CheckBox { Name = "PathStyle", Text = "Put bucket in the URL path", Checked = existing?.PathStyle ?? true, AutoSize = true };
        var advanced = new CheckBox { Name = "Advanced", Text = "Advanced settings", AutoSize = true };
        Row("Bucket / path", scope); Row("Endpoint", endpoint); Row("Credentials", source);
        Row("Access key ID", keyId); Row("Secret access key", secret); Row("", advanced);
        var advancedPanel = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
        void AdvancedField(string label, Control control) { advancedPanel.Controls.Add(new Label { Text = label, AutoSize = true }); control.Width = 330; advancedPanel.Controls.Add(control); }
        advancedPanel.Controls.Add(tls); advancedPanel.Controls.Add(pathStyle);
        AdvancedField("Session token (optional)", token); AdvancedField("Profile name (profile credentials only)", profile);
        Row("", advancedPanel);
        var help = new Label { AutoSize = true, MaximumSize = new(350, 0) };
        string? credentialNotice = null;
        Row("", help);
        void UpdateFields()
        {
            var chosen = Sources[source.SelectedIndex]; var accessKey = chosen == S3CredentialSource.AccessKey;
            keyId.Enabled = secret.Enabled = token.Enabled = accessKey;
            profile.Enabled = chosen is S3CredentialSource.AwsProfile or S3CredentialSource.AwsSso;
            if (profile.Enabled) advanced.Checked = true;
            advancedPanel.Visible = advanced.Checked;
            help.Text = credentialNotice ?? (chosen == S3CredentialSource.Environment
                ? "Use AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY and optional AWS_SESSION_TOKEN. These environment-variable names also apply to S3-compatible servers."
                : "Endpoint: your storage server's hostname and optional port. Leave it blank only for Amazon S3.");
        }
        source.SelectedIndexChanged += (_, _) => UpdateFields(); advanced.CheckedChanged += (_, _) => UpdateFields();
        if (existing?.Credentials == S3CredentialSource.AccessKey)
        {
            try
            {
                var saved = editedKey ?? CredentialStore.Read(existing.CredentialId!);
                keyId.Text = saved.KeyId; secret.Text = saved.Secret; token.Text = saved.SessionToken ?? "";
            }
            catch (InvalidOperationException) { credentialNotice = "The saved key is unavailable. Enter it again to repair this connection."; }
        }
        UpdateFields(); scroll.Controls.Add(fields); outer.Controls.Add(scroll, 0, 0);
        var buttons = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
        var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
        var done = new Button { Name = "Done", Text = "Done", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
        done.Click += (_, _) =>
        {
            try
            {
                var chosen = Sources[source.SelectedIndex];
                var credentialId = chosen == S3CredentialSource.AccessKey ? existing?.CredentialId ?? Guid.NewGuid().ToString("N") : null;
                // Signing is a protocol detail. Keep provider overrides already set
                // through host configuration; ordinary users need only their endpoint.
                var connection = new S3Connection(scope.Text.Trim(), chosen, existing?.Region ?? S3Connection.DefaultSigningRegion, NullIfEmpty(endpoint.Text),
                    chosen is S3CredentialSource.AwsProfile or S3CredentialSource.AwsSso ? NullIfEmpty(profile.Text) : null, tls.Checked, pathStyle.Checked, credentialId);
                connection.Validate();
                var key = chosen == S3CredentialSource.AccessKey ? new S3AccessKey(keyId.Text.Trim(), secret.Text, string.IsNullOrEmpty(token.Text) ? null : token.Text) : null;
                key?.Validate(); Connection = connection; Key = key; DialogResult = DialogResult.OK; Close();
            }
            catch (Exception error) { MessageBox.Show(this, error.Message, "Connection settings", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        };
        buttons.Controls.Add(cancel); buttons.Controls.Add(done); outer.Controls.Add(buttons, 0, 1);
        Controls.Add(outer); AcceptButton = done; CancelButton = cancel;
        AutoScaleDimensions = new(96, 96); ResumeLayout(true);
    }
    private static string? NullIfEmpty(string value) => string.IsNullOrWhiteSpace(value) ? null : value.Trim();
}
