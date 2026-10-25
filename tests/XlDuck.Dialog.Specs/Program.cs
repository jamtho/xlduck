// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using XlDuck.Core;
using XlDuck.Host;

namespace XlDuck;

// Only the connection dialog is hosted here. These stand-ins never automate Excel or start a worker.
internal sealed class OpenWorkbook(string name, WorkbookSettings settings)
{
    internal string Name = name;
    internal WorkbookSettings Settings = settings;
}
internal static class WorkbookHost
{
    internal static void ApplySettings(OpenWorkbook book, WorkbookSettings settings) => book.Settings = settings;
}
internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
        Application.EnableVisualStyles();
        try
        {
            var output = args.FirstOrDefault() ?? Path.GetFullPath("artifacts/dialog-review"); Directory.CreateDirectory(output);
            using (var editor = new ConnectionEditor(null, null))
            {
                ShowOwned(editor);
                Assert(editor.ClientSize.Width >= 578 * editor.DeviceDpi / 96, "Dialog dimensions did not scale with the display DPI");
                Get<TextBox>(editor, "Scope").Text = "s3://research/";
                Get<TextBox>(editor, "Endpoint").Text = "storage.example.test:9000";
                Get<TextBox>(editor, "KeyId").Text = "fixture-access-key";
                Get<TextBox>(editor, "Secret").Text = "fixture-secret";
                Assert(Get<TextBox>(editor, "Secret").UseSystemPasswordChar, "Secret is not masked");
                Render(editor, Path.Combine(output, "connection.png"));
                Get<CheckBox>(editor, "Advanced").Checked = true; Application.DoEvents();
                Render(editor, Path.Combine(output, "connection-advanced.png"));
                Get<Button>(editor, "Done").PerformClick();
                Assert(editor.Connection is { Credentials: S3CredentialSource.AccessKey, UseSsl: true, PathStyle: true, Region: "us-east-1" }, "Incorrect connection defaults");
                Assert(editor.Connection!.Endpoint == "storage.example.test:9000" && editor.Key?.Secret == "fixture-secret", "Editor lost connection fields");
                // Done stages the key; only the outer dialog's Apply writes Credential Manager.
                try { CredentialStore.Read(editor.Connection.CredentialId!); throw new Exception("Editor saved a key before Apply"); }
                catch (InvalidOperationException) { }
                Console.WriteLine("PASS access-key editor masks secrets and stages settings without persisting credentials");
            }
            using (var editor = new ConnectionEditor(null, null))
            {
                ShowOwned(editor);
                Get<TextBox>(editor, "Scope").Text = "s3://public/";
                Get<ComboBox>(editor, "Credentials").SelectedIndex = 2;
                Assert(!Get<TextBox>(editor, "KeyId").Enabled && !Get<TextBox>(editor, "Secret").Enabled, "Anonymous mode still requests credentials");
                Get<CheckBox>(editor, "Advanced").Checked = true;
                Get<CheckBox>(editor, "Tls").Checked = false;
                Get<TextBox>(editor, "Region").Text = "local-1";
                Get<Button>(editor, "Done").PerformClick();
                Assert(editor.Connection is { Credentials: S3CredentialSource.Anonymous, UseSsl: false, Region: "local-1", CredentialId: null } && editor.Key is null, "Advanced or anonymous settings were lost");
                Console.WriteLine("PASS anonymous and advanced settings round-trip without an AWS profile");
            }
            using (var dialog = new ConnectionsDialog(new("Research.xlsx", new([new("s3://research/", S3CredentialSource.Environment, Endpoint: "storage.example.test:9000", PathStyle: true)]))))
            {
                ShowOwned(dialog); Render(dialog, Path.Combine(output, "connections.png")); dialog.Close();
            }
            Console.WriteLine("PASS connection list renders; three visual review images saved");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }
    private static T Get<T>(Control parent, string name) where T : Control => (T)parent.Controls.Find(name, true).Single();
    private static void ShowOwned(Form form)
    {
        form.ShowInTaskbar = false; form.StartPosition = FormStartPosition.Manual; form.Location = new(-30000, -30000);
        form.Show(); Application.DoEvents();
    }
    private static void Render(Form form, string path)
    {
        using var bitmap = new Bitmap(form.Width, form.Height);
        form.DrawToBitmap(bitmap, new(0, 0, bitmap.Width, bitmap.Height)); bitmap.Save(path);
    }
    private static void Assert(bool condition, string message) { if (!condition) throw new Exception(message); }
}
