# Install XlDuck for Windows Excel

XlDuck is a Windows x64 Excel add-in for read-and-compute SQL research. Install it once and Excel loads it on normal startup. Workbooks remain ordinary `.xlsx` files. Your AI provider's Excel add-in is optional and installed separately.

Download the complete Windows x64 ZIP from [1.0.1](https://github.com/jamtho/xlduck/releases/tag/v1.0.1), or build from source with `scripts/Build-Package.ps1`. This is an **unsigned release**. Do not use an older packed XLL as a substitute for this distribution. The release lists its tested scope and deferred compatibility trials.

## Install in a few steps

1. Use **64-bit desktop Excel on Windows x64**, with the **x64 .NET 10 Desktop Runtime** and **WebView2 Evergreen Runtime**. Microsoft 365 / Excel 2021 or 2024 use the Office 16.x registration supported by setup. Mac, browser, 32-bit and ARM Excel are not supported. Install and open Excel at least once before setup.
2. Review the download's origin. If Windows marks the ZIP as downloaded, use **Properties > General > Unblock** only after deciding to trust it, then extract the entire folder. XlDuck setup does not unblock files or change Trust Center. An organization's policies still apply.
3. Open **START-HERE.html** in the extracted folder for prerequisite links and optional agent setup. Save your work and **close Excel yourself**.
4. Double-click **Install.cmd** from **Windows File Explorer**. No administrator rights are needed for XlDuck itself. If a prerequisite is missing, setup prints a link and stops without changing registration. Install the prerequisite and rerun it.
5. Open Excel from the Windows Start menu or taskbar and perform the smoke test below. It should also work after closing and restarting Excel.

Download prerequisites from Microsoft: [.NET 10 Desktop Runtime](https://dotnet.microsoft.com/en-us/download/dotnet/10.0) (Windows x64, under “Run apps”; no SDK needed), and [WebView2 Evergreen Runtime](https://developer.microsoft.com/en-us/microsoft-edge/webview2/). Your organization may manage their installation. An existing .NET 8/9 runtime or a private development runtime is not the public prerequisite.

Setup verifies the complete package, copies it to a stable directory, and registers it for the current user's Excel startup:

```text
%LOCALAPPDATA%\XlDuck\install\
    state.json
    versions\package-<content-id>\
        XlDuck-AddIn64.xll
        worker\
        readers\
        preview\
        ...the remaining packaged files
```

The content ID identifies the entire package manifest, including its documentation. Installing the same package again reuses its directory. Do not move the installed folder or copy just the XLL. Keep the extracted source ZIP until verification succeeds; Excel loads the stable copy, so the source ZIP is not required afterwards.

Registration uses Excel-DNA's documented `HKCU\Software\Microsoft\Office\16.0\Excel\Options` `OPEN`, `OPEN1`, ... mechanism. Setup preserves other entries' values, types and relative load order and keeps the numbered list contiguous. It saves a recovery snapshot before changing it, checks for Excel reopening, and restores the prior list if committing installation state fails. It does not touch security settings, Disabled Items, calculation mode, RTD, workbooks or credentials. See [Excel-DNA installation guidance](https://excel-dna.net/docs/guides-basic/installing-your-add-in/).

If an older XlDuck was registered manually, setup stops rather than loading both. Open **File > Options > Add-ins > Manage: Excel Add-ins > Go**, uncheck that entry, save and fully close Excel, then rerun Install.cmd. Do not load two versions in the same Excel process.

For a read-only check from the extracted package directory:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\Setup.ps1 -Action Check
```

This checks package integrity, installed prerequisites and whether Excel is closed. A failure returns a nonzero exit code with the next step. Hashes detect incomplete or changed files; they do not authenticate an unsigned publisher. Setup does not download or silently install dependencies. If managed policy blocks these scripts or the XLL, use your organization's deployment route; do not change that policy to run setup.

Run setup and its startup check outside packaged development apps. On this workstation, setup launched from Codex wrote an Excel startup entry visible to Codex and its child processes but absent from the registry view used by Excel launched from Windows Explorer. Reading the entry back or launching Excel again from the same development app gave a false positive. If this occurs, close Excel and rerun **Install.cmd** from File Explorer, then start Excel from its Windows icon. Windows supports private registry views for packaged apps; see [Microsoft's registry virtualization documentation](https://learn.microsoft.com/en-us/windows/msix/desktop/flexible-virtualization). No Trust Center change is needed.

## Verify installation

In a blank ordinary workbook, enter these formulas in separate cells:

```excel
=DuckVersion()
=DuckQueryOutScalar("SELECT 42")
=DuckHelp()
```

The query should settle to `42` after initialization; help should show the installed documentation topics. To check charts, enter this in another cell, select its handle and choose **XlDuck > Preview Pane**:

```excel
=DuckPlot(DuckQuery("SELECT range AS n FROM range(5)"),"line","x","n","y","n")
```

A five-point line should appear. Save the workbook, close Excel normally and reopen it to verify persistent startup loading. Ordinary workbooks initialize automatically; Protected View prevents queries until the user chooses Excel's **Enable Editing**. Results rebuild when a saved workbook opens. Source connections are configured through **XlDuck > Connections**.

## Optional agent setup

For a first investigation, open `examples/research-demo/XlDuck Research Demo.xlsx` from the extracted package and save a working copy under Documents. It contains synthetic data and needs no connection settings. Select C11 or C12 on Research and open the preview pane; edit the amber inputs to explore the dependency graph.

Open **START-HERE.html** or the [agent integration guide](AGENT-INTEGRATION.md#install-the-bootstrap-skill). The package includes `agent/xlduck-skill.zip` for supported provider uploads and **Install Codex Skill.cmd** for local Codex. These are separate actions from installing the SQL add-in. The tiny skill reads detailed guidance from the running add-in; it does not contain a frozen copy of the SQL API. Verify it in a fresh conversation before relying on automatic discovery.

## Upgrade and roll back

Save and close Excel, then run **Install.cmd** from a new extracted package. Setup switches the registered path only after copying and verifying the new distribution. The preceding version remains available. Open Excel, repeat the smoke test and reopen a saved research workbook.

To undo an upgrade, close Excel and run **Rollback.cmd** from an installed or extracted package. It switches to the preceding managed installation and retains both versions. This does not undo changes you made to your workbook. Handles are session values; saved formulas rebuild them.

There is no automatic updater. Run setup as the normal Windows account that uses Excel, not under a different administrator account. Setup refuses to operate while any Excel process is running; it never terminates one or suppresses save prompts.

## Uninstall

Close Excel and run **Uninstall.cmd**. It removes startup entries for versions managed by this installer. It preserves unrelated add-ins, including entries installed since XlDuck was installed. Recovery packages and installation history remain under `%LOCALAPPDATA%\XlDuck\install`; after checking for files you placed there, you can delete that folder yourself. It is deliberately separate from your connections and other data. There is no Windows Installed Apps entry yet.

For a manually registered version, uncheck it in Excel's Add-ins dialog before closing Excel and deleting its package folder. Do not delete an XLL while it is still registered or loaded.

Agent skills are removed separately through the provider's UI, or by removing the local `%USERPROFILE%\.agents\skills\xlduck` folder after checking for your customizations. XlDuck setup does not remove them automatically.

Workbook connection metadata, local credential associations under `%LOCALAPPDATA%\XlDuck\credential-bindings`, legacy settings under `%LOCALAPPDATA%\XlDuck\connections`, saved keys in Windows Credential Manager, logs and WebView data remain intact. Remove those only deliberately. Never remove workbooks, source data, AWS profiles or shared runtimes as part of add-in removal.

## Manual installation

If you prefer Excel's UI, extract the complete package to a stable versioned folder. Open **File > Options > Add-ins**, select **Excel Add-ins** and **Go**, then **Browse** to `XlDuck-AddIn64.xll`. Leave it checked. This is not the COM Add-ins list or Office Store. Follow normal Excel trust prompts after verifying the package. Perform the same smoke test and restart check.

To upgrade manually, uncheck the old entry, save and fully close Excel, then restart and browse to the new version. Unchecking does not unload all code from the process. Keep the old folder for rollback. Merely opening an XLL or running the repository's demo/development launcher is temporary and does not install it for startup. The development launcher's private portable runtime is not part of this public route.

## Troubleshooting and release status

| Symptom | Check |
| --- | --- |
| Setup says Excel is open | Save and close your Excel windows, then retry. Setup never closes them. |
| Another XlDuck build is registered | Uncheck the manual/development entry in Excel Add-ins, fully exit Excel, then retry. |
| No ribbon or `#NAME?` | Check the registered XLL path, x64 Excel/runtime, complete package and any Excel security or Disabled Items notice. |
| Works when an agent launches Excel, but not from the Windows icon | Close Excel and rerun Install.cmd from File Explorer. A packaged development app can see a private startup registration that ordinary Excel cannot see. Verify from the Windows icon, not the agent's launcher. |
| Runtime/load error | Install the x64 Desktop Runtime. Other .NET add-ins can affect the runtime Excel selects; see [coexistence](RELEASE.md#runtime-and-coexistence). |
| Preview is blank | Check WebView2 and the packaged preview assets. |
| Query is pending or fails | Inspect `DuckCellStatus(producingCell)` and `DuckStatus()`; source configuration, Protected View and manual calculation are separate from installation. |
| Skill missing inside Excel | Verify the provider account, plan and skill picker. A local Codex skill does not install into the Excel sidebar. Use the documented starting prompt if uploads are unavailable. |

From the installed or extracted package directory:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\Get-Diagnostics.ps1 -PackageDirectory .
```

This prints a report without attaching to Excel or collecting workbook contents, credentials or connection settings. Include the package version and relevant `DuckStatus()` output when reporting a problem. Review error text before sharing because it may contain source paths or data.

This is an unsigned per-user installer. Installation and ordinary Windows Excel startup have been exercised on the development workstation using the system runtime. Clean-machine certification, code signing and an automatic updater remain follow-up work; see [tested scope](RELEASE.md#tested-scope-and-follow-up-work).
