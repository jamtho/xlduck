# Release process and tested scope

**XlDuck 1.0.0** is the public release of the research rewrite. The Windows x64 distribution contains the Excel-DNA host, separate DuckDB worker, pinned readers, bundled chart libraries, synthetic demo and agent bootstrap. Keep the complete folder intact; an XLL alone is not a distribution. The release is unsigned.

## Runtime and coexistence

The public prerequisites are **64-bit desktop Excel on Windows x64**, the **x64 .NET 10 Desktop Runtime**, and **WebView2 Evergreen Runtime** for the task pane. Mac, browser, 32-bit and ARM Excel are not supported. The installer recognizes Office 16.x registration; the broader Excel version matrix has not been certified.

The binaries retain the .NET 8 API baseline and use `LatestMajor` roll-forward. `DuckStatus()` reports the runtime actually selected inside Excel; the worker selects its runtime independently. Development probes passed with .NET 10 and .NET Framework 4.8 companions in both load orders, and with a .NET 8 companion loaded first. Loading XlDuck first and then a companion pinned to .NET 8 blocked in native loading. These are synthetic probes, not certification of third-party add-ins. An incompatible runtime requirement may need a separate Excel instance or a vendor update. See [Excel-DNA runtime guidance](https://excel-dna.net/docs/guides-basic/dotnet-runtime-support/).

## Build and verify a package

Build from a clean checkout of the intended release commit:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/Get-ReaderExtensions.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/Get-PreviewAssets.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Run-CoreTests.ps1 -WithReaders
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/Build-Package.ps1 -Destination artifacts/release-package
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-Package.ps1 -PackageDirectory artifacts/release-package
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-Setup.ps1 -PackageDirectory artifacts/release-package
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Run-Tests.ps1 -NoBuild -WithPreview -AddInPath artifacts/release-package/XlDuck-AddIn64.xll
```

Also run the browser and agent-helper checks described in [tests](../tests/README.md), and the synthetic demo regression. Excel tests create their own processes and workbooks, preserve unrelated work, and verify clean process exit. They do not change Trust Center, global calculation mode or RTD latency.

Packaging builds into a fresh directory using NuGet lockfiles, verifies reader and preview dependency hashes, and includes pinned upstream licence notices. It verifies both folder and ZIP against a SHA-256 manifest and an independent required-file baseline. Missing dependencies, changed or unlisted files, links and noncanonical archive entries fail verification. Hashes detect incomplete or changed distributions; they do not authenticate an unsigned publisher.

`build-info.json` records the source commit, tracked/untracked changes, SDK and build time. Published packages must identify the clean tagged commit. This makes provenance reviewable; it does not promise byte-identical ZIP timestamps. The original user-supplied skill archive and private research workbook edits must stay outside release changes. The distributed skill ZIP is generated from `skill/xlduck/SKILL.md`.

## Installation, upgrade and removal contract

Follow the [installation guide](INSTALL.md) and packaged `START-HERE.html`. `Install.cmd` runs a per-user PowerShell installer, checks prerequisites and the complete package, copies it into `%LOCALAPPDATA%/XlDuck/install/versions/package-<manifest-id>`, and registers it for Excel startup. It requires Excel closed, preserves other add-ins and retains the previous package for rollback. It does not download prerequisites, unblock files, change Trust Center or install skills into provider accounts.

Run installation from **File Explorer**, then verify Excel launched from its ordinary Windows icon. A packaged development app can expose a private registry view: reading back registration and launching Excel as its child can both succeed while ordinary desktop Excel sees no entry. This was reproduced and corrected on the development workstation. Acceptance requires an independently verified ordinary Windows launch.

Connection definitions, stable IDs, readable names and resource budgets live in workbook custom XML metadata. Save after applying settings. Moves, copies and Save As retain those definitions. Keys remain in Windows Credential Manager; local associations under `%LOCALAPPDATA%/XlDuck/credential-bindings` bind a connection ID and destination to a vault reference. Neither keys nor vault references are stored in the workbook. A new machine or changed destination may require entering credentials again.

Legacy path-based settings under `%LOCALAPPDATA%/XlDuck/connections` migrate on open, including legacy OneDrive URL records, and remain for recovery. Save the workbook to retain migration. Invalid or duplicate metadata and conflicting legacy records fail explicitly. Use `.xlsx`, `.xlsm` or `.xlsb`; CSV and legacy `.xls` cannot retain this metadata. Package rollback does not undo workbook changes, and an older add-in may not understand the new metadata.

Uninstall removes managed startup registrations while preserving package history, workbooks, connection metadata, local associations, keys and agent skills. There is no MSI, Windows Installed Apps entry or automatic updater yet. Agent setup is a separate action: see [agent integration](AGENT-INTEGRATION.md). The tiny skill obtains current instructions from the loaded add-in through `DuckHelp`.

## Tested scope and follow-up work

Development verification covers core types and leases, direct and worker execution, read-only SQL enforcement, authenticated HTTP/S3 fixtures, native Excel DAGs, reuse and restart, manual calculation, cancellation, Protected View, exact values, workbook save/copy/reopen, connection migration and preview updates. Browser checks cover all eight chart templates offline under CSP. Package and installer regressions exercise damaged distributions, registration preservation, upgrade, rollback and recovery using disposable fixtures.

The public installation route, upgrade, rollback, removal and normal startup have been exercised on the existing development workstation. The latest installed development build used the system .NET 10.0.14 runtime and loaded from an independently verified ordinary Windows launch. This is not clean-machine certification.

Live S3-compatible research trials include Hetzdev/Enron and AIS data. An owned AIS fixture listed files, inspected schema and sampled rows, then saved, copied and reopened successfully using its local credential association. The user's AIS workbook was separately migrated and saved, with formulas and chart identities preserved and a live S3 query verified. No credentials are included in release examples.

OpenAI Excel trials exercised bootstrap instructions, installed help, connection discovery and bounded Office.js exploration without worksheet writes. Account-side automatic skill discovery is a separate concern; a live Claude trial has not been completed. README screenshots show the running add-in with synthetic data.

The preview uses explicit WebView2 rasterization matching its Office child window, with separate browser profiles for incompatible DPI awareness contexts. This follows [Microsoft's Office DPI guidance](https://learn.microsoft.com/en-us/office/client-developer/ddpi/handle-high-dpi-and-dpi-scaling-in-your-office-solution). Native tests compare the actual browser pixel ratio and viewport with the child window, rather than treating 100% page zoom as sufficient evidence. Map requests identify XlDuck and follow the [OpenStreetMap tile policy](https://operations.osmfoundation.org/policies/tiles/); tile availability remains a provider dependency.

The following are accepted follow-up work, not claims made by v1:

- Clean Windows/VM installation and a broader Excel, Windows and third-party add-in compatibility matrix.
- Provider-account skill discovery improvements and a live Claude execution trial.
- More monitor/DPI transitions and S3-compatible providers, credential renewal and enterprise-network configurations. Optional AWS profile/SSO providers require separate live validation before advertising that coverage.
- Code signing and a more integrated installer/updater.
- In-sheet charts, report distribution and saved snapshots for recipients without the add-in.

## Support diagnostics

Run packaged `scripts/Get-Diagnostics.ps1 -PackageDirectory <package>` for integrity, source and binary versions, architecture and available runtimes. Its fixed failure codes distinguish missing files, corruption and runtime-discovery problems. It does not attach to Excel or collect workbook contents, connection settings, credentials, private paths or logs. Include relevant `DuckStatus()` and `DuckCellStatus(cell)` output separately; review engine error text before sharing because it may contain source paths or query details.

## Publishing

1. Update `Directory.Build.props`, README/download links and [release notes](RELEASE-NOTES.md). Commit only reviewed product, documentation and test changes.
2. Run local verification against the intended package, including native Excel and the demo. Push the branch, update its PR with the final scope and evidence, and wait for GitHub Verify to pass before merging.
3. Tag the merged commit `v<version>` and push the tag. The Verify workflow rebuilds from a clean checkout and repeats the core, reader, browser, helper, package and isolated installer checks. Native Excel tests remain a workstation check.
4. The publishing job requires matching tag/version/source commit, a clean source tree and valid download hashes. Stable tags create a normal latest release; `-preview.N` tags create prereleases. Only this job receives repository write permission. Branch and PR builds do not publish.
5. Verify the published asset inventory, download checksums and package manifest. Release assets are `xlduck-<version>-windows-x64.zip`, `xlduck-skill.zip`, `XlDuck-Research-Demo.xlsx`, `build-info.json` and `SHA256SUMS.txt`.

Publication does not imply code signing or compatibility beyond the tested scope above.
