# Public release preparation

The current package is an unsigned development candidate. It is a Windows x64 folder containing the Excel-DNA XLL, managed host, separate DuckDB worker, verified readers and bundled preview assets. Keep the folder intact. It is not a single-file XLL and contains no 32-bit add-in.

## Runtime and coexistence

The intended public prerequisite is the **x64 .NET 10 Desktop Runtime** and WebView2 for the task pane. The binaries retain the .NET 8 API baseline and use `LatestMajor` roll-forward. This prefers the newest available compatible runtime and can load into an older runtime already selected by another add-in. `DuckStatus()` reports the runtime actually used by the Excel host. The worker selects its runtime independently.

Portable .NET 10.0.13 tests pass with .NET 10 and .NET Framework 4.8 companions in both load orders. A .NET 8 companion loaded first also passes. Loading XlDuck first and then a companion pinned to .NET 8 blocked in the native loading path; that case was terminated using only the recorded owned test PID/start time. These are synthetic compatibility probes, not certification of third-party add-ins.

Excel permits only one modern .NET runtime in a process, alongside .NET Framework. Another add-in pinned to an incompatible runtime can therefore require a separate Excel instance or its vendor's runtime update. Roll-forward cannot guarantee compatibility for arbitrary add-ins. .NET 8/9 support ends on 10 November 2026; an older-runtime compatibility test is not a supported public deployment baseline after that date.

Sources: [Excel-DNA runtime guidance](https://excel-dna.net/docs/guides-basic/dotnet-runtime-support/), [Microsoft support policy](https://dotnet.microsoft.com/en-us/platform/support/policy/dotnet-core).

## Reproducible candidate

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/Get-ReaderExtensions.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/Get-PreviewAssets.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/Build-Package.ps1
```

The build uses a fresh output directory and NuGet lockfiles. Dependency acquisition checks package integrity and reader hashes. Packaging selects the intended x64 files explicitly, includes licence metadata/notices and creates a SHA-256 manifest and ZIP. `build-info.json` records the source commit, whether tracked changes or untracked files were present, SDK and build time. This makes dependencies and packaged contents reviewable; it does not promise byte-identical ZIP timestamps. Do not publish a candidate with missing tests or notices.

Run `tests/Run-Tests.ps1 -AddInPath <package>/XlDuck-AddIn64.xll -WithPreview` against the actual package. The portable-runtime test options use a test-only .NET Framework companion to set the runtime root inside an owned Excel process; they never change user or machine environment settings.

## Installation, upgrade and removal contract

For development, extract into a new version-specific directory and load its `XlDuck-AddIn64.xll`. For persistent installation, use Excel's Add-ins dialog to select that exact file. Follow the organization's normal trusted-publisher/location policy. Do not disable Trust Center protections or bypass a blocked download.

An upgrade must close Excel instances using the old XlDuck installation before switching the registered path. Keep the previous version until the new version passes a workbook smoke test. Runtime session handles are deliberately rebuilt, so saved files must retain formulas. Connection settings are separate under `%LOCALAPPDATA%/XlDuck/connections`; access keys are in Windows Credential Manager. The earlier development connection records are no longer read. Package replacement must not copy credentials or change Excel trust settings.

Removal must unregister the add-in before removing its versioned folder after Excel exits. User-local connection settings, saved credentials, logs and WebView data require a separate, explicit remove-data choice. Never remove workbooks, AWS profiles, shared runtimes or data-source files. An installer implementing and testing this contract remains a release gate.

## Support diagnostics

Run `scripts/Get-Diagnostics.ps1 -PackageDirectory <package>` from the repository or packaged folder to print a JSON report with package integrity, binary versions, Windows architecture and available .NET runtimes. It does not attach to Excel or collect workbook contents, local paths, connections, credentials or logs. Include `DuckStatus()` separately when reporting the runtime actually loaded inside Excel and the affected workbook's computation state. Review any error text before sharing it; engine errors shown in cells may contain source paths or query details.

## Gates before public publication

- Clean supported Windows machine: runtime/WebView prerequisites, first install and workbook/preview smoke test.
- Actual upgrade, rollback and removal, including files locked by an open Excel instance.
- Representative third-party add-ins and the supported Excel/Windows matrix.
- Visual review of connections and preview at common DPI/scale settings.
- Live S3-compatible endpoints, access-key/session-token renewal and enterprise-network behaviour. Certify optional AWS profile/SSO providers separately before advertising their support.
- Complete redistribution notice review, including native extension dependencies; signing identity and installer design.
- Validate release artifacts against the final tested commit. Crash-orphan spill recovery is implemented and tested against active host/worker leases and unmarked directories.

No public upload, signing purchase, signing identity or clean-machine result is implied by a local development ZIP.
