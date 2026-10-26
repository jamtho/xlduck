# Public preview release process

Version **0.2.0-preview.1** is the first unsigned public preview of the rewrite. It is a Windows x64 folder containing the Excel-DNA XLL, managed host, separate DuckDB worker, verified readers, bundled preview assets, a synthetic demo and agent bootstrap. Keep the folder intact. It is not a single-file XLL and contains no 32-bit add-in.

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

The build uses a fresh output directory and NuGet lockfiles. Dependency acquisition checks package integrity and reader hashes. Packaging selects the intended x64 files explicitly, includes licence metadata/notices and creates a SHA-256 manifest and ZIP. Before reporting success, it verifies both the folder and archive against the required host, worker, reader and preview files. The verifier rejects missing dependencies even if their manifest entries are also missing, changed or unlisted files, noncanonical paths, links/junctions and unexpected ZIP entries. These checks detect incomplete or damaged distributions; they do not authenticate an unsigned release.

`build-info.json` records the source commit, whether tracked changes or untracked files were present, SDK and build time. This makes dependencies and packaged contents reviewable; it does not promise byte-identical ZIP timestamps. Do not publish a candidate with missing tests or notices. `tests/Test-Package.ps1 -PackageDirectory <package>` exercises release failures using disposable file fixtures without starting Excel; CI runs it after packaging.

Run `tests/Run-Tests.ps1 -AddInPath <package>/XlDuck-AddIn64.xll -WithPreview` against the actual package. The portable-runtime test options use a test-only .NET Framework companion to set the runtime root inside an owned Excel process; they never change user or machine environment settings.

## Installation, upgrade and removal contract

The user-facing [installation guide](INSTALL.md) and packaged `START-HERE.html` give prerequisite, installation, agent setup, smoke-test and recovery steps. The package now includes a per-user PowerShell installer behind `Install.cmd`, plus `Rollback.cmd`, `Uninstall.cmd` and `Install Codex Skill.cmd`. It requires Excel closed, checks installed x64 Excel/.NET 10 Desktop/WebView2, verifies and copies the complete package, and manages current-user Excel startup registration. It does not download dependencies, unblock files, change Trust Center or automatically install skills into provider accounts.

Managed packages live under `%LOCALAPPDATA%/XlDuck/install/versions/package-<manifest-id>`. Identical installs reuse their directory; upgrades keep the preceding version for rollback. Startup registration preserves unrelated entries and their relative order, with recovery snapshots and rollback on state-commit failure. Manual registration through Excel's Add-ins dialog remains supported. Follow the organization's normal trusted-publisher/location policy.

An upgrade must close Excel instances using the old XlDuck installation before switching the registered path. Keep the previous version until the new version passes a workbook smoke test. Runtime session handles are deliberately rebuilt, so saved files must retain formulas. Connection settings are separate under `%LOCALAPPDATA%/XlDuck/connections`; access keys are in Windows Credential Manager. The earlier development connection records are no longer read. Package replacement must not copy credentials or change Excel trust settings.

Uninstall removes managed startup registrations and keeps package copies/history for deliberate manual removal. It preserves connection settings, credentials, agent skills, workbooks and unrelated registrations. There is no automatic updater, MSI or Windows Installed Apps entry yet. `tests/Test-Setup.ps1` uses disposable files and a separate test registry key to exercise install/reinstall, upgrade/rollback, third-party entries, failed commit recovery, blocked/corrupt packages and local skill preservation. It never changes the user's Office options or installed skills. CI runs it after packaging. Clean-machine validation is deferred for this preview and remains necessary before claiming that coverage.

The skill ZIP contains only `xlduck/SKILL.md`. Detailed workflow/help is embedded from `docs/AGENT-WORKFLOW.md`, `docs/FORMULAS.md` and the JavaScript helper, preventing a bootstrap recursion or stale account-side API copy. Agent-provider upload/discovery and Claude execution are separate validation steps; see [agent integration](AGENT-INTEGRATION.md).

On 26 October 2026, candidate 13 was installed with the public setup script on the development machine after installing Microsoft's x64 .NET 10.0.14 Desktop Runtime. Normal Excel startup loaded the XLL and worker from the managed installation and the system runtime, rebuilt the synthetic research graph to its expected 58-row result, and exited with code 0. The actual package passed all 30 production, 3 Protected View and 3 preview checks, plus 47 package and 14 isolated installer checks. This is a real startup test on an existing development machine, not a clean-machine certification.

That trial found and fixed a native crash in the old build's ribbon shutdown: it called Excel's C API to remove a shortcut outside an allowed callback context. Shortcut registration/removal now belongs to Excel-DNA's `ExcelCommand.ShortCut` lifecycle. The integration harness retains the process handle and checks its exit code, so a crash can no longer count as successful cleanup. See [Excel-DNA's shortcut lifecycle](https://github.com/Excel-DNA/ExcelDna/blob/v1.9.0/Source/ExcelDna.Loader/XlRegistration.cs).

The actual launchers also passed upgrade from candidate 13 to 14, rollback, uninstall and reinstall against the current user's Excel startup registration. Candidate 14 passed the same 36 native Excel checks and normal startup/shutdown. A subsequent shell test reproduced incompatible module discovery when launching a `.cmd` from PowerShell 7. Launchers now select Windows PowerShell's built-in module directory for their child process only. The installer regression exercises all four launchers with an incompatible inherited module, a path containing spaces and a nonzero probe exit code; it touches no Office registration.

## Support diagnostics

Run `scripts/Get-Diagnostics.ps1 -PackageDirectory <package>` from the repository or packaged folder to print a JSON report with package integrity, source commit, binary versions, Windows architecture and available .NET runtimes. Schema version 2 includes a fixed `packageIssue` code on failure (for example `missing_required_file` or `file_mismatch`), and distinguishes failed runtime discovery from a missing `dotnet` command. Unreadable binary metadata does not prevent the rest of the report. It does not attach to Excel or collect workbook contents, local paths, connections, credentials or logs. Include `DuckStatus()` separately when reporting the runtime actually loaded inside Excel and the affected workbook's computation state. Review any error text before sharing it; engine errors shown in cells may contain source paths or query details.

## Preview scope and follow-up trials

Development validation includes live reads from the user's Hetzdev S3-compatible endpoint, Enron workbook edits/cache reuse, simulated authenticated S3 outages and recovery, and native preview checks at DPI 192 and 240. These are local results, not certification of every provider or display configuration.

The first preview ships with the development-machine installation evidence above. Actual upgrade, rollback, uninstall/reinstall and refusal to install while Excel is open have passed. Candidate 15 additionally passed all 36 native checks, 15 installer checks, normal startup on the system .NET 10.0.14 runtime, and a clean exit after recalculating the saved demo. Screenshots in the README show the actual running add-in with synthetic data. Native redistribution notices are now included from pinned upstream sources; see [third-party notices](../THIRD-PARTY-NOTICES.md).

The following work is deliberately deferred, rather than a blocker for this exploratory preview:

- Clean supported Windows machine: runtime/WebView prerequisites, first install and workbook/preview smoke test. No VM or cloud spend is required for this milestone.
- Provider-account skill upload/discovery and a live Claude execution trial. The bootstrap-to-help-to-query path has been exercised through OpenAI's Excel integration; automatic discovery is a separate check.
- Representative third-party add-ins and the broader Excel/Windows matrix.
- More display scales and monitor transitions for connections and preview.
- Broader S3-compatible provider coverage, access-key/session-token renewal and enterprise-network behaviour. Optional AWS profile/SSO providers need separate live validation before their support is advertised.
- Code signing and a more integrated installer/updater, once the public preview demonstrates value.

## Publishing a preview

Run native Excel tests against a package built from the intended commit, and run the demo regression. Preserve user workbooks and original skill archives outside the staged release changes. Push the reviewed branch and check GitHub's Verify workflow before tagging.

The Verify workflow builds from a clean checkout, runs core/read-policy/browser/package/installer checks and retains verified downloads. A pushed `v…-preview.…` tag also enables the publish job. That job requires the package version to match the tag and its source commit to match the clean tagged checkout; it checks download hashes and creates a GitHub **prerelease** with the ZIP, skill ZIP, demo, build metadata and SHA-256 list. It has write permission only in that publishing job. A branch push or pull request does not publish a release. Update `docs/PREVIEW-RELEASE.md` and README links when preparing the next version.

Publication does not imply code signing or clean-machine certification. Crash-orphan spill recovery is implemented and tested against active host/worker leases and unmarked directories.
