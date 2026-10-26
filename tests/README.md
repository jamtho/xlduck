# XlDuck tests

The production host, typed core, engine, worker and preview have separate checks. All Excel scripts create and verify ownership of their own Excel process and workbook. They never attach through the Running Object Table, choose ActiveWorkbook as a test target, close unrelated workbooks or kill all Excel processes.

## Production integration

After acquiring the pinned readers and preview libraries, run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Run-Tests.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Run-Tests.ps1 -NoBuild -WithPreview
```

The production suite covers automatic startup, typed DAGs and captures, reuse, nested functions, exact output, registry restrictions, two-workbook isolation, dates, pause/cancel/retry, manual calculation, cancelled closure, Save As, fresh reopen, ordinary copies and moved-formula ownership. The preview smoke test opens a visible owned window, waits for local WebView startup, and verifies a selected plot update without reselection or clearing the previous state. It does not assert pixel-perfect rendering or different DPI layouts.

Cleanup retains the owned process handle and requires a zero exit code after Quit. A disappearing window is not proof of clean shutdown: this check catches native crashes in ribbon shutdown callbacks. The Protected View fixture releases temporary Application references before waiting for exit.

Installer tests execute all four packaged `.cmd` launchers against a harmless setup stub. They check x64 Windows PowerShell selection, action arguments, paths containing spaces, exit-code propagation and isolation from incompatible inherited PowerShell modules. These probes do not invoke installation or change Office settings.

Use `-AddInPath <absolute-or-relative-xll>` to test a clean isolated build or package. A development test loads a private copy, leaving the original build available. `Test-PreviewHost.ps1 -InspectionSeconds 60` optionally keeps its owned window visible before cleanup. No Trust Center settings are changed and security dialogs are not dismissed automatically.

`Test-ProtectedView.ps1` marks only its own generated workbook as downloaded, opens it in Protected View, verifies zero new loopback source requests even after full calculation, then exercises Enable Editing on that fixture and checks fresh computation. It never changes Trust Center or dismisses a security prompt.

`Test-HarnessIsolation.ps1` creates its own sentinel instance, runs the suite in another instance, and verifies that unrelated and subsequently added workbooks survive cleanup. It also cleans up its own sentinel workbooks.

## Core, host, engine and worker

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Run-CoreTests.ps1 -WithReaders -WithExcel
```

Use `-NoRestore` after dependencies have been restored, or `-NoBuild` for existing binaries. The core tests cover typed identities, SQL placeholders, leases, dependency retention, cancellation, bounds and protocol frames. Host tests cover path-based connection settings, Windows credential persistence, exact dates, formula selection and WebView access policy. Dialog tests render off-screen, exercise authentication and advanced fields, and verify that editing alone never saves a credential. Review images are written to `artifacts/dialog-review`. Engine tests run directly and through a worker; worker failure tests terminate only their own child process. A narrow separate Excel test add-in preserves the execution prototype's latency and ownership checks.

The agent-feedback regressions cover per-cell state, last-input reuse, cancellation/timeout ownership, bounded exploration, SELECT-list/recursive UNNEST, schema binding without evaluating rows, precise typed samples, structured memory/source/binding failures, resource persistence and its dialog. Native production tests call `Worksheet.Evaluate` on status/exploration without editing cells and distinguish Excel spill failure from query failure. They preserve manual calculation and verify per-workbook resource restarts. The separate live Office.js trials are recorded in [agent integration](../docs/AGENT-INTEGRATION.md#live-validation-and-limits).

Run `node tests/agent-exploration.test.cjs` for the standalone Office.js helper's idempotent polling, bounded waits, session-change handling and manual-calculation handoff. Regressions from the live trial cover scalar/one-cell JSON responses and hosts without timers, where pending status returns for a later poll. This test uses a synthetic Office.js context and does not connect to Excel; [agent integration](../docs/AGENT-INTEGRATION.md#live-validation-and-limits) records the separate live trial. Production tests also check embedded `DuckHelp` topics through worksheet formulas and Evaluate, including text limits and absence of query execution.

Long requests are tested through native Evaluate using bounded staged parts, followed by explicit submission. Repeating a request or submit does not rerun it; changing inputs with a live request id fails. Helper tests preserve Unicode, quotes and backslashes across formula/JSON encoding and assert every expression stays within native Evaluate's 255-character limit.

Reader tests use a loopback server and dummy credentials, never a real bucket or user credential files. They cover HTTP/S3 reads, listing, explicit saved keys, named profiles, headers, errors and extension integrity. They do not verify AWS signatures cryptographically or establish live AWS/SSO/proxy compatibility. Omit `-WithReaders` to run local-only policy checks.

## Browser rendering

Acquire libraries once with `scripts/Get-PreviewAssets.ps1`. Install the pinned Playwright dependency in `tests/preview`, or supply an existing installation through `NODE_PATH`, then run:

```powershell
node tests/preview/preview.test.cjs
```

The suite uses an isolated headless Edge profile and serves the actual local assets at the production virtual origin. External network requests are blocked. It checks injection, exact conversions, interval fractions, stale states and rendering of all eight templates under CSP. It confirms non-map charts make no external requests. Maps are tested with unavailable tiles and invalid coordinates. CI sets `XLDUCK_TEST_BROWSER=chromium` after installing Playwright's pinned Chromium build.

## Packages and runtimes

`scripts/Build-Package.ps1` builds into a fresh directory using locked dependencies. `scripts/Verify-Package.ps1 -Directory <package>` verifies its x64 files and SHA-256 manifest. `scripts/Get-Diagnostics.ps1 -PackageDirectory <package>` emits version and integrity information without workbook data, paths, settings or credentials. Test the actual package with `Run-Tests.ps1 -AddInPath <package>/XlDuck-AddIn64.xll -WithPreview`.

Run `powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-Package.ps1 -PackageDirectory <package>` for file-only release regressions. It creates tiny disposable fixtures from the package inventory; it never executes their placeholder binaries, opens Excel, or modifies the source package. Cases cover omitted runtime dependencies, file and archive corruption, malformed manifests, path aliases/traversal, duplicate entries, junctions, and diagnostic privacy and failure reporting. The build verifies its own output before reporting success; CI then runs these negative cases. Native workbook tests are separate and are unnecessary for changes confined to the packaging/diagnostics scripts.

`XlDuck.RuntimeProbe` builds separate .NET Framework 4.8, .NET 8 and .NET 10 companions using `ProbeFramework` and `ProbeId` MSBuild properties, with a separate artifacts directory per build. `Test-RuntimeCoexistence.ps1` accepts their XLL paths and the companion function (`XlDuckProbeEight` or `XlDuckProbeTen`). It defaults to both load orders; `-Orders companion-first` restricts a case known to block in the reverse order. See [release preparation](../docs/RELEASE.md) for the observed results and runtime constraints. These synthetic probes do not certify unrelated add-ins.

Verification for the access/handle update on 25 October 2026: 22 core, 8 host, 3 dialog, 14 direct-engine, 17 worker and 18 policy/reader cases in each mode pass. Native checks cover 24 production cases, 3 Protected View cases and 3 preview cases; the existing demo passes all 6 groups. S3 credentials and protocol reads use local fixtures. Real storage-provider/proxy validation and display transitions remain release checks.

## Evidence

On 24 October 2026: 21 core, 6 host, 14 direct-engine, 17 worker, 17 policy/reader cases in each engine mode, 11 narrow Excel core checks, 22 production checks, 2 native preview checks and 21 browser checks passed. The production eight-step chain settled in roughly 30–40 ms on the development machine; this is an observation, not a latency guarantee.

The old 95-assertion Excel suite and legacy source-linked defect probes are preserved in Git history at `f031d7a`. Their relevant behaviours and identified defects have been ported to the new suites; intentionally removed mutation/configuration functions now have rejection tests. The old executable is no longer a production test target.

More visual configurations, representative runtime coexistence, clean-machine installation and broader credential-provider checks remain follow-up trials for the public preview; see [release scope](../docs/RELEASE.md#preview-scope-and-follow-up-trials).

## Installation and skill distribution

After building a package, run `powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-Setup.ps1 -PackageDirectory <package>`. It tests real filesystem/registry operations under disposable `artifacts/setup-tests` and `HKCU\Software\XlDuck\SetupTests` locations. It does not start Excel, change Office options, install a global skill or install prerequisites. It covers first/repeat install, upgrade, rollback, third-party load order/types, failed state-commit recovery, corrupted and downloaded payloads, and preservation of customized local skills. CI runs it with package regressions. The skill ZIP must contain only the bootstrap file matching the packaged source.

Use `scripts/Setup.ps1 -Action Check` **from the built package** for the real machine's read-only prerequisites check. A clean-machine install, normal Excel startup, upgrade and removal trial still needs a machine with installed prerequisites and closed Excel. The private runtime used by development tests must not be mistaken for a system installation.

The 0.2.0-preview.1 release review passed the core/host/dialog/engine/worker and both reader-policy suites, 36 native Excel checks, all six synthetic demo groups, 23 browser checks, four agent-helper groups, 53 package checks and 15 installer checks. Additional native licence texts are verified against 69 pinned upstream hashes before packaging. README images are actual Excel captures of the synthetic demo. GitHub's Verify workflow repeats the non-Excel checks on the pushed source and publishes only matching preview tags.
