# XlDuck tests

## New core and worker

The new implementation has UI-independent core specifications, real DuckDB cases in both direct and worker modes, and a separate Excel test add-in. From the repository root:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Run-CoreTests.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Run-CoreTests.ps1 -WithExcel
```

Use `-NoBuild` for an existing build, or `-NoRestore` when dependencies have already been restored. The default build needs NuGet access. `-WithExcel` creates and closes an owned Excel instance; it never selects an existing user workbook. Worker crash and blocked-pipe cases deliberately terminate only their own child processes.

For HTTP/S3 integration, acquire the pinned official development readers once, then opt into loopback network tests:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/Get-ReaderExtensions.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Run-CoreTests.ps1 -WithExcel -WithReaders
```

The acquisition script needs internet access and verifies exact extension hashes. Queries never install extensions. Reader tests use an owned local server, temporary files and dummy AWS credentials in the test process and its child. They do not access a real bucket or change user credential files. Named-profile and authentication-header tests are not live AWS/SSO validation. The policy suite without `-WithReaders` still tests local readers and prohibited capabilities.

The suites can also run individually after building `XlDuck.Worker`, the three specification projects and (for Excel) `tests/XlDuck.TestAddIn`:

```powershell
dotnet run --project tests/XlDuck.Core.Specs --no-build --no-restore
dotnet run --project tests/XlDuck.Engine.Specs --no-build --no-restore
dotnet run --project tests/XlDuck.Engine.Specs --no-build --no-restore -- --worker
dotnet run --project tests/XlDuck.Policy.Specs --no-build --no-restore -- --readers
dotnet run --project tests/XlDuck.Policy.Specs --no-build --no-restore -- --worker --readers
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-NewCore.ps1 -NoBuild
```

The [core contract and limitations](../XlDuck.Core/README.md) describe what this validates. The Excel test add-in is not the public host. Its preview test checks plot data and retention, not rendered charts. Production `XlDuck/` still uses the legacy implementation.

## Existing add-in and regression baseline

The Excel suite creates a new Excel process and workbook, checks that it owns that process, and closes only its own workbook. It does not attach to the active workbook. Run it from the repository root on Windows with 64-bit Excel installed:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Run-Tests.ps1
```

Use `-NoBuild` after a successful build. To verify that the suite leaves another workbook untouched, run the isolation test. It creates both the sentinel workbook and the test instance itself:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-HarnessIsolation.ps1
```

The core regression executable links the existing computational source with minimal UI host stubs. It needs no Excel process. Its normal command exits unsuccessfully while the documented defects remain:

```powershell
dotnet run --project tests/XlDuck.Core.Tests
```

For recording the current baseline, append `-- --allow-known-defects`. This permits only cases explicitly marked as known defects; unexpected failures still fail the run. The output reports known defects separately. Keep these legacy tests until the production entry points are migrated; the new `XlDuck.Core.Specs` project independently verifies the corrected implementation. Do not remove legacy known-defect markers merely because the new core passes.

Preview tests run the real embedded HTML in a separate headless Edge profile with all page network requests blocked. Install the pinned Playwright dependency in `tests/preview`, or supply an existing Playwright installation through `NODE_PATH`:

```powershell
node tests/preview/preview.test.cjs
node tests/preview/preview.test.cjs --allow-known-defects
```

The browser's normal command also fails while known defects remain. These checks exercise table rendering and duration conversion without downloading chart libraries. They do not replace chart and WebView2 integration checks.

Cell readiness is observed with bounded polling. Test assertions read `Value2`; formatting tests should assert formatting explicitly. Formula writes use `Formula2` so Excel can spill arrays. Excel busy responses are retried for a bounded time rather than handled with fixed sleeps.
