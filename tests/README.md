# XlDuck tests

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

For recording the current baseline, append `-- --allow-known-defects`. This permits only cases explicitly marked as known defects; unexpected failures still fail the run. The output reports known defects separately. Remove each known-defect marker when its fix lands. Milestone 3 will replace the legacy source links with the extracted core project.

Preview tests run the real embedded HTML in a separate headless Edge profile with all page network requests blocked. Install the pinned Playwright dependency in `tests/preview`, or supply an existing Playwright installation through `NODE_PATH`:

```powershell
node tests/preview/preview.test.cjs
node tests/preview/preview.test.cjs --allow-known-defects
```

The browser's normal command also fails while known defects remain. These checks exercise table rendering and duration conversion without downloading chart libraries. They do not replace chart and WebView2 integration checks.

Cell readiness is observed with bounded polling. Test assertions read `Value2`; formatting tests should assert formatting explicitly. Formula writes use `Formula2` so Excel can spill arrays. Excel busy responses are retried for a bounded time rather than handled with fixed sleeps.
