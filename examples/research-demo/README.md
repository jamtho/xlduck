# Research demo

The [demo workbook](XlDuck%20Research%20Demo.xlsx) uses 84 generated observations from three fictional sites. It needs no credentials or external sources. After installing XlDuck, open it normally in Excel and save your working copy under Documents. The Research tab exposes three inputs, the formula DAG, a site summary, a schema inspection, a bounded spill and two plot handles. The Observations table preserves text sample IDs, numeric measurements and Boolean acceptance flags.

For development, from the source repository root, open a disposable working copy in a new Excel process:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/Open-Demo.ps1
```

The launcher uses the newest verified local package, or accepts `-AddInPath` and `-WorkbookPath`. It waits for automatic initialization, opens the preview pane on the line-chart handle and leaves the new instance under the user's control. Existing Excel instances are left alone. Working copies live under `artifacts/demo-sessions`; close the instance normally when finished. The local portable .NET 10 development bootstrap is used when present, without installing a runtime or changing machine settings. Other machines need the prerequisites in `docs/RELEASE.md`.

To keep your edits in a particular copy, for example under Documents, supply its `-WorkbookPath` and `-UseWorkbookInPlace`. This opens and saves that file directly in the new instance. The default still creates a disposable copy. These launchers load XlDuck only for that session; persistent startup registration uses Install.cmd as described in `docs/INSTALL.md`.

Try `Research!C7:C9`, then change a response or accepted flag on Observations. Select C11 for the line chart or C12 for the bar chart, then press Ctrl+Shift+P. Handle cells display their actual values. Select C19 to preview the retained table. C18 is a lazy recipe and C19 materialises the result shared by its branches. Recalculation reuses retained results; Restart workbook rebuilds the session.

## Rebuild and verify

`build.mjs` authors the file with the bundled `@oai/artifact-tool` library. It requires a local `node_modules` link to the bundled dependencies, not a new project dependency. Run it from the repository root. The default output directory is the one linked above; an optional first argument overrides it. It also writes a test contract under `artifacts/demo-authoring`.

The artifact renderer cannot execute Excel-DNA functions. Complete the build by running the native test with `-Prepare`, which writes the authored formulas through Formula2, calculates them, restores its test edits and saves native cached results:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-Demo.ps1 -WorkbookPath 'examples/research-demo/XlDuck Research Demo.xlsx' -Prepare
```

Without `-Prepare`, the test opens a fresh working copy and verifies the saved formulas without rewriting them. It checks all controls against independent aggregates, empty results, source edits, leading-zero IDs, reuse and dynamic spills. The `--render-values` builder mode reads the native calculation snapshot for layout review and never exports that rendering copy.
