# Enron research demo

The [workbook](../../outputs/01a16b0a-0074-7383-ba33-19bcf4c27b0f/XlDuck%20Enron%20Demo.xlsx) reads the user's Enron Parquet collection from their S3-compatible service. It requires that service to be reachable and a connection configured locally through **XlDuck > Connections**. Credentials are not in the workbook or repository.

Open the configured demo in a separate development Excel instance:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/Open-EnronDemo.ps1
```

This opens the demo at its existing path, where XlDuck has saved its connection settings. It never opens the legacy workbook and leaves existing Excel instances alone. An optional `-ConnectionFile` supplies non-secret `WorkbookSettings` JSON referencing an existing Windows Credential Manager entry. `-WorkbookPath` and `-AddInPath` select another demo or build. A file copied to a different path needs its own connection settings.

## Research workflow

- **Research!C6:C8** select the year, a literal subject substring and an exact sender identifier. Use `All` for an unrestricted text filter. Dates and year boundaries are UTC.
- **Queries!C5/C6** contain source SQL and its `DuckQuery` formula. The retained table contains the entire selected year, projected to five metadata columns. It does not fetch email bodies.
- The remaining SQL cells and formulas filter that table, summarise it, group months, rank senders and select the earliest 25 matching records. Editing text filters does not reread S3. Changing the year does.
- **Research!C12/C13** are live plot handles. Select one and use Ctrl+Shift+P. **C14** previews the first 25 matching records. Handles retain their actual visible tokens.
- Counts are source records, including mailbox copies. Sender identifiers use a lowercased email address, falling back to the display name. They are not deduplicated people. Reply share uses the source `is_reply` flag. Months with no matches are omitted.
- Unchanged calculation reuses results. Reopening or restarting the workbook starts a fresh session and reads the source again.

## Rebuild and verify

`build.mjs` uses the bundled `@oai/artifact-tool` runtime, through an ignored `node_modules` junction. Run from the repository root. The optional first argument selects the output directory. It writes a test contract to `artifacts/enron-demo`.

Native Excel supplies the XLL calculation and spill metadata. With your local non-secret connection JSON and saved credential:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-EnronDemo.ps1 -WorkbookPath 'outputs/01a16b0a-0074-7383-ba33-19bcf4c27b0f/XlDuck Enron Demo.xlsx' -ConnectionFile artifacts/enron-demo/connection.json -Prepare
```

The opt-in live test checks the default and changed filters against independent source aggregates, reconciles monthly totals, verifies both plots, tests empty results and year changes, and confirms unchanged recalculation reuse. It restores the controls, removes its own verification sheet, scans for errors and saves native cached values. `-ReopenOnly` verifies the saved formulas and persisted local connection without supplying settings again. Without either switch, it tests a disposable copy.

For layout review, `build.mjs <output-directory> --render-values artifacts/enron-demo/calculated.json` renders captured native results without exporting a static replacement. The source workbook and its credentials are never copied into the demo.
