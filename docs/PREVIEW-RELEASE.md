First public preview of the rewritten XlDuck: read-and-compute SQL research in Windows desktop Excel, with workbook-local DuckDB workers, retained table handles and a chart/table preview pane.

Download **xlduck-0.2.0-preview.1-windows-x64.zip**, extract the whole folder and open **START-HERE.html**. Install the x64 .NET 10 Desktop Runtime and WebView2 if needed, close Excel, and run **Install.cmd**. XlDuck then loads on normal Excel startup. This preview is **unsigned**; setup does not change Excel's security settings.

The ZIP includes a synthetic research workbook and the tiny agent bootstrap skill. The workbook and skill ZIP are also available as separate downloads. XlDuck executes SQL; your chosen Excel AI add-in remains a separate installation. The skill reads current guidance from `DuckHelp()` in the installed XlDuck build.

Validated on the development workstation: native Excel computation, Protected View, live preview updates, startup/shutdown, installation, upgrade, rollback and removal. The Office.js exploration path has also passed live trials with OpenAI's Excel integration. GitHub CI builds this package from the tagged source and runs the core, reader, browser, packaging and isolated installer checks. `build-info.json` identifies the source; `SHA256SUMS.txt` checks download integrity.

Deferred trials: clean Windows VM installation, automatic account-side skill discovery, Claude execution, broader Excel/add-in coexistence and S3-provider coverage. These are follow-up compatibility work, not claims made by this preview. The 0.1 API is not backward compatible; saved analyses should keep their SQL and formulas, and results rebuild on workbook open.

See the packaged installation and agent guides for setup and troubleshooting. Please report your Excel/Windows versions, `DuckVersion()` and the smallest reproducible example, without credentials or private research data.
