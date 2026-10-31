XlDuck 1.0 brings DuckDB SQL research into Windows desktop Excel. An in-sheet coding agent can inspect data, write SQL and leave a reproducible investigation built from ordinary cells, retained table handles and charts. Handwritten SQL uses the same workflow.

- Read CSV, Parquet and S3-compatible storage; keep large intermediate tables in a separate DuckDB worker for each workbook.
- Compose queries through Excel's dependency graph. Unchanged SQL and inputs reuse results; opening a workbook or explicitly restarting XlDuck rebuilds them.
- Inspect tables, schemas, query state and errors, and view charts in the preview pane. Agents can discover configured connections and run bounded exploration without first writing formulas into the sheet.
- Select a handle or any cell in its `DuckOut` spill to inspect the underlying table. Short handles use `xlduck://t/ABCD/123`. The preview follows its Office window's DPI, including bitmap-scaled desktop/RDP contexts.
- Save connection definitions and resource budgets with the workbook. Keep keys in Windows Credential Manager, associated locally with the connection and its destination.
- Read current guidance through `DuckHelp()`. The small optional agent skill delegates to that installed help.

Download **xlduck-1.0.0-windows-x64.zip**, extract the complete folder and open **START-HERE.html**. Install the x64 .NET 10 Desktop Runtime and WebView2 if needed, save and close Excel, then run **Install.cmd from Windows File Explorer**. Verify startup by opening Excel from its Windows icon. This release is **unsigned**; setup preserves Excel's security settings. See the [installation guide](https://github.com/jamtho/xlduck/blob/v1.0.0/docs/INSTALL.md).

The package includes a synthetic research demo and `agent/xlduck-skill.zip`; both are also separate downloads. Install your chosen AI provider's Excel add-in separately and follow the [agent setup guide](https://github.com/jamtho/xlduck/blob/v1.0.0/docs/AGENT-INTEGRATION.md). Account-side skill setup is manual. A fallback starting prompt is included.

GitHub builds the downloads from the clean tagged source and runs the non-Excel verification suites. `build-info.json` identifies that source; `SHA256SUMS.txt` checks download integrity. Native Excel computation, Protected View, preview updates and the demo are checked on the development workstation. Installation lifecycle, ordinary Windows startup, live S3-compatible reads and OpenAI's Office.js exploration path have also been exercised there.

Clean-machine certification, code signing, broader add-in/provider compatibility and a live Claude trial remain follow-up work. This release targets Windows x64 desktop Excel; Mac, browser, ARM and 32-bit Excel are not supported. See [tested scope](https://github.com/jamtho/xlduck/blob/v1.0.0/docs/RELEASE.md#tested-scope-and-follow-up-work).

The rewrite is not backward compatible with the 0.1 API. Keep SQL and formulas in saved analyses; runtime handles rebuild on open. Existing development-era local connection settings migrate into workbook metadata when opened and saved. Credentials remain local, so another computer may need its own credential setup.
