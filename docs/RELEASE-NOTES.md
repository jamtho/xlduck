XlDuck 1.0.1 puts the workbook session first in worksheet handles: `xlduck://ABCD/t/123`. The hierarchy is workbook tag, result kind (`t`, `f` or `p`), then result number.

The four-letter tag allocation, workbook isolation and computation reuse are unchanged. Existing formulas rebuild their handles when the workbook opens; no formula edits or connection migration are needed. Reference the producing cells rather than saving or constructing handle text. Tokens from the previous format are not accepted as live handles.

The parser, generated table/fragment/plot handles, embedded help, public documentation and native Excel tests use the new order. Core/readers, browser, agent-helper, package and installer checks run in CI. Native Excel computation, Protected View, spill previews, DPI behaviour and the research demo are verified on the development workstation.

Download **xlduck-1.0.1-windows-x64.zip**, extract the complete folder and open **START-HERE.html**. Save and close Excel, then run **Install.cmd from Windows File Explorer**. Setup preserves the previous package for rollback, workbook connection settings and local credentials. Verify startup by opening Excel from its Windows icon. See the [installation guide](https://github.com/jamtho/xlduck/blob/v1.0.1/docs/INSTALL.md).

The package includes the synthetic research demo and `agent/xlduck-skill.zip`, also available as separate downloads. The bootstrap skill still reads current guidance from `DuckHelp()`; existing agent installations do not need a new skill for this patch. See [agent setup](https://github.com/jamtho/xlduck/blob/v1.0.1/docs/AGENT-INTEGRATION.md).

This is an **unsigned Windows x64 desktop Excel release**, requiring the x64 .NET 10 Desktop Runtime and WebView2. Clean-machine certification, signing, broader compatibility and a live Claude trial remain follow-up work. See [tested scope](https://github.com/jamtho/xlduck/blob/v1.0.1/docs/RELEASE.md#tested-scope-and-follow-up-work). Published assets include source provenance and SHA-256 download checksums.
