---
name: xlduck
description: Use XlDuck in a live Excel workbook for SQL research, table handles and query previews. Applies when the user requests XlDuck or the workbook uses Duck functions or xlduck:// handles.
---

Read `DuckHelp()` and `DuckHelp("workflow")` from the loaded Excel add-in, then follow that build's guidance and fetch other help topics as needed.

Use the host's formula evaluation tool, or Office.js `worksheet.evaluate('DuckHelp()')` followed by `ctx.sync()` and read the returned `value` (it may be a one-cell array). If only cell tools are available, use an agreed empty cell for the help formula. If help is unavailable, report that XlDuck must be installed or updated; do not invent its API.
