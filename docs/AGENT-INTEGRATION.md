# Using XlDuck with an Excel agent

XlDuck is the SQL execution add-in. The agent's Excel add-in is a separate product: it supplies the model and tools for reading and editing the workbook. XlDuck does not run an AI model, install another agent, expose a network server, or automatically inject instructions into an agent's conversation.

## Give the agent one starting hint

> Use XlDuck for this analysis. First read DuckHelp() and DuckHelp("workflow"). Consult DuckHelp("reference") as needed. Inspect the actual schema and a small sample before choosing an analysis. Leave the final SQL and formulas in the workbook so I can inspect and change them.

The package includes `agent/xlduck-skill.zip`, containing only a small `SKILL.md` bootstrap. It tells the agent how to read installed help, without freezing the function catalogue, workflow or JavaScript helper in its account. Install it using the provider-specific steps below, or paste the hint above when skill installation is unavailable.

XlDuck 1.0 includes `DuckHelp`. It returns plain text from resources embedded in the installed binary, without running queries or reading external data. The same documentation sources generate the installed help and the repository guidance:

| Formula | Content |
| --- | --- |
| `=DuckHelp()` | Short topic index and version |
| `=DuckHelp("workflow")` | Research workflow from `docs/AGENT-WORKFLOW.md`, independent of the bootstrap skill |
| `=DuckHelp("reference")` | Function signatures, SQL capabilities, types, limits and recovery from `docs/FORMULAS.md` |
| `=DuckHelp("officejs")` | JavaScript helpers from `docs/agent-exploration.js`; the version header is a comment |

An agent with formula evaluation can fetch help directly. An agent limited to cell tools can write a help formula into an agreed empty cell, then read its value. A visible help cell or the starting prompt above solves the initial discovery step; a function cannot make an arbitrary agent call it automatically. Do not silently insert hidden instruction sheets or defined names into users' workbooks.

The bootstrap requires a build with `DuckHelp`. It reports missing help instead of guessing an older API. The original user-provided `skill/xlduck.zip` is a historical archive, not the current installable skill. New builds generate a separate ZIP from `skill/xlduck/SKILL.md` and ship detailed documents with the add-in. Do not embed the bootstrap itself as the workflow: that would send the agent back to the same help call indefinitely.

## Install the bootstrap skill

Provider support checked **26 October 2026**. Account, organization and host availability can differ. Installing a skill supplies instructions; it does not add Office.js or macro tools to the agent.

### OpenAI's Excel sidebar

In an eligible ChatGPT workspace, open **Plugins > Skills > Create > Upload from your computer**, upload the package's `agent/xlduck-skill.zip`, and complete installation. In Excel, open the signed-in OpenAI add-in, select **+** to browse skills or use **@**, and look for XlDuck. The [Excel documentation](https://help.openai.com/en/articles/20001063-chatgpt-for-excel-and-google-sheets) documents that picker. The [custom skill guide](https://help.openai.com/en/articles/20001066-skills-in-chatgpt) currently lists eligible Business, Enterprise, Healthcare and Edu accounts; a personal account may lack uploads. Verify that the uploaded skill appears in this particular Excel host. Otherwise use the starting prompt rather than claiming installation succeeded.

### Claude for Excel

Open **Claude > Customize > Skills > + > Create skill > Upload a skill**, upload the same ZIP, and enable it. Code execution and any organization skill controls must permit it. Sign into the same account in Claude for Excel. Anthropic explicitly documents that enabled account skills are available in Excel, selected automatically when relevant or through the **/** menu. See [Use skills in Claude](https://support.claude.com/en/articles/12512180-use-skills-in-claude). The XlDuck execution workflow still needs its own live Claude compatibility trial; do not infer tool compatibility from skill availability.

### Codex desktop controlling Excel

Double-click **Install Codex Skill.cmd** in the package. It copies the single bootstrap file to `%USERPROFILE%\.agents\skills\xlduck\SKILL.md`, the documented [user skill location](https://learn.chatgpt.com/docs/build-skills). It is idempotent for an identical skill and refuses to overwrite a different existing one. It changes no Codex configuration. Start a new Codex chat if discovery has not refreshed; select XlDuck and connect the open workbook through the signed-in OpenAI Excel add-in. This does not install a skill into the separate Excel sidebar.

### Verify and update

Start a fresh agent conversation, select the installed XlDuck skill, and ask it to read the installed help and run `SELECT 42 AS answer` through the documented exploration interface. It should retrieve the workflow from the add-in, use the helper if its host supports it, poll the same request if pending, and return 42. If only cell tools exist, it should explain that and use an agreed scratch area. Separately try an ordinary XlDuck request without naming the skill to check automatic discovery.

Most upgrades need only the add-in: the stable bootstrap fetches that build's guidance. If the bootstrap itself changes, update it through the provider's skill UI. For a customized local skill, review/move the old folder before rerunning the local installer. Remove account skills in the provider UI and a local skill by removing its dedicated folder. XlDuck uninstall does not silently edit agent accounts or delete user-customized skills.

`START-HERE.html` accompanies every package with setup links and the fallback prompt. No supported unattended registration into either Excel sidebar account has been established. Microsoft's separate [Copilot custom-functions skill preview](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-copilot-custom-functions-skill) does not establish discovery for another provider's add-in.

## Where the Office.js path is used

Normal worksheet formulas use Excel's calculation engine and XlDuck directly. They do not call the JavaScript helper.

The optional helper is for an agent whose actual tools include `run_officejs` (or an equivalent supported Office.js host). It uses desktop [`Worksheet.evaluate`](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheet#excel-excel-worksheet-evaluate-member(1)), requiring ExcelApiDesktop 1.1, to call XlDuck functions without inserting temporary worksheet formulas:

```text
Agent's Excel tool
  -> Office.js Worksheet.evaluate + ctx.sync()
  -> Excel calls DuckConnections, DuckExplore or DuckCellStatus
  -> XlDuck reads workbook metadata or uses its existing worker/session
  -> JSON status or bounded results return to the agent
```

The helper uses blank `AZ1000` on the explicitly named sheet as its workbook anchor. It never writes that cell. Inspect the workbook first and choose a different blank single-cell anchor in the helper if necessary. Connections, resources and the SQL policy are the same as for worksheet queries; this is not an unrestricted SQL endpoint.

Load the helper declarations from `DuckHelp("officejs")` or the supplied file, then use the existing `ctx` supplied by the agent's host. Do not nest `Excel.run()` inside a host that already supplies it. Discover configured S3 scopes with `return await xlduckConnections(ctx, "Research");`. It returns endpoints and authentication methods without keys or source reads. Follow a non-null `nextOffset` with the helper's third argument. Use `DuckExplore` to list a narrow prefix, inspect a discovered file's schema and retrieve a small sample; the installed workflow gives the SQL. The tiny bootstrap skill does not need updating when these instructions change.

For a source-independent query, after the helper declarations:

```js
return await xlduckExplore(ctx, "Research", {
    op: "query", id: "sample-001",
    sql: "SELECT unnest([1,2,3]) AS n", maxRows: 2
});
```

If the result is pending, a later tool call should include the helper declarations again as required by the host, then:

```js
return await xlduckExplore(ctx, "Research", {op: "poll", id: "sample-001"});
```

The first call starts one operation; polls retrieve that same operation. Do not generate another ID or resubmit SQL just because work is running. An `initializing` response means no operation was admitted: retry the original start request. If the workbook session changes, stop and review before deliberately starting again. Terminal states are `succeeded`, `failed` and `cancelled`.

For an existing worksheet computation:

```js
return await xlduckWaitCell(ctx, "Research", "C19");
```

`ctx.sync()` sends queued Excel API calls and returns their responses. It does not wait for the DuckDB query to finish. The tested agent sandbox has no JavaScript `setTimeout`; in that environment the helper returns pending status for the next tool call. Hosts with timers can use its short bounded wait. It never busy-loops or changes Excel's RTD interval or calculation mode.

Office.js can return a scalar string or a one-cell array for an XLL result. The helper normalizes these shapes and rejects missing/multi-cell results. It also handles native Evaluate's 255-character expression limit by staging longer JSON requests before explicit submission. See the [reference](FORMULAS.md#read-only-agent-inspection-prototype) for bounds, types, errors and expiry. Returned big integers and decimals are exact text with type tags; do not coerce them to JavaScript numbers.

After exploration, use the agent's ordinary worksheet tools to leave an inspectable SQL/formula graph. If Office.js evaluation is unavailable, use the worksheet functions and an agreed scratch area. Formula tools alone do not grant COM/VBA or macro access: connection setup, resources and restart may still require a ribbon action by the user.

## Live validation and limits

On **26 October 2026**, candidate 12 passed a further live trial through Codex and the signed-in OpenAI Excel add-in in a separately owned workbook. Starting from the bootstrap instructions, Codex retrieved `DuckHelp()` and the independent workflow, fetched the JavaScript helper from `DuckHelp("officejs")`, and used that returned helper to run `SELECT 42 AS answer`. The result was one typed integer, `42`, with `hasMore: false`. Worksheet names and used ranges were unchanged, and the explicit anchor remained blank. The owned trial process was closed afterwards. This verifies the bootstrap-to-installed-help execution path; it is not evidence that either provider automatically discovers an uploaded skill. Claude's live trial remains pending.

On 25 October 2026, a live trial used the installed OpenAI Excel add-in's advertised `run_officejs` tool in Windows desktop Excel, with candidate 9 loaded in a separate research-demo process. The Enron research workbook was left open; only read-only version/capability probes were made there.

The trial reproduced and fixed two assumptions in the original helper: XLL status was returned as `[json]`, and the Office.js sandbox did not define `setTimeout`. Regression tests cover both, including absent/multi-cell responses and later-call polling without query resubmission.

Verified through the live bridge:

- A producing cell's completed status and its retained table's schema/row count.
- A bounded query returning two of three rows with `hasMore: true`.
- Zero-row schema inspection containing `error('must not execute')`, without evaluating the error expression.
- A query bound to an existing retained table, returning its independently known 58-row count.
- Long-request staging with quotes, backslashes and Unicode; exact `9007199254740993` and text `0007` survived the roundtrip.
- A missing SQL column returned the structured `sql_binding` failure category.
- Worksheet names/used ranges were unchanged, and the anchor cell remained blank.

This establishes the bridge in the observed host, not compatibility with every Excel build or agent. The trial used synthetic/local retained data; existing native tests cover source access. No macro-command capability, automatic skill installation, global settings change or external-source refresh was inferred from Office.js access.
