# Execution architecture checkpoint — 24 October 2026

Use one lazily started worker process per enabled workbook, with one serial DuckDB execution queue inside each worker. Keep the Excel host responsible for formula arguments, workbook/session identity, result subscriptions, and publication on Excel's main thread. Keep intermediate tables in the worker; transfer only captures and bounded output/preview data.

The experiment supports this decision: the process boundary adds a small cost to these short workloads, while a worker failure can be contained and recovered without terminating Excel. This is fault isolation, not an operating-system security sandbox. The agreed read-and-compute capability policy remains separate work.

## Measured comparison

These are exploratory measurements on one Windows development machine, using installed .NET 8, Excel-DNA 1.9.0, and DuckDB.NET 1.5.6.1. Runs were sequential. They are not distribution-wide performance guarantees or percentile benchmarks. Raw measurements are in [results/2026-10-24](results/2026-10-24).

The Excel workload contains a source query, eight dependent query cells, and a scalar output. Ordinary notification uses Excel-DNA's RTD path. Targeted publication additionally queues recalculation of the requesting cell on Excel's main thread when its result is ready. It marks that cell dirty and evaluates its current formula; it does not write a result over a formula or change Excel's application-wide RTD throttle, which remained 2,000 ms.

| Excel workload | In-process | Worker |
| --- | ---: | ---: |
| Ordinary RTD, cold chain | 18.16 s | 18.23 s |
| Ordinary RTD, two edits | 19.99–20.01 s | 19.99–20.00 s |
| Targeted publication, cold chain | 168 ms | 252 ms |
| Targeted publication, three edits | 22–33 ms | 26–43 ms |
| Unchanged full recalculation, targeted publication | 7.3 ms | 5.7 ms |
| Two branches and an Excel sum, targeted publication | 37 ms | 30 ms |

Notification delay dominates the ordinary RTD chain. The worker and in-process edit timings overlap at this sample size. The cold timings start after loading the XLL and include lazy engine/worker initialization; they do not measure Excel launch or installation.

The ordinary RTD column is the fully asynchronous prototype, not the existing production add-in. Production already waits briefly (up to one second) for a new topic to finish and can return short-query results immediately. These measurements therefore do not establish a 20-second-to-milliseconds speedup over production. The new approach removes that blocking wait while avoiding repeated RTD delays in a dependent chain: a completion queues targeted cell recalculation, and the UDF then returns the cached completed result directly. It does not force RTD to deliver faster than its global throttle.

The separate engine harness excludes Excel and measures a chain of 20 queries:

| Engine workload | In-process | Worker |
| --- | ---: | ---: |
| First query and scalar readback | 68.0 ms | 154.2 ms |
| 20-query chain, total | 11.7 ms | 14.4 ms |
| DuckDB execution within that chain | 10.7 ms | 11.4 ms |
| Capture 10,000 numeric rows | 8.3 ms | 27.9 ms |
| Return 10,000 numeric rows | 4.1 ms | 15.6 ms |

The difference between engine and total chain time includes scheduling, serialization, transport, and bookkeeping. It does not isolate each component. JSON is adequate for this prototype; measure wider/larger research data before fixing the production transfer format.

An initialized worker used about 43 MiB working set in the lifecycle run. Excel used about 293 MiB with the worker and 306 MiB with in-process DuckDB. These are single process snapshots, not a memory limit or a leak test. Each worker was configured with two DuckDB threads and a 256 MiB engine memory limit for these small workloads; that limit does not bound total process memory.

## Behaviour established in real Excel

- Unchanged full calculation and exact formula reentry reused results without executing queries again.
- Scalar, spilled, and preview-data requests completed through the same session. Preview tests here retrieve data, not a rendered chart.
- Identical formulas in separate workbooks received separate handles. Restarting one workbook left the other's handle unchanged.
- Replacing and deleting pending formulas cancelled obsolete work and allowed a following short query to finish. A deleted cell was not overwritten by a late completion.
- Two cells shared one pending computation. Removing one subscriber did not cancel the other subscriber's work. A deterministic core test also checks this ownership rule.
- Restart during execution created a new session and retired the old worker. Closing a workbook during execution retired its session/worker while another workbook's result survived.
- Deliberately terminating the worker produced a visible formula error, left Excel running, and recovered after workbook restart. Requests sent after worker exit failed promptly.
- Prototype cleanup exited its owned Excel processes. The harness now treats a process that remains after Quit as a failed cleanup, rather than silently leaving it behind or killing it indiscriminately.
- A worker trunk query completed in 38.5 seconds and produced the expected scalar result. While it was pending, an unrelated ordinary Excel formula was entered and calculated in 15 ms.
- The current add-in and the worker prototype both calculated successfully in the same .NET 8 Excel process in both load orders. This is a limited coexistence check, not evidence about arbitrary third-party add-ins or mixed runtime versions.

Automatic calculation is the supported v1 workflow. In the manual-calculation experiment, a source plus asynchronous scalar output required three explicit calculation passes to settle in both backends. The prototype did not change calculation mode or force dependent calculations in manual mode. Document manual mode as limited; do not build a second dependency scheduler or reinterpret F9/Shift F9 to hide this behaviour.

## Production implementation consequences

Use immutable, session-qualified result handles, typed request identities, shared in-flight work, and subscriber-aware cancellation. Carry session and request identity through every completion. Publish only if the session and subscriber are still current, and let Excel evaluate the current formula. Batch/coalesce publication where useful, with tests for large fan-out and ordinary Excel dependents.

Keep the serial per-workbook queue initially. DuckDB can parallelize a query internally, and different workbook workers can execute independently. Add actual bounded admission and overload behaviour; a semaphore alone does not bound all queued requests. Retain live results and dependencies, and bound historical cache entries. The prototype deliberately retains history until restart and must not be promoted unchanged into production.

Workbook restart is the explicit refresh operation. It retires the old session/worker and rebuilds only that workbook's computations. Normal recalculation, formula reentry, preview selection, and asynchronous notification must not become implicit source refreshes.

Preserve the existing chart pane UX. The worker boundary does not require any in-sheet chart feature. Add a real source-to-chart path after extracting the typed core and enforcing read-and-compute SQL.

## Runtime and release gate still open

The installed .NET 8 runtime is a measurement baseline, not the intended public-release runtime. Microsoft ends .NET 8 and 9 support on 10 November 2026; .NET 10 LTS is the candidate for production. Excel-DNA 1.9 supports .NET 10. See [Microsoft's lifecycle policy](https://learn.microsoft.com/en-ie/lifecycle/products/microsoft-net-and-net-core) and [Excel-DNA 1.9 release notes](https://excel-dna.net/docs/release-notes-1-9/).

Only one modern .NET runtime can be loaded into an Excel process. Moving DuckDB outside Excel does not remove the Excel host's own runtime compatibility constraint. Before changing the production target, test .NET 10 loading, both load orders with representative other add-ins, and the native/worker deployment layout. [Excel-DNA documents this constraint](https://excel-dna.net/docs/guides-basic/dotnet-runtime-support/). The current two-add-in .NET 8 test does not close that gate.

At the time of this prototype checkpoint the production core, host/runtime packaging, supported readers, WebView2 update, and clean-machine release tests remained unimplemented. Subsequent core work is tracked in [the modernization plan](../../PLAN-modernization.md) and [core contract](../../XlDuck.Core/README.md). The prototype itself does not cover Save As, a user cancelling workbook closure, workbook reopen/trust persistence, full chart rendering, or large-data retention pressure. These are explicit remaining work, not implied successes of the lifecycle samples.

Milestone 1 is complete. Milestone 2 has established the worker/publication decision and calculation-mode limit; its runtime/deployment compatibility gate remains open. The next production slice is the typed core and owned workbook sessions, followed by the read-and-compute boundary and one integrated research workflow.
