# XlDuck execution prototype

This separate add-in compares an in-process DuckDB backend with one worker process per workbook. It is an experiment for the modernization plan, not the public add-in. The normal XlDuck project and its worksheet functions are unchanged.

The measured decision is to use a worker per workbook with targeted result publication. Read [DECISION.md](DECISION.md) for the measurements, calculation-mode limit, and remaining runtime/release gate.

Both variants use the same engine, protocol, cache, and worksheet integration. The experiment uses Excel-DNA 1.9.0 and DuckDB.NET 1.5.6.1. It currently targets the installed .NET 8 runtime; the public-release runtime remains a separate compatibility decision. External access is disabled for these synthetic local workloads. This prototype does not implement the production SQL capability policy.

Build the worker and prototype add-in, then run the independent engine checks:

```powershell
dotnet build prototypes/Execution/Worker
dotnet build prototypes/Execution/AddIn
dotnet run --project prototypes/Execution/Checks
```

The checks cover a chain of 20 queries, typed input and cache identity, reuse, cancellation and recovery, shared subscribers, worker termination, foreign-handle rejection, and transfer of a 10,000-row capture and output. They print timing measurements separately from their assertions.

Run the Excel comparison in owned, isolated Excel processes:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File prototypes/Execution/Run-Experiment.ps1 -NoBuild -Mode inproc -Publisher rtd
powershell -NoProfile -ExecutionPolicy Bypass -File prototypes/Execution/Run-Experiment.ps1 -NoBuild -Mode worker -Publisher rtd
powershell -NoProfile -ExecutionPolicy Bypass -File prototypes/Execution/Run-Experiment.ps1 -NoBuild -Mode inproc -Publisher targeted
powershell -NoProfile -ExecutionPolicy Bypass -File prototypes/Execution/Run-Experiment.ps1 -NoBuild -Mode worker -Publisher targeted
```

The RTD variant uses ordinary Excel-DNA result notifications. The targeted variant also queues calculation of the requesting cell when its work finishes, and only in automatic calculation mode. It evaluates the cell's current formula and never writes a result over a formula. Neither variant changes Excel's RTD throttle or keyboard shortcuts. The comparison establishes whether publication latency, rather than transport, dominates short query chains.

Measurements are written to the ignored `artifacts/execution` directory. Excel checks cover dependent results, edits, full-calculation reuse, unchanged formula reentry, scalar/spilled/preview data, two workbooks, and restarting one workbook without restarting the other.

Run lifecycle checks after building. They use their own Excel instances and change calculation mode only within those owned instances:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File prototypes/Execution/Test-Lifecycle.ps1 -Mode inproc
powershell -NoProfile -ExecutionPolicy Bypass -File prototypes/Execution/Test-Lifecycle.ps1 -Mode worker
powershell -NoProfile -ExecutionPolicy Bypass -File prototypes/Execution/Test-Coexistence.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File prototypes/Execution/Test-SlowTrunk.ps1 -Mode worker
```

The coexistence check also requires the normal XlDuck project to have been built. The slow-trunk check intentionally runs a longer query (about 38 seconds on the development machine). Do not run concurrent workloads when collecting comparative timing samples. Cleanup verifies process exit and reports a failure if Excel remains; it does not forcibly kill an unresponsive instance.

## Prototype constraints

The query API intentionally has one input relation named `__input` and one optional named scalar parameter `$value`. Captures accept one numeric column. Results use a small typed JSON protocol. Scalar, sheet, and preview requests retrieve bounded data; intermediate tables stay in the engine. These are measurement conveniences, not the final worksheet API or a commitment to JSON for large transfers.

Session identity currently follows the workbook name and does not implement Save As or a user cancelling workbook closure. The cache retains all successful results until session disposal, so use bounded experiments; production needs dependency ownership and bounded retention. Shared in-flight subscribers have cancellation tests, but this is not the final ownership system for retained fragments and plots. Preview checks exercise data retrieval, not full chart rendering or a replacement pane.

Failures, workbook closure, and cancellation must be evaluated alongside latency. A worker process is fault isolation, not a security sandbox. Do not load this add-in into a research workbook or distribute it as a release build.
