# Development and Excel Safety

This document is for contributors and reviewers. It records how the runtime is validated: the
four test gates, the popup-adaptive Excel harness, benchmark baselines, and the demo rebuild
pipeline. Using the runtime requires none of this tooling.

## Test layers

ROneCOne uses four independent gates:

1. Python source-contract tests enforce the one-file invariant, public API, ASCII portability,
   IntelliSense metadata, and absence of runtime VBIDE/process dependencies.
2. pyVBAanalysis checks the runtime and all VBA fixtures as one project.
3. pyOpenVBA builds test and demo workbooks and verifies byte-for-byte module round trips.
4. Microsoft Excel compiles and executes the VBA suite, records worksheet-observed assertions,
   and runs delegate and collection benchmarks.

The live suite exercises explicit and inferred lambda creation, `Var`/`VarLike`, unary and binary
calls, explicit and default invocation, comparisons, short-circuit behavior, typed failures,
object-member expressions, universal `Func`/`Action` adapters, callable objects, procedure calls,
object returns, `DynamicInvoke`, metadata, multicast ordering/removal, native function pointers,
true `ByRef`, fail-closed ABI boundaries, composition, typed events, structured exceptions, and
unbound-parameter rejection.

The collection suite adds strict primitive/user-class lists, atomic mutation failures, zero-based
indexing, populated/inferred initializers, atomic heterogeneous inputs, deferred source mutation,
universal procedure delegates in LINQ, `ForEach(Action)`, text joining, query chaining, numeric
terminals, predicate algebra, collection membership, null-safe paths, custom comparers, nested
quantifiers, stable composite ordering, indexed hash collections, capacity control, and enumerator
refresh after mutation. Dedicated advanced-collection and task/data/provider suites cover
specialized families, scheduler state, cancellation, indexed data relations, and provider
capabilities. The current live total is 505 assertions across all four suites.

The invocation benchmark has a configurable release ceiling (`-MaxBenchmarkSeconds`, default
`0.5` for 10,000 calls). The v0.1.0 measurements are stored in
`benchmarks/v0.1.0-baseline.json`.

The v0.2.0 collection gate requires a 10,000-element `Range.Where.ToList` pipeline to complete in
at most `0.75` seconds. Measurements are stored in `benchmarks/v0.2.0-baseline.json`.

The v0.3.0 baseline re-runs both gates through the concise `AsFunc` and implicit `Where` forms.
Measurements are stored in `benchmarks/v0.3.0-baseline.json`.

The v0.5.0 baseline runs the same hot paths after the universal delegate kernel, native ABI, and
multicast changes. Measurements and live assertion totals are stored in
`benchmarks/v0.5.0-baseline.json`.

The v0.6.0 baseline repeats those gates after collection initializers, typed events, and structured
exceptions. It records three fresh-process samples plus the expanded live assertion totals in
`benchmarks/v0.6.0-baseline.json`.

The v0.7.0 baseline repeats them after contextual LINQ, inferred sequence predicates, key-based
operators, and native bang-member expressions. The collection suite now includes fluent condition
builders, composed stable-parameter expressions, dotted object paths, member-name selectors, and
procedure predicate inference. Five fresh-process samples are recorded in
`benchmarks/v0.7.0-baseline.json`.

The v0.8.0 baseline adds a repeated object-member query to measure the cached member plan and
repeats both established release gates across five fresh Excel processes. Measurements are stored
in `benchmarks/v0.8.0-baseline.json`.

The v1.0.0 baseline retains the 10,000-element `OrderBy.ThenByDescending.ToList` scenario after the
stable O(n log n) ordering rewrite. Three fresh Excel processes established its initial gate.
Measurements are recorded in
`benchmarks/v1.0.0-baseline.json`.

The v1.1.0 baseline adds a 10,000-item dictionary build-and-read scenario and 100,000 indexed
lookups over a reserved 10,000-item dictionary. The live release gates are 1.5 and 2.0 seconds.
The composite-ordering gate is 2.0 seconds to cover directly observed fresh-process host variance;
its implementation and 10,000-element workload are unchanged. Three samples and all 414 live
assertions are recorded in `benchmarks/v1.1.0-baseline.json`.

The v1.2.0 baseline adds 1,000 complete native `Task.Run(...).Await` lifecycles, including
expression verification, bytecode marshaling, Windows thread-pool dispatch, typed result recovery,
and deterministic resource cleanup. Its 1.5-second release gate is based on three fresh-process
samples with a 0.5234375-second median. The same three processes repeat every established
performance gate and all 425 live assertions. Measurements are recorded in
`benchmarks/v1.2.0-baseline.json`.

The v1.3.0 baseline follows the removal of the native execution slice, the renaming of the
cooperative scheduling call to `Task.Run` ([ADR 0001](decisions/0001-remove-native-task-run.md),
[ADR 0002](decisions/0002-task-run-names-the-cooperative-scheduler.md)), lossless numeric
widening, the worksheet Range bridge, and array-backed element storage
([ADR 0003](decisions/0003-lossless-numeric-widening.md) through
[ADR 0005](decisions/0005-array-backed-collection-storage.md)). The task scenario is 1,000
cooperative `Task.Run(...).Await` lifecycles against the retained 1.5-second gate, roughly ten
times faster than the removed native path measured on the same machine. The collection suite adds
numeric-widening, indexed-access-scaling, and expression-display contracts; the task and data
suite adds the Range bridge contract. Three fresh processes repeat every established gate and all
444 live assertions. Measurements are recorded in `benchmarks/v1.3.0-baseline.json`.

The v1.4.0 baseline follows the in-place hash maintenance
([ADR 0009](decisions/0009-in-place-hash-index-maintenance.md)), version-cached snapshots and
O(1) positional access ([ADR 0010](decisions/0010-version-cached-positional-access.md)), and the
awaitable HTTP client ([ADR 0011](decisions/0011-awaitable-http-client-over-winhttp.md)). Four
scenarios join the gated set: 12,000 keyed mutations, 10,000 indexed list writes, 2,000
positional row reads through `Rows`, and 10,000 positional hash-set reads, each against a
1.5-second gate. The task and data suite adds twenty HTTP assertions against
https://pokeapi.co. Three fresh processes repeat every gate and all 505 live assertions.
Measurements are recorded in `benchmarks/v1.4.0-baseline.json`.

The v1.5.0 baseline follows the incremental constraint indexes
([ADR 0012](decisions/0012-incremental-constraint-indexes.md)) and the JSON layer
([ADR 0013](decisions/0013-json-in-the-spirit-of-system-text-json.md)). Two scenarios join the
gated set: 6,000 operations on a table with a primary key and a unique column against a
1.5-second gate, and a 1,000-row table round-tripped through `ToJson` and `DeserializeTable`
against a 2.5-second gate. The task and data suite adds the JSON contracts, and the eighth demo
workbook exercises the JSON surface offline. Three fresh processes repeat every gate and all
548 live assertions. Measurements are recorded in `benchmarks/v1.5.0-baseline.json`.

The HTTP contract in the task and data suite, and the HTTP demo workbook, make live requests to
https://pokeapi.co, so those runs need internet access; every other gate runs offline. No other
host is contacted.

The SQL Server contract in the task and data suite ([ADR
0014](decisions/0014-native-provider-async-over-adodb.md)) connects to the local default
instance as `Provider=MSOLEDBSQL;Data Source=localhost;Initial Catalog=tempdb;Integrated
Security=SSPI;`. Running the live suite therefore requires a reachable localhost SQL Server
with the Microsoft OLE DB driver installed and a Windows login allowed into tempdb; every
scratch object is a session temp table, so nothing persists. No credential is stored in the
repository.

pyVBAanalysis 1.2.0 treats a VBA bang identifier such as `!Age` as an ordinary variable token.
Live bang examples therefore declare the token name in their local scope even though VBA uses it
as a default-member name, not a variable value. This keeps both the normal and
`--no-inline-suppression` complete-project gates clean. Excel compilation and execution remain the
decisive host-level checks for that syntax.

## Popup-adaptive Excel harness

Excel COM calls block while an Office or VBE modal surface is open. Every live macro run therefore
creates a fresh task-owned Excel process and starts a watcher keyed to that exact process ID.

The watcher:

- inventories every visible non-primary window owned by that Excel process;
- captures Win32 class names, titles, child text, buttons, and handles;
- captures UI Automation names, control types, IDs, and selected VBE code;
- prefers `Cancel` or `Close`, or `OK` when all other choices are non-decision helpers;
- closes an unknown modal conservatively after recording its complete surface;
- detects VBE `[break]` mode after a popup and terminates only the task-owned Excel process;
- writes local JSONL diagnostics under ignored test/demo output directories.

An independent outer process imposes a 30-second default deadline. If the worker, COM cleanup, or
watcher stalls without a visible popup, the outer watchdog terminates the recorded Excel, watcher,
and worker PIDs. It never enumerates or closes user-open Excel instances.

This harness is development infrastructure. The deployed `ROneCOne.cls` neither inspects windows
nor accesses the VBIDE.

## Rebuild the living demo

The visible workbook is authored with the artifact-tool spreadsheet runtime, rendered to PNG for
visual inspection, converted by a bounded task-owned Excel process, and then populated through
pyOpenVBA.

```powershell
# Set NODE_PATH to the artifact-tool node_modules directory reported by the
# workspace dependency loader, or expose that directory as the local node_modules.
node tools\build_demo_workbook.cjs
node tools\build_collections_demo_workbook.cjs
node tools\build_capability_demo_workbooks.cjs
powershell -ExecutionPolicy Bypass -File tools\convert_demo_workbook.ps1
powershell -ExecutionPolicy Bypass -File tools\convert_demo_workbook.ps1 `
    -InputPath demo\.working\ROneCOne_Collections_Demo.xlsx `
    -OutputPath demo\ROneCOne_Collections_Demo.xlsm
powershell -ExecutionPolicy Bypass -File tools\convert_demo_workbook.ps1 `
    -InputPath demo\.working\ROneCOne_Events_Demo.xlsx `
    -OutputPath demo\ROneCOne_Events_Demo.xlsm
powershell -ExecutionPolicy Bypass -File tools\convert_demo_workbook.ps1 `
    -InputPath demo\.working\ROneCOne_Exceptions_Demo.xlsx `
    -OutputPath demo\ROneCOne_Exceptions_Demo.xlsm
powershell -ExecutionPolicy Bypass -File tools\convert_demo_workbook.ps1 `
    -InputPath demo\.working\ROneCOne_Http_Demo.xlsx `
    -OutputPath demo\ROneCOne_Http_Demo.xlsm
powershell -ExecutionPolicy Bypass -File tools\convert_demo_workbook.ps1 `
    -InputPath demo\.working\ROneCOne_Json_Demo.xlsx `
    -OutputPath demo\ROneCOne_Json_Demo.xlsm
.venv\Scripts\python.exe tools\package_demo_workbook.py
powershell -ExecutionPolicy Bypass -File tools\run_demo_workbook.ps1
powershell -ExecutionPolicy Bypass -File tools\run_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Collections_Demo.xlsm `
    -MacroName RunROneCOneCollectionsDemo
powershell -ExecutionPolicy Bypass -File tools\run_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Events_Demo.xlsm `
    -MacroName RunROneCOneEventsDemo
powershell -ExecutionPolicy Bypass -File tools\run_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Exceptions_Demo.xlsm `
    -MacroName RunROneCOneExceptionsDemo
powershell -ExecutionPolicy Bypass -File tools\run_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Http_Demo.xlsm `
    -MacroName RunROneCOneHttpDemo
powershell -ExecutionPolicy Bypass -File tools\run_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Json_Demo.xlsm `
    -MacroName RunROneCOneJsonDemo
powershell -ExecutionPolicy Bypass -File tools\render_demo_workbook.ps1
powershell -ExecutionPolicy Bypass -File tools\render_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Collections_Demo.xlsm `
    -OutputPrefix collections
powershell -ExecutionPolicy Bypass -File tools\render_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Events_Demo.xlsm `
    -OutputPrefix events
powershell -ExecutionPolicy Bypass -File tools\render_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Exceptions_Demo.xlsm `
    -OutputPrefix exceptions
powershell -ExecutionPolicy Bypass -File tools\render_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Http_Demo.xlsm `
    -OutputPrefix http
powershell -ExecutionPolicy Bypass -File tools\render_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Json_Demo.xlsm `
    -OutputPrefix json
```

Development-only VBIDE trust is used once during each conversion to seed an otherwise empty
`vbaProject.bin`. pyOpenVBA replaces that seed. Neither final workbook nor the runtime uses VBIDE
automation. Every core capability gets a separate workbook with its own macro, examples,
benchmark, live execution gate, and all-sheet render pass. Renders accumulate in the ignored
`demo\.working` directory across runs; pass `-Clean` to the first render call of a batch to drop
every stale PNG set first.

The collections workbook packages one demo-only class, `DemoCustomer`, as the user model. Typed
member expressions remove the need for a predicate/selector adapter class. The public
`RunROneCOneCollectionsDemo` macro is only an orchestrator; primitive examples, user-class LINQ,
benchmarking, reporting, and helpers are kept in small commented procedures. The delegates demo
uses the same organization and leads with inferred `AsFunc` expressions.

[Back to the documentation index](README.md)
