# Development and Excel Safety

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
capabilities. The current live total is 414 assertions across all four suites.

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
```

Development-only VBIDE trust is used once during each conversion to seed an otherwise empty
`vbaProject.bin`. pyOpenVBA replaces that seed. Neither final workbook nor the runtime uses VBIDE
automation. Every core capability gets a separate workbook with its own macro, examples,
benchmark, live execution gate, and all-sheet render pass.

The collections workbook packages one demo-only class, `DemoCustomer`, as the user model. Typed
member expressions remove the need for a predicate/selector adapter class. The public
`RunROneCOneCollectionsDemo` macro is only an orchestrator; primitive examples, user-class LINQ,
benchmarking, reporting, and helpers are kept in small commented procedures. The delegates demo
uses the same organization and leads with inferred `AsFunc` expressions.
