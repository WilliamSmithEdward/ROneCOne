# Development and Excel Safety

This document is for contributors and reviewers. It records how the runtime is validated: the
five test gates, the popup-adaptive Excel harness, benchmark baselines, and the demo rebuild
pipeline. Using the runtime requires none of this tooling.

## Test layers

ROneCOne uses five independent gates:

1. Python source-contract tests enforce the one-file invariant, public API, ASCII portability,
   IntelliSense metadata, identifier casing, absence of runtime VBIDE/process dependencies, and
   the security scan's rules.
2. pyVBAanalysis checks the runtime on its own, and with all VBA fixtures as one project.
3. pyOpenVBA builds test and demo workbooks and verifies byte-for-byte module round trips.
4. Microsoft Excel compiles and executes the VBA suite, records worksheet-observed assertions,
   runs delegate and collection benchmarks, and exports a host project through the VBE to prove
   the runtime recases none of its names.
5. olevba and MacroRaptor scan `ROneCOne.cls` and every demo workbook, and
   `tools/security_scan.py` holds their results to the reviewed baseline in
   `tools/security_baseline.json`.

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

Every benchmark reads VBA's `Timer`, which returns a Single, so it moves in 1/256-second steps
before 18:12 and 1/128-second steps after. A scenario that finishes inside one step reads as zero,
and the live suite and the demo runner fail a zero reading, because a scenario that never ran
leaves the same value. Each gated scenario therefore runs for several steps. Since 1.10.2 the
hash-set read scenario makes ten passes over its 10,000 items, 100,000 reads, and the Zip demo
inflates a 50,000-line file; at one pass and 1,000 lines both took about one step.

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

pyVBAanalysis 2.2.0 reports locals, module-private variables, and constants that nothing uses, and
variables that are assigned but never read. The gate fails on these like any other diagnostic. VBA
has no discard syntax, so a Function called only for its effect is written as a statement, such as
`dictionary.EnsureCapacity 10000&`, rather than assigned to a throwaway variable. Where VBA
requires the assignment, because the value comes from a property read, a conversion probe, or a
`delegate(args)` call, the code reads the value it got: a test asserts on it, and the runtime's CSV
classifier returns its probe's type. Bang expressions such as `!Age` are member access to the
analyzer, so they need no local declaration, which versions before 2.2.0 required.

`requirements-dev.txt` pins the pyVBAanalysis a local checkout runs, while CI installs the newest
release, so each new rule meets this code as soon as it ships and can fail CI with no change here.

`ROneCOne.cls` ships as one file, so it is also analyzed alone:

```powershell
.venv\Scripts\pyvbaanalysis.exe --whole-project src\ROneCOne.cls `
    --no-inline-suppression --format text
```

In the run over `src`, `tests\vba`, and `demo\vba` together, a call from the runtime into a test or
demo module resolves. A single file is otherwise analyzed as a project fragment, which skips the
whole-project checks; `--whole-project`, added in pyVBAanalysis 2.3.1, analyzes the runtime as the
complete project it is when a user imports it, and a call like that fails.

## Security scan

olevba and MacroRaptor (mraptor), both from oletools, report what the shipped VBA could do. The
runtime reads and writes files, runs processes, calls native code, and talks HTTP by design, so
olevba always reports Suspicious keywords and IOCs for it and mraptor always flags it `-WX`, which
[SECURITY.md](../SECURITY.md) explains. `tools/security_scan.py` scans `src/ROneCOne.cls` and every
`demo/*.xlsm` and compares each file's olevba findings and mraptor flags with
`tools/security_baseline.json`. It fails on a finding that appears or disappears and on flags that
change. It also fails on code that runs on its own, which olevba reports as AutoExec and mraptor
flags A, and on P-code its source does not explain, and `--update-baseline` refuses to record
either. olevba's own VBA stomping flag is set aside, because pcodedmp prints eight runtime names
with a trailing type character the source never spells; the script repeats olevba's comparison
with that suffix allowed and fails on anything else it finds. `tests/python/test_security_scan.py`
runs each of those failures on small modules written for it.

```powershell
.venv\Scripts\python.exe tools\security_scan.py
.venv\Scripts\python.exe tools\security_scan.py --update-baseline
```

Update the baseline only after reviewing what it records, in the same commit as the change that
caused it. Repackaging the demo workbooks can change their findings, so a release runs the scan
once the workbooks are final. The Security workflow runs the check on every push and pull request.
When a release is published, the release-security workflow downloads its assets, writes
`<tag>-security-report.md` with each asset's olevba findings, mraptor verdict, SHA-256 hash, and
comparison with the baseline at the tag, and attaches the report to the release.

The same Security workflow runs `tools/malware_scan.py` using ClamAV's official signatures and a
pinned YARA Forge Core package through YARA-X. The scan covers the shipped runtime, every demo
workbook, and each workbook ZIP member. See [SECURITY.md](../SECURITY.md#malware-signatures) for
the reviewed, exact-hash exception format and failure behavior. To run it locally, install ClamAV,
update its signatures with `freshclam`, install `yara-x==1.20.0` in the project environment, and
download the Core ZIP from the pinned YARA Forge release:

```powershell
.venv\Scripts\python.exe tools\malware_scan.py yara-forge-rules-core.zip
```

## Identifier casing

VBA keeps one spelling per identifier across a whole project, and a declaration in any module
sets it. A parameter named `value` in the runtime turns every `.Value` in the host's modules
into `.value`, so each export of the host's code carries the change as diff noise
([issue #6](https://github.com/WilliamSmithEdward/ROneCOne/issues/6)). The runtime therefore
declares no name whose spelling differs from the one the default references use: VBA, Excel,
stdole, and Office. Public parameters take the reference spelling (`Value`, `Index`,
`Predicate`), because IntelliSense shows them and VBA matches named arguments regardless of case.
Everything else takes a name no reference defines, such as `itemValue`, `idx`, or `outcome`.
`Guid` and `Xml` are the two deliberate exceptions, kept at their .NET spelling over Office's
`GUID` and `XML`. The demo modules ship beside the runtime and follow the same rule, spelling
ROneCOne's members as the runtime does, and the VBA snippets on their worksheets quote those
names rather than a placeholder such as `path` or `text` that would recase a reader's module.
The VBA examples in README.md and docs/ follow the same rule, because readers copy them into
their own modules: no name a library or ROneCOne spells differently, and no variable named like
the procedure it wraps.

`tests/python/test_casing.py` holds the runtime and the demo modules to one spelling per name
everywhere, and to the reference spelling where pywin32 can read the registered type libraries,
so CI runs the first half and a local Windows run adds the second. The live check builds a
workbook holding the runtime, every demo module, and a host module that uses Excel and ROneCOne
members without declaring any of them, opens it in a task-owned Excel, exports every module
through the VBE, and requires every token back as written:

```powershell
powershell -ExecutionPolicy Bypass -File tools\run_casing_roundtrip.ps1
```

`-RuntimePath` points it at another copy of `ROneCOne.cls`. The export uses development-only
VBIDE trust, the same as demo conversion. The runtime itself never touches the VBIDE.

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
    -InputPath demo\.working\ROneCOne_Tasks_Demo.xlsx `
    -OutputPath demo\ROneCOne_Tasks_Demo.xlsm
powershell -ExecutionPolicy Bypass -File tools\convert_demo_workbook.ps1 `
    -InputPath demo\.working\ROneCOne_Data_Demo.xlsx `
    -OutputPath demo\ROneCOne_Data_Demo.xlsm
powershell -ExecutionPolicy Bypass -File tools\convert_demo_workbook.ps1 `
    -InputPath demo\.working\ROneCOne_Http_Demo.xlsx `
    -OutputPath demo\ROneCOne_Http_Demo.xlsm
powershell -ExecutionPolicy Bypass -File tools\convert_demo_workbook.ps1 `
    -InputPath demo\.working\ROneCOne_Json_Demo.xlsx `
    -OutputPath demo\ROneCOne_Json_Demo.xlsm
powershell -ExecutionPolicy Bypass -File tools\convert_demo_workbook.ps1 `
    -InputPath demo\.working\ROneCOne_Files_Demo.xlsx `
    -OutputPath demo\ROneCOne_Files_Demo.xlsm
powershell -ExecutionPolicy Bypass -File tools\convert_demo_workbook.ps1 `
    -InputPath demo\.working\ROneCOne_Process_Demo.xlsx `
    -OutputPath demo\ROneCOne_Process_Demo.xlsm
powershell -ExecutionPolicy Bypass -File tools\convert_demo_workbook.ps1 `
    -InputPath demo\.working\ROneCOne_Text_Demo.xlsx `
    -OutputPath demo\ROneCOne_Text_Demo.xlsm
powershell -ExecutionPolicy Bypass -File tools\convert_demo_workbook.ps1 `
    -InputPath demo\.working\ROneCOne_DateTime_Demo.xlsx `
    -OutputPath demo\ROneCOne_DateTime_Demo.xlsm
powershell -ExecutionPolicy Bypass -File tools\convert_demo_workbook.ps1 `
    -InputPath demo\.working\ROneCOne_Xml_Demo.xlsx `
    -OutputPath demo\ROneCOne_Xml_Demo.xlsm
powershell -ExecutionPolicy Bypass -File tools\convert_demo_workbook.ps1 `
    -InputPath demo\.working\ROneCOne_Zip_Demo.xlsx `
    -OutputPath demo\ROneCOne_Zip_Demo.xlsm
powershell -ExecutionPolicy Bypass -File tools\convert_demo_workbook.ps1 `
    -InputPath demo\.working\ROneCOne_Query_Demo.xlsx `
    -OutputPath demo\ROneCOne_Query_Demo.xlsm
powershell -ExecutionPolicy Bypass -File tools\convert_demo_workbook.ps1 `
    -InputPath demo\.working\ROneCOne_ListObject_Demo.xlsx `
    -OutputPath demo\ROneCOne_ListObject_Demo.xlsm
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
    -WorkbookPath demo\ROneCOne_Tasks_Demo.xlsm `
    -MacroName RunROneCOneTasksDemo
powershell -ExecutionPolicy Bypass -File tools\run_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Data_Demo.xlsm `
    -MacroName RunROneCOneDataDemo
powershell -ExecutionPolicy Bypass -File tools\run_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Http_Demo.xlsm `
    -MacroName RunROneCOneHttpDemo
powershell -ExecutionPolicy Bypass -File tools\run_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Json_Demo.xlsm `
    -MacroName RunROneCOneJsonDemo
powershell -ExecutionPolicy Bypass -File tools\run_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Files_Demo.xlsm `
    -MacroName RunROneCOneFilesDemo
powershell -ExecutionPolicy Bypass -File tools\run_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Process_Demo.xlsm `
    -MacroName RunROneCOneProcessDemo
powershell -ExecutionPolicy Bypass -File tools\run_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Text_Demo.xlsm `
    -MacroName RunROneCOneTextDemo
powershell -ExecutionPolicy Bypass -File tools\run_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_DateTime_Demo.xlsm `
    -MacroName RunROneCOneDateTimeDemo
powershell -ExecutionPolicy Bypass -File tools\run_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Xml_Demo.xlsm `
    -MacroName RunROneCOneXmlDemo
powershell -ExecutionPolicy Bypass -File tools\run_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Zip_Demo.xlsm `
    -MacroName RunROneCOneZipDemo
powershell -ExecutionPolicy Bypass -File tools\run_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Query_Demo.xlsm `
    -MacroName RunROneCOneQueryDemo
powershell -ExecutionPolicy Bypass -File tools\run_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_ListObject_Demo.xlsm `
    -MacroName RunROneCOneListObjectDemo
powershell -ExecutionPolicy Bypass -File tools\render_demo_workbook.ps1 -Clean
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
    -WorkbookPath demo\ROneCOne_Tasks_Demo.xlsm `
    -OutputPrefix tasks
powershell -ExecutionPolicy Bypass -File tools\render_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Data_Demo.xlsm `
    -OutputPrefix data
powershell -ExecutionPolicy Bypass -File tools\render_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Http_Demo.xlsm `
    -OutputPrefix http
powershell -ExecutionPolicy Bypass -File tools\render_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Json_Demo.xlsm `
    -OutputPrefix json
powershell -ExecutionPolicy Bypass -File tools\render_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Files_Demo.xlsm `
    -OutputPrefix files
powershell -ExecutionPolicy Bypass -File tools\render_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Process_Demo.xlsm `
    -OutputPrefix process
powershell -ExecutionPolicy Bypass -File tools\render_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Text_Demo.xlsm `
    -OutputPrefix text
powershell -ExecutionPolicy Bypass -File tools\render_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_DateTime_Demo.xlsm `
    -OutputPrefix datetime
powershell -ExecutionPolicy Bypass -File tools\render_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Xml_Demo.xlsm `
    -OutputPrefix xml
powershell -ExecutionPolicy Bypass -File tools\render_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Zip_Demo.xlsm `
    -OutputPrefix zip
powershell -ExecutionPolicy Bypass -File tools\render_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_Query_Demo.xlsm `
    -OutputPrefix query
powershell -ExecutionPolicy Bypass -File tools\render_demo_workbook.ps1 `
    -WorkbookPath demo\ROneCOne_ListObject_Demo.xlsm `
    -OutputPrefix listobject
```

Development-only VBIDE trust is used once during each conversion to seed an otherwise empty
`vbaProject.bin`. pyOpenVBA replaces that seed. Neither final workbook nor the runtime uses VBIDE
automation. The render pass re-evaluates the status formulas in the artifact-tool spreadsheet
engine, which coerces `0 = ""` to true the way JavaScript does; Excel itself computes PASS and
the run gate is authoritative, but demo examples avoid a raw `0` as an expected value so the
review renders stay truthful. Excel also coerces time-shaped text (such as `36:7`) into a date
serial on write, so demo values never look like clock times either. Every core capability gets a separate workbook with its own macro, examples,
benchmark, live execution gate, and all-sheet render pass. Renders accumulate in the ignored
`demo\.working` directory across runs; pass `-Clean` to the first render call of a batch to drop
every stale PNG set first.

The renderer draws text in narrower fonts than Excel, so a render can show a line whole that Excel
splits or clips. Column widths and row heights in the builders therefore follow Excel's own
measurements, not the renders. A code line holds at most 57 characters in the capability demos'
Consolas column, and longer ones break with a VBA line continuation. Each example row takes its
height from its longest cell, and every table row gets an explicit height: Excel keeps the height
a row was saved with, so a row left unsized shows only the first line of wrapped text. The
artifact tool also stores a JavaScript Boolean as an Excel checkbox, so builders write an expected
Boolean as `=TRUE` or `=FALSE` to match the plain value the macro writes; the renderer draws both
as 1 and 0, where Excel shows TRUE and FALSE.

The collections workbook packages one demo-only class, `DemoCustomer`, as the user model. Typed
member expressions remove the need for a predicate/selector adapter class. The public
`RunROneCOneCollectionsDemo` macro is only an orchestrator; primitive examples, user-class LINQ,
benchmarking, reporting, and helpers are kept in small commented procedures. The delegates demo
uses the same organization and leads with inferred `AsFunc` expressions.

## Release headers

Every module that ships opens with its release and the full MIT license, directly below
`Option Explicit`: the runtime and each module the demo workbooks package. `ROneCOne.cls`
downloaded alone, or a module copied out of a workbook, travels without the repository's
`LICENSE` file, so the notice travels inside it.

The version and date come from the newest dated heading in `CHANGELOG.md`, and the notice from
`LICENSE`. After dating a release in the changelog, restamp every module and repackage the
workbooks:

```powershell
.venv\Scripts\python.exe tools\stamp_release_headers.py
.venv\Scripts\python.exe tools\package_demo_workbook.py
```

The source contract rebuilds the expected header from those two files and fails on any module
that differs, locally and in CI. The header carries no repository URL, because the offline demo
contracts reject `https://` anywhere in their source.

[Back to the documentation index](README.md)
