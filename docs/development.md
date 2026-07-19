# Development and Excel Safety

## Test layers

ROneCOne uses four independent gates:

1. Python source-contract tests enforce the one-file invariant, public API, ASCII portability,
   IntelliSense metadata, and absence of runtime VBIDE/process dependencies.
2. pyVBAanalysis checks the runtime and all VBA fixtures as one project.
3. pyOpenVBA builds test and demo workbooks and verifies byte-for-byte module round trips.
4. Microsoft Excel compiles and executes the VBA suite, records worksheet-observed assertions,
   and runs the invocation benchmark.

The live suite currently exercises lambda creation, unary and binary calls, explicit and default
invocation, comparisons, short-circuit behavior, typed failures, method/action delegates, object
returns, composition, and unbound-parameter rejection.

The invocation benchmark has a configurable release ceiling (`-MaxBenchmarkSeconds`, default
`0.5` for 10,000 calls). The v0.1.0 measurements are stored in
`benchmarks/v0.1.0-baseline.json`.

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
$env:NODE_PATH = "C:\Users\William\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\node_modules"
node tools\build_demo_workbook.cjs
powershell -ExecutionPolicy Bypass -File tools\convert_demo_workbook.ps1
.venv\Scripts\python.exe tools\package_demo_workbook.py
powershell -ExecutionPolicy Bypass -File tools\run_demo_workbook.ps1
powershell -ExecutionPolicy Bypass -File tools\render_demo_workbook.ps1
```

Development-only VBIDE trust is used once during conversion to seed an otherwise empty
`vbaProject.bin`. pyOpenVBA replaces that seed. Neither the final workbook nor the runtime uses
VBIDE automation.
