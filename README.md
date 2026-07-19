# ROneCOne

ROneCOne is a one-file VBA runtime that makes ordinary Excel VBA feel more like modern C#.
Version 0.4.0 uses `Element` as the clear, LINQ-aligned sequence-parameter name. Inferred
`Func` expressions and concise LINQ remain dependency-free,
with no runtime install, add-in, network access, external library, or trusted VBIDE access.

## Quick start

Each core capability has its own verified demo workbook:

- [`ROneCOne_Delegates_Demo.xlsm`](demo/ROneCOne_Delegates_Demo.xlsm) runs six delegate and
  expression examples through `RunROneCOneDemo`.
- [`ROneCOne_Collections_Demo.xlsm`](demo/ROneCOne_Collections_Demo.xlsm) runs fourteen `List<T>`
  and LINQ examples through `RunROneCOneCollectionsDemo`. Its dedicated **User Class LINQ** sheet
  filters, projects, orders, quantifies, and aggregates `DemoCustomer` objects.

Both include a capability-specific 10,000-operation benchmark and execute in one Excel process.

To use the runtime in another workbook, import [`src/ROneCOne.cls`](src/ROneCOne.cls) through the
VBE's **File > Import File** command. That one class is the entire deployed runtime.

```vba
Dim x As ROneCOne
Dim square As ROneCOne

Set x = ROneCOne.Var(vbLong)
Set square = x.Multiply(x).AsFunc

Debug.Print square(CLng(9))      ' 81: default-member delegate call
Debug.Print square.Run(CLng(9))  ' 81: explicit form
```

```vba
Dim numbers As ROneCOne
Dim x As ROneCOne

Set numbers = ROneCOne.ListOf(vbLong)
numbers.Add CLng(5)
numbers.Add CLng(20)
Set x = numbers.Element

Set numbers = numbers _
    .Where(x.GreaterThan(CLng(10))) _
    .ToList

Debug.Print numbers.GenericTypeName  ' List<Long>
Debug.Print numbers(0)               ' 20
```

```vba
Dim customer As ROneCOne
Dim customers As ROneCOne
Dim names As ROneCOne
Dim prototype As Customer

Set prototype = New Customer
Set customers = ROneCOne.ListOf(prototype)
Set customer = customers.Element
Set names = customers _
    .Where(customer("Age").AtLeast(CLng(40))) _
    .Map(customer("CustomerName"), vbString) _
    .Sorted _
    .ToList
```

## Version 0.4.0

- immutable expression trees and string-free anonymous lambdas
- `Var` and `VarLike` typed argument sugar
- automatic parameter inference through `.AsFunc` or `Lambda(body)`
- runtime-typed parameters and deterministic contract errors
- unary and binary arithmetic, comparisons, concatenation, and Boolean negation
- short-circuit `AndAlso` and `OrElse` semantics
- object method delegates through `FromMethod`
- scalar and object return values
- delegate composition through `PipeTo`
- callable default-member syntax such as `square(9)`
- strict primitive and exact user-defined class `List<T>` values
- zero-based default and explicit indexers, mutation, and nested `For Each`
- deferred `Where`, `SelectItems`, ordering, slicing, distinct, append/prepend, and reverse
- LINQ-aligned sequence expressions through `Element`
- concise `Map`, `Exists`, `Sorted`, `SortedDescending`, `AtLeast`, and `AtMost` forms
- immediate query terminals, materialization, `Range`, and `Repeat`
- IntelliSense descriptions embedded in the exported class
- a living user-defined-class LINQ tutorial with six formula-verified scenarios
- small, purpose-focused, commented procedures throughout both demo VBA projects

VBA classes cannot declare a public member named `Invoke` because it collides with inherited COM
`IDispatch.Invoke`. ROneCOne uses `Run` as the explicit name and marks it as the default member,
which enables the more C#-like `square(9)` call form.

VBA reserves `Select` and `Any`. The concise surface uses `Map` and `Exists`; the canonical
`SelectItems` and `AnyItem` members remain available. See
[`docs/collections.md`](docs/collections.md) for concise and canonical typed user-class examples,
semantics, and the supported collection surface.

## Runtime contract

- Windows x64 Microsoft 365 Excel
- `.xlsm`, `.xlsb`, and `.xlam` deployment targets
- one imported runtime file and one Excel application process
- no runtime code generation or VBIDE trust
- no elevation, telemetry, or implicit network traffic
- existing VBA remains usable unchanged
- local diagnostics are opt-in and never transmit data

## Quality gates

Every release is exercised at the source, workbook, and Excel-host layers. Python contracts enforce
the one-file API and repository invariants; pyVBAanalysis scans complete VBA projects; pyOpenVBA
verifies module round trips; and bounded Excel workers compile and execute the live suites while a
popup watcher prevents modal Office or VBE windows from hanging automation. Performance baselines
run in the same Excel process as the code under test.

## Development

Use Python 3.10 or later. Development dependencies are isolated from the shipped VBA runtime.

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.venv\Scripts\python.exe -m unittest discover -s tests\python -v
.venv\Scripts\pyvbaanalysis.exe src\ROneCOne.cls tests\vba\DelegateFixture.cls `
    tests\vba\GenericCustomer.cls tests\vba\TestDelegates.bas `
    tests\vba\TestCollections.bas --no-inline-suppression --format text
.venv\Scripts\python.exe tools\build_test_workbook.py
powershell -ExecutionPolicy Bypass -File tools\run_excel_tests.ps1
```

The Excel harness observes Win32 and UI Automation surfaces owned by its exact task process. It
captures and dismisses modal dialogs, records selected VBE code for compiler faults, fails fast on
break mode, and enforces a hard deadline so a hidden Excel instance cannot hang indefinitely. See
[`docs/development.md`](docs/development.md).

The real `%USERPROFILE%\.ronecone.env` is private and optional. The committed
`.ronecone.env.example` defines the reserved local-only diagnostics schema. Version 0.4.0 does not
read this file or emit runtime logs.

## Release roadmap

Capabilities enter the supported surface only after their behavioral, live-host, and performance
gates pass. Delegates, expression lambdas, runtime-generic `List<T>`, and foundational LINQ are
available now. The release sequence continues with structured exceptions, additional generic
collections, tasks and async/await, events, disposables, and native-safe parallel operations. See
[`docs/architecture.md`](docs/architecture.md) for the governing invariants.

## License

MIT. See [`LICENSE`](LICENSE).
