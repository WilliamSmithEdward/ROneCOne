# ROneCOne

ROneCOne is a one-file runtime that brings typed expressions, universal delegates, runtime-generic
collections, and LINQ-style queries to ordinary Excel VBA. Version 0.5.0 unifies expression
lambdas, object methods, callable objects, workbook procedures, multicast chains, and
signature-bound native entry points behind one immutable `Func` / `Action` model.

The deployed runtime is one [`ROneCOne.cls`](src/ROneCOne.cls) file. It requires no install,
add-in, runtime code generation, network access, external package, or trusted VBIDE access.

## Quick start

Each core capability has its own verified workbook:

- [`ROneCOne_Delegates_Demo.xlsm`](demo/ROneCOne_Delegates_Demo.xlsm) runs eleven expression,
  object, procedure, dynamic, multicast, metadata, composition, and true-`ByRef` examples.
- [`ROneCOne_Collections_Demo.xlsm`](demo/ROneCOne_Collections_Demo.xlsm) runs fourteen `List<T>`
  and LINQ examples, including a dedicated **User Class LINQ** tutorial.

Both include a 10,000-operation benchmark and execute in one Excel process. Import
[`ROneCOne.cls`](src/ROneCOne.cls) through the VBE's **File > Import File** command to use the
runtime in another workbook.

```vba
Dim square As ROneCOne
Dim x As ROneCOne

Set x = ROneCOne.Var(vbLong)
Set square = x.Multiply(x).AsFunc

Debug.Print square(CLng(9))      ' 81: natural delegate call
Debug.Print square.Run(CLng(9))  ' 81: explicit form
```

Object methods, callable objects, and workbook procedures share the same typed contract:

```vba
Dim maximum As ROneCOne

Set maximum = ROneCOne.Func(Application.WorksheetFunction, "Max")
Set maximum = maximum.Takes(vbLong, vbLong).Returns(vbDouble)

Debug.Print maximum(CLng(4), CLng(7))
Debug.Print maximum.DynamicInvoke(Array(CLng(4), CLng(7)))
Debug.Print maximum.Signature  ' Func<Long, Long, Double>
```

`Action` delegates combine and remove immutably, matching C# multicast ordering:

```vba
Dim changed As ROneCOne
Dim ignored As Variant

Set changed = ROneCOne.Combine(firstHandler, secondHandler)
ignored = changed("ready")
Set changed = changed.Remove(secondHandler)
```

Typed collections and user-defined classes use the same delegate kernel:

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

## Version 0.5.0

- universal `Func` and `Action` factories for expressions, objects, and workbook procedures
- default callable-object binding through an object's `Run` method
- immutable `.Takes(...)` and `.Returns(...)` signatures for primitives and user classes
- natural call syntax, explicit `Run`, and C#-aligned `DynamicInvoke`
- immutable multicast `Combine`, `Remove`, `GetInvocationList`, and invocation ordering
- `Target`, `MethodName`, `Arity`, `IsAction`, `InvocationCount`, and `Signature` metadata
- fail-closed x64 `Native` / `NativeAction` dispatch through the Windows Automation ABI
- true native `ByRef` with typed reference descriptors and variable wrappers
- expression lambdas, captured values, short-circuit Boolean logic, and `PipeTo` composition
- strict primitive and exact user-defined class `List<T>` values
- deferred LINQ-style filtering, projection, ordering, slicing, quantifiers, and aggregation
- IntelliSense descriptions embedded in the exported class

VBA classes cannot declare `Invoke` because that name collides with their inherited COM dispatch
surface. ROneCOne therefore makes `Run` the explicit member and the default member, enabling the
closer `square(9)` form. VBA also reserves `Select` and `Any`; ROneCOne uses concise `Map` and
`Exists` names while retaining the explicit `SelectItems` and `AnyItem` primitives.

See [`docs/delegates.md`](docs/delegates.md) for the concise and canonical delegate forms, ABI and
`ByRef` safety boundaries, and complete invocation model. See
[`docs/collections.md`](docs/collections.md) for `List<T>` and LINQ semantics.

## Runtime contract

- Windows x64 Microsoft 365 Excel
- `.xlsm`, `.xlsb`, and `.xlam` deployment targets
- one imported runtime file and one Excel application process
- no runtime code generation or VBIDE trust
- no elevation, telemetry, or implicit network traffic
- local diagnostics are opt-in and never transmit workbook data

## Quality gates

Every release is exercised at the source, workbook, and Excel-host layers. Python contracts enforce
the one-file API and repository invariants; pyVBAanalysis scans complete VBA projects; pyOpenVBA
verifies module round trips; and bounded Excel workers compile and execute the live suites. A
process-scoped popup watcher records and dismisses modal Office or VBE surfaces before they can hang
automation. Performance baselines run in the same Excel process as the code under test.

## Development

Use Python 3.10 or later. Development dependencies are isolated from the shipped VBA runtime.

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.venv\Scripts\python.exe -m unittest discover -s tests\python -v
.venv\Scripts\pyvbaanalysis.exe src\ROneCOne.cls tests\vba\DelegateFixture.cls `
    tests\vba\GenericCustomer.cls tests\vba\DelegateProcedures.bas `
    tests\vba\TestDelegates.bas tests\vba\TestCollections.bas `
    --no-inline-suppression --format text
.venv\Scripts\python.exe tools\build_test_workbook.py
powershell -ExecutionPolicy Bypass -File tools\run_excel_tests.ps1
```

The harness observes Win32 and UI Automation surfaces owned by its exact disposable Excel process,
captures selected VBE code for compiler faults, fails fast on break mode, and enforces a hard
deadline. It never enumerates or closes user-open Excel instances. See
[`docs/development.md`](docs/development.md).

The optional `%USERPROFILE%\.ronecone.env` is private. The committed
`.ronecone.env.example` defines the reserved local-only diagnostics schema; version 0.5.0 does not
read it or emit runtime logs.

## Release roadmap

Delegates, expression lambdas, runtime-generic `List<T>`, and foundational LINQ are available now.
The release sequence continues with structured exceptions, additional generic collections, tasks
and async/await, events, disposables, and native-safe parallel operations. See
[`docs/architecture.md`](docs/architecture.md).

## License

MIT. See [`LICENSE`](LICENSE).
