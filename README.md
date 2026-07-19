# ROneCOne

ROneCOne is a one-file runtime that brings typed expressions, universal delegates, runtime-generic
collections, LINQ-style queries, typed events, and structured exceptions to ordinary Excel VBA.
Version 0.8.0 adds a composable predicate algebra, collection-membership expressions, null-safe
member paths, predicate-aware terminals, nested quantifiers, and custom comparer support.

The deployed runtime is one [`ROneCOne.cls`](src/ROneCOne.cls) file. It requires no install,
add-in, runtime code generation, network access, external package, or trusted VBIDE access.

## Quick start

Each core capability has its own verified workbook:

- [`ROneCOne_Delegates_Demo.xlsm`](demo/ROneCOne_Delegates_Demo.xlsm) runs eleven expression,
  object, procedure, dynamic, multicast, metadata, composition, and true-`ByRef` examples.
- [`ROneCOne_Collections_Demo.xlsm`](demo/ROneCOne_Collections_Demo.xlsm) runs twenty-five `List<T>`
  and LINQ examples, including a dedicated **User Class LINQ** tutorial.
- [`ROneCOne_Events_Demo.xlsm`](demo/ROneCOne_Events_Demo.xlsm) demonstrates typed subscription,
  deterministic emission, removal, and handler metadata.
- [`ROneCOne_Exceptions_Demo.xlsm`](demo/ROneCOne_Exceptions_Demo.xlsm) demonstrates catch-all and
  filtered catches, rethrow behavior, and guaranteed finalization.

Each includes a same-process benchmark and executes in one Excel process. Import
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

Set maximum = ROneCOne.Func(Application.WorksheetFunction, "Max") _
    .Takes(vbLong, vbLong) _
    .Returns(vbDouble)

Debug.Print maximum(CLng(4), CLng(7))
Debug.Print maximum.DynamicInvoke(Array(CLng(4), CLng(7)))
Debug.Print maximum.Signature  ' Func<Long, Long, Double>
```

`Action` delegates combine and remove immutably, matching C# multicast ordering:

```vba
Dim changed As ROneCOne
Set changed = ROneCOne.Combine(firstHandler, secondHandler)
changed.Execute "ready"
Set changed = changed.Remove(secondHandler)
```

Typed collections and user-defined classes use the same delegate kernel:

```vba
Dim customers As ROneCOne
Dim names As ROneCOne
Set customers = ROneCOne.ListFrom(ada, grace, katherine)
Set names = customers _
    .Where("Age").AtLeast(CLng(40)) _
    .Map("CustomerName", vbString) _
    .Sorted _
    .ToList
```

Predicates stay declarative across collections and nullable object graphs:

```vba
Set allowedCities = ROneCOne.ListOf(vbString, "London", "Cleveland")

Set selected = customers.Where(allowedCities.Contains(customers!City))
Set managed = customers.Where("Manager?.Age").AtLeast(CLng(40))
Set parents = customers.WhereAny("Reports", reportPredicate)
```

VBA's native bang syntax can remove even the member-name string when desired:

```vba
With customers
    Set names = .Where(!Age.AtLeast(CLng(40))) _
        .Map("CustomerName", vbString) _
        .ToList
End With
```

Typed events and structured exceptions use the same Actions:

```vba
Set changed = ROneCOne.EventOf(vbString) _
    .Subscribe(firstHandler) _
    .Subscribe(secondHandler)
changed.Emit "ready"

Set attempt = ROneCOne.Try(work) _
    .Catch(errorHandler) _
    .Finally(cleanup)
attempt.Execute
```

## Version 0.8.0

- collection membership through `value.IsIn(sequence)`, contextual
  `Where("Member").IsIn(sequenceOrArray)`, and `sequence.Contains(valueExpression)`
- null-safe paths such as `Where("Manager?.Age").AtLeast(40)`
- `Both`, `Either`, `Negated`, and `WhereNot` Boolean composition sugar
- case-insensitive equality, prefix, suffix, containment, and pattern helpers
- predicate-aware `Count`, `FirstOrDefault`, `LastOrDefault`, `SingleItem`,
  `SingleOrDefault`, and `None`
- nested `AnyMatch`, `AllMatch`, `NoneMatch`, `WhereAny`, `WhereAll`, and `WhereNone`
- reusable `Always`, `Never`, `Match`, and `NotMatch` predicate factories
- typed equality and ordering comparer factories across containment, distinctness, sorting,
  extremes, and `SequenceEqual`
- cached member-access plans that retain normalized names directly on expression nodes

Previously delivered capabilities include:

- deferred `Where("Member").AtLeast(...)` contextual filters
- reusable `Condition("Member")` expressions with a stable typed element parameter
- `Between`, `OneOf`, `StartsWith`, `EndsWith`, `Contains`, `ContainsText`, `MatchesPattern`,
  `IsNothing`, `IsNotNothing`, `IsNullOrEmpty`, `IsTrue`, and `IsFalse`
- member-name `Map`, `OrderBy`, `OrderByDescending`, `Sum`, `Average`, `Min`, `Max`, and `JoinText`
- `DistinctBy`, `MinBy`, and `MaxBy`
- dotted object paths with explicit `Nothing` guards
- `Predicate` and `WhereMethod` inference of `Func<T, Boolean>` from the source sequence
- native `sequence!Member` expression selection through the existing default member
- the canonical `Element`, `Member`, expression, and universal delegate forms remain available

- `Execute` for Action calls without dummy return variables
- typed `EventOf`, `Subscribe`, `Unsubscribe`, `Emit`, and `HandlerCount`
- structured `Try`, catch-all or error-number `Catch`, `Finally`, captured exception metadata,
  rethrow semantics, and cleanup precedence
- populated `ListOf(T, items...)`, inferred `ListFrom(first, rest...)`, and tested `ListLike`
- atomic `AddRange` from typed sequences, VBA arrays, and `Collection` values
- `ForEach(Action)`, predicate-optional `Exists`, and `JoinText`
- every universal delegate role admitted across LINQ selectors and predicates
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
closer `square(9)` form. VBA also reserves `Select`, `Any`, `In`, and `Single`; ROneCOne uses
`Map`, `Exists`, `IsIn`, and `SingleItem` while retaining `SelectItems` and `AnyItem` as explicit
primitives.

See [`docs/delegates.md`](docs/delegates.md) for the concise and canonical delegate forms, ABI and
`ByRef` safety boundaries, and complete invocation model. See
[`docs/collections.md`](docs/collections.md) for `List<T>` and LINQ semantics,
[`docs/events.md`](docs/events.md) for typed events, and
[`docs/exceptions.md`](docs/exceptions.md) for structured error flow.

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
.venv\Scripts\pyvbaanalysis.exe src tests\vba demo\vba --format text
.venv\Scripts\python.exe tools\build_test_workbook.py
powershell -ExecutionPolicy Bypass -File tools\run_excel_tests.ps1
```

The harness observes Win32 and UI Automation surfaces owned by its exact disposable Excel process,
captures selected VBE code for compiler faults, fails fast on break mode, and enforces a hard
deadline. It never enumerates or closes user-open Excel instances. See
[`docs/development.md`](docs/development.md).

The optional `%USERPROFILE%\.ronecone.env` is private. The committed
`.ronecone.env.example` defines the reserved local-only diagnostics schema; the runtime does not
read it or emit runtime logs.

## Release roadmap

Delegates, expression lambdas, runtime-generic `List<T>`, foundational LINQ, events, and structured
exceptions are available now. The release sequence continues with additional generic collections,
tasks and async/await, disposables, and native-safe parallel operations. See
[`docs/architecture.md`](docs/architecture.md).

## License

MIT. See [`LICENSE`](LICENSE).
