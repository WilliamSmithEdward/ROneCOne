# ROneCOne Architecture

## Runtime invariant

The deployed runtime is one imported class module, `src/ROneCOne.cls`. Test modules, fixture
workbooks, Python tooling, documentation, and benchmark drivers are development artifacts and are
not runtime dependencies.

The class uses `VB_PredeclaredId = True` so `ROneCOne` provides a zero-setup factory surface while
explicit `New ROneCOne` instances remain available. Every runtime value returned by the library is
also a `ROneCOne` instance with an internal role tag. This provides one extensible representation
for delegates, expression nodes, collections, tasks, cancellation tokens, and errors without
additional shipped class modules.

## Developer experience invariant

Each capability is proven through canonical behavioral and live-host tests before syntax sugar is
added. The next pass then removes avoidable ceremony without weakening correctness,
discoverability, performance, privacy, or the one-file runtime contract.
Living demos lead with the shortest clear form; deeper documentation pairs it with the canonical
primitive it represents.

## Universal delegate and expression slice

The first slice is an immutable expression-tree interpreter:

```vba
Dim x As ROneCOne
Dim square As ROneCOne

Set x = ROneCOne.Var(vbLong)
Set square = x.Multiply(x).AsFunc

Debug.Print square(CLng(9))
```

This is ordinary compilable VBA. It does not inject source, require VBIDE trust, or ask developers
to write a lambda as a string. `Var` aliases the canonical `Parameter`; `AsFunc` traverses the
immutable expression once and captures unique parameters in left-to-right order. The explicit
equivalent remains `ROneCOne.Lambda(x.Multiply(x), x)`. The expression tree is evaluated only when
the delegate is invoked.

The `Func` and `Action` factories normalize expression trees, object methods, callable objects, and
workbook procedures. Immutable `Takes` and `Returns` descriptors enforce runtime signatures, while
`DynamicInvoke`, multicast invocation lists, removal, metadata, and composition operate on the same
delegate representation.

```vba
Set transform = ROneCOne.Func(target, "Transform") _
    .Takes(vbLong) _
    .Returns(vbString)
```

Native delegates add a narrower Windows x64 boundary. `Native` and `NativeAction` require complete
signatures before `DispCallFunc` dispatch, and typed reference wrappers permit true `ByRef`
mutation. Late-bound object and workbook-procedure paths reject `ByRef` because `CallByName` and
`Application.Run` do not preserve variable identity. See [`delegates.md`](delegates.md) for the
full contract and the irreducible name/address boundaries.

## Generic collection and query slice

The same tagged class now represents materialized typed lists and immutable deferred queries.
Primitive `T` values are represented by an exact `VbVarType`; user-defined class `T` values are
captured from a non-retained prototype's exact `TypeName`. Internal values use one wrapper format,
so scalars, object identity, and object `Nothing` flow through delegates and query operators without
conversion or another runtime class.

Sequence-returning LINQ operators link a source and one operation. Consumption recursively
materializes the current source, which preserves deferred behavior after list mutation. Lists carry
a mutation version; queries propagate that version so nested enumerators share stable backing while
unchanged and refresh after mutation.

See [`collections.md`](collections.md) for the public surface and examples.

## Event slice

An event is a mutable typed publisher with an owned handler list. `EventOf` builds its parameter
contract through the same descriptors as `Action.Takes`; `Subscribe` verifies the complete Action
signature before mutation. `Emit` validates arguments once and invokes a snapshot in deterministic
order. This preserves .NET-like duplicate subscription and last-match removal semantics without
turning event state into process-global state. See [`events.md`](events.md).

## Structured exception slice

`Try` operations are immutable builders over zero-argument Actions. Catch clauses retain an error
number filter and an Action; `Finally` retains one cleanup Action. Execution captures the local VBA
`Err` state into a tagged error value, selects the first matching catch, always runs cleanup, and
rethrows the final pending error. Catch and cleanup invocation are isolated in separate VBA frames
so secondary errors can be captured without an already-active error handler. See
[`exceptions.md`](exceptions.md).

## State and isolation

Expression nodes, delegates, query nodes, and Try builders are immutable after construction.
Explicit lists and events own their mutable values; the predeclared instance remains a stateless
factory. Scheduler and diagnostics components follow the same ownership rule: explicit runtime
contexts are keyed to one workbook, and process-global mutable state never couples workbooks.

## Release sequence

| Release status | Capability |
|---|---|
| Available | Universal delegates, multicast, native calls, and expression lambdas |
| Available | Runtime-generic `List<T>` and foundational query operators |
| Available | Inferred `Func` and LINQ syntax sugar |
| Available | Typed events over universal Actions |
| Available | Structured `Try/Catch/Finally` over callable blocks |
| Scheduled | Tasks, async/await, cancellation, progress, and captured exceptions |
| Scheduled | Additional generic collections, disposables, and native-safe parallelism |

Each capability must pass its full behavioral, live-host, and performance gates before it enters
the supported surface.
