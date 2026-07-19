# ROneCOne Architecture

## Runtime invariant

The deployed runtime is one imported class module, `src/ROneCOne.cls`. Test modules, fixture
workbooks, Python tooling, documentation, and benchmark drivers are development artifacts and are
not runtime dependencies.

The class uses `VB_PredeclaredId = True` so `ROneCOne` provides a zero-setup factory surface while
explicit `New ROneCOne` instances remain available. Every runtime value returned by the library is
also a `ROneCOne` instance with an internal role tag. This permits delegates, expression nodes,
collections, tasks, cancellation tokens, errors, and later runtime concepts to coexist without
additional shipped class modules.

## Delegate and expression slice

The first slice is an immutable expression-tree interpreter:

```vba
Dim x As ROneCOne
Dim square As ROneCOne

Set x = ROneCOne.Parameter(vbLong)
Set square = ROneCOne.Lambda(x.Multiply(x), x)

Debug.Print square(CLng(9))
```

This is ordinary compilable VBA. It does not inject source, require VBIDE trust, or ask developers
to write a lambda as a string. The expression tree is evaluated only when the delegate is invoked.

Method delegates use VBA's late-bound method-name boundary because the language exposes no general
first-class procedure value:

```vba
Set transform = ROneCOne.FromMethod(target, "Transform", 1)
```

That string boundary is secondary to expression-tree lambdas and remains validated at invocation.

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

## State and isolation

Expression nodes, delegates, and query nodes are immutable after construction. Explicit list
instances own their mutable values; the predeclared instance remains a stateless factory. Later
schedulers and logs must be owned by explicit runtime contexts keyed to one workbook; no
process-global mutable state may couple workbooks.

## Feature order

1. Delegates and expression-tree lambdas.
2. Runtime-generic `List<T>` and foundational query operators.
3. Structured `Try/Catch/Finally` over callable blocks.
4. Tasks, async/await, cancellation, progress, and captured exceptions.
5. Additional generic collections, events, disposables, parallel native-safe operations, and
   other C#-inspired types.

Each feature must pass its full behavioral, live-host, and performance gates before the next one
begins.
