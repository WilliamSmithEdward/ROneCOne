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

## State and isolation

Expression nodes and delegates are immutable after construction. The predeclared instance is a
stateless factory for the first slice. Later schedulers and logs must be owned by explicit runtime
contexts keyed to one workbook; no process-global mutable state may couple workbooks.

## Feature order

1. Delegates and expression-tree lambdas.
2. Structured `Try/Catch/Finally` over callable blocks.
3. Runtime-generic collections and query operators.
4. Tasks, async/await, cancellation, progress, and captured exceptions.
5. Events, disposables, parallel native-safe operations, and additional C#-inspired types.

Each feature must pass its full behavioral, live-host, and performance gates before the next one
begins.
