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

Each sequence lazily owns one stable typed element parameter. `Condition(memberPath)` builds an
expression from that parameter, while `Where(memberName)` returns a short-lived contextual builder
whose terminal comparison creates the ordinary immutable `Where` query node. Member-name selectors
are normalized into the same unary lambda representation as canonical expressions, so concise and
explicit forms share evaluation, validation, and performance behavior.

Predicate nodes retain normalized member names directly instead of storing a second captured-value
node. Repeated evaluation therefore follows a cached member-access plan and performs only the
unavoidable host property dispatch. Membership and nested-quantifier nodes retain their typed
sequence or predicate directly, so they compose without source generation or global state.

Contiguous ordering nodes form an ordered-query chain. A primary `Order` or `OrderBy` establishes
the chain; `ThenBy` adds levels only while that chain remains the immediate receiver. Materializing
the chain evaluates each selector once into a key matrix, then applies a stable bottom-up merge
sort. This keeps composite ordering O(n log n), preserves equal-key source order, and avoids
repeated property or delegate calls inside the comparison loop.

Default-equality hash collections use open addressing with power-of-two slot arrays and direct
key/value references. Capacity reservation sizes the table for its load factor before mutation.
Keyed reads can bypass `Collection.Item`, whose numeric access is linear in VBA, while canonical
collections remain the ordered enumeration source. Deletion and replacement rebuild the compact
index so canonical positions and direct slots cannot diverge. Custom equality delegates use the
linear path because the runtime cannot infer a hash function that honors an arbitrary comparer.
Sorted maps and sets use binary search over their maintained order.

The sequence default member distinguishes numeric indices from String member names. Numeric values
retain zero-based indexing; names return `Condition(name)`. This makes VBA's native `sequence!Age`
syntax an expression selector without code generation, parsing a new language, or accessing VBIDE.
`Predicate` uses the sequence's existing `T` descriptor to complete dynamic procedure and object
method delegates as `Func<T, Boolean>`.

See [`collections.md`](collections.md) for the public surface and examples.

## Task, data, and provider slice

Tasks are explicit cooperative state machines driven on Excel's owning thread. Deadlines use
`GetTickCount64`; bounded waits pump events, sleep briefly, and reject same-task reentrancy.
Combinators retain child tasks, cancellation registrations own removable callback entries, and
fault sets are represented as AggregateException values without losing VBA error identity.

Data tables maintain primary-key hash indexes. Relations cache parent and child lookup indexes and
invalidate them from table version changes. Provider objects remain late-bound ADO wrappers;
task-returning calls report cooperative capability honestly, and deterministic scopes centralize
cleanup and dual-failure aggregation.

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
| Available | Contextual member LINQ, bang expressions, and key operators |
| Available | Predicate algebra, membership, null-safe paths, nested quantifiers, and comparers |
| Available | Stable composite `Order`/`OrderBy`/`ThenBy` queries with per-level comparers |
| Available | Typed events over universal Actions |
| Available | Structured `Try/Catch/Finally` over callable blocks |
| Available | Tasks, await-style coordination, cancellation, progress, and completion sources |
| Available | Mutable, concurrent-style, immutable, and specialized generic collections |
| Available | DataTable, DataSet, DataView, relations, readers, adapters, and providers |
| Constrained by host | CPU-parallel VBA execution; COM and workbook state remain on Excel's thread |

Each capability must pass its full behavioral, live-host, and performance gates before it enters
the supported surface.
