# Typed events

New to ROneCOne? Start with the
[Events and exceptions user guide](user-guide/events-and-exceptions.md).

ROneCOne events provide C#-style typed subscription and deterministic multicast delivery without
another class module. An event is a mutable publisher role inside the same `ROneCOne.cls` runtime.

## Frictionless form

```vba
Dim changed As ROneCOne

Set changed = ROneCOne.EventOf(vbString) _
    .Subscribe(firstHandler) _
    .Subscribe(secondHandler)

changed.Emit "ready"
changed.Unsubscribe secondHandler
```

`EventOf` accepts the same primitive and exact-class type tokens as delegate `Takes`. Every handler
must be an Action with a complete matching signature:

```vba
Set handler = ROneCOne.Action("Demo.HandleChanged").Takes(vbString)
```

Signature mismatch is rejected before the subscription list changes. Events are ByVal because
late-bound event handlers cannot preserve VBA variable identity; native `ByRef` remains available
through `NativeAction` and `RefOf` when that narrower boundary is required.

## Delivery contract

- `Subscribe` appends a handler and returns the same event for fluent construction.
- Duplicate subscriptions are retained, matching .NET event behavior.
- `Unsubscribe` removes the last matching handler and reports whether it found one.
- `HandlerCount` exposes the current subscription count.
- `Emit` validates arguments before delivery and invokes handlers in subscription order.
- Emission uses a snapshot, so subscription changes made by a handler affect the next emission.
- A handler error stops the remaining snapshot and propagates to the caller.
- Emitting an event with no handlers is a valid no-op.

The event façade owns its mutable handler list. Existing `Combine`, `Remove`, and
`GetInvocationList` remain the immutable delegate primitives documented in
[`delegates.md`](delegates.md).

The independently executable
[`ROneCOne_Events_Demo.xlsm`](../demo/ROneCOne_Events_Demo.xlsm) demonstrates subscription,
emission, removal, metadata, and the same-process dispatch benchmark.
