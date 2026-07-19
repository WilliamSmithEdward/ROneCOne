# Events and exceptions

Events coordinate work that should happen in response to a change. Structured exceptions make
success, recovery, and cleanup visible in one flow. Both use the same typed Action surface as the
rest of ROneCOne.

## Publish a typed event

First create compatible handlers:

```vba
Dim updateStatus As ROneCOne
Dim writeAudit As ROneCOne

Set writeAudit = ROneCOne.Action("Audit.WriteEntry") _
    .Takes(vbString)
Set updateStatus = ROneCOne.Action("Status.Show") _
    .Takes(vbString)
```

Then subscribe and emit:

```vba
Dim changed As ROneCOne

Set changed = ROneCOne.EventOf(vbString) _
    .Subscribe(writeAudit) _
    .Subscribe(updateStatus)

changed.Emit "ready"
changed.Unsubscribe updateStatus
```

Handlers run in subscription order. A signature mismatch is rejected before the event's handler
list changes.

## Handle a failure and always clean up

Create the three actions that define the flow:

```vba
Dim cleanup As ROneCOne
Dim handleError As ROneCOne
Dim work As ROneCOne

Set work = ROneCOne.Action("ImportJobs.Run")
Set handleError = ROneCOne.Action("ImportJobs.HandleError") _
    .Takes(ROneCOne.Exception)
Set cleanup = ROneCOne.Action("ImportJobs.Cleanup")
```

Build and execute the operation:

```vba
Dim attempt As ROneCOne

Set attempt = ROneCOne.Try(work) _
    .Catch(handleError) _
    .Finally(cleanup)

attempt.Execute
```

`Finally` runs after success, a handled error, an unhandled error, or a failed catch.

## Read the captured error

The catch procedure may accept one captured exception:

```vba
Public Sub HandleError(ByVal errorInfo As Variant)
    Debug.Print errorInfo.ErrorNumber
    Debug.Print errorInfo.ErrorSource
    Debug.Print errorInfo.Message
End Sub
```

The captured value contains the local VBA `Err` state. ROneCOne does not transmit or implicitly log
that information.

## Catch one known error

The two-argument `Catch` form matches an exact VBA error number:

```vba
Set attempt = ROneCOne.Try(work) _
    .Catch(vbObjectError + 100, expectedHandler) _
    .Catch(fallbackHandler) _
    .Finally(cleanup)
```

Catches are checked in construction order. The one-argument form is the catch-all.

## Delivery and failure rules worth knowing

- Event emission uses a handler snapshot. Subscription changes affect the next emission.
- An event handler error stops the remaining handlers and returns to the caller.
- An error raised by `Finally` takes precedence over an earlier pending error.
- The same structured operation can be executed repeatedly or nested inside another operation.

## Where to go deeper

- [Events technical reference](../events.md) defines subscription and delivery semantics.
- [Exceptions technical reference](../exceptions.md) defines catch ordering, rethrow behavior,
  cleanup precedence, and captured metadata.
- [Guide index](README.md) returns to the full learning path.
