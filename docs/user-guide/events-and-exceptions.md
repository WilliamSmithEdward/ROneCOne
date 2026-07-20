# Events and exceptions

An event sends one update to every interested part of a workbook. Structured exceptions show what
to do when work fails and what cleanup must always happen. Both use the same checked `Action`
shape, so mistakes are caught before the workflow starts.

## Send one order update to several features

First create compatible handlers:

```vba
Dim updateStatus As ROneCOne
Dim writeAudit As ROneCOne

Set writeAudit = ROneCOne.Action("Orders.WriteAudit") _
    .Takes(vbString)
Set updateStatus = ROneCOne.Action("Orders.UpdateDashboard") _
    .Takes(vbString)
```

Then subscribe and emit:

```vba
Dim orderStatusChanged As ROneCOne

Set orderStatusChanged = ROneCOne.EventOf(vbString) _
    .Subscribe(writeAudit) _
    .Subscribe(updateStatus)

orderStatusChanged.Emit "Order 1042 shipped"
orderStatusChanged.Unsubscribe updateStatus
```

`Emit` sends the same text to both procedures. They run in subscription order. `Unsubscribe`
removes one procedure from future updates. A handler expecting the wrong kind of value is rejected
before the subscriber list changes.

## Recover from a bad import row and always close the file

Create the three actions that define the flow:

```vba
Dim closeFile As ROneCOne
Dim importSales As ROneCOne
Dim skipBadRow As ROneCOne

Set importSales = ROneCOne.Action("SalesImport.ImportSales")
Set skipBadRow = ROneCOne.Action("SalesImport.SkipBadRow") _
    .Takes(ROneCOne.Exception)
Set closeFile = ROneCOne.Action("SalesImport.CloseFile")
```

Build and execute the operation:

```vba
Dim importAttempt As ROneCOne

Set importAttempt = ROneCOne.Try(importSales) _
    .Catch(INVALID_AMOUNT_ERROR, skipBadRow) _
    .Finally(closeFile)

importAttempt.Execute
```

Read this as: try the import; if the amount-error occurs, skip that row; then close the file no
matter what happened. `Finally` runs after success, a handled error, or an unhandled error.

## Read the captured error

The catch procedure may accept one captured exception:

```vba
Public Sub SkipBadRow(ByVal errorInfo As Variant)
    Debug.Print errorInfo.ErrorNumber
    Debug.Print errorInfo.ErrorSource
    Debug.Print errorInfo.Message
End Sub
```

The captured value contains the local VBA `Err` state. ROneCOne does not transmit or implicitly log
that information.

### If Excel pauses before Catch

> [!NOTE]
> The VBA editor has an **Error Trapping** preference. If it is set to **Break on All Errors**,
> Excel pauses on every error before any handler - including `Catch` - can respond. In the VBA
> editor, open **Tools > Options > General**, then choose **Break on Unhandled Errors** to test
> normal handled-error behavior. The packaged demo uses a valid sample import, so running it does
> not intentionally trigger this editor setting.

## Catch one known error

The two-argument `Catch` form matches an exact VBA error number:

```vba
Set importAttempt = ROneCOne.Try(importSales) _
    .Catch(INVALID_AMOUNT_ERROR, skipBadRow) _
    .Catch(reportUnexpectedError) _
    .Finally(closeFile)
```

Catches are checked in construction order. The one-argument form is the catch-all.

## Delivery and failure rules worth knowing

- Event emission uses a handler snapshot. Subscription changes affect the next emission.
- An event handler error stops the remaining handlers and returns to the caller.
- An error raised by `Finally` takes precedence over an earlier pending error.
- The same structured operation can be executed repeatedly or nested inside another operation.

## Where next

- [Tasks and async](tasks-and-async.md) continues the learning path.
- [Events technical reference](../events.md) defines subscription and delivery semantics.
- [Exceptions technical reference](../exceptions.md) defines catch ordering, rethrow behavior,
  cleanup precedence, and captured metadata.
- [Guide index](README.md) returns to the full learning path.
