# Tasks and async

A Task represents work that will produce a result. ROneCOne can run safe calculations on Windows
worker threads, wait for several answers together, keep Excel responsive, add a timeout, report
progress, or respond to cancellation. Everything stays local and inside the current Excel process.

## Run two calculations at the same time

Suppose a planning workbook needs a revenue forecast and a reorder point. Build each calculation,
start both Tasks, and wait once:

```vba
Dim forecastTask As ROneCOne
Dim forecastWork As ROneCOne
Dim reorderTask As ROneCOne
Dim reorderWork As ROneCOne
Dim results As ROneCOne

Set forecastWork = ROneCOne.Value(125000#).Multiply(1.08).AsFunc
Set reorderWork = ROneCOne.Value(80#).Multiply(1.65).Add(20#).AsFunc

Set forecastTask = ROneCOne.Task.Run(forecastWork)
Set reorderTask = ROneCOne.Task.Run(reorderWork)
Set results = ROneCOne.Task.WhenAll(forecastTask, reorderTask).Await

Debug.Print results.Item(0)   ' 135000
Debug.Print results.Item(1)   ' 152
```

In plain language:

1. `AsFunc` turns each formula into a self-contained expression lambda.
2. `Task.Run` checks that the expression is safe, then starts it on the Windows thread pool.
3. `WhenAll` represents the pair.
4. `Await` waits for both answers and returns them in the original order.

The usual numeric and Boolean result types are inferred. Add `.Returns(vbLong)` or another
supported scalar type when you want an exact result type. The current native scalar set is Byte,
Integer, Long, Single, Double, Boolean, and Date.

## Keep workbook work on Excel's thread

Excel objects and ordinary VBA procedures belong to Excel's owning thread. Wrap that work as a
delegate, then use the explicit safe path:

```vba
Dim countOpenOrders As ROneCOne
Dim ordersTask As ROneCOne

Set countOpenOrders = ROneCOne.Func( _
    "Operations.CountOpenOrders").Takes().Returns(vbLong)

Set ordersTask = ROneCOne.Task.RunOnExcel(countOpenOrders)
Debug.Print ordersTask.Await
```

`RunOnExcel` uses cooperative scheduling. It keeps composition and await-style flow, but it does
not claim that a VBA function or workbook call is running in parallel.

> [!NOTE]
> Passing an Excel-owned delegate to `Task.Run` raises a clear error instead of quietly changing
> the execution mode. Work never falls back to a different thread silently.

## Know what may run on a worker

The first native expression surface accepts numeric or Boolean constants with addition,
subtraction, multiplication, division, negation, comparisons, `AndAlso`, `OrElse`, and `Not`.
It does not accept strings, objects, member access, VBA calls, Excel calls, COM, or external
addresses. This boundary is what makes worker-thread execution predictable.

You can inspect it when needed:

```vba
Debug.Print forecastTask.ExecutionMode     ' NativeThreadPool
Debug.Print forecastTask.WorkerThreadId    ' Nonzero worker thread
Debug.Print ROneCOne.CurrentThreadId        ' Excel's current thread
```

Normal workbook code rarely needs the thread identifiers; they are useful for diagnostics and
proof that work did not execute on Excel's thread.

## Turn the result into the next step

A continuation is work that starts after an earlier Task finishes. It can update a dashboard or
build a summary after `WhenAll` completes:

```vba
Dim buildSummary As ROneCOne
Dim summaryTask As ROneCOne

Set buildSummary = ROneCOne.Func( _
    "Operations.BuildSummary") _
    .Takes(ROneCOne.Task) _
    .Returns(vbString)

Set summaryTask = allWork.ContinueWith(buildSummary)
Debug.Print summaryTask.Await
```

Continuations run cooperatively on Excel's thread, so they can safely hand the finished results to
ordinary workbook code.

## Keep Excel responsive and put a limit on waiting

```vba
Dim ignored As Variant

ignored = ROneCOne.Task.Delay(250&).Await
ignored = RefreshDataAsync().WaitAsync(10000&).Await
ignored = ROneCOne.Task.YieldOnce.Await
```

`Delay` waits without launching another Excel. `WaitAsync(10000&)` gives the refresh ten seconds
to finish and raises a clear timeout error if it does not. `YieldOnce` lets Excel process its
message queue before the workflow continues.

## Let a user cancel and show progress

```vba
Dim cancelSource As ROneCOne
Dim progress As ROneCOne

Set cancelSource = ROneCOne.CancellationTokenSource
Set progress = ROneCOne.ProgressOf( _
    vbLong, ROneCOne.Action("ImportScreen.ShowRowsProcessed").Takes(vbLong))

progress.Report 250&
cancelSource.Cancel
```

Pass `cancelSource.Token` to work that supports cancellation. Native work checks between
instructions; cooperative work checks at safe scheduler points. A typed progress reporter rejects
the wrong kind of value before it reaches the screen.

## When the answer is already available

`Task.FromResult` creates a Task that is already finished:

```vba
Set customerTask = ROneCOne.Task.FromResult(cachedCustomer)
```

Use it when an API must return a Task but a cache or default value already contains the answer. It
is mainly an adapter and testing tool. If ordinary code only needs the value, use the value
directly.

## Where next

- [Data and providers](data-and-providers.md) continues the learning path.
- [Task technical reference](../tasks.md) defines exact failure, cancellation, memory,
  completion-source, and aggregate-error rules.
- [Guide index](README.md) returns to the full learning path.
