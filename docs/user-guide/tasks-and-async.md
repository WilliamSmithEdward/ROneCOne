# Tasks and async

A Task represents work that will produce a result. ROneCOne coordinates that work with the
familiar .NET shape: start several calculations, wait for them together, keep Excel responsive,
add a timeout, report progress, or respond to cancellation. Every Task runs cooperatively on
Excel's own thread, inside your one Excel process.

> [!NOTE]
> Tasks coordinate work; they do not run it in parallel. ROneCOne schedules every Task on
> Excel's thread, which keeps workbook objects safe and never launches a second Excel.

## Coordinate two calculations

Suppose a planning workbook needs a revenue forecast and a reorder point. Build each calculation,
schedule both Tasks, and wait once:

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
2. `Task.Run` schedules it on the cooperative scheduler.
3. `WhenAll` represents the pair.
4. `Await` runs the scheduled work and returns both answers in the original order.

## Schedule workbook procedures the same way

Ordinary VBA functions, workbook objects, and COM calls use the identical shape. Wrap the
procedure as a delegate and schedule it:

```vba
Dim countOpenOrders As ROneCOne
Dim ordersTask As ROneCOne

Set countOpenOrders = ROneCOne.Func( _
    "Operations.CountOpenOrders").Takes().Returns(vbLong)

Set ordersTask = ROneCOne.Task.Run(countOpenOrders)
Debug.Print ordersTask.Await
```

Because everything already lives on Excel's thread, delegates, workbook reads, and Excel object
calls need no special handling.

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

Continuations receive the finished antecedent Task and can safely hand its results to ordinary
workbook code.

## Keep Excel responsive and put a limit on waiting

```vba
Dim ignored As Variant

ignored = ROneCOne.Task.Delay(250).Await
ignored = RefreshDataAsync().WaitAsync(10000).Await
ignored = ROneCOne.Task.YieldOnce.Await
```

`Delay` waits without launching another Excel. `WaitAsync(10000)` gives the refresh ten seconds
to finish and raises a clear timeout error if it does not. `YieldOnce` lets Excel process its
message queue before the workflow continues.

## Let a user cancel and show progress

```vba
Dim cancelSource As ROneCOne
Dim progress As ROneCOne

Set cancelSource = ROneCOne.CancellationTokenSource
Set progress = ROneCOne.ProgressOf( _
    vbLong, ROneCOne.Action("ImportScreen.ShowRowsProcessed").Takes(vbLong))

progress.Report 250
cancelSource.Cancel
```

Pass `cancelSource.Token` to work that supports cancellation. The scheduler checks the token at
safe points before and during waits. A typed progress reporter rejects the wrong kind of value
before it reaches the screen.

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
- [Task technical reference](../tasks.md) defines exact failure, cancellation,
  completion-source, and aggregate-error rules.
- [Guide index](README.md) returns to the full learning path.
