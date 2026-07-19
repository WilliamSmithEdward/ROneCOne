# Tasks and async

ROneCOne brings the shape of .NET task coordination into one Excel process. Work is scheduled
cooperatively: no second Excel instance is launched, and VBA objects never cross unsafe thread
boundaries.

## Start with a completed task

```vba
Dim answer As Long

answer = ROneCOne.Task.FromResult(42&).Await
```

`Await`, `Result`, and `Wait` observe the same terminal state. A fault is re-raised with its
captured number, source, description, help file, and help context.

## Coordinate several tasks

```vba
Dim results As ROneCOne

Set results = ROneCOne.Task.WhenAll( _
    ROneCOne.Task.FromResult(10&), _
    ROneCOne.Task.FromResult(20&)).Await
Debug.Print results.JoinText(",")
```

`WhenAll` preserves task order. `WhenAny` returns the first completed task. `ContinueWith`
accepts a typed `Func<Task, TResult>`.

## Delay, yield, and add a timeout

```vba
Dim ignored As Variant

ignored = ROneCOne.Task.Delay(25&).Await
ignored = ROneCOne.Task.Delay(25&).WaitAsync(1000&).Await
ignored = ROneCOne.Task.YieldOnce.Await
```

`Await` is the primary composition surface. `Wait(timeout)` remains useful when the result needed
is simply whether a task completed in time. `WaitAsync` is the task-returning form: it composes
with continuations and raises a typed timeout or cancellation error. VBA will not resolve a class
member named `Yield`, so the closest legal form is `YieldOnce`.

## Cancellation, progress, and completion sources

```vba
Dim registration As ROneCOne
Dim source As ROneCOne

Set source = ROneCOne.CancellationTokenSource
Set registration = source.Token.Register( _
    ROneCOne.Action("Module1.OnCanceled").Takes)
source.CancelAfter 1000&
registration.Dispose
```

Use `ProgressOf(T, handler)` for typed progress reports. Use `TaskCompletionSourceOf(T)` when an
event or callback owns completion; `SetResult`, `SetException`, and `SetCanceled` are paired with
non-throwing `TrySet...` forms.

`WhenAll` retains every child failure in `task.Exception.InnerExceptions`. Use `Flatten` for nested
aggregates and `Handle(predicate)` when a caller deliberately marks matching failures handled.
`Task.Exception` never starts or waits for a task.

## What async means in VBA

Task work runs when observed and yields through `DoEvents` while waiting. Delay, cancellation,
continuations, provider calls, and completion sources compose without another Excel application.
This keeps COM and workbook state on Excel's owning thread. It does not claim CPU-parallel VBA
execution, which the host cannot safely provide.

See the [task contracts](../tasks.md) or return to the [guide index](README.md).
