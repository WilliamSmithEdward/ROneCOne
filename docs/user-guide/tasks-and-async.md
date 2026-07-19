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
Dim tasks As ROneCOne

Set tasks = ROneCOne.ListOf( _
    ROneCOne.Task, _
    ROneCOne.Task.FromResult(10&), _
    ROneCOne.Task.FromResult(20&))

Set results = ROneCOne.Task.WhenAll(tasks).Await
Debug.Print results.JoinText(",")
```

`WhenAll` preserves task order. `WhenAny` returns the first completed task. `ContinueWith`
accepts a typed `Func<Task, TResult>`.

## Cancellation, progress, and completion sources

```vba
Dim source As ROneCOne

Set source = ROneCOne.CancellationTokenSource
source.Token.Register ROneCOne.Action("Module1.OnCanceled").Takes
source.CancelAfter 1000&
```

Use `ProgressOf(T, handler)` for typed progress reports. Use `TaskCompletionSourceOf(T)` when an
event or callback owns completion; `SetResult`, `SetException`, and `SetCanceled` are paired with
non-throwing `TrySet...` forms.

## What async means in VBA

Task work runs when observed and yields through `DoEvents` while waiting. Delay, cancellation,
continuations, provider calls, and completion sources compose without another Excel application.
This keeps COM and workbook state on Excel's owning thread. It does not claim CPU-parallel VBA
execution, which the host cannot safely provide.

See the [task contracts](../tasks.md) or return to the [guide index](README.md).
