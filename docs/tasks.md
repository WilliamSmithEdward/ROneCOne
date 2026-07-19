# Tasks, cancellation, and progress

ROneCOne tasks are explicit state machines inside the one-file runtime. They provide .NET-shaped
coordination while preserving Excel's single-threaded COM ownership.

## Surface

- `Task.Run`, `Task.FromResult`, `Task.CompletedTask`, `Task.Delay`, and `Task.YieldOnce`
- `Await`, `Wait`, `Result`, `Status`, `Exception`, and terminal-state properties
- `WhenAll`, `WhenAny`, `WaitAsync`, and `ContinueWith`
- `CancellationTokenSource`, token registration, `Cancel`, `CancelAfter`, and
  `ThrowIfCancellationRequested`
- `ProgressOf(T, Action<T>)`
- `TaskCompletionSourceOf(T)` with `Set...` and `TrySet...` completion
- `OpenAsync`, `ExecuteReaderAsync`, `ExecuteNonQueryAsync`, `ExecuteScalarAsync`, `ReadAsync`,
  `FillAsync`, and `UpdateAsync`

## Execution contract

A task begins in `Created` and executes when observed. Delay and pending completion sources pump
Excel events while waiting. Cancellation is cooperative and checked before work and during waits.
Continuations receive the antecedent task. `WhenAll` fails or cancels if any child does and returns
ordered results only after every child completes.

`WhenAll` and `WhenAny` accept tasks directly or one sequence of tasks. A faulted `WhenAll` keeps
every child failure in an `AggregateException`; `InnerExceptions`, `Flatten`, and `Handle` expose
the structured error set. `Await` rethrows the primary VBA error so ordinary error handlers remain
natural. `Task.Exception` is nonblocking and returns `Nothing` until a task faults.

`WaitAsync` adds a composable timeout and optional cancellation token without changing the source
task. Timeouts surface as `TimeoutException`; cancellation surfaces as
`OperationCanceledException`. Scheduler deadlines use a monotonic Windows clock, and reentrant
waiting on the same task fails deterministically rather than deadlocking Excel. Cancellation
registrations are disposable and invoke every callback even when another callback fails.

No task launches Excel, creates a second application instance, transmits data, or moves workbook
objects to a worker thread. Provider tasks use typed internal dispatch to avoid COM default-member
coercion. Providers report `SupportsNativeAsync = False` and `AsyncMode = "Cooperative"`; the
task-returning surface composes consistently but does not misrepresent synchronous ADO calls as
native asynchronous I/O.

[Task user guide](user-guide/tasks-and-async.md)
