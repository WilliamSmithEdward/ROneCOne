# Tasks, cancellation, and progress

> [!TIP]
> New to ROneCOne? Start with the [Tasks and async user guide](user-guide/tasks-and-async.md).

A Task represents work that can finish, fail, or be canceled. ROneCOne uses the familiar .NET
shape to coordinate that work while keeping every workbook object inside the current Excel
process. The sections below define the exact technical contract.

## Surface

- `Task.Run`, `Task.FromResult`, `Task.CompletedTask`, `Task.Delay`, and
  `Task.YieldOnce`
- `Await`, `Wait`, `Result`, `Status`, `Exception`, and terminal-state properties
- `WhenAll`, `WhenAny`, `WaitAsync`, and `ContinueWith`
- `CancellationTokenSource`, token registration, `Cancel`, `CancelAfter`, and
  `ThrowIfCancellationRequested`
- `ProgressOf(T, Action<T>)`
- `TaskCompletionSourceOf(T)` with `Set...` and `TrySet...` completion
- `OpenAsync`, `ExecuteReaderAsync`, `ExecuteNonQueryAsync`, `ExecuteScalarAsync`, `ReadAsync`,
  `FillAsync`, and `UpdateAsync`

## Execution model

Every Task executes on the cooperative scheduler on Excel's owning thread. `Task.Run` schedules
any zero-argument delegate: expression lambdas, ordinary VBA procedures, workbook work, and COM
calls. Delay, continuations, provider calls, and pending completion sources use the same
scheduler. Bounded waits pump Excel events and sleep briefly so the application remains
responsive. There is no other execution mode; a Task never moves work to another thread and
never launches a second Excel application. The name follows C# because only one execution model
exists; the single-thread boundary is documented here rather than encoded in a bespoke name
(see [ADR 0002](decisions/0002-task-run-names-the-cooperative-scheduler.md)).

## Coordination and failure

Cancellation is checked before work and during waits. Continuations receive the antecedent task.
`WhenAll` waits for every child and preserves result order. `WhenAny` returns the first terminal
child and does not silently cancel the remaining tasks.

`WhenAll` and `WhenAny` accept tasks directly or one sequence of tasks. A faulted `WhenAll` keeps
every child failure in an `AggregateException`; `InnerExceptions`, `Flatten`, and `Handle` expose
the structured error set. `Await` rethrows the primary VBA error so ordinary error handlers remain
natural. `Task.Exception` is nonblocking and returns `Nothing` until a task faults.

## Timeouts, deadlines, and registrations

`WaitAsync` adds a composable timeout and optional cancellation token without changing the source
task. Timeouts surface as `TimeoutException`; cancellation surfaces as
`OperationCanceledException`. Scheduler deadlines use a monotonic Windows clock, and reentrant
waiting on the same task fails deterministically rather than deadlocking Excel. Cancellation
registrations are disposable and invoke every callback even when another callback fails.

## Process safety

No task launches Excel, creates a second application instance, transmits data, or moves workbook
objects to another thread.

## Provider tasks

Provider tasks use typed internal dispatch to avoid COM default-member coercion. `OpenAsync`,
`ExecuteReaderAsync`, `ExecuteScalarAsync`, and `FillAsync` start natively inside ADO with the
async option and poll provider state from the scheduler, mirroring the HTTP transport pattern:
the work overlaps inside the provider, never on a second VBA thread. `ExecuteNonQueryAsync`,
`UpdateAsync`, and `ReadAsync` complete inside one cooperative step and are documented as such.

[Back to the documentation index](README.md)
