# Tasks, cancellation, and progress

ROneCOne tasks are explicit state machines inside the one-file runtime. They provide .NET-shaped
coordination while preserving Excel's single-threaded COM ownership.

## Surface

- `Task.Run`, `Task.FromResult`, `Task.CompletedTask`, and `Task.Delay`
- `Await`, `Wait`, `Result`, `Status`, `Exception`, and terminal-state properties
- `WhenAll`, `WhenAny`, and `ContinueWith`
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

No task launches Excel, creates a second application instance, transmits data, or moves workbook
objects to a worker thread. Provider tasks use typed internal dispatch to avoid COM default-member
coercion.

[Task user guide](user-guide/tasks-and-async.md)
