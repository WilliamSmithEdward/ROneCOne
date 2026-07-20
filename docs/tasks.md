# Tasks, cancellation, and progress

> [!TIP]
> New to ROneCOne? Start with the [Tasks and async user guide](user-guide/tasks-and-async.md).

A Task represents work that can finish, fail, or be canceled. ROneCOne uses the familiar .NET
shape to coordinate that work while keeping every workbook object inside the current Excel
process. The sections below define the exact technical contract.

## Surface

- `Task.Run`, `Task.RunOnExcel`, `Task.FromResult`, `Task.CompletedTask`, `Task.Delay`, and
  `Task.YieldOnce`
- `Await`, `Wait`, `Result`, `Status`, `Exception`, and terminal-state properties
- `ExecutionMode`, `WorkerThreadId`, and `CurrentThreadId`
- `WhenAll`, `WhenAny`, `WaitAsync`, and `ContinueWith`
- `CancellationTokenSource`, token registration, `Cancel`, `CancelAfter`, and
  `ThrowIfCancellationRequested`
- `ProgressOf(T, Action<T>)`
- `TaskCompletionSourceOf(T)` with `Set...` and `TrySet...` completion
- `OpenAsync`, `ExecuteReaderAsync`, `ExecuteNonQueryAsync`, `ExecuteScalarAsync`, `ReadAsync`,
  `FillAsync`, and `UpdateAsync`

## Execution modes

`Task.Run` is hot: it verifies a zero-argument expression lambda, converts it to a private numeric
bytecode, and submits it immediately to the Windows thread pool. The worker can evaluate numeric
constants, arithmetic, comparisons, and Boolean operators. It cannot call VBA, touch Excel or COM,
read objects, or invoke an arbitrary address. Unsupported work is rejected before submission;
there is no silent fallback to Excel's thread. The result type is inferred, or it can be declared
with `Returns`.

`Task.RunOnExcel` is the deliberate path for ordinary VBA procedures, workbook work, and COM. It
uses the cooperative scheduler on Excel's owning thread. Delay, continuations, provider calls, and
pending completion sources also use this scheduler. Bounded waits pump Excel events and sleep
briefly so the application remains responsive.

## Coordination and failure

Native cancellation is checked between bytecode instructions. Cooperative cancellation is checked
before work and during waits. Continuations receive the antecedent task. `WhenAll` waits for every
child, preserves result order, and allows already-started native children to progress in parallel.
`WhenAny` returns the first terminal child and does not silently cancel the remaining tasks.
Native `AndAlso` and `OrElse` preserve C#-style short-circuit behavior.

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

## Memory and process safety

Native code and data use separate memory regions. The worker kernel becomes read/execute only
before submission; bytecode and context remain read/write and non-executable. Completion waits for
the callback before closing its work handle and releasing all task-owned memory.

No task launches Excel, creates a second application instance, transmits data, or moves workbook
objects to a worker thread. `ExecutionMode` reports `NativeThreadPool` or `ExcelCooperative`, and
`WorkerThreadId` can verify native placement.

## Provider tasks

Provider tasks use typed internal dispatch to avoid COM default-member coercion. Providers report
`SupportsNativeAsync = False` and `AsyncMode = "Cooperative"`; the task-returning surface composes
consistently without describing synchronous ADO calls as native asynchronous I/O.

[Back to the documentation index](README.md)
