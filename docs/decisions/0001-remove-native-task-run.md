# ADR 0001: Remove the native Task.Run execution slice

Status: accepted, 2026-07-19

## Context

Version 1.2.0 shipped a native execution path: `Task.Run` verified a zero-argument scalar
expression lambda, compiled it to a private bytecode, and ran an embedded x64 interpreter on the
Windows thread pool with W^X memory discipline. The path was correct, fully tested, and
performance-gated, but empirical use showed that parallel execution is feasible only for narrow
pure-scalar cases and delivers no practical performance gain there. The project heuristic that
governs frontier features requires a material improvement to end-user developer experience or
performance; the native path no longer met it. The slice also carried the runtime's largest
liabilities: roughly 790 lines including embedded machine-code bytes that must byte-match a
Python builder, virtual-memory and thread-pool API declares, and a behavior profile (a macro
allocating executable memory) that antivirus heuristics are known to flag.

## Decision

Remove the native execution slice from `src/ROneCOne.cls`: the expression verifier, bytecode
compiler, kernel hex, memory and thread-pool declares, native task state, and the `TaskRun`,
`ExecutionMode`, `WorkerThreadId`, and `CurrentThreadId` members, plus
`tools/build_native_task_kernel.py` and the kernel source contract. `Task.RunOnExcel` remains
the single execution path and now accepts the same expression lambdas the native path served.
Signature-bound native delegates (`Native`, `NativeAction`, `RefOf`, `DispCallFunc` dispatch,
true `ByRef`) are a separate feature and are unchanged.

## Alternatives considered

Repointing `Task.Run` to the cooperative scheduler was rejected: the .NET name promises
thread-pool execution, and keeping it while executing on Excel's thread would misstate the
semantics. `RunOnExcel` states the truth in its name; the C#-alignment rule permits departure
where the host imposes a demonstrated constraint, documented here. Keeping the slice as an
undocumented internal was rejected under the no-dormant-capability and no-compatibility-hack
heuristics.

## Consequences

The runtime shrinks by about 790 lines and no longer contains executable machine code, virtual
memory allocation, or thread-pool submission, which simplifies audit and should reduce
antivirus friction. Parallel speedups are no longer claimed or possible; the Task surface is
coordination, cancellation, progress, timeouts, and continuations. This is a breaking API
removal for the next major version. Demos, tests, benchmarks, and docs moved in the same change;
the task benchmark now measures cooperative `Task.RunOnExcel(...).Await` lifecycles.
