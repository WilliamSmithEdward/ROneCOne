# ADR 0002: Task.Run names the cooperative scheduler

Status: accepted, 2026-07-19; supersedes the naming portion of
[ADR 0001](0001-remove-native-task-run.md)

## Context

ADR 0001 removed the native execution slice and kept `RunOnExcel` as the single scheduling
call, reasoning that the .NET name `Task.Run` promises thread-pool execution. The owner
reopened the naming question. The honest-naming argument dated from the era of two execution
paths, when the name had to distinguish where work ran. With exactly one execution model, the
distinction is gone: `RunOnExcel` now implies an alternative that does not exist, and the
project's alignment rule prefers the C# name unless the host imposes a demonstrated
constraint. The constraint was the native path; it no longer exists. Single-threaded runtimes
using the standard async vocabulary are a well-established pattern.

## Decision

`ROneCOne.Task.Run(work, optional token)` is the single public scheduling call, dispatched
through the Task factory's default member to a Friend implementation. `RunOnExcel` is removed
outright rather than aliased. The single-thread boundary is stated prominently in the task
documentation: Tasks coordinate work, they do not run it in parallel, and everything executes
on Excel's thread.

## Alternatives considered

Keeping `RunOnExcel` was rejected because the name's information content became negative once
only one mode existed. Keeping both names was rejected under the no-inferior-aliases rule.

## Consequences

C# muscle memory works: `Task.Run` is the form every .NET example uses, and the demo,
tests, and guides read closer to their C# equivalents. A developer who assumes background
execution learns the single-thread model from the same documentation note that governed
`RunOnExcel`; the rename does not change semantics, blocking behavior, or safety. The
cooperative lifecycle benchmark now carries the `Task.Run` name.
