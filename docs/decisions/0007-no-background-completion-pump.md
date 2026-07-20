# ADR 0007: No background task-completion pump

Status: accepted, 2026-07-20

## Context

A non-blocking `task.OnCompleted(handler)` was proposed so cooperative work could finish and
notify a handler without the caller blocking on `Await`. Delivering that needs the runtime to
re-check and advance a task instance repeatedly in the background, driven by a timer such as
`Application.OnTime` or the Win32 `SetTimer` callback.

## Decision

Do not add a background completion pump. Both timer mechanisms can only call back into a
procedure in a standard module (`Application.OnTime` targets a named module procedure; `AddressOf`
resolves only standard-module functions), never a method on a class instance. A pump would
therefore require a companion `.bas` module to host the callback shim and a registry mapping
timers to task instances. The shipped runtime is a single predeclared class by hard invariant, so
it cannot own that shim. The one-file contract outranks the feature.

## Consequences

Tasks remain cooperatively scheduled. Excel already stays responsive during a bounded `Await` or
`WaitAsync` because those pump the message queue while waiting, which covers the common
"keep the UI alive during a long operation" need. What is not offered is fire-and-forget
background completion with a later callback and no `Await` at all. If a future design admits a
companion module or an external host loop, revisit; nothing in the current runtime is blocked by
this decision.
