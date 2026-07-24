# ADR 0025: Awaitable file watching over polled snapshots

Status: accepted, 2026-07-24

## Context

Drop-folder automation is a classic Excel pattern: wait for an export to land, then load it. The
.NET answer is `FileSystemWatcher`, which raises events from a background thread. VBA has no
safe equivalent; a true event source would need a second thread the runtime deliberately does
not have, and the whole task model is cooperative on Excel's thread. But the shape already
exists: start an operation, hold a pollable handle, let the Task scheduler poll it (ADR 0011,
0014, 0017). A watcher is that shape applied to a folder snapshot.

## Decision

`ROneCOne.FileWatcher(folder, [filter])` captures a name-size-modified snapshot of the matching
files, and `WaitForChangeAsync([cancellationToken])` returns a hot task of a new
`TASK_FILE_WATCH` kind. The advance step rescans the folder, throttled to one scan every 250
milliseconds so an await loop does not spin the disk, and resolves on the first difference to a
change value carrying `ChangeType` (`Created`, `Changed`, or `Deleted`) and `Name`. The baseline
is taken when `WaitForChangeAsync` is called, not when the task is first polled, so a change that
lands between the call and the await is still observed. The filter uses the same escape-and-`Like`
matching as `Directory.GetFiles`, so a watcher pattern behaves exactly like an enumeration
pattern. Cancellation and `WaitAsync` timeouts compose like every other task; a missing folder
is refused up front with the typed `ROneCOne.IOError`.

## Consequences

No new prog-id, `Declare`, or error number: the watcher reuses the file system object, the
`GetTickCount64` already declared for task timing, and the directory surface's pattern matching.
The contract is honestly a poll, not an event subscription: a change is seen on the next scan
while the task is awaited, not the instant it happens, and a file created and deleted between two
scans is invisible. This is documented rather than dressed up as push notification. The live
suite adds a watcher contract: created, changed, and deleted detection, the filter honored
against a non-matching sibling, a change staged before the await, timeout and cancellation
composition, and the missing-folder failure.
