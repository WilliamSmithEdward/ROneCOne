# ADR 0028: Interactive process sessions over Win32 pipes

Status: accepted, 2026-07-28

## Context

`Process.RunAsync` (ADR 0017) runs a command to completion and hands back everything it printed.
That shape cannot hold a conversation. Driving `sqlcmd`, a Python interpreter, or any other
read-evaluate-print tool means writing a line, reading the answer, and writing again against one
long-lived process. It also cannot show a long command's output as it happens, because nothing is
readable until the command exits.

`RunAsync` redirects both streams to scratch files precisely so that polling never touches a
blocking pipe reader. A session cannot do that: the whole point is to read while the command is
still running. So the question was whether VBA can read and write pipes without ever blocking
Excel, and a probe answered it in detail. `PeekNamedPipe` reports the byte count waiting, so a
following `ReadFile` sized to that count returns immediately; an echoed line came back in 16
milliseconds. Standard error arrives on its own pipe. `GetExitCodeProcess` reports 259 while
running and the real code afterwards, and a drained pipe then reports error 109 for a broken pipe,
which is the end-of-output signal. Writes were the interesting half: an anonymous pipe blocks the
caller once its buffer fills, so the probe tested an overlapped named pipe instead, and a 64KB
write into a 4KB buffer returned `ERROR_IO_PENDING` and then polled as incomplete rather than
freezing the host.

## Decision

`Process.StartSession(command, [workingDirectory], [encodingName])` starts the command under
`cmd.exe /q /k`, which keeps the interpreter alive between writes. `WriteAsync` and
`WriteLineAsync` queue input and return a Task that completes when the payload reaches the pipe.
`CloseInput` closes standard input so a filter sees end of file. `ReadLineAsync` and
`ReadErrorLineAsync` await the next line on each stream separately, `ReadAvailable` and
`ReadErrorAvailable` take whatever is buffered without waiting, `ReadToEndAsync` collects the rest
after exit, and `WaitForExitAsync` resolves to the exit code. `HasExited`, `ExitCode`,
`ProcessId`, and `KillProcess` complete the surface.

Nothing blocks Excel. Reads size themselves from `PeekNamedPipe`. Writes ride an overlapped named
pipe and poll `GetOverlappedResult`, so a payload larger than the buffer pends instead of freezing
the host; the queue drains in order and a completed counter resolves every waiting Task.

Two behaviors exist because a live run demanded them. A prompt arrives with no trailing newline,
so an unbounded line read against an idle command waits forever; `ReadLineAsync` therefore takes
an optional timeout that resolves to empty text, and `ReadAvailable` exists so a caller can see
the prompt at all. A read the caller abandoned used to leave the session permanently unreadable,
so the single read slot per stream is returned on timeout and on cancellation as well as on
completion.

Output decodes through the machine's OEM code page by default, which is what console programs
emit; the probe reported 437 on this machine and confirmed `ADODB.Stream` accepts it. Bytes
accumulate and the whole buffer is decoded against a character cursor, so a chunk boundary can
never split a multibyte character if a caller overrides the encoding.

## Decisions declined

`WScript.Shell.Exec` was rejected for sessions even though `RunAsync` uses it: its `StdOut.Read`
blocks, and its `Status` property is the only nonblocking signal, which is not enough to read a
line while a command runs. Anonymous pipes were rejected for standard input alone, because a write
into a full buffer would freeze Excel, which is exactly the promise this library makes.

## Consequences

One new role, one new task operation, sixteen new `Declare` statements, and four new private
types, all kernel32. No new prog-id and no reference. The existing `Process.RunAsync` is unchanged
and remains the right call for a command that simply runs and finishes.

VBA reserves `Kill` for its file-deletion statement, so the member is `KillProcess`. That joins
`Connect`, `Disconnect`, `YieldOnce`, `IsIn`, and `SingleItem` on the host-boundary list, and the
live compiler is the only thing that catches it; static analysis passed the illegal name.

Teardown closes handles without terminating the command, so a session that goes out of scope
leaves its process running. That is deliberate: silently killing a child on garbage collection
would be surprising, and `KillProcess` states the intent.
