# ADR 0017: Awaitable processes over polled WScript Exec

Status: accepted, 2026-07-23

## Context

VBA's native `Shell` returns a bare process id: no exit code, no output, no completion
signal. The Windows Script Host `Exec` object exposes `Status`, `ExitCode`, and the standard
streams, but its stream readers block the calling thread, which is exactly the classic
deadlock: a process that fills its output pipe waits for a reader while VBA waits for the
process. The runtime already has the shape for this problem: start the operation, hold a
pollable handle, and let the Task scheduler poll (ADR 0011 for WinHTTP, ADR 0014 for ADO).

A live probe established the mechanics: `cmd.exe /d /s /c` with the user command
parenthesized and both streams redirected to scratch files preserves inner quoting (paths
with spaces) and exit codes (`exit 7` arrives as 7); `Status` polls cleanly from 0 to 1;
`Terminate` ends a running process; and the redirect files can remain locked for a beat after
`Status` turns 1, so the first read attempt after completion may fail before succeeding.

## Decision

`ROneCOne.Process.RunAsync(command, [workingDirectory], [cancellationToken])` starts the
command immediately through `WScript.Shell.Exec` with both streams redirected to scratch
files, and returns a hot task of a new `TASK_PROCESS_RUN` kind. The advance step polls
`Status`; on completion it reads both scratch files (retrying on the transient lock by simply
polling again), captures `ExitCode`, deletes the scratch files, and completes with a
process-result value exposing `ExitCode`, `StandardOutput`, and `StandardError`. The blocking
`StdOut`/`StdErr` readers are never touched, so no output volume can deadlock Excel.

Failure semantics copy `System.Diagnostics.Process`: a command that runs and fails reports
through its exit code and standard error without faulting the task; only a start failure
raises, with the typed `ROneCOne.ProcessError` from source `ROneCOne.ProcessException`.
Cancellation terminates the process and removes the scratch files. A working directory is
injected as `cd /d "..." &&` inside the command line rather than by mutating the process-wide
current directory. Captures decode as windows-1252 because console redirection writes in the
OEM/ANSI code page family; ASCII output, the overwhelmingly common case, is always exact, and
the limit is documented rather than hidden.

## Consequences

`WScript.Shell` joins the prog-id whitelist as a const-based `CreateObject`. The live suite
adds a process contract: echo capture, exit-code passthrough, standard-error routing, a
missing command's failure shape, working-directory placement, two commands overlapped under
`WhenAll`, cancellation mid-ping against the loopback address, and the argument and role
guards. Everything runs offline against `cmd.exe` built-ins plus loopback `ping`.
