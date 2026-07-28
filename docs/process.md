# Processes

Exact semantics for the awaitable process surface. For the workflow-first introduction, read
[Processes and commands](user-guide/processes-and-commands.md).

## Surface

`ROneCOne.Process.RunAsync(command, [workingDirectory], [cancellationToken], [standardInput])`
starts a shell command immediately and returns a Task. The result is a process-result value
with `ExitCode`, `StandardOutput`, and `StandardError`. A command that cannot start raises error
number `ROneCOne.ProcessError` from source `ROneCOne.ProcessException`; a command that runs and
fails does not fault the task, exactly like `System.Diagnostics.Process`: inspect `ExitCode`.

## Mechanics

The command runs through `cmd.exe /d /s /c` with the user command parenthesized, so its own
operators and redirections behave normally, and with standard output and standard error
redirected to scratch files in the temporary folder. The Windows Script Host `Exec` object
exposes a pollable `Status`, which the Task scheduler polls cooperatively; the runtime never
touches the blocking pipe readers, so a chatty process cannot deadlock Excel. After the
process exits, the scratch files are read (decoded as windows-1252, the console code page
family; ASCII output is always exact) and deleted. Cancellation terminates the process and
removes the scratch files.

`workingDirectory` is applied inside the command line as `cd /d "..." &&`, so nothing mutates
Excel's own current directory. An empty command raises `InvalidArgumentError`.

`standardInput`, when supplied, is written to the process at start; the input stream is always
closed afterward, so a filter such as `sort` or `findstr` sees end-of-file and completes rather
than waiting forever. A command given no input still has its input closed.

## Composition

The returned Task composes like any other: `Await`, `Wait(timeout)`, `WaitAsync`, `WhenAll`
across several commands, continuations, and cancellation tokens all apply. The processes
themselves run outside Excel concurrently; VBA stays single-threaded and only polls.

## Interactive sessions

`RunAsync` runs a command to completion. When you need a conversation instead, or want to see
output while the command is still running, `ROneCOne.Process.StartSession(command,
[workingDirectory], [encodingName])` keeps one `cmd.exe` alive and hands back a session.

| Member | Behavior |
|---|---|
| `WriteAsync` / `WriteLineAsync` | Queues input; the Task completes when it reaches the pipe |
| `CloseInput` | Closes standard input so a filter sees end of file |
| `ReadLineAsync([timeoutMs])` | Awaits the next standard output line |
| `ReadErrorLineAsync([timeoutMs])` | The same for standard error, which stays separate |
| `ReadAvailable` / `ReadErrorAvailable` | Takes what is buffered right now, without waiting |
| `ReadToEndAsync` | Awaits exit, then resolves to the rest of standard output |
| `WaitForExitAsync` | Resolves to the exit code |
| `HasExited` / `ExitCode` / `ProcessId` | State, without waiting |
| `KillProcess` | Terminates the command now |

```vba
Dim session As ROneCOne

Set session = ROneCOne.Process.StartSession("sort")
session.WriteLineAsync("banana").Await
session.WriteLineAsync("apple").Await
session.CloseInput
Debug.Print session.ReadLineAsync(4000).Await   ' apple
Debug.Print session.WaitForExitAsync.Await      ' 0
```

### Reading without hanging

A prompt arrives with no trailing newline, so a line read against an idle command has nothing to
return and would wait indefinitely. Two things address that. `ReadLineAsync` takes an optional
timeout in milliseconds and resolves to empty text when it elapses, and `ReadAvailable` returns
whatever is buffered including a partial line, which is how you see a prompt at all.

One read may be pending per stream; a second raises `ROneCOne.ProcessError`. The slot is returned
when the read completes, times out, or is canceled, so a session cannot become stuck.

### Mechanics

Standard output and standard error ride anonymous pipes. Each poll asks `PeekNamedPipe` how many
bytes are waiting and reads exactly that many, so a read never blocks. Standard input rides an
overlapped named pipe, so a payload larger than the pipe buffer pends and polls instead of
freezing Excel while the command catches up.

Output decodes through the machine's OEM code page, which is what console programs emit; pass
`encodingName` to override. Bytes accumulate and the whole buffer is decoded against a character
cursor, so a chunk boundary cannot split a multibyte character.

VBA reserves `Kill` for its file-deletion statement, so the member is `KillProcess`. A session
that goes out of scope releases its handles without terminating the command; call `KillProcess`
when you mean to stop it.

See [ADR 0028](decisions/0028-interactive-process-sessions.md) for the reasoning.

[Back to the documentation index](README.md)
