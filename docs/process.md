# Processes

Exact semantics for the awaitable process surface. For the workflow-first introduction, read
[Processes and commands](user-guide/processes-and-commands.md).

## Surface

`ROneCOne.Process.RunAsync(command, [workingDirectory], [cancellationToken])` starts a shell
command immediately and returns a Task. The result is a process-result value with `ExitCode`,
`StandardOutput`, and `StandardError`. A command that cannot start raises error number
`ROneCOne.ProcessError` from source `ROneCOne.ProcessException`; a command that runs and
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

## Composition

The returned Task composes like any other: `Await`, `Wait(timeout)`, `WaitAsync`, `WhenAll`
across several commands, continuations, and cancellation tokens all apply. The processes
themselves run outside Excel concurrently; VBA stays single-threaded and only polls.

[Back to the documentation index](README.md)
