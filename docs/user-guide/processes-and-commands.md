# Processes and commands

Run a command line from VBA and await it like any other task, without `Shell`'s
fire-and-forget blindness or hand-rolled pipe plumbing. The process runs outside Excel while
your workbook stays responsive, and the result carries the exit code and both output streams.

## Run one command

```vba
Dim outcome As ROneCOne

Set outcome = ROneCOne.Process.RunAsync("git --version").Await
If outcome.ExitCode = 0 Then
    MsgBox outcome.StandardOutput
Else
    MsgBox "Failed: " & outcome.StandardError
End If
```

A failing command does not raise; it reports through `ExitCode` and `StandardError`, the way
`System.Diagnostics.Process` behaves. Only a command that cannot start at all raises the
typed `ROneCOne.ProcessError`.

Pass a working directory as the second argument when the command should run somewhere
specific:

```vba
Set outcome = ROneCOne.Process.RunAsync("git status --short", "C:\repos\project").Await
```

## Overlap several commands

Each `RunAsync` starts its process immediately, so several commands genuinely run at the same
time outside Excel while one `WhenAll` collects the results:

```vba
Dim outcomes As ROneCOne

Set outcomes = ROneCOne.Task.WhenAll( _
    ROneCOne.Process.RunAsync("ipconfig"), _
    ROneCOne.Process.RunAsync("systeminfo")).Await
Debug.Print outcomes.Item(0).StandardOutput
```

Cancellation tokens work too: cancel the token and the runtime terminates the process and
cleans up after it.

## Hold a conversation

`RunAsync` waits for a command to finish. When you want to keep talking to one process, or watch
its output arrive, start a session instead:

```vba
Dim session As ROneCOne

Set session = ROneCOne.Process.StartSession("sort")
session.WriteLineAsync("banana").Await
session.WriteLineAsync("apple").Await
session.CloseInput                              ' sort reads until end of file
Debug.Print session.ReadLineAsync(4000).Await   ' apple
Debug.Print session.WaitForExitAsync.Await      ' 0
```

Standard error stays on its own stream through `ReadErrorLineAsync`. Give a read a timeout in
milliseconds so it resolves to empty text rather than waiting forever, because a prompt arrives
with no trailing newline; `ReadAvailable` returns whatever is buffered, prompt included.
`KillProcess` stops the command (VBA reserves `Kill` for deleting files).

## Where next

- [Process technical reference](../process.md) defines the transport, decoding, sessions, and
  failure contract exactly.
- [Tasks and async](tasks-and-async.md) covers the Task surface these results ride on.
- [Guide index](README.md) returns to the full learning path.
