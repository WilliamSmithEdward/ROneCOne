# Processes and commands

Run a command line from VBA and await it like any other task, without `Shell`'s
fire-and-forget blindness or hand-rolled pipe plumbing. The process runs outside Excel while
your workbook stays responsive, and the result carries the exit code and both output streams.

## Run one command

```vba
Dim result As ROneCOne

Set result = ROneCOne.Process.RunAsync("git --version").Await
If result.ExitCode = 0 Then
    MsgBox result.StandardOutput
Else
    MsgBox "Failed: " & result.StandardError
End If
```

A failing command does not raise; it reports through `ExitCode` and `StandardError`, the way
`System.Diagnostics.Process` behaves. Only a command that cannot start at all raises the
typed `ROneCOne.ProcessError`.

Pass a working directory as the second argument when the command should run somewhere
specific:

```vba
Set result = ROneCOne.Process.RunAsync("git status --short", "C:\repos\project").Await
```

## Overlap several commands

Each `RunAsync` starts its process immediately, so several commands genuinely run at the same
time outside Excel while one `WhenAll` collects the results:

```vba
Dim results As ROneCOne

Set results = ROneCOne.Task.WhenAll( _
    ROneCOne.Process.RunAsync("ipconfig"), _
    ROneCOne.Process.RunAsync("systeminfo")).Await
Debug.Print results.Item(0).StandardOutput
```

Cancellation tokens work too: cancel the token and the runtime terminates the process and
cleans up after it.

## Where next

- [Process technical reference](../process.md) defines the transport, decoding, and failure
  contract exactly.
- [Tasks and async](tasks-and-async.md) covers the Task surface these results ride on.
- [Guide index](README.md) returns to the full learning path.
