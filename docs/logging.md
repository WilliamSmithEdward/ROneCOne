# Logging

Exact semantics for the logger. For the workflow-first introduction, read
[Files and folders](user-guide/files-and-folders.md).

## Surface

`ROneCOne.Logger(path, [minimumLevel])` returns a logger that writes to one file.
`minimumLevel` defaults to `Information`.

| Member | Behavior |
|---|---|
| `LogTrace` / `LogDebug` / `LogInformation` | Writes at the named level when it is enabled |
| `LogWarning` / `LogError` / `LogCritical` | The higher severities, in the M.E.Logging order |
| `IsEnabled(levelName)` | True when a level would be written |
| `MinimumLevel` | The current minimum severity name |

Each log call takes a composite template and its arguments, formatted through the same grammar
as [`Strings.Format`](strings.md).

## Line format

Every written line is one physical line:

```
2026-07-24T16:30:05.123Z [INF] user Ada scored 12.3
```

A UTC timestamp to the millisecond with a `Z` suffix, a three-letter severity code in brackets
(`TRC`, `DBG`, `INF`, `WRN`, `ERR`, `CRT`), then the formatted message. The format is fixed and
greppable rather than configurable.

## Level policy and durability

The levels order `Trace < Debug < Information < Warning < Error < Critical`. A call below the
minimum level is dropped before any disk access, so leaving `LogDebug` calls in place costs
almost nothing in production. Each call that does write appends the line and closes the file
immediately, so a line is on disk the moment the call returns and survives a crash later in the
run.

An unknown level name (to the constructor or `IsEnabled`) raises `ROneCOne.FormatError`. A write
to a path whose folder does not exist raises the file layer's `ROneCOne.IOError`.

[Back to the documentation index](README.md)
