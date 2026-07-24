# ADR 0024: File logger in the Microsoft.Extensions.Logging spirit

Status: accepted, 2026-07-24

## Context

Every serious workbook application ends up hand-rolling logging: a mix of `Debug.Print` that
vanishes when the IDE closes and ad-hoc `Open path For Append` blocks with inconsistent
timestamps and no level filtering. The pieces for a real logger already shipped: `Strings.Format`
does composite message templating with invariant output, the DateTime surface knows universal
time, and the file layer appends UTF-8 with typed failures. A logger is mostly composition of
what exists, plus a level policy.

## Decision

`ROneCOne.Logger(path, [minimumLevel])` returns a logger over one file. `LogTrace`, `LogDebug`,
`LogInformation`, `LogWarning`, `LogError`, and `LogCritical` follow the
`Microsoft.Extensions.Logging` severity order; each takes a composite template and its
arguments, formats through the same core as `Strings.Format`, and writes one line: a UTC
timestamp to the millisecond with a `Z` suffix, a three-letter severity code in brackets, and
the message. `IsEnabled(levelName)` and the `MinimumLevel` getter expose the policy; a call
below the minimum never touches the disk. Each write appends and closes the file immediately
through the file layer, so lines are durable the instant they return and survive a crash mid-run,
at the cost of one open-append-close per line (correct for diagnostic volumes, and the honest
trade for a logger whose whole value is being readable after a failure). An unknown level name
is refused with the typed `ROneCOne.FormatError`; a write to a missing folder surfaces the file
layer's typed `ROneCOne.IOError`.

## Consequences

No new prog-id, `Declare`, or error number: the logger reuses the composite formatter, the
kernel32 time already declared for DateTime, and the file layer's append path. The three-letter
codes (`TRC`, `DBG`, `INF`, `WRN`, `ERR`, `CRT`) and the fixed timestamp shape are a deliberate,
greppable line format rather than a configurable template, keeping the surface small. The live
suite adds a logger contract: level filtering that drops sub-minimum calls, invariant message
formatting, the timestamp and code shape checked by regular expression, the full verbose range,
and the invalid-level and missing-folder failures.
