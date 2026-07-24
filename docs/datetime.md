# Dates, times, and durations

Exact semantics for the DateTime and TimeSpan surfaces. For the workflow-first introduction,
read [Dates and times](user-guide/dates-and-times.md).

## Model

A DateTime value mirrors `DateTimeOffset`: one instant, stored as milliseconds since
1970-01-01T00:00:00Z, viewed at one clock offset. Two values with different offsets can name the
same instant; `CompareTo` and `Subtract` always compare instants, while the component getters
(`Year`, `Month`, `Day`, `Hour`, `Minute`, `Second`, `Millisecond`, `DayOfWeek` with Sunday as
zero, `DayOfYear`) always read the clock time. Precision is one millisecond. Supported clock
years run 0100 through 9999; Windows zone conversion additionally requires years 1601 and later.

A TimeSpan value is a signed duration in milliseconds. `TotalDays` through `TotalMilliseconds`
express the whole duration in one unit; `Days`, `Hours`, `Minutes`, `Seconds`, and
`Milliseconds` are components signed like the duration, exactly as in .NET.

## Creating values

| Need | Member |
|---|---|
| Parse ISO 8601 text | `DateTime.Parse(text)`, `DateTime.TryParse(text, result)` |
| Current instant | `DateTime.UtcNow`, `DateTime.Now`, `DateTime.Today` |
| From epoch numbers | `DateTime.FromUnixTimeSeconds(n)`, `DateTime.FromUnixTimeMilliseconds(n)` |
| From a VBA `Date` | `DateTime.FromLocal(dateValue)`, `DateTime.FromUtc(dateValue)` |
| Durations | `TimeSpan.FromDays/FromHours/FromMinutes/FromSeconds/FromMilliseconds(n)`, `TimeSpan.Zero`, `TimeSpan.Parse(text)` |

`Parse` accepts ISO 8601 extended text: `yyyy-MM-dd`, an optional time introduced by `T` or a
space (`HH:mm`, `HH:mm:ss`, or `HH:mm:ss.f` with one through seven fraction digits, extra digits
beyond milliseconds discarded), and an optional offset (`Z`, `+HH:mm`, `+HHmm`, or `+HH`, at
most 14 hours). Text without an offset assumes the machine's local offset for that clock time,
like `DateTimeOffset.Parse`. Impossible dates, hours, and offsets are rejected. Failures raise
error number `ROneCOne.FormatError` from source `ROneCOne.FormatException`; `TryParse` returns
False and leaves its result `Nothing`.

`TimeSpan.Parse` reads `[-][d.]hh:mm[:ss[.fff]]` with hours 0 through 23, minutes and seconds 0
through 59, and round-trips `ToString` output.

## Reading and converting

| Need | Member |
|---|---|
| Clock components | `Year` .. `Millisecond`, `DayOfWeek`, `DayOfYear` |
| VBA dates | `UtcDateTime`, `LocalDateTime` |
| Epoch numbers | `ToUnixTimeSeconds`, `ToUnixTimeMilliseconds` |
| Re-view the instant | `ToUniversalTime`, `ToLocalTime`, `ToOffset(minutesOrTimeSpan)` |
| The offset itself | `Offset` (a TimeSpan) |
| Text | `ToIsoString`, `ToString([pattern])` |

`ToIsoString` round-trips: `yyyy-MM-ddTHH:mm:ss`, a `.fff` fraction only when nonzero, then `Z`
for a zero offset or `+hh:mm`/`-hh:mm` otherwise. `ToString` with no pattern is `ToIsoString`;
with a pattern it formats through the token subset `yyyy`, `MM`, `dd`, `HH`, `mm`, `ss`, and
`fff`. Any other alphabetic token raises `FormatError`; separators must be non-alphabetic.

## Arithmetic and comparison

| Need | Member |
|---|---|
| Calendar steps | `AddYears(n)`, `AddMonths(n)` (month ends clamp, like .NET) |
| Linear steps | `AddDays/AddHours/AddMinutes/AddSeconds(x)` (fractions welcome), `AddMilliseconds(n)` |
| Combine with durations | `Add(timeSpan)`, `Subtract(timeSpan)` |
| Difference of instants | `Subtract(otherDateTime)` returns a TimeSpan |
| Ordering | `CompareTo(other)` returns -1, 0, or 1 by instant |
| Duration algebra | `Add`, `Subtract`, `Negate`, `Duration`, `CompareTo`, `ToString` |

## Time zone behavior

Every local conversion is delegated to Windows through kernel32 with a null zone pointer, so
daylight saving applies per instant against the machine's current zone; the runtime stores no
offset table and no rules of its own. During a daylight saving gap or overlap, a local clock
time maps to whichever instant Windows chooses; the mapping is deterministic on a given machine
but is Windows's decision, not the runtime's.

[Back to the documentation index](README.md)
