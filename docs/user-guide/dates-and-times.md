# Dates and times

Read the timestamps your APIs actually send. `ROneCOne.DateTime` and `ROneCOne.TimeSpan` mirror
`DateTimeOffset` and `TimeSpan`, so ISO 8601 text, Unix epoch numbers, universal time, and
durations all become one-line conversions instead of string surgery.

## Parse what the API sent

```vba
Dim posted As ROneCOne

Set posted = ROneCOne.DateTime.Parse("2026-07-24T18:30:05.123+02:00")
Debug.Print posted.Year, posted.Hour          ' 2026  18
Debug.Print posted.ToUniversalTime.ToIsoString ' 2026-07-24T16:30:05.123Z
Debug.Print posted.LocalDateTime               ' the same instant on your clock
```

Offsets are part of the value, so two views of one instant compare equal:

```vba
Debug.Print posted.CompareTo( _
    ROneCOne.DateTime.Parse("2026-07-24T16:30:05.123Z"))   ' 0
```

Text without an offset assumes your machine's local clock, and `TryParse` reports failure
instead of raising:

```vba
Dim moment As ROneCOne

If Not ROneCOne.DateTime.TryParse(cellText, moment) Then Exit Sub
```

## Unix epoch numbers

Half the JSON in the wild carries `1784910605` instead of readable text:

```vba
Debug.Print ROneCOne.DateTime.FromUnixTimeSeconds(1784910605).ToIsoString
Debug.Print ROneCOne.DateTime.UtcNow.ToUnixTimeMilliseconds
```

## Arithmetic that respects the calendar

```vba
Dim due As ROneCOne

Set due = ROneCOne.DateTime.Parse("2026-01-31T12:00:00Z")
Debug.Print due.AddMonths(1).ToIsoString   ' 2026-02-28T12:00:00Z, clamped like .NET
Debug.Print due.AddDays(0.5).ToIsoString   ' fractions welcome
```

Subtracting two instants yields a duration; durations have totals and components:

```vba
Dim wait As ROneCOne

Set wait = due.Subtract(ROneCOne.DateTime.UtcNow)
Debug.Print wait.TotalHours
Debug.Print ROneCOne.TimeSpan.FromMinutes(90).ToString   ' 01:30:00
```

## Bridge to plain VBA dates

`FromLocal` and `FromUtc` wrap a VBA `Date`; `LocalDateTime` and `UtcDateTime` hand one back,
so worksheet cells and the typed world convert freely. Windows performs every local
conversion, with daylight saving applied per instant; the runtime never guesses an offset.

## Format for people

`ToIsoString` round-trips; `ToString` takes the token subset `yyyy MM dd HH mm ss fff`:

```vba
Debug.Print posted.ToString("yyyy-MM-dd HH:mm")   ' 2026-07-24 18:30
```

Bad text raises the typed `ROneCOne.FormatError`, catchable like every other runtime error.

## Where next

- [Dates, times, and durations reference](../datetime.md) defines parsing, precision, ranges,
  and zone behavior exactly.
- [JSON and typed objects](json-and-objects.md) covers the payloads these timestamps arrive in.
- [Guide index](README.md) returns to the full learning path.
