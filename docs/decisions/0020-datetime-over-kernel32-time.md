# ADR 0020: DateTime and TimeSpan over kernel32 time conversion

Status: accepted, 2026-07-24

## Context

Every transport the runtime ships hands timestamps to VBA that VBA cannot read. REST APIs and
SQL Server speak ISO 8601 with fractional seconds and offsets (`2026-07-24T18:30:05.123+02:00`),
JSON APIs also speak Unix epoch numbers, and `CDate` parses neither. VBA's `Date` has no concept
of universal time, no offset, and no epoch converters, so workbook code that touches the HTTP,
JSON, CSV, or provider surfaces ends up treating instants as strings.

The design question was where time zone knowledge should live. Hard-coding offsets or daylight
saving rules in VBA is wrong by construction; the rules change by machine and by year. A probe
confirmed that kernel32 answers everything: `GetSystemTime` returns the UTC instant with
milliseconds, and `SystemTimeToTzSpecificLocalTime` / `TzSpecificLocalTimeToSystemTime` with a
null zone pointer convert against the machine's current zone with daylight saving applied per
instant (the probe measured -480 minutes for a January instant and -420 for a July instant on a
Pacific machine, with exact round trips and sane pre-1970 behavior).

## Decision

`ROneCOne.DateTime` mirrors `DateTimeOffset`: an instant is milliseconds since
1970-01-01T00:00:00Z plus the clock offset it displays in. `Parse`/`TryParse` accept ISO 8601
extended text (date, optional time to millisecond precision, optional `Z` or `+hh:mm` offset);
text without an offset assumes the machine's local offset for that clock time, exactly like
`DateTimeOffset.Parse`. `UtcNow`, `Now`, and `Today` read `GetSystemTime`. Converters cover both
worlds: `FromUnixTimeSeconds`/`FromUnixTimeMilliseconds` and `ToUnixTimeSeconds`/
`ToUnixTimeMilliseconds` for epoch numbers, `FromLocal`/`FromUtc` and `UtcDateTime`/
`LocalDateTime` for plain VBA `Date` values. Calendar arithmetic (`AddYears` through
`AddMilliseconds`) keeps the offset and clamps month ends like .NET; `ToUniversalTime`,
`ToLocalTime`, and `ToOffset` re-view the same instant; `CompareTo` and `Subtract` order and
difference instants. `ToIsoString` round-trips, and `ToString(pattern)` formats with the token
subset `yyyy MM dd HH mm ss fff`.

`ROneCOne.TimeSpan` is a signed millisecond duration with `FromDays` through
`FromMilliseconds`, the `Total*` and component getters (components signed like .NET), `Add`,
`Subtract`, `Negate`, `Duration`, `CompareTo`, and a `[-][d.]hh:mm:ss[.fff]` `ToString` that
`Parse` round-trips.

Every local conversion is delegated to Windows; the runtime stores no offset table and no
daylight saving rule. Precision is one millisecond, matching what VBA dates and JSON APIs can
carry; extra ISO fraction digits are discarded. The supported clock years are 0100 through 9999
(VBA's `Date` range); Windows zone conversion additionally requires years 1601 and later.
Text-format failures raise the new typed `ROneCOne.FormatError` from source
`ROneCOne.FormatException`.

## Consequences

Three kernel32 `Declare`s and a private `SYSTEMTIME` type join the class; no prog-id is added.
Adding members named `Year`, `Month`, `Day`, and `Now` shadows the VBA intrinsics class-wide, so
the runtime's own intrinsic calls are qualified with `VBA.`; the source contract pins the
surface and forbids `GetTimeZoneInformation` so conversions can never drift away from the
per-instant Windows path. Daylight saving edge cases (nonexistent and ambiguous local times)
resolve to whatever Windows returns, deterministically, and are documented rather than
re-derived. The live suite asserts parse round trips, epoch vectors, offset re-views, calendar
clamping, machine-zone round trips, and rejection of impossible dates, hours, and offsets.
