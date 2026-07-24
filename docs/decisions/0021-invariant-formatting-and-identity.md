# ADR 0021: Invariant composite formatting and identity helpers

Status: accepted, 2026-07-24

## Context

VBA's text tooling fights the runtime's determinism rules. `Format$` speaks the machine locale,
so the same workbook prints `1,234.50` on one machine and `1.234,50` on another; string
concatenation in a loop is quadratic; and VBA offers no GUID or cryptographic random source at
all. C# code leans on `String.Format`, `StringBuilder`, `Guid.NewGuid`, and
`RandomNumberGenerator` for all four, and each has a clean one-file answer: pure VBA arithmetic
for formatting, the JSON writer's existing doubling buffer for building, and two more `Declare`s
(`CoCreateGuid` in ole32, `BCryptGenRandom` in the already-declared bcrypt.dll) for identity and
randomness. A probe verified both functions: version 4 GUIDs with the correct variant bits, and
status-zero random fills that never repeat.

## Decision

`ROneCOne.Strings.Format` implements the `String.Format` grammar: `{index}`,
`{index,alignment}`, `{index:format}`, and `{{`/`}}` escapes. Numeric specifiers `G`, `N`, `F`,
`D`, `X`/`x`, and `P` are computed in pure VBA and always emit invariant text (a period decimal
separator, comma grouping) on every machine; fixed-point work is capped at 15 significant
digits. Date arguments format through the same `yyyy MM dd HH mm ss fff` token subset the
DateTime surface uses and default to the `yyyy-MM-ddTHH:mm:ss` stamp the JSON and CSV writers
already emit. `DateTime` and `TimeSpan` values format through their own `ToString`; plain
strings pass through untouched, exactly as `String.Format` treats non-`IFormattable` arguments.
Unknown specifiers, unmatched item indexes, and unbalanced braces raise the typed
`ROneCOne.FormatError`.

`ROneCOne.StringBuilder()` exposes the runtime's doubling text buffer as a value with `Append`,
`AppendLine`, `AppendFormat`, `Length`, `Clear`, and `ToString`. `ROneCOne.Guid.NewGuid` formats
`CoCreateGuid` output in canonical lowercase 8-4-4-4-12 form (the first three groups read
little-endian, exactly like `Guid.ToString`), with `EmptyGuid` as the all-zero constant.
`ROneCOne.RandomNumberGenerator` mirrors the .NET class of the same name: `GetBytes(count)`
fills from the system-preferred CNG provider and `GetInt32(fromInclusive, toExclusive)` stays
uniform through rejection sampling. The seeded `System.Random` API is deliberately not mirrored;
the crypto source is the one that has no VBA answer.

## Consequences

A member named `Format` shadows VBA's `Format$` class-wide, so the runtime's two remaining
intrinsic calls are qualified with `VBA.` and the source contract forbids unqualified
`Format$(` forever. VBA cannot forward a `ParamArray`, so `Format` and `AppendFormat` copy
their arguments into a plain array before sharing the composite core; the live suite caught
this as a compile error the static analyzer missed. The factory is named
`RandomNumberGenerator` rather than `Random` both for fidelity to the mirrored .NET class and
because `Random` is a VBA `Open`-statement keyword that trips source tooling. The live suite
asserts grouped, fixed, padded, hex, percent, aligned, and escaped formatting, date and
DateTime arguments, builder composition and reuse, the GUID shape by regular expression, and
random-range uniformity bounds.
