# Formatting, building, and identity

Exact semantics for `Strings.Format`, `StringBuilder`, `Guid`, and `RandomNumberGenerator`.
For the workflow-first introduction, read
[Text, hashing, and encoding](user-guide/text-and-hashing.md).

## Composite formatting

`ROneCOne.Strings.Format(template, values...)` implements the `String.Format` grammar. A format
item is `{index}`, `{index,alignment}`, `{index:format}`, or `{index,alignment:format}`;
`{{` and `}}` write literal braces. A positive alignment right-aligns in that many characters, a
negative alignment left-aligns; short text is padded with spaces, long text is never truncated.
An index without a matching argument, a stray or unclosed brace, or an unknown format raises
error number `ROneCOne.FormatError` from source `ROneCOne.FormatException`.

Output is invariant on every machine: a period decimal separator and comma grouping, never the
machine locale.

| Specifier | Meaning | `{0:...}` of `1234.567` |
|---|---|---|
| none or `G` | Shortest invariant number | `1234.567` |
| `N` or `Nx` | Grouped, x decimals (default 2) | `1,234.57` |
| `F` or `Fx` | Fixed point, x decimals (default 2) | `1234.57` |
| `D` or `Dx` | Whole number zero-padded to x digits | `D6` of `42` is `000042` |
| `X`, `x`, `Xx` | Hexadecimal, upper or lower, zero-padded | `X` of `255` is `FF` |
| `P` or `Px` | Percent: value times 100, grouped, `%` suffix | `P1` of `0.123` is `12.3%` |

Rounding is half away from zero. Fixed-point and integer text support up to 15 significant
digits; hexadecimal accepts the `LongLong` range and prints negatives in two's complement, like
.NET. `D` and `X` require whole numbers.

Argument handling by type:

- Strings pass through untouched and ignore any format, exactly as `String.Format` treats
  arguments that do not implement `IFormattable`.
- Booleans print `True` or `False`.
- VBA dates default to the `yyyy-MM-ddTHH:mm:ss` stamp and accept the date token subset
  (`yyyy MM dd HH mm ss fff`) as a format.
- DateTime and TimeSpan values format through their own `ToString`; a DateTime accepts the
  token subset, a TimeSpan accepts no format.
- `Null`, `Empty`, and `Nothing` print as empty text. Other objects are refused with
  `TypeMismatchError`.

## StringBuilder

`ROneCOne.StringBuilder()` creates a growable text buffer (the same doubling buffer the JSON
writer uses, so appends stay linear). `Append(value)` accepts anything the formatter's general
form accepts and returns the builder for chaining; `AppendLine([value])` adds a CRLF;
`AppendFormat(template, values...)` appends a composite-formatted piece; `Length` reads the
text length; `Clear` empties without releasing capacity; `ToString` returns the built text.

## Guid

`ROneCOne.Guid.NewGuid` returns a version 4 GUID from `CoCreateGuid` in canonical lowercase
8-4-4-4-12 text, byte-ordered exactly like `Guid.ToString`. `EmptyGuid` is
`00000000-0000-0000-0000-000000000000`.

## RandomNumberGenerator

`ROneCOne.RandomNumberGenerator` mirrors `System.Security.Cryptography.RandomNumberGenerator`,
drawing from Windows CNG's system-preferred provider. `GetBytes(count)` returns a new `Byte`
array (`count` zero returns an empty array; negative counts are refused). `GetInt32(
fromInclusive, toExclusive)` returns a uniform integer in the half-open range, using rejection
sampling so no value is favored; an empty range is refused with `InvalidArgumentError`. There
is no seed: this is the cryptographic source, not a repeatable sequence.

[Back to the documentation index](README.md)
