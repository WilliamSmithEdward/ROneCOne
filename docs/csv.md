# CSV

Exact semantics for the CSV exchange surface. For the workflow-first introduction, read the
CSV section of [Data and providers](user-guide/data-and-providers.md).

## Surface

`ROneCOne.Csv.Serialize(value)` and `table.ToCsv` write a `DataTable` or `DataView` to RFC
4180 text. `ROneCOne.Csv.DeserializeTable(text, [tableName])` parses CSV text into a typed
`DataTable` (named `Csv` by default). CSV has no indented form and no array path; passing
either raises `InvalidArgumentError`. Parse and write violations raise error number
`ROneCOne.CsvError` from source `ROneCOne.CsvException` with the row or character position in
the message.

## Writing

The header row carries the column names. Rows terminate with CRLF, including the last. A
field is quoted only when RFC 4180 requires it: embedded quotes (doubled), commas, line
breaks, leading or trailing spaces, or an empty string. Cell values follow the JSON writer's
conventions: `true`/`false` for Boolean, invariant-culture numbers (never a locale comma),
ISO 8601 timestamps for dates, and deleted rows are skipped. A database null writes an empty
unquoted field; an empty string writes a quoted pair, so the two survive a round trip
distinctly. Non-scalar cells raise.

## Parsing

Parsing is strict RFC 4180 with tolerant line endings (CRLF, LF, or CR): quoted fields may
contain anything with quotes doubled, a quote inside an unquoted field raises, text after a
closing quote raises, and every row must match the header's field count. Header names must be
non-empty and unique (case-insensitive).

## Type inference

Each column unifies to one deterministic type from its non-null cells:

| Cell shape (unquoted) | Type |
|---|---|
| JSON-grammar integer (no leading zeros) | `Long`, widening to `LongLong` then `Double` |
| JSON-grammar fraction or exponent | `Double` |
| `true` / `false` (case-insensitive) | `Boolean` |
| ISO 8601 `yyyy-mm-dd` or `yyyy-mm-ddThh:nn:ss` (validated, no rollover) | `Date` |
| Anything else, or any quoted cell | `String` |

A quoted cell is always text, so deliberately quoted values such as `"00042"` or `"true"`
stay intact. Matching numeric types widen across rows; any other mixture falls the column
back to `String` and every cell keeps its original characters. An unquoted empty field is a
database null in any column; an all-null column is `String`.

[Back to the documentation index](README.md)
