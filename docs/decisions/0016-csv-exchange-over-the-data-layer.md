# ADR 0016: RFC 4180 CSV exchange over the data layer

Status: accepted, 2026-07-23

## Context

CSV is the most-handled data format in Excel work, and VBA offers nothing for it: `Split` on
commas breaks on the first quoted field, `Workbooks.OpenText` changes the active workbook and
guesses types by locale, and hand-rolled parsers rot. The runtime already round-trips tables
through JSON (ADR 0013) and reads and writes real files (ADR 0015); CSV completes data-in,
data-out for tabular text.

## Decision

One factory surface, `ROneCOne.Csv`, reuses the shared member names by role dispatch:
`Serialize(value)` writes a `DataTable` or `DataView`, `DeserializeTable(text, [tableName])`
parses into a typed table, and `table.ToCsv` mirrors `table.ToJson`. Writing follows RFC 4180
with minimal quoting and CRLF-terminated rows, and reuses the JSON writer's invariant value
conventions (lowercase Boolean literals, dot-decimal numbers, ISO 8601 timestamps). A
database null writes an empty unquoted field and an empty string writes a quoted pair, so the
distinction survives a round trip.

Parsing is strict where the format is defined (quote discipline, uniform field counts, unique
non-empty headers) and tolerant where reality varies (CRLF, LF, or CR line endings). Type
inference is column-level and deterministic: unquoted cells classify by JSON's number grammar
(leading zeros disqualify), `true`/`false`, or a validated ISO 8601 timestamp; quoted cells
are always text; matching numeric types widen; any other mixture falls the column back to
`String` with every cell's original characters preserved. Cell-level Variant typing (JSON's
choice) was rejected for CSV because the format carries no type markers: a column that mixes
`1` and `abc` almost always means text, and silently converting some cells would corrupt
values like `00042`.

The parse errors are typed (`ROneCOne.CsvError`, source `ROneCOne.CsvException`) with row or
character positions, mirroring the JSON contract. As part of this change the number reader
behind both layers now converts validated non-integral text with `Val`, which always parses
the invariant dot separator, instead of the locale-sensitive `CDbl`.

## Consequences

`DeserializeTable` is now a shared member: its text parameter is named `text` rather than
`jsonText`, its default table name resolves per role (`Json` or `Csv`), and the CSV role
rejects an array path. The live suite adds a CSV contract covering the exact serialized
document, typed round trips including null-versus-empty-string, quoting and embedded
newlines, inference and widening cases, view serialization, file composition through
`ROneCOne.File`, and the typed failure paths. All of it runs offline.
