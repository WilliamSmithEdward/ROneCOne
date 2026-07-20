# ADR 0004: Worksheet Range bridge

Status: accepted, 2026-07-19

## Context

ROneCOne's audience lives in worksheets, yet moving data between a range and a typed collection
or DataTable was the one everyday task the library did not help with. Users hand-rolled cell
loops, which are both verbose and slow: each cell read or write is a separate COM round trip, so
a few thousand cells cost seconds. The Excel object model already exposes a fast path, reading or
writing a whole rectangular block as a single 2D array through `Range.Value`, but nothing in the
runtime surfaced it.

## Decision

Add a Range bridge built entirely on single bulk `Range.Value` calls, never per-cell access:

- `ROneCOne.DataTableFromRange(range, headers, name)` reads a block into a new Variant-typed
  DataTable, taking column names from the header row when present.
- `DataTable.LoadFromRange(range, headers)` appends a block into an existing typed table,
  coercing each value through the table's column types (including numeric widening).
- `ROneCOne.ListFromRange(range)` reads a single row or column into a `List<Variant>`.
- `ToRange(target, writeHeaders)` writes a DataTable grid, a DataView's visible grid, or a
  scalar sequence's column vector back to a range in one assignment.

Range parameters are typed `Object` and used through late binding so the runtime keeps no
compile-time Excel reference and the one-file, host-agnostic source contract holds.

## Alternatives considered

Per-cell iteration was rejected on performance. A typed-inference reader that guesses each
column's type from its data was rejected as the default because mixed columns make it
unpredictable; `DataTableFromRange` uses Variant columns, and `LoadFromRange` gives full type
control against a schema the caller defines. Writing through `Range.Value2` was rejected in favor
of `Range.Value` so dates, currency, and Booleans keep their types across a round trip.

## Consequences

Worksheet data crosses into the typed world and back in one call each way, orders of magnitude
faster than the cell loops users wrote before, and object-valued cells raise a clear error rather
than corrupting a sheet. The bridge is covered by a live round-trip test. The runtime gains no
external dependency.
