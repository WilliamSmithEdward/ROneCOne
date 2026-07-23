# ADR 0010: Version-cached snapshots and O(1) positional access

Status: accepted, 2026-07-20

## Context

Three more rebuild-per-operation cliffs shipped alongside the one ADR 0009 fixed, all invisible
to the gates because no benchmark exercised writes or repeated property access:

1. Every indexed list write rebuilt the For Each mirror Collection (`ReplaceItem` ended with a
   full `MaterializeValues`), and every list add, insert, and removal maintained that mirror
   eagerly. Ten thousand `list(i) = value` assignments did on the order of one hundred million
   Collection operations.
2. `Rows`, `Columns`, `Tables`, `Relations`, and `PrimaryKey` built a fresh snapshot object on
   every access, so the natural read loop `For i: table.Rows.Item(i)` was quadratic in
   allocation.
3. Positional reads on materialized generic collections (sets, queues, dictionaries) went
   through full per-access materialization, and the returned Collection was itself indexed by
   an O(index) walk.

A live scenario combining these exceeded the harness's 30-second deadline before the fix.

## Decision

Make the element array the single authority and rebuild derived structures lazily behind
version keys:

- Lists join the version-checked lazy branch `NewEnum` already used for queries and generic
  collections; all eager mirror maintenance is deleted. Mutations bump `mVersion`; the next
  For Each materializes once.
- The five data snapshot properties cache their built object against the structural version and
  return it until the version moves. To keep read-write row loops linear, the table version is
  split: membership and schema changes advance `mVersion` (invalidating snapshots), while field
  edits and row-state changes advance a new `mDataVersion` that feeds only
  `InternalSequenceVersion`, so views and enumerations refresh without discarding snapshots.
- `WrappedItem` indexes materialized generic collections directly (dictionaries construct the
  positional entry from the key and value arrays in O(1)); read-only wrappers delegate to their
  source. Deferred queries and live views still materialize per read on purpose: the documented
  contract is that a query always sees the latest data, and a version key cannot observe object
  field mutations. Positional loops over queries stay O(n) per access, exactly like `ElementAt`
  over a lazy sequence in .NET; `ToList` remains the documented answer.

Two defects found while pinning the semantics are fixed in the same change: replacing a view's
filter or sort now advances the view's version, so an already-enumerated view refreshes instead
of serving stale results, and `AddColumn` on a populated table backfills every existing row with
the column's default cell instead of leaving rows one cell short of the schema.

One behavior change is accepted and intended: repeated `Rows`-style access returns the same
cached snapshot object until the owner structurally changes, rather than a fresh copy per call.

## Consequences

The live scenarios went from exceeding the 30-second deadline to 0.100 seconds for 10,000
indexed list writes, 0.049 seconds for 2,000 positional row reads, and 0.018 seconds for 10,000
positional set reads, each against a new 1.5-second release gate wired into the harness
(`Collection Benchmarks` B17:B25). Live consistency contracts cover mirror laziness across every
list mutation, snapshot refresh on membership and schema changes, view refresh after refilter
and after field edits, and the late-column backfill. The keyed update path also lost a redundant
probe (the found slot is reused), and custom-comparer membership scans read the element array
directly instead of allocating a snapshot Collection per test.
