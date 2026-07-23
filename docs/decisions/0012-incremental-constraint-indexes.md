# ADR 0012: Incremental, lazily rebuilt constraint indexes

Status: accepted, 2026-07-23

## Context

The data layer's constraint machinery was the last rebuild-per-operation surface. Every
`AddRow` rebuilt the whole primary-key index; every field write on an attached row rebuilt it
again, key column or not; `ValidateDataRowConstraints` ignored its changed-column argument and
re-validated every column on every edit; and each `Unique` column was enforced by scanning all
rows per validation. Bulk-loading a keyed table, editing fields in a loop, and adapter `Fill`
were all quadratic, twice over when a unique column existed. The benchmark table defined no
primary key and no unique column, so three releases of gates never saw it; a live scenario of
2,000 loads plus 2,000 edits plus 2,000 finds on a constrained table exceeded the harness's
30-second deadline.

## Decision

Constraint enforcement becomes probes over token indexes that maintain themselves the same way
every other derived structure in the runtime now does:

- One rebuild pass constructs the primary-key index and one token index per `Unique` column.
- `AddRow` indexes the new row incrementally; validation has just proven its tokens absent.
- A field write validates only the changed column, and only a write to a primary-key or unique
  column marks the indexes dirty; `Find` and validation flush the flag with one rebuild.
- Deleting a row marks the indexes dirty, so its key and unique values free up on the next
  probe, and adding a unique column (or `AsUnique` on an attached column) invalidates too.
- Uniqueness checks probe the column's token index, tolerating a hit on the row itself so an
  unchanged value never conflicts with its own entry.

Semantics are unchanged: the same duplicates are rejected with the same errors, failed edits
still restore the previous value, and deleted rows still fall out of every index.

## Consequences

The 6,000-operation constrained scenario went from exceeding the 30-second deadline to well
inside a new 1.5-second release gate wired into the harness (`Collection Benchmarks` B26:B28,
`-MaxConstraintBenchmarkSeconds`). Live contracts pin duplicate-key and duplicate-unique
rejection on add and on edit, restoration after a rejected edit, lazy reindexing after key
edits and deletes, and unique-value reuse after deletion. The source contract pins the
incremental wiring. `AcceptChanges` and `RejectChanges` keep their single full rebuild, since
they restructure membership wholesale.
