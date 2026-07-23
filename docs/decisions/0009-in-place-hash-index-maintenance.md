# ADR 0009: In-place hash index maintenance for keyed mutation

Status: accepted, 2026-07-20

## Context

Through 1.3.0, every keyed mutation rebuilt the entire hash index. Replacing a value on an
existing key (`d.Item(k) = v`), removing a key, and each iteration of `RemoveWhere` all called
`RebuildHashIndex`, which re-hashes and re-inserts every live entry. Reads were O(1) but the
canonical tally pattern `d(k) = d(k) + 1` paid a full O(n) rebuild per increment. A live scenario
of 10,000 in-place updates plus 2,000 keyed removals on a 10,000-entry dictionary exceeded the
harness's 30-second hard deadline. The established benchmark gates only measured adds and reads,
which is why the cliff shipped unseen.

A dirty-flag alone (defer one rebuild until the next probe) was implemented and measured first:
8.39 seconds for the same scenario. It fails because removal by key probes the index before every
removal, so interleaved probe-remove sequences still rebuild once per operation.

## Decision

Maintain the index in place, with three cooperating mechanisms:

1. Replacing the value of an existing key updates the one slot the probe already identified
   (`RefreshHashSlotValue`). Keys and positions are untouched, so nothing else changes.
2. Removing an entry deletes its slot and repairs the probe cluster with the standard linear-probe
   backward-shift (`DeleteHashSlot`), then slides every stored canonical index above the removed
   position down by one (`ShiftHashIndexesAbove`). Removing the last position skips the slide
   entirely, so tail removal does only cluster-local work.
3. Bulk removal loops that never probe between removals (`RemoveWhere`, `ExceptWith`,
   `IntersectWith`) mark the index dirty once (`MarkHashIndexDirty`); the next probe or insert
   performs a single rebuild (`EnsureHashIndexCurrent`). Every probe and insert path flushes the
   flag first, so a stale index is never read.

A pending-removals list with adjusted probes (tombstone-style deferral) was considered and
rejected: it is faster only for interleaved workloads that are already fast enough, and it puts
materially more state and invariants into the runtime's most central structure.

## Consequences

The 12,000-operation mutation scenario went from exceeding the 30-second deadline to 0.586 and
1.281 seconds on a normal and a heavily loaded sample, inside its new 1.5-second release gate.
The scenario is wired into the live harness (`Collection Benchmarks` B14:B16, worker gate
`-MaxHashMutationBenchmarkSeconds`), a live consistency contract covers update, removal, re-add,
enumeration, the concurrent surface, and hash-set element replacement after mutation, and the
source contract pins the in-place wiring. Removing from the front of a large keyed collection
still pays one O(capacity) index slide; that matches the O(n) element shift the dense ordered
arrays already require, so removal's asymptotic class is unchanged while its constant no longer
includes re-hashing every entry.
