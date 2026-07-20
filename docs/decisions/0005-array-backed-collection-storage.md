# ADR 0005: Array-backed collection storage

Status: accepted, 2026-07-20

## Context

Every list, generic collection, DataTable row set, DataRow, and aggregate error stored its
elements in a VBA `Collection`. A `Collection` is a linked list: `Add` and `For Each` are cheap,
but numeric `Item(i)` access is linear. The public indexer resolved a position by materializing
the whole sequence into a fresh Collection and then indexing it, so a plain
`For i = 0 To n - 1: total = total + list(i)` loop was O(n squared) and, at ten thousand
elements, exceeded the test harness deadline outright. Value replacement rebuilt the entire
Collection. This was the deepest performance ceiling in the runtime.

## Decision

Replace the element and key stores (`mItems`, `mKeys`, `mPriorities`, `mOriginalItems`) with
doubling dynamic `ROneCOne` arrays paired with an explicit count, accessed through one set of
helpers (`ArrAppend`, `ArrInsert`, `ArrRemoveAt`, `ArrReplaceAt`, `ArrReset`, `ArrSnapshot`).
A positional read is now O(1), an append amortized O(1), and a value replacement O(1). The hot
find, sort-insert, and hash-rebuild helpers operate on the arrays directly rather than through a
Collection. Cross-instance `Internal...` accessors return a right-sized snapshot Collection so
every existing consumer is unchanged, and the enumerator cache that backs VBA `For Each` stays a
Collection because it exposes the hidden `[_NewEnum]` member an array cannot.

The delegate, task, and relational machinery (`mParameters`, `mInvocationList`, catch handlers,
DataSet tables and relations, primary-key column list) keeps using `Collection`. Those hold small
fixed sets that are never randomly indexed in a hot loop, so array backing would add churn with
no benefit; "array backing" means the indexed element stores, and all of them were converted.

## Alternatives considered

Keeping the Collection and adding a parallel array index was rejected as two structures that can
desync. Exact-size `ReDim Preserve` on every append was rejected because it makes append O(n) and
bulk builds O(n squared); over-allocation with a tracked count is why `For Each` had to move to
counted loops or snapshots. Converting the auxiliary Collections too was rejected as scope with
no performance payoff.

## Consequences

`list(i)` in a loop drops from quadratic to linear; the ten-thousand-element indexed-access test
that timed out under the Collection now completes instantly, and every other performance gate
holds. The public API and all 442 live assertions are unchanged. A regression test locks in the
scaling behavior.
