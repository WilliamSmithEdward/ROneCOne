# ADR 0006: Defer compiled query plans and DISPID member-access caching

Status: accepted, 2026-07-20

## Context

Two optimizations were proposed for the query path: compiling each lambda's expression tree once
into a flat instruction stream to cut the per-node allocation the recursive evaluator does, and
caching object member DISPIDs so property access in object LINQ bypasses `CallByName`'s
name resolution. Both are real optimizations. The engineering rule here is to measure before
optimizing and to pursue a frontier change only when it delivers a material end-user gain.

## Measurement

A probe filtered a 10,000-element sequence (about 5,000 matches), best of three runs, on the
target host:

- Raw hand-written VBA loop: 0.049 s
- Numeric predicate through the evaluator, no member access: 0.127 s
- Object-property predicate, adding `CallByName` per element: 0.179 s

The evaluator adds roughly 0.078 s over a raw loop across 10,000 elements; `CallByName` member
access adds roughly 0.052 s on top. Compiled plans could recover part of the first, DISPID
caching part of the second.

## Decision

Do not implement compiled query plans or DISPID caching now. At 0.13 to 0.18 s for a
ten-thousand-object query, neither path is a user-felt bottleneck, so their real but small
constant-factor gains do not justify their cost: compiled plans require re-implementing the
correctness-critical operator set of the evaluator, and DISPID caching requires raw
`IDispatch` vtable dispatch with fail-closed ABI handling. The risk outweighs an unfelt gain.

## Consequences

The query evaluator and `CallByName` member access are unchanged. This decision is data-backed
and revisitable: if a workload shows user-felt slowness in large object-property queries (tens of
thousands of elements or more), revisit with a fresh measurement, and the DISPID path would be
the higher-value of the two because member access is the larger and more targetable share. The
measurement method (numeric predicate isolates the evaluator; object predicate isolates member
access; both against a raw-loop baseline) is recorded here to repeat.
