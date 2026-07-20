# ADR 0003: Lossless numeric widening at type admission

Status: accepted, 2026-07-19

## Context

Typed primitive collections, signatures, keys, and DataColumns admitted a value only when its
runtime `VarType` matched the declared type exactly. Because VBA integer literals arrive as
`Integer` (or `Long` when large) and floating literals as `Double`, every example and every
call site had to wrap literals in `CLng`, `CDbl`, and the like:
`ROneCOne.ListOf(vbLong, CLng(90), CLng(72))`. This ceremony was the first friction a new
developer met, and it added nothing: a `Long` list initialised from integer literals has one
obviously correct meaning. The product direction is maximal safe syntax sugar, so the ceremony
had to go without weakening the strictness that catches real mistakes.

## Decision

Admit a scalar whose type is losslessly widenable to the declared type, and coerce the stored
value to the declared type so its runtime type stays exact. Widening is applied at every
admission point: list and collection elements, dictionary and keyed-collection keys, lambda
parameter binding, delegate arguments, delegate return values, DataColumn values, primary-key
`Find`, progress values, completion-source results, and out-parameter references. The lattice
admits only provably safe promotions: within the integer family (Byte to Integer to Long to
LongLong), Single to Double, integers into Double or Single or Currency where the whole range
fits, and nothing else. Narrowing, cross-family, and any Boolean, Date, or String conversion
remain hard rejects, so genuinely wrong values still raise `TypeMismatchError` atomically.

## Alternatives considered

Accepting any convertible value via `CLng`-style coercion was rejected: it would silently
truncate `1.9` to `1` and mask real errors. A per-collection "loose mode" flag was rejected as
configuration where a single correct default suffices. Widening only list elements was rejected
under the no-partial rule; the ceremony appears at every typed boundary, so the fix must too.

## Consequences

`ROneCOne.ListOf(vbLong, 90, 72, 88)`, `dict.Add 101, 95`, `applyDiscount(100)`, and
`row.Item("Age") = 40` all work without conversion wrappers, and stored values keep their exact
declared type so `GenericTypeName`, hashing, ordering, and equality are unchanged. Strictness is
preserved where it protects. Explicit conversions remain valid and are never required. The
widening lattice is covered by accept-and-reject edge tests in the live collection suite.
