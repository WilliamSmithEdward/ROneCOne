# ADR 0029: JSON partial reads

Status: accepted, 2026-07-28

## Context

Consuming a large JSON API cost a full parse even when the caller wanted a handful of members.
`Deserialize` materialized every dictionary, list, and string in the document before the caller
could select anything, and the `arrayPath` parameter on `DeserializeTable` and
`DeserializeObjects` did not help: both parsed the whole document first and then navigated the
materialized tree, so the path located rows rather than avoiding work.

Measured on PokeAPI `/pokemon/{id}`, consumed by a downstream project that renders `id`, `name`,
`height`, `weight`, `types`, `stats`, and `sprites.front_default`:

| Entry | Raw bytes | Bytes needed | Share used |
|---|---|---|---|
| #1 Bulbasaur | 271,521 | 1,813 | 0.7% |
| #25 Pikachu | 290,335 | 2,397 | 0.8% |
| #143 Snorlax | 405,282 | 1,754 | 0.4% |
| #151 Mew | 664,610 | 1,745 | 0.3% |

The bulk is `moves` and the `sprites.versions` and `sprites.other` trees. Over 99% of the parse
was wasted. The workaround living downstream was a textual section stripper: a denylist that
removed known-unwanted keys with a bracket-depth scan before parsing. It worked, but it encoded
one API's response shape, and every consumer of a fat API would otherwise write its own copy.

## Decision

`JsonSkipValue` advances the reader past one complete value without building anything, mirroring
`JsonReadValue`'s dispatch with dedicated scanners for strings and numbers so a skipped subtree
allocates nothing. Three entry points ride it.

`Json.DeserializeOnly(jsonText, paths)` walks the document once and materializes only the
requested paths, skipping every other member. The result is structure preserving: keeping
`"$.sprites.front_default"` returns `sprites` as an object holding only that member, so
navigation is identical to `Deserialize` and the result is a true subset of the source.

`Json.DeserializeAt(jsonText, path)` returns the single value a path addresses, skipping every
sibling on the way down.

`DeserializeTable` and `DeserializeObjects` now apply `arrayPath` during the scan rather than
after a full parse, so an envelope's unrelated members are stepped over instead of built and then
discarded. The parameter finally means what its name implies.

Paths use the grammar `ResolveJsonPath` already defined (`$`, `$.a.b`, `$.rows[0].values`), so the
library keeps one path syntax rather than two.

The four open questions are settled as follows. A missing path is omitted silently by
`DeserializeOnly`, because an allowlist over a varying API should tolerate absence, while
`DeserializeAt` raises `ROneCOne.JsonError`, because asking for one value and receiving nothing is
a question that deserves an answer; this mirrors `HasAttribute` against `GetAttribute` in the XML
layer. Index steps are accepted at any depth including the root. Duplicate members keep the last
value, which is what the full parser already did, so the two paths cannot disagree.
`DeserializeOnly` keeps its name: it has no direct BCL analogue, but neither did `StringBuilder`
or `RandomNumberGenerator` in this runtime.

## Consequences

No new role, error number, prog-id, or `Declare`. The skip functions are pure reader arithmetic
over the byte snapshot the parser already holds.

Binding changed behavior, and this is a fix rather than an optimization. `BindJsonObjectInto`
called `CallByName` with no error trap, so a scalar member the target class did not model raised
438; `DeserializeInto` could not touch a real API response at all. Unknown members are now
ignored, matching how System.Text.Json treats them. Nested objects and arrays were already left
to the caller, so that part of the contract is unchanged.

The live suite gained twenty-seven assertions, including a benchmark that requires the partial
read to beat a whole parse on a document whose bulk is one unwanted member. That is the claim the
issue left open: the byte reduction was measured, but the parse-time saving was not, and asserting
it in the versioned harness is what keeps it true.

The downstream textual stripper in ReDim's `PokeDex.bas` can now be replaced by an allowlist call.
That replacement is deliberately not part of this change, since it lives in another repository.
