# ADR 0013: JSON in the spirit of System.Text.Json

Status: accepted, 2026-07-23

## Context

Exchanging JSON from VBA means hand-rolled parsing or importing a third-party module, exactly
the dependency class this project exists to remove, and the HTTP client made the gap loud:
the runtime could download an API response but not read it. The goal was
`JsonSerializer.Serialize` / `Deserialize` in spirit, within two hard VBA limits: no runtime
reflection (property names of a user class cannot be enumerated) and no `Activator`-style
instantiation (a class cannot be created from its name).

## Decision

The parsing and writing mechanics are adapted from
[ModernJsonInVBA](https://github.com/WilliamSmithEdward/ModernJsonInVBA), with its author's
permission: character codes are scanned from a one-time UTF-16 byte snapshot, strings are read
in chunks bounded by the candidate closing quote with an escape-free fast path, integers of up
to nine digits are accumulated during the scan, documents are written into one pre-allocated
buffer, numbers are formatted with an invariant decimal separator, and RFC 8259 strictness is
enforced with position-carrying errors (`ROneCOne.JsonError`, source
`ROneCOne.JsonException`).

The model, however, is runtime-native rather than the library's tagged Collections: a JSON
object deserializes to an ordered String-to-Variant dictionary (insertion-ordered, hash-keyed)
and a JSON array to a Variant list, so a parsed tree is immediately navigable and queryable
with the existing collection surface. Integers beyond `Long` widen to `LongLong` before
falling to `Double`, dates serialize as ISO 8601, and duplicate members keep the last value,
matching System.Text.Json's dictionary behavior.

The two VBA limits are answered with two explicit contracts instead of magic:

- A zero-argument factory delegate with a declared object result
  (`ROneCOne.Func("Module.NewItem").Takes().Returns(vbObject)`) stands in for instantiation;
  the declaration is validated up front, because without it the delegate machinery would take
  its scalar invocation path and fail opaquely. `DeserializeObjects` and `ToObjects` call the
  factory per element and bind scalar members by property name through VBA's own member
  dispatch; `DeserializeInto` binds onto an instance the caller already has.
- An explicit property-name list stands in for reflection. `DataTableFromObjects` reads the
  named properties from each object and types each column from its first non-null value.

`DeserializeTable` lands an array of objects (optionally at a `$.data.items`-style path)
directly in a typed `DataTable`: dotted columns for nested objects, JSON text cells for nested
arrays, database nulls for absences, and numeric column types that widen across rows. With
`ToRange`, a web response becomes a worksheet in three calls.

## Consequences

The surface is `ROneCOne.Json` (`Serialize`, `Deserialize`, `DeserializeTable`,
`DeserializeInto`, `DeserializeObjects`), `DataTableFromObjects`, `table.ToObjects`, and a
`ToJson` convenience on serializable roles. Live contracts cover the model, escapes and
surrogate validation, RFC strictness rejections, compact and indented round-trips, table
mapping with dotted columns and type unification, and every binding direction. Serialization
of an arbitrary unmapped object raises with a pointer to `DataTableFromObjects` rather than
guessing; nested object and array members are skipped by the flat binders by design. The
parser and writer run entirely in-process with no new dependencies.
