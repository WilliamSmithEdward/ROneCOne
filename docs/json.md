# JSON

Exact semantics for the System.Text.Json-style surface. For the workflow-first introduction,
read [JSON and typed objects](user-guide/json-and-objects.md).

## Model

`Deserialize` maps JSON to runtime-native values: an object becomes an ordered
String-to-Variant dictionary (insertion-ordered, O(1) keyed access, case-sensitive keys per
RFC 8259), an array becomes a Variant list, `true`/`false` become Boolean, `null` becomes
`Null`, strings become String, and numbers become the narrowest lossless type: `Long` for
integers up to nine digits, `LongLong` beyond that, and `Double` for fractions, exponents, and
overflow. A duplicate member name keeps the last value, matching System.Text.Json's dictionary
deserialization.

## Strictness and errors

Parsing follows RFC 8259: no leading zeros, no trailing commas, no unescaped control
characters, validated `\uXXXX` escapes and surrogate pairs, and no text after the document.
Violations raise error number `ROneCOne.JsonError` from source `ROneCOne.JsonException`, with
the character position in the message. The reader scans UTF-16 code units from a one-time byte
snapshot, returns escape-free strings with a single copy, and accumulates short integers during
the scan; the strategy is adapted, with its author's permission, from
[ModernJsonInVBA](https://github.com/WilliamSmithEdward/ModernJsonInVBA).

## Serialization

| Input | Output |
|---|---|
| Ordered or plain dictionary | Object, members in insertion order, keys as text |
| List, query, set, queue, or other sequence | Array |
| `DataTable` / `DataView` | Array of row objects, skipping deleted rows |
| `DataRow` | Object keyed by column names |
| VBA `Collection` / one-dimensional array | Array (two-dimensional arrays raise) |
| String, Boolean, integer types | JSON native |
| `Double`, `Single`, `Currency`, `Decimal` | Invariant-culture number, never a locale comma |
| `Date` | ISO 8601 text, `"2026-07-23T10:00:00"` |
| `Null`, `Empty`, `Nothing` | `null` |
| Any other object | Raises: map it with `DataTableFromObjects` first |

`Serialize(value, True)` and `ToJson(True)` produce two-space-indented output that parses back
to the identical document. The whole document is written into one pre-allocated buffer;
escaping copies clean runs in chunks.

## Tables and objects

`DeserializeTable(json, [tableName], [arrayPath])` expects an array of objects at `arrayPath`
(`"$"`, member steps `"$.data.items"`, and index steps `"[0]"`). Nested objects contribute
dotted columns, nested arrays are serialized into text cells, absent members become database
nulls, and column types unify across rows: matching numeric types widen (`Long` to `LongLong`
to `Double`), anything else mixes to Variant.

Factory-based binding stands in for reflection: the factory is a zero-argument Func declaring an
object result (`ROneCOne.Func("Module.NewItem").Takes().Returns(vbObject)`).
`DeserializeObjects(json, factory, [arrayPath])` invokes it per element and assigns scalar
members by property name; `DeserializeInto(json, instance)` binds one object, skipping null, object, and
array members; `ToObjects(factory)` maps table rows to instances (database nulls skip the
property); `DataTableFromObjects(source, propertyNames, [tableName])` reads the named
properties from each object, typing each column from its first non-null value. Property names
bind through VBA's own member dispatch, which is case-insensitive.

[Back to the documentation index](README.md)
