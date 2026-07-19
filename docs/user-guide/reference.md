# Practical reference

This page is the everyday map. The focused technical documents remain the authority for exact
contracts and edge cases.

## Create values and collections

| Need | Member |
|---|---|
| Typed primitive list | `ListOf(typeToken, items...)` |
| Typed object list inferred from a real object | `ListFrom(first, rest...)` |
| Empty list inferred from an object | `ListLike(example)` |
| Numeric sequence | `Range(start, count)` |
| Repeated value | `Repeat(value, count)` |
| Typed expression variable | `Var(typeToken)` or `VarLike(example)` |
| Captured expression value | `Value(value)` |

```vba
Set numbers = ROneCOne.ListOf( _
    vbLong, CLng(10), CLng(20), CLng(30))
Set customers = ROneCOne.ListFrom(ada, grace, katherine)
```

## Query a sequence

| Need | Member |
|---|---|
| Filter | `Where`, `WhereNot` |
| Project | `Map`, `SelectItems` |
| Order | `Order`, `OrderDescending`, `OrderBy`, `OrderByDescending`, `ThenBy`, `ThenByDescending` |
| Page | `Take`, `Skip` |
| Remove duplicates | `Distinct`, `DistinctBy` |
| Add around a sequence | `Append`, `Prepend` |
| Reverse | `Reverse` |
| Materialize | `ToList`, `ToArray` |

## Build a condition

| Kind | Members |
|---|---|
| Equality | `EqualTo`, `NotEqualTo`, `Match`, `NotMatch` |
| Ordering | `LessThan`, `AtMost`, `GreaterThan`, `AtLeast`, `Between` |
| Membership | `OneOf`, `IsIn`, `NotIn`, `ContainsMember` |
| Text | `StartsWith`, `EndsWith`, `ContainsText`, `MatchesPattern` |
| Text ignoring case | `EqualToIgnoreCase`, `StartsWithIgnoreCase`, `ContainsIgnoreCase` |
| Null and Boolean | `IsNothing`, `IsNotNothing`, `IsNullOrEmpty`, `IsTrue`, `IsFalse` |
| Composition | `Both`, `Either`, `Negated` |
| Constants | `Always`, `Never` |
| Nested sequences | `AnyMatch`, `AllMatch`, `NoneMatch` |

## Ask for a result

| Need | Member |
|---|---|
| Existence | `Exists`, `AnyItem`, `None`, `All` |
| Number of matches | `Count(predicate)` |
| First or last | `First`, `Last`, `FirstOrDefault`, `LastOrDefault` |
| Exactly one | `SingleItem`, `SingleOrDefault` |
| Numeric result | `Sum`, `Average`, `Min`, `Max`, `MinBy`, `MaxBy` |
| Join text | `JoinText(separator)` |
| Compare sequences | `SequenceEqual` |
| Run an action per item | `ForEach(action)` |

## Functions, actions, events, and exceptions

| Need | Member |
|---|---|
| Function | `Func`, `AsFunc` |
| Action | `Action`, `Execute` |
| Signature | `Takes`, `Returns` |
| Dynamic call | `DynamicInvoke` |
| Multicast | `Combine`, `Remove`, `GetInvocationList` |
| Composition | `PipeTo` |
| Typed event | `EventOf`, `Subscribe`, `Unsubscribe`, `Emit` |
| Structured failure flow | `Try`, `Catch`, `Finally` |

## Type tokens

Primitive collections and signatures use VBA's ordinary `VbVarType` constants, including
`vbBoolean`, `vbByte`, `vbInteger`, `vbLong`, `vbLongLong`, `vbSingle`, `vbDouble`, `vbCurrency`,
`vbDate`, `vbString`, `vbObject`, and `vbVariant`.

For an exact user-defined class, pass one non-`Nothing` instance as the type example. ROneCOne
records its concrete class name; it does not retain the example merely because it was used as a
type token.

## Default values

The `OrDefault` terminals follow `default(T)` behavior:

| Element type | Default |
|---|---|
| Object or user-defined class | `Nothing` |
| String | Empty string |
| Boolean | `False` |
| Numeric | Typed zero |
| Date | Zero date |
| Variant | `Empty` |

## Names VBA does not allow

VBA reserves several names that C# uses for LINQ. ROneCOne uses the closest legal, readable form:

| C# name | ROneCOne name |
|---|---|
| `Select` | `Map` or `SelectItems` |
| `Any` | `Exists` or `AnyItem` |
| `In` | `IsIn` |
| `Single` | `SingleItem` |
| Explicit delegate `Invoke` | `Run` or natural call syntax |

## Compatibility and deployment

- Windows x64 Microsoft 365 Excel
- `.xlsm`, `.xlsb`, and `.xlam`
- One imported `ROneCOne.cls` runtime file
- One Excel application process
- No runtime source generation or VBIDE trust
- No implicit network traffic or telemetry

## Technical documentation

- [Collections and LINQ](../collections.md)
- [Delegates and native invocation](../delegates.md)
- [Typed events](../events.md)
- [Structured exceptions](../exceptions.md)
- [Architecture and roadmap](../architecture.md)
- [Development and verification](../development.md)

[Back to the guide index](README.md)
