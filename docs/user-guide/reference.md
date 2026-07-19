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
| Dictionary | `DictionaryOf(keyType, valueType)` |
| Set | `HashSetOf(typeToken)` or `SortedSetOf(typeToken)` |
| Queue or stack | `QueueOf(typeToken)` or `StackOf(typeToken)` |
| Linked list | `LinkedListOf(typeToken)` |
| Sorted map | `SortedListOf` or `SortedDictionaryOf` |
| Priority queue | `PriorityQueueOf(elementType, priorityType)` |
| Observable/read-only/keyed | `ObservableCollectionOf`, `ReadOnlyCollectionOf`, `KeyedCollectionOf` |
| Concurrent-style | `ConcurrentDictionaryOf`, `ConcurrentQueueOf`, `ConcurrentStackOf`, `ConcurrentBagOf` |
| Blocking collection | `BlockingCollectionOf(typeToken, capacity)` |
| Immutable | `ImmutableListOf`, `ImmutableDictionaryOf`, `ImmutableHashSetOf`, and related factories |
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
| Combine | `Concat`, `Union`, `Intersect`, `Except` and `...By` forms |
| Flatten or pair | `SelectMany`, `Zip`, `Join`, `GroupJoin` |
| Partition | `TakeWhile`, `SkipWhile`, `TakeLast`, `SkipLast`, `Chunk` |
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
| Group or index | `GroupBy`, `ToLookup`, `ToDictionary`, `ToHashSet`, `Index` |
| Modern aggregates | `CountBy`, `Aggregate`, `AggregateBy` |
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

## Tasks and data

| Need | Member |
|---|---|
| Completed or delayed task | `Task.FromResult`, `Task.CompletedTask`, `Task.Delay` |
| Coordinate tasks | `Await`, `Wait`, `WhenAll`, `WhenAny`, `ContinueWith` |
| Cancellation | `CancellationTokenSource`, `Cancel`, `CancelAfter`, `Register` |
| Progress or external completion | `ProgressOf`, `TaskCompletionSourceOf` |
| In-memory data | `DataTable`, `DataColumn`, `DataSet`, `DataRelation`, `DataView` |
| Provider access | `DbConnection`, `DbCommand`, `DbParameter`, `DbDataAdapter` |
| Async provider access | `OpenAsync`, `ExecuteReaderAsync`, `FillAsync`, `UpdateAsync` |

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
| `DbConnection.Open` / `Close` | `Connect` / `Disconnect` |

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
- [Tasks, cancellation, and progress](../tasks.md)
- [DataTable, DataSet, and providers](../data.md)
- [Architecture and release capabilities](../architecture.md)
- [Development and verification](../development.md)

[Back to the guide index](README.md)
