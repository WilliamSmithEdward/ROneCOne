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
| Reserve or compact hash storage | `Capacity`, `EnsureCapacity`, `TrimExcess` |
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
    vbLong, 10, 20, 30)
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
| Readable form | `ToDisplayString` |
| Dynamic call | `DynamicInvoke` |
| Multicast | `Combine`, `Remove`, `GetInvocationList` |
| Composition | `PipeTo` |
| Typed event | `EventOf`, `Subscribe`, `Unsubscribe`, `Emit` |
| Structured failure flow | `Try`, `Catch`, `Finally` |

## Tasks and data

| Need | Member |
|---|---|
| Completed or delayed task | `Task.FromResult`, `Task.CompletedTask`, `Task.Delay` |
| Coordinate tasks | `Await`, `Wait`, `WaitAsync`, `WhenAll`, `WhenAny`, `ContinueWith`, `YieldOnce` |
| Observe task failures | `Exception`, `InnerExceptions`, `Flatten`, `Handle` |
| Cancellation | `CancellationTokenSource`, `Cancel`, `CancelAfter`, `Register`, `Dispose` |
| Progress or external completion | `ProgressOf`, `TaskCompletionSourceOf` |
| In-memory data | `DataTable`, `DataColumn`, `Row`, `DBNull`, `Find`, `DataSet`, `DataRelation`, `DataView` |
| Worksheet range | `DataTableFromRange`, `LoadFromRange`, `ListFromRange`, `ToRange` |
| Provider access | `DbConnection`, `DbCommand`, `DbParameter`, `DbDataAdapter` |
| Async provider access | `OpenAsync`, `ExecuteReaderAsync`, `FillAsync`, `UpdateAsync` |
| Provider capability | `AsyncMode`, `State` |
| Deterministic cleanup | `Using(resource).Run(work)` |
| CSV exchange | `Csv.Serialize`, `Csv.DeserializeTable`, `table.ToCsv` |
| Text files | `File.ReadAllText`, `WriteAllText`, `AppendAllText`, `ReadAllLines`, `WriteAllLines` |
| Binary files and management | `ReadAllBytes`, `WriteAllBytes`, `Exists`, `Copy`, `Move`, `Delete` |
| Folders | `Directory.CreateDirectory`, `Exists`, `Delete`, `GetFiles`, `GetDirectories` |
| Path text | `Path.Combine`, `GetFileName`, `GetDirectoryName`, `GetExtension`, `ChangeExtension`, `GetFullPath`, `GetTempPath` |
| Shell commands | `Process.RunAsync`, `ExitCode`, `StandardOutput`, `StandardError` |
| Regular expressions | `Regex`, `IsMatch`, `Match`, `Matches`, `Replace`, `Split` |
| Hashing | `Hash.Sha256`, `Sha512`, `Sha1`, `Md5`, `HmacSha256` |
| Byte encoding | `Convert.ToBase64String`, `FromBase64String`, `ToHexString`, `FromHexString` |
| Download a file | `HttpClient.DownloadFileAsync` |
| Instants with offsets | `DateTime.Parse`, `TryParse`, `UtcNow`, `Now`, `Today`, `FromLocal`, `FromUtc` |
| Epoch numbers | `FromUnixTimeSeconds`, `FromUnixTimeMilliseconds`, `ToUnixTimeSeconds`, `ToUnixTimeMilliseconds` |
| Instant views and text | `ToUniversalTime`, `ToLocalTime`, `ToOffset`, `ToIsoString`, `ToString(pattern)` |
| Calendar arithmetic | `AddYears` .. `AddMilliseconds`, `Add`, `Subtract`, `CompareTo` |
| Durations | `TimeSpan.FromDays` .. `FromMilliseconds`, `Zero`, `Parse`, `Total*`, components, `Duration`, `Negate` |
| Invariant formatting | `Strings.Format("{0,8:N2}", ...)` with `G N F D X P` and date tokens |
| Text building | `StringBuilder().Append`, `AppendLine`, `AppendFormat`, `Length`, `Clear`, `ToString` |
| Identifiers | `Guid.NewGuid`, `Guid.EmptyGuid` |
| Crypto randomness | `RandomNumberGenerator.GetBytes(count)`, `GetInt32(fromInclusive, toExclusive)` |
| XML documents | `Xml.Parse`, `Xml.Load`, `Name`, `Value`, `GetAttribute`, `HasAttribute`, `OuterXml` |
| XML queries | `Elements([name])`, `SelectNodes(xpath)`, `SelectSingleNode(xpath)` |
| XML and tables | `Xml.DeserializeTable(text, [name], [rowsPath])`, `table.ToXml([root], [row])` |

## Type tokens

Primitive collections and signatures use VBA's ordinary `VbVarType` constants, including
`vbBoolean`, `vbByte`, `vbInteger`, `vbLong`, `vbLongLong`, `vbSingle`, `vbDouble`, `vbCurrency`,
`vbDate`, `vbString`, `vbObject`, and `vbVariant`.

For an exact user-defined class, pass one non-`Nothing` instance as the type example. ROneCOne
records its concrete class name; it does not retain the example merely because it was used as a
type token.

## Numeric literals widen automatically

A typed numeric collection, signature, key, or DataColumn accepts any value that promotes to its
declared type without loss, and stores it as the declared type. Plain literals need no `CLng`,
`CDbl`, or similar wrapper.

| Declared type | Also accepts | Never accepts |
|---|---|---|
| `vbInteger` | `Byte` | wider or fractional numbers |
| `vbLong` | `Byte`, `Integer` | `LongLong`, `Single`, `Double`, `Currency` |
| `vbLongLong` | `Byte`, `Integer`, `Long` | `Single`, `Double`, `Currency` |
| `vbCurrency` | `Byte`, `Integer`, `Long` | floating-point sources |
| `vbSingle` | `Byte`, `Integer` | `Long`, `LongLong`, `Double` |
| `vbDouble` | `Byte`, `Integer`, `Long`, `Single` | `LongLong`, `Currency` |

Any narrowing, cross-family, or Boolean, Date, or String conversion is refused with
`TypeMismatchError`, atomically, before the collection changes. Explicit conversions stay valid
and are never required.

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
| `XElement.Attribute` | `GetAttribute` (with `HasAttribute`) |
| `System.Random` | `RandomNumberGenerator` (the crypto source; no seed) |

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

[Back to the guide index](README.md) | [Documentation index](../README.md)
