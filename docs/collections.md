# Runtime-generic collections and LINQ

New to ROneCOne? Start with the
[Collections and LINQ user guide](user-guide/collections-and-linq.md).

ROneCOne provides strictly typed generic collection families and deferred query pipelines in
the same single `ROneCOne.cls` runtime file. The design follows the behavior of
[`List<T>`](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1) and
[`Enumerable`](https://learn.microsoft.com/en-us/dotnet/api/system.linq.enumerable) where VBA's
syntax and type system permit it.

## Generic collection families

The one-file runtime includes List, Dictionary, HashSet, Queue, Stack, LinkedList, SortedList,
SortedDictionary, SortedSet, OrderedDictionary, PriorityQueue, ObservableCollection,
ReadOnlyCollection, KeyedCollection, concurrent-style dictionary/queue/stack/bag,
BlockingCollection, and immutable list/dictionary/hash-set/queue/stack/sorted variants.

Factories accept VBA type tokens or a real user-class instance. Specialized operations follow
their .NET counterparts, including set algebra, linked nodes, priority pairs, builders, snapshots,
bounded completion, and atomic-style `Try...` methods within Excel's single execution thread.

## Indexed collections and capacity

Default-equality dictionaries, hash sets, ordered/keyed collections, concurrent and immutable
hash collections, lookups, primary keys, and relations share an open-address hash kernel with
direct key/value slots. Average keyed lookup is O(1); removals rebuild a compact valid index.
Sorted maps and sets use binary search. Supplying a custom equality comparer deliberately uses a
linear equality path because VBA cannot derive a compatible hash code from an arbitrary delegate.

`Capacity`, `EnsureCapacity(count)`, and `TrimExcess` expose allocation control on hash-backed
collections. Reserving capacity before bulk insertion avoids repeated rehashing. These members
reject collections whose comparer or storage strategy cannot honor the contract rather than
pretending that a reservation occurred.

```vba
Dim scores As ROneCOne
Dim reserved As Long

Set scores = ROneCOne.DictionaryOf(vbLong, vbLong)
reserved = scores.EnsureCapacity(10000&)
scores.Add 1&, 95&
Debug.Print scores.Item(1&)
scores.TrimExcess
```

## Primitive lists

Use a `VbVarType` token as `T`. Values must have that exact VBA runtime type; ROneCOne does not
silently coerce them.

```vba
Dim numbers As ROneCOne

Set numbers = ROneCOne.ListOf(vbLong, CLng(5), CLng(10))

Debug.Print numbers.GenericTypeName  ' List<Long>
Debug.Print numbers(0)               ' 5: default zero-based indexer
Debug.Print numbers.Item(1)          ' 10: explicit indexer
```

An invalid mutation raises `ROneCOne.TypeMismatchError` before changing the list.

## User-defined class lists

`ListFrom(first, rest...)` infers the exact concrete class from the first item. `ListLike(example)`
creates an empty list of that inferred type. When an empty list has no real example yet, pass one
non-`Nothing` prototype to `ListOf`; the prototype is not retained. Members may contain instances
of that class or `Nothing`; unrelated object types are rejected.

```vba
Dim ada As Customer
Dim customers As ROneCOne

Set ada = New Customer
ada.CustomerName = "Ada"
Set customers = ROneCOne.ListFrom(ada)

Debug.Print customers.GenericTypeName       ' List<Customer>
Debug.Print customers.Item(0).CustomerName  ' Ada
```

This prototype token provides dependency-free `List<Customer>` semantics without code generation,
VBIDE trust, or a second runtime class.

## List surface

| C# idea | ROneCOne member | Behavior |
|---|---|---|
| `new List<T> { ... }` | `ListOf(typeToken, items...)` | Creates an explicitly typed list |
| inferred collection expression | `ListFrom(first, rest...)` | Infers `T` and populates the list |
| `new List<T>()` | `ListLike(example)` | Infers `T` and creates an empty list |
| `list[index]` | `list(index)` / `list.Item(index)` | Zero-based get; `Item` also supports set |
| `Add` / `AddRange` | `Add` / `AddRange` | Accepts typed sequences, arrays, or Collections atomically |
| `Insert` | `Insert` | Inserts at a zero-based position |
| `Remove` / `RemoveAt` | `Remove` / `RemoveAt` | Removes by value or index |
| `Contains` / `IndexOf` | `Contains` / `IndexOf` | Equality, identity, or custom comparer |
| `Count` / `Clear` | `Count` / `Clear` | Total count, predicate count, and mutation |
| `foreach` | `For Each` | Supports independent nested enumeration |
| `List<T>.ForEach` | `ForEach(Action)` | Runs one typed Action per element |
| `String.Join` | `JoinText(separator)` | Joins scalar elements without a helper loop |
| `ToArray` | `ToArray` | Returns a zero-based Variant array |

## Deferred LINQ

Sequence-returning operators create immutable query nodes. They do not enumerate until a terminal,
indexer, `Count`, or `For Each` consumes the query. A query therefore observes later source-list
mutations, matching LINQ's deferred-execution model.

```vba
Dim filtered As ROneCOne
Dim numbers As ROneCOne
Dim x As ROneCOne

Set numbers = ROneCOne.ListOf(vbLong, CLng(5), CLng(20))

Set x = numbers.Element
Set filtered = numbers.Where(x.GreaterThan(CLng(10)))

numbers.Add CLng(30)
Set filtered = filtered.ToList

Debug.Print filtered.Count   ' 2
Debug.Print filtered.Last    ' 30
```

The sequence operators include `Where`, `SelectItems`, `Take`, `Skip`, `Distinct`, `DistinctBy`,
`Order`, `OrderDescending`, `OrderBy`, `OrderByDescending`, `ThenBy`, `ThenByDescending`, `Append`,
`Prepend`, `Reverse`, `Concat`, `Union`, `Intersect`, `Except`, their `...By` forms, `SelectMany`,
`Cast`, `OfType`, `Chunk`, `TakeWhile`, `SkipWhile`, `TakeLast`, and `SkipLast`. `Range`, `Repeat`,
and `EmptyOf` create typed source sequences. Immediate terminals
are `Exists`, `AnyItem`, `All`, `None`, `First`,
`FirstOrDefault`, `Last`, `LastOrDefault`, `SingleItem`, `SingleOrDefault`, `Sum`, `Average`,
`Min`, `Max`, `MinBy`, `MaxBy`, `Count`, `ForEach`, `JoinText`, `SequenceEqual`, `ToList`, and
`ToArray`, `ToDictionary`, `ToHashSet`, `ToLookup`, `GroupBy`, `Join`, `GroupJoin`, `Zip`,
`CountBy`, `Aggregate`, `AggregateBy`, `Index`, and `TryGetNonEnumeratedCount`.

```vba
Set result = ROneCOne.Range(CLng(1), CLng(6)) _
    .Where(x.Modulo(CLng(2)).EqualTo(CLng(0))) _
    .Map(x.Multiply(CLng(10)), vbLong) _
    .OrderDescending _
    .Take(CLng(2)) _
    .ToList

Debug.Print result(0)  ' 60
Debug.Print result(1)  ' 40
```

## Concise syntax and canonical core

The concise surface removes repetition after the underlying capability has a passing behavioral
contract. The canonical members remain public for explicit construction, teaching, and debugging.

| Concise form | Canonical expansion | Meaning |
|---|---|---|
| `Set element = values.Element` | `Set element = ROneCOne.Parameter(vbLong)` | Parameter typed from the sequence |
| `values.Where(x.AtLeast(10))` | `values.Where(ROneCOne.Lambda(x.GreaterThanOrEqual(10), x))` | Inferred unary predicate |
| `values.Where("Age").AtLeast(40)` | `values.Where(values.Condition("Age").AtLeast(40))` | Contextual member filter |
| `values.Where("City").IsIn(cities)` | `values.Where(values.Condition("City").IsIn(cities))` | Collection membership |
| `cities.Contains(values!City)` | `values.Condition("City").IsIn(cities)` | C#-shaped membership expression |
| `values.Where("Manager?.Age")` | `values.Condition("Manager?.Age")` | Null-safe nested access |
| `values!Age` | `values.Condition("Age")` | Native VBA default-member expression |
| `values.Map(x.Multiply(2), vbLong)` | `values.SelectItems(ROneCOne.Lambda(x.Multiply(2), x), vbLong)` | Typed projection |
| `values.Map("Name", vbString)` | `values.Map(values.Condition("Name"), vbString)` | Member-name projection |
| `values.OrderBy("Age")` | `values.OrderBy(values.Condition("Age"))` | Member-name key selector |
| `values.ThenBy("Name")` | `values.ThenBy(values.Condition("Name"))` | Adds a secondary key |
| `values.WhereMethod("Queries.IsActive")` | `values.Where(values.Predicate("Queries.IsActive"))` | Inferred `Func<T, Boolean>` |
| `values.Exists(predicate)` | `values.AnyItem(predicate)` | Predicate-based existence test |
| `values.Exists` | `values.AnyItem` | Tests whether the sequence has an element |
| `values.None(predicate)` | `Not values.AnyItem(predicate)` | Tests that no element matches |
| `values.Count(predicate)` | `values.Where(predicate).Count` | Counts matching elements |
| `values.WhereAny("Items", p)` | `values.Where(values.Condition("Items", True).AnyMatch(p))` | Nested Any |
| `values.ForEach(action)` | explicit `For Each` plus `action.Execute` | Typed side effects |
| `values.JoinText("|")` | explicit string-building loop | Scalar text joining |
| `values.Order` | `values.OrderBy(values.Element)` | Identity-key ascending order |
| `values.OrderDescending` | `values.OrderByDescending(values.Element)` | Identity-key descending order |

`Where`, `Map`, ordering, quantifiers, element terminals, and selector-based numeric terminals all
accept either any universal unary `Func` or an expression containing exactly one parameter. Member
selectors additionally accept a member-path string. That includes object, callable-object,
workbook-procedure, native, composition, and multicast delegates. The runtime infers expression
parameters once when the query is built, not once per element.

## Stable composite ordering

The ordering surface mirrors .NET's distinction between primary and secondary keys. `Order` and
`OrderDescending` compare each element directly. `OrderBy` and `OrderByDescending` select a new
primary key. `ThenBy` and `ThenByDescending` extend only the immediately active ordered query:

```vba
Set ranked = customers _
    .OrderBy("City") _
    .ThenByDescending("Age") _
    .ThenBy("CustomerName") _
    .ToList
```

Starting another `Order` or `OrderBy` replaces the active ordering chain. An intervening operator
such as `Where`, `Map`, or `Take` ends the ordered-query capability, so a following `ThenBy` raises
`InvalidOperationError`. This makes an invalid chain fail at construction instead of quietly
changing its meaning.

Ordering is deferred and stable. Every key selector runs exactly once per element per enumeration,
and a stable O(n log n) merge sort preserves source order when all keys compare equal. Null scalar
keys come first in ascending order and last in descending order. Boolean order is `False`, then
`True`; default string order is binary and ordinal. Incompatible mixed Variant key types raise
`TypeMismatchError` instead of using VBA's implicit coercion. User-defined objects can be ordered
directly only when an explicit comparer is supplied.

## Contextual conditions

`Where("Member")` starts a deferred condition builder. The comparison completes the query, so
there is no adapter class, explicit element variable, lambda wrapper, or predicate string language.

```vba
Set adults = customers.Where("Age").AtLeast(CLng(18))
Set selected = customers _
    .Where("Age").Between(CLng(18), CLng(65)) _
    .Where("City").OneOf("London", "Paris")
Set matching = customers.Where("Name").StartsWith("Gr")
```

The fluent condition vocabulary is `EqualTo`, `NotEqualTo`, `LessThan`, `AtMost`, `GreaterThan`,
`AtLeast`, inclusive or exclusive `Between`, `OneOf`, `IsIn`, `NotIn`, `StartsWith`, `EndsWith`,
`Contains`, `ContainsText`, `MatchesPattern`, `IsNothing`, `IsNotNothing`, `IsNullOrEmpty`,
`IsTrue`, and `IsFalse`. `OneOf` and `IsIn` accept ROneCOne sequences, VBA arrays, and
`Collection` values. String comparisons default to binary comparison; `EqualToIgnoreCase`,
`NotEqualToIgnoreCase`, `StartsWithIgnoreCase`, `EndsWithIgnoreCase`, `ContainsIgnoreCase`, and
`MatchesPatternIgnoreCase` provide the common text forms directly. Repeated `Where` calls are
logical AND and retain deferred execution.

Use `Condition` when a predicate combines multiple members:

```vba
Set predicate = customers.Condition("Age").AtLeast(CLng(40)) _
    .Both(customers.Condition("City").EqualTo("London"))
Set selected = customers.Where(predicate)
```

`Both`, `Either`, and `Negated` are concise aliases for the canonical `AndAlso`, `OrElse`, and
`NotExpression` nodes. `WhereNot(predicate)` filters the inverse directly. `Always`, `Never`,
`Match(member, value)`, and `NotMatch(member, value)` create reusable parameterized predicates.

`Element` is stable for each sequence, so separately created `Condition` expressions share one
typed parameter and compose as a unary predicate. Dotted paths traverse object-valued intermediate
members. The `?.` path operator propagates `Null` when an intermediate object is `Nothing`:

```vba
Set selected = customers.Where("Manager?.Age").AtLeast(CLng(40))
```

Ordinary `.` access still raises `MemberAccessError` on `Nothing`. Null-safe relational and string
conditions evaluate False, equality follows deterministic Null equality, and `IsNullOrEmpty`
evaluates True for propagated scalar Null. Once `?.` encounters `Nothing`, every remaining segment
in that path short-circuits, matching C# null-conditional chaining.

`Predicate` and `WhereMethod` infer the input descriptor directly from `List<T>`:

```vba
Set isExperienced = customers.Predicate("Queries.IsExperienced")
Debug.Print isExperienced.Signature  ' Func<Customer, Boolean>
Set selected = customers.WhereMethod("Queries.IsExperienced")
```

An object method uses `customers.Predicate(target, "IsExperienced")` or
`customers.WhereMethod(target, "IsExperienced")`. Existing typed delegates remain valid.

## Collection membership predicates

Membership is an ordinary expression node, so it can sit anywhere a predicate can. The clearest
contextual form accepts a typed sequence, VBA array, or `Collection`:

```vba
Set selected = customers.Where("City").IsIn(allowedCities)
Set selected = customers.Where("City").IsIn(Array("London", "Paris"))
```

The form closest to C# reverses the receiver and passes the current member expression to the
collection:

```vba
Set selected = customers.Where(allowedCities.Contains(customers!City))
```

`List.Contains(literal)` retains normal immediate containment semantics. Passing a ROneCOne value,
parameter, or expression returns a composable predicate instead. `ContainsMember(source,
memberPath)` provides an explicit reusable form when bang syntax is not desirable.

## Predicate terminals and nested collections

`Count`, `FirstOrDefault`, `LastOrDefault`, `SingleItem`, `SingleOrDefault`, and `None` accept the
same expression, member context, or universal `Func` predicates as `Where`. Defaults follow
`default(T)`: `Nothing` for object types, an empty string for String, False for Boolean, zero for
numeric types, and Empty for Variant.

For a member that returns a ROneCOne sequence, use `AnyMatch`, `AllMatch`, or `NoneMatch` on the
object-valued condition. `WhereAny`, `WhereAll`, and `WhereNone` provide the primary parent-query
form:

```vba
Set selected = customers.WhereAny("Reports", reportPredicate)
Set selected = customers.WhereAll("Reports", reportPredicate)
Set selected = customers.WhereNone("Reports", reportPredicate)
```

As in LINQ, `AllMatch` is True for an empty nested sequence, while `AnyMatch` is False and
`NoneMatch` is True. A `Nothing` nested sequence follows the same empty-sequence behavior.

## Equality and ordering comparers

`EqualityComparer(callable)` builds a `Func<T, T, Boolean>` and `Comparer(callable)` builds a
`Func<T, T, Long>`. Both accept the universal object, callable-object, workbook-procedure, native,
or expression delegate surface. Equality comparers are accepted by `Contains`, `IndexOf`,
`Distinct`, `DistinctBy`, and `SequenceEqual`. Ordering comparers are accepted independently by
`Order`, `OrderDescending`, `OrderBy`, `OrderByDescending`, `ThenBy`, `ThenByDescending`, `Min`,
`Max`, `MinBy`, and `MaxBy`.

```vba
Set equality = ROneCOne.EqualityComparer("Queries.TextEqualsIgnoreCase")
Set ordering = ROneCOne.Comparer("Queries.CompareTextIgnoreCase")

Set unique = names.Distinct(equality)
Set ordered = names.Order(ordering)
Debug.Print names.Contains("ada", equality)
```

## LINQ over user-defined classes

The collections demo includes a normal `DemoCustomer` class with `CustomerName`, `Age`, `City`,
`Manager`, and nested `Reports` properties. No predicate adapter class is required. Contextual
member names are the primary readable surface.

```vba
Dim customers As ROneCOne
Dim experienced As ROneCOne
Dim names As ROneCOne
Dim firstCustomer As DemoCustomer
Set customers = ROneCOne.ListFrom(ada, grace, katherine)

Set experienced = customers.Where("Age").AtLeast(CLng(40))
Set names = experienced _
    .Map("CustomerName", vbString) _
    .Order _
    .ToList
Set firstCustomer = customers _
    .OrderBy("City") _
    .ThenByDescending("Age") _
    .First

Debug.Print names.GenericTypeName  ' List<String>
Debug.Print firstCustomer.CustomerName
```

The same query expressed through canonical primitives is:

```vba
Dim age As ROneCOne
Dim customer As ROneCOne
Dim predicate As ROneCOne

Set customer = ROneCOne.ParameterLike(ada)
Set age = customer.Member("Age")
Set predicate = ROneCOne.Lambda( _
    age.GreaterThanOrEqual(CLng(40)), customer)
Set experienced = customers.Where(predicate)
```

VBA has no `nameof` operator or first-class property reference, so ordinary contextual selectors
use a member name. VBA's existing bang operator provides an optional identifier-shaped form because
ROneCOne's default member maps a sequence member to `Condition`:

```vba
Set experienced = customers.Where(customers!Age.AtLeast(CLng(40)))

With customers
    Set experienced = .Where(!Age.AtLeast(CLng(40)))
End With
```

This is native VBA syntax, not source rewriting. The string form remains the primary surface
because it is clearer to developers outside Access-style VBA and naturally supports dotted paths.
The canonical scalar forms are `customer("Age")` and `customer.Member("Age")`. For an object-valued
property, use `customer.Member("Manager", True)`. Missing or invalid access raises the deterministic
`ROneCOne.MemberAccessError`.

The workbook proves seventeen object-oriented cases, including deferred contextual filtering,
membership expressions, null-safe paths, predicate terminals, nested quantifiers, custom
comparers, composition, bang syntax, and key operators. The customer model is demo application
code; the deployed runtime remains the single `ROneCOne.cls` file.

## VBA-constrained API names

VBA reserves `Select`, `Any`, `In`, and `Single`, and the VBE rejects those words as class member
declarations. ROneCOne therefore uses `Map`, `Exists`, `IsIn`, and `SingleItem`; `SelectItems` and
`AnyItem` remain the explicit core members. Legal .NET names are used directly throughout the
ordering surface.

The supported surface includes dictionaries, lookups, groupings, joins, sets, queues, stacks,
specialized collections, immutable collections, and concurrent-style collections. Each family
uses the same runtime type contracts and one-file implementation as `List<T>`.
