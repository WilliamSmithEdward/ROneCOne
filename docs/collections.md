# Runtime-generic List<T> and LINQ

ROneCOne provides a strictly typed, zero-based `List<T>` model and deferred query pipelines in
the same single `ROneCOne.cls` runtime file. The design follows the behavior of
[`List<T>`](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1) and
[`Enumerable`](https://learn.microsoft.com/en-us/dotnet/api/system.linq.enumerable) where VBA's
syntax and type system permit it.

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
| `Contains` / `IndexOf` | `Contains` / `IndexOf` | Scalar equality or object identity |
| `Count` / `Clear` | `Count` / `Clear` | Immediate count and mutation |
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

The sequence operators are `Where`, `SelectItems`, `Take`, `Skip`, `Distinct`, `DistinctBy`,
`OrderBy`, `OrderByDescending`, `Append`, `Prepend`, and `Reverse`. `Range` and `Repeat` create typed
source sequences. Immediate terminals are `Exists`, `AnyItem`, `All`, `First`, `Last`, `Sum`,
`Average`, `Min`, `Max`, `MinBy`, `MaxBy`, `Count`, `ForEach`, `JoinText`, `ToList`, and `ToArray`.

```vba
Set result = ROneCOne.Range(CLng(1), CLng(6)) _
    .Where(x.Modulo(CLng(2)).EqualTo(CLng(0))) _
    .Map(x.Multiply(CLng(10)), vbLong) _
    .SortedDescending _
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
| `values!Age` | `values.Condition("Age")` | Native VBA default-member expression |
| `values.Map(x.Multiply(2), vbLong)` | `values.SelectItems(ROneCOne.Lambda(x.Multiply(2), x), vbLong)` | Typed projection |
| `values.Map("Name", vbString)` | `values.Map(values.Condition("Name"), vbString)` | Member-name projection |
| `values.OrderBy("Age")` | `values.OrderBy(values.Condition("Age"))` | Member-name key selector |
| `values.WhereMethod("Queries.IsActive")` | `values.Where(values.Predicate("Queries.IsActive"))` | Inferred `Func<T, Boolean>` |
| `values.Exists(predicate)` | `values.AnyItem(predicate)` | Predicate-based existence test |
| `values.Exists` | `values.AnyItem` | Tests whether the sequence has an element |
| `values.ForEach(action)` | explicit `For Each` plus `action.Execute` | Typed side effects |
| `values.JoinText("|")` | explicit string-building loop | Scalar text joining |
| `values.Sorted` | `values.OrderBy(values.Element)` | Identity-key ascending order |
| `values.SortedDescending` | `values.OrderByDescending(values.Element)` | Identity-key descending order |

`Where`, `Map`, ordering, quantifiers, element terminals, and selector-based numeric terminals all
accept either any universal unary `Func` or an expression containing exactly one parameter. Member
selectors additionally accept a member-path string. That includes object, callable-object,
workbook-procedure, native, composition, and multicast delegates. The runtime infers expression
parameters once when the query is built, not once per element.

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
`AtLeast`, inclusive or exclusive `Between`, `OneOf`, `StartsWith`, `EndsWith`, `Contains`,
`ContainsText`, `MatchesPattern`, `IsNothing`, `IsNotNothing`, `IsNullOrEmpty`, `IsTrue`, and
`IsFalse`. String comparisons default to binary comparison and accept `vbTextCompare` explicitly.
Repeated `Where` calls are logical AND and retain deferred execution.

Use `Condition` when a predicate combines multiple members:

```vba
Set predicate = customers.Condition("Age").AtLeast(CLng(40)) _
    .AndAlso(customers.Condition("City").EqualTo("London"))
Set selected = customers.Where(predicate)
```

`Element` is stable for each sequence, so separately created `Condition` expressions share one
typed parameter and compose as a unary predicate. Dotted paths traverse object-valued intermediate
members. Guard a nullable reference before consuming a deeper path, matching C# null-safety:

```vba
Set selected = customers _
    .Where("Manager").IsNotNothing _
    .Where("Manager.Age").AtLeast(CLng(40))
```

`Predicate` and `WhereMethod` infer the input descriptor directly from `List<T>`:

```vba
Set isExperienced = customers.Predicate("Queries.IsExperienced")
Debug.Print isExperienced.Signature  ' Func<Customer, Boolean>
Set selected = customers.WhereMethod("Queries.IsExperienced")
```

An object method uses `customers.Predicate(target, "IsExperienced")` or
`customers.WhereMethod(target, "IsExperienced")`. Existing typed delegates remain valid.

## LINQ over user-defined classes

The collections demo includes a normal `DemoCustomer` class with `CustomerName`, `Age`, `City`, and
`Manager` properties. No predicate adapter class is required. Contextual member names are the
primary readable surface.

```vba
Dim customers As ROneCOne
Dim experienced As ROneCOne
Dim names As ROneCOne
Dim oldest As DemoCustomer
Set customers = ROneCOne.ListFrom(ada, grace, katherine)

Set experienced = customers.Where("Age").AtLeast(CLng(40))
Set names = experienced _
    .Map("CustomerName", vbString) _
    .Sorted _
    .ToList
Set oldest = customers.OrderByDescending("Age").First

Debug.Print names.GenericTypeName  ' List<String>
Debug.Print oldest.CustomerName
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

The workbook proves eleven object-oriented cases, including deferred contextual filtering,
member-name projection and ordering, composed predicates, inferred procedure predicates, string
conditions, key distinctness, bang syntax, and safe object paths. The customer model is demo
application code; the deployed runtime remains the single `ROneCOne.cls` file.

## VBA-constrained API names

VBA reserves `Select` and `Any`, and the VBE rejects those words as class member declarations.
ROneCOne therefore uses `Map` and `Exists` as its concise names and exposes `SelectItems` and
`AnyItem` as the explicit core members. `AtLeast`, `AtMost`, `Sorted`, and
`SortedDescending` similarly expand to the longer comparison and ordering primitives without
changing their behavior.

The feature roadmap covers the standard LINQ and generic-collection surface. Operators that
need new result abstractions, including dictionaries, lookups, groupings, joins, sets, queues, and
stacks, enter the supported surface as independently tested release slices so `List<T>` contracts
remain stable.
