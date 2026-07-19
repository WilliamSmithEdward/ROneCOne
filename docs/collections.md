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

The sequence operators are `Where`, `SelectItems`, `Take`, `Skip`, `Distinct`, `OrderBy`,
`OrderByDescending`, `Append`, `Prepend`, and `Reverse`. `Range` and `Repeat` create typed source
sequences. Immediate terminals are `Exists`, `AnyItem`, `All`, `First`, `Last`, `Sum`, `Average`,
`Min`, `Max`, `Count`, `ForEach`, `JoinText`, `ToList`, and `ToArray`.

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
| `values.Map(x.Multiply(2), vbLong)` | `values.SelectItems(ROneCOne.Lambda(x.Multiply(2), x), vbLong)` | Typed projection |
| `values.Exists(predicate)` | `values.AnyItem(predicate)` | Predicate-based existence test |
| `values.Exists` | `values.AnyItem` | Tests whether the sequence has an element |
| `values.ForEach(action)` | explicit `For Each` plus `action.Execute` | Typed side effects |
| `values.JoinText("|")` | explicit string-building loop | Scalar text joining |
| `values.Sorted` | `values.OrderBy(values.Element)` | Identity-key ascending order |
| `values.SortedDescending` | `values.OrderByDescending(values.Element)` | Identity-key descending order |

`Where`, `Map`, ordering, quantifiers, element terminals, and selector-based numeric terminals all
accept either any universal unary `Func` or an expression containing exactly one parameter. That
includes object, callable-object, workbook-procedure, native, composition, and multicast delegates.
The runtime infers expression parameters once when the query is built, not once per element.

## LINQ over user-defined classes

The collections demo includes a normal `DemoCustomer` class with `CustomerName`, `Age`, and `City`
properties. No predicate adapter class is required. `Element` creates the typed customer
parameter, and its default member turns `customer("Age")` into a scalar property-access expression.

```vba
Dim customer As ROneCOne
Dim customers As ROneCOne
Dim experienced As ROneCOne
Dim names As ROneCOne
Dim oldest As DemoCustomer
Set customers = ROneCOne.ListFrom(ada, grace, katherine)
Set customer = customers.Element

Set experienced = customers.Where(customer("Age").AtLeast(CLng(40)))
Set names = experienced _
    .Map(customer("CustomerName"), vbString) _
    .Sorted _
    .ToList
Set oldest = customers.OrderByDescending(customer("Age")).First

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

VBA has no `nameof` operator or first-class property reference, so the member name is the one
irreducible string boundary. Scalar properties use `customer("Age")` or
`customer.Member("Age")`. For an object-valued property, use
`customer.Member("Manager", True)` so ROneCOne preserves object assignment semantics. A missing,
invalid, or scalar-on-non-object member access raises the deterministic
`ROneCOne.MemberAccessError`.

The workbook proves six object-oriented cases: exact `List<DemoCustomer>` typing, deferred
filtering after source mutation, projection to `List<String>`, ordering while preserving the
customer type, `Exists`/`All` predicates, and an average over a projected `List<Long>`. The customer
model is demo application code; the deployed runtime remains the single `ROneCOne.cls` file.

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
