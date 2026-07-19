# Runtime-generic List<T> and LINQ

ROneCOne v0.2.0 adds a strictly typed, zero-based `List<T>` model and deferred query pipelines to
the same single `ROneCOne.cls` runtime file. The design follows the behavior of
[`List<T>`](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1) and
[`Enumerable`](https://learn.microsoft.com/en-us/dotnet/api/system.linq.enumerable) where VBA's
syntax and type system permit it.

## Primitive lists

Use a `VbVarType` token as `T`. Values must have that exact VBA runtime type; ROneCOne does not
silently coerce them.

```vba
Dim numbers As ROneCOne

Set numbers = ROneCOne.ListOf(vbLong)
numbers.Add CLng(5)
numbers.Add CLng(10)

Debug.Print numbers.GenericTypeName  ' List<Long>
Debug.Print numbers(0)               ' 5: default zero-based indexer
Debug.Print numbers.Item(1)          ' 10: explicit indexer
```

An invalid mutation raises `ROneCOne.TypeMismatchError` before changing the list.

## User-defined class lists

Pass one non-`Nothing` prototype to capture the exact concrete class name. The prototype is not
retained. Members may contain instances of that class or `Nothing`; unrelated object types are
rejected.

```vba
Dim prototype As Customer
Dim ada As Customer
Dim customers As ROneCOne

Set prototype = New Customer
Set customers = ROneCOne.ListOf(prototype)
Set prototype = Nothing

Set ada = New Customer
ada.CustomerName = "Ada"
customers.Add ada

Debug.Print customers.GenericTypeName       ' List<Customer>
Debug.Print customers.Item(0).CustomerName  ' Ada
```

This prototype token provides dependency-free `List<Customer>` semantics without code generation,
VBIDE trust, or a second runtime class.

## List surface

| C# idea | ROneCOne member | Behavior |
|---|---|---|
| `new List<T>()` | `ListOf(typeToken)` / `ListLike(example)` | Creates an empty strict list |
| `list[index]` | `list(index)` / `list.Item(index)` | Zero-based get; `Item` also supports set |
| `Add` / `AddRange` | `Add` / `AddRange` | Validates before mutation |
| `Insert` | `Insert` | Inserts at a zero-based position |
| `Remove` / `RemoveAt` | `Remove` / `RemoveAt` | Removes by value or index |
| `Contains` / `IndexOf` | `Contains` / `IndexOf` | Scalar equality or object identity |
| `Count` / `Clear` | `Count` / `Clear` | Immediate count and mutation |
| `foreach` | `For Each` | Supports independent nested enumeration |
| `ToArray` | `ToArray` | Returns a zero-based Variant array |

## Deferred LINQ

Sequence-returning operators create immutable query nodes. They do not enumerate until a terminal,
indexer, `Count`, or `For Each` consumes the query. A query therefore observes later source-list
mutations, matching LINQ's deferred-execution model.

```vba
Dim filtered As ROneCOne
Dim numbers As ROneCOne
Dim x As ROneCOne

Set numbers = ROneCOne.ListOf(vbLong)
numbers.Add CLng(5)
numbers.Add CLng(20)

Set x = ROneCOne.Parameter(vbLong)
Set filtered = numbers.Where( _
    ROneCOne.Lambda(x.GreaterThan(CLng(10)), x))

numbers.Add CLng(30)
Set filtered = filtered.ToList

Debug.Print filtered.Count   ' 2
Debug.Print filtered.Last    ' 30
```

The v0.2.0 sequence operators are `Where`, `SelectItems`, `Take`, `Skip`, `Distinct`, `OrderBy`,
`OrderByDescending`, `Append`, `Prepend`, and `Reverse`. `Range` and `Repeat` create typed source
sequences. Immediate terminals are `AnyItem`, `All`, `First`, `Last`, `Sum`, `Average`, `Min`,
`Max`, `Count`, `ToList`, and `ToArray`.

```vba
Set result = ROneCOne.Range(CLng(1), CLng(6)) _
    .Where(ROneCOne.Lambda(x.Modulo(CLng(2)).EqualTo(CLng(0)), x)) _
    .SelectItems(ROneCOne.Lambda(x.Multiply(CLng(10)), x), vbLong) _
    .OrderByDescending(ROneCOne.Lambda(x, x)) _
    .Take(CLng(2)) _
    .ToList

Debug.Print result(0)  ' 60
Debug.Print result(1)  ' 40
```

## LINQ over user-defined classes

The collections demo includes a normal `DemoCustomer` class with `CustomerName`, `Age`, and `City`
properties. A small application-side `DemoCustomerQuery` class supplies named predicates and
selectors, which `FromMethod` adapts into the unary delegate contract used by `Where`,
`SelectItems`, ordering, and terminals.

```vba
Dim agePredicate As ROneCOne
Dim ageSelector As ROneCOne
Dim customers As ROneCOne
Dim names As ROneCOne
Dim nameSelector As ROneCOne
Dim oldest As DemoCustomer
Dim prototype As DemoCustomer
Dim query As DemoCustomerQuery

Set prototype = New DemoCustomer
Set customers = ROneCOne.ListOf(prototype)
Set query = New DemoCustomerQuery
query.MinimumAge = 40

Set agePredicate = ROneCOne.FromMethod(query, "MeetsMinimumAge", 1)
Set nameSelector = ROneCOne.FromMethod(query, "SelectName", 1)
Set ageSelector = ROneCOne.FromMethod(query, "SelectAge", 1)

Set names = customers _
    .Where(agePredicate) _
    .SelectItems(nameSelector, vbString) _
    .ToList
Set oldest = customers.OrderByDescending(ageSelector).First

Debug.Print names.GenericTypeName  ' List<String>
Debug.Print oldest.CustomerName
```

The workbook proves six object-oriented cases: exact `List<DemoCustomer>` typing, deferred
filtering after source mutation, projection to `List<String>`, ordering while preserving the
customer type, `AnyItem`/`All` predicates, and an average over a projected `List<Long>`. The helper
classes are demo application code; the deployed runtime remains the single `ROneCOne.cls` file.

## VBA-compatible API names

VBA reserves `Select` and `Any`, and the VBE rejects those words as class member declarations.
ROneCOne therefore exposes `SelectItems` and `AnyItem`. These are compile-time language limits,
not string-based aliases or runtime dispatch tricks.

The compatibility roadmap covers the standard LINQ and generic-collection surface. Operators that
need new result abstractions, including dictionaries, lookups, groupings, joins, sets, queues, and
stacks, enter the supported surface as independently tested release slices so `List<T>` contracts
remain stable.
