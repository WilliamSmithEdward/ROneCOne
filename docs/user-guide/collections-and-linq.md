# Collections and LINQ

ROneCOne lets VBA work with a collection as a typed sequence instead of a pile of loops, counters,
temporary arrays, and helper procedures. Queries are built in readable stages and evaluated when
you ask for a result.

## The high-impact example

Assume `ada`, `grace`, and `katherine` are ordinary instances of your existing `Customer` class.
ROneCOne can infer that concrete class, filter its properties, project names, and sort the result:

```vba
Dim customers As ROneCOne
Dim names As ROneCOne

Set customers = ROneCOne.ListFrom(ada, grace, katherine)

Set names = customers _
    .Where("Age").AtLeast(40) _
    .Map("CustomerName", vbString) _
    .Order _
    .ToList
```

`customers` behaves as `List<Customer>`. `names` behaves as `List<String>`. No adapter class is
needed, and the original `Customer` class does not need to know about ROneCOne.

## Create a typed list

Use a VBA type token for primitive values:

```vba
Dim scores As ROneCOne

Set scores = ROneCOne.ListOf(vbLong, 75, 90, 82)

scores.Add 95
Debug.Print scores(0)
Debug.Print scores.Count
```

Plain numeric literals are enough. A typed list accepts any number that fits its declared type
without loss and stores it as that type, so an integer literal drops straight into a
`List<Long>` or `List<Double>`. A value that would lose information is still refused before the
list changes.

Use `ListFrom` when real objects are already available:

```vba
Dim customers As ROneCOne

Set customers = ROneCOne.ListFrom(ada, grace, katherine)
Debug.Print customers.GenericTypeName
```

ROneCOne enforces the list's element type before a mutation changes the list.

## Use indexed dictionaries and sets

```vba
Dim reserved As Long
Dim scoresById As ROneCOne

Set scoresById = ROneCOne.DictionaryOf(vbLong, vbLong)
reserved = scoresById.EnsureCapacity(10000)
scoresById.Add 101, 95

Debug.Print scoresById.Item(101)
```

Default-equality dictionaries and hash sets use indexed lookup. `EnsureCapacity` is useful before
a bulk load; `Capacity` reports the current reservation and `TrimExcess` compacts it afterward.
Custom equality comparers retain their exact delegated semantics through a deliberate linear path.

## Filter by an object property

The member name begins a condition, and the comparison completes the query:

```vba
Dim experienced As ROneCOne

Set experienced = customers _
    .Where("Age").AtLeast(40) _
    .ToList
```

Common conditions read the same way:

```vba
Set selected = customers _
    .Where("Age").Between(18, 65) _
    .Where("City").OneOf("London", "Paris") _
    .ToList

Set matching = customers _
    .Where("CustomerName").StartsWithIgnoreCase("gr") _
    .ToList
```

## Filter by another collection

The direct contextual form is the easiest to read:

```vba
Dim allowedCities As ROneCOne
Dim selected As ROneCOne

Set allowedCities = ROneCOne.ListOf( _
    vbString, "London", "Cleveland")

Set selected = customers _
    .Where("City").IsIn(allowedCities) _
    .ToList
```

The receiver can also follow the C# shape:

```vba
Set selected = customers.Where( _
    allowedCities.Contains(customers!City)).ToList
```

Passing a literal to `Contains` performs an immediate lookup. Passing a ROneCOne expression builds
a reusable membership predicate.

## Navigate nullable relationships

Use `?.` when an intermediate object might be `Nothing`:

```vba
Dim managed As ROneCOne

Set managed = customers _
    .Where("Manager?.Age").AtLeast(40) _
    .ToList
```

If `Manager` is `Nothing`, the remaining path short-circuits. The query continues without a member
access error.

## Combine conditions

Build named conditions when a rule spans more than one property:

```vba
Dim predicate As ROneCOne
Dim selected As ROneCOne

Set predicate = customers.Condition("Age").AtLeast(40) _
    .Both(customers.Match("City", "London"))

Set selected = customers.Where(predicate).ToList
```

Use `Both` for AND, `Either` for OR, and `Negated` for NOT. `WhereNot` filters the inverse directly.

## Query nested collections

If each customer exposes a `Reports` property containing another ROneCOne sequence, parent records
can be filtered by their children:

```vba
Dim experiencedReport As ROneCOne
Dim managers As ROneCOne

Set experiencedReport = ada.Reports _
    .Condition("Age").AtLeast(40)

Set managers = customers _
    .WhereAny("Reports", experiencedReport) _
    .ToList
```

The matching forms are `WhereAny`, `WhereAll`, and `WhereNone`.

## Shape and order a result

`Map` is ROneCOne's concise projection operator because `Select` is reserved by VBA:

```vba
Dim names As ROneCOne

Set names = customers _
    .OrderByDescending("Age") _
    .Map("CustomerName", vbString) _
    .Distinct _
    .ToList

Debug.Print names.JoinText(", ")
```

Build a stable multi-key order by starting with `OrderBy` and extending it with `ThenBy`:

```vba
Dim ranked As ROneCOne

Set ranked = customers _
    .OrderBy("City") _
    .ThenByDescending("Age") _
    .ThenBy("CustomerName") _
    .ToList
```

Use `Order` or `OrderDescending` when each element is already the value to compare. Every ordering
method can receive its own comparer.

Queries can also use `Take`, `Skip`, `Append`, `Prepend`, and `Reverse`.

## Ask for one answer

Terminals return a value instead of another query:

```vba
Dim oldest As Customer

Set oldest = customers.OrderByDescending("Age").First

Debug.Print customers.Count( _
    customers.Condition("Age").AtLeast(40))
Debug.Print customers.Exists(customers.Match("City", "London"))
Debug.Print customers.None(customers.Match("Age", 100))
```

Use `FirstOrDefault`, `LastOrDefault`, or `SingleOrDefault` when no match is a normal outcome.

## Where next

- [Delegates and expressions](delegates-and-expressions.md) continues the learning path.
- [Practical reference](reference.md) lists the everyday operators.
- [Collections technical reference](../collections.md) covers exact typing, deferred execution,
  custom comparers, canonical expansions, and edge-case semantics.
- [Guide index](README.md) returns to the full learning path.
