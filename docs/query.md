# Queryable

Exact semantics for translating expression trees into SQL. For the workflow-first introduction,
read [Data and providers](user-guide/data-and-providers.md).

## Surface

`connection.Queryable(tableName)` returns a deferred query over one table of an open or lazy
connection. Every composing member returns a new query, so a base query can be shared safely.

| Member | Behavior |
|---|---|
| `Where(predicate)` | Adds a condition, joined to the others with `AND` |
| `Where("Column")` | Begins a fluent condition (`.AtLeast`, `.EqualTo`, `.ContainsText`) |
| `OrderBy` / `OrderByDescending` | Sets the ordering, replacing any previous one |
| `ThenBy` / `ThenByDescending` | Adds a secondary ordering; requires `OrderBy` first |
| `SelectColumns(names...)` | Projects onto the named columns instead of `*` |
| `Take(count)` / `Skip(count)` | Limits and offsets the rows |
| `Count` / `AnyItem` | Runs a server-side `COUNT`, optionally with a predicate |
| `FirstOrDefault` | Returns the first `DataRow`, or `Nothing` when there is none |
| `ToDataTable` / `ToDataTableAsync` | Runs the query and lands the rows in a typed DataTable |
| `CountAsync` | Starts the count natively async; the Task resolves to the number |
| `ToSqlString` / `SqlParameterValues` | Renders the statement and its bound values |

```vba
Dim recent As ROneCOne

Set recent = connection.Queryable("Orders") _
    .Where("Total").AtLeast(100) _
    .OrderByDescending("Placed") _
    .Take(50) _
    .ToDataTable
```

## Values are always parameters

Every captured constant leaves as a `?` marker bound through the provider's typed parameter path.
A value can never be read as SQL, so an injection-shaped string matches literally:

```vba
' Finds the one customer actually named this. It cannot widen the result set.
connection.Queryable("Orders").Where("Customer").EqualTo("x' OR '1'='1").Count
```

`ToSqlString` shows the statement with its markers and `SqlParameterValues` lists the values in
binding order, so what will be sent is inspectable before it runs.

## Translation rules

| Expression | SQL |
|---|---|
| `EqualTo(Null)` / `NotEqualTo(Null)` | `IS NULL` / `IS NOT NULL` |
| `StartsWith` / `EndsWith` / `ContainsText` | `LIKE` with the text's own wildcards escaped |
| `EqualToIgnoreCase` | `UPPER(...) = UPPER(?)`, or `UCASE` on ACE |
| `OneOf` / `IsIn` | `IN (?, ?, ...)` |
| `Between` | The two comparisons it already expands to |

Use the plain VBA `Null` for a null comparison. `ROneCOne.DBNull` is a DataTable cell value, not
an expression operand, and the expression layer refuses it.

## Dialects

Two dialects are recognized from the connection string, and the differences are real rather than
cosmetic.

| | SQL Server | ACE (Excel and Access) |
|---|---|---|
| Row limit | `TOP (?)`, parameterized | `TOP n`, literal only |
| Offset | `OFFSET ... FETCH NEXT` | not available, so `Skip` refuses |
| LIKE escaping | `ESCAPE '\'` | bracketed (`[%]`, `[_]`) |
| Uppercase | `UPPER` | `UCASE` |
| Modulo, concatenation | `%`, `+` | `MOD`, `&` |

A connection string naming neither dialect raises `ROneCOne.QueryError` at `Queryable`.

## Refusal over fallback

Anything the translator cannot express raises the typed `ROneCOne.QueryError` rather than quietly
running client-side. That includes a nested member path (it would need a join), a custom comparer,
a `Like` pattern, a nested quantifier, `Skip` on ACE, `Skip` without an `OrderBy`, and an
identifier containing brackets or control characters. A silent fallback would produce a query that
looks server-side while loading the whole table.

Joins, grouping, and aggregates beyond `Count` are not translated. Compose those in SQL with
[`DbCommand`](data.md), or land the rows and use the in-memory
[LINQ surface](collections.md).

See [ADR 0027](decisions/0027-linq-to-sql-over-expression-trees.md) for the reasoning.

[Back to the documentation index](README.md)
