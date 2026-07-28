# ADR 0027: LINQ to SQL over the existing expression trees

Status: accepted, 2026-07-28

## Context

The runtime already owned the two halves nobody else in VBA has: real expression trees, complete
with a renderer that prints them as readable pseudocode, and late-bound OLE DB and ODBC providers
with typed parameters. They had never been joined. Every database read went out as a hand-built
SQL string, which meant string concatenation at the call site, no protection against a value that
contains a quote, and whole tables pulled into Excel so that a `Where` could discard most of the
rows in memory. For the data-heavy workbooks this library exists to serve, that last point is the
dominant cost.

A live probe settled the dialect questions before any runtime code was written, because the two
providers the suite can reach disagree in ways that matter. SQL Server through MSOLEDBSQL accepts
a parameterized `TOP (?)`, `OFFSET`/`FETCH`, and `LIKE ... ESCAPE '\'`, and it uses `%` for
modulo, `+` for concatenation, and `UPPER`. ACE accepts only a literal `TOP n`, rejects both
`OFFSET` and `ESCAPE`, escapes wildcards by bracketing them (`[%]`, `[_]`), and uses `MOD`, `&`,
and `UCASE`. Guessing either way would have produced a layer that worked on one machine.

## Decision

`connection.Queryable(tableName)` returns a deferred query that compiles the runtime's own
expression trees into parameterized SQL. `Where`, `OrderBy`, `OrderByDescending`, `ThenBy`,
`ThenByDescending`, `SelectColumns`, `Take`, and `Skip` each return a new query rather than
mutating the receiver, so a base query can be shared and specialized. `Count`, `AnyItem`,
`FirstOrDefault`, `ToDataTable`, `ToDataTableAsync`, and `CountAsync` execute. `ToSqlString` and
`SqlParameterValues` render exactly what will be sent, so the statement is inspectable before it
runs and in a debugger afterwards.

Every captured constant leaves as a `?` marker bound through the existing typed parameter path, so
a value can never be read as SQL. A `Null` constant becomes `IS NULL` or `IS NOT NULL` rather than
a comparison that silently matches nothing. `StartsWith`, `EndsWith`, and `ContainsText` become
`LIKE` with the pattern's own wildcards escaped for the target dialect. `IsIn` becomes an `IN`
list of markers. Identifiers are refused if they contain brackets or control characters, and are
then bracket-quoted.

Anything the translator cannot express refuses with the typed `ROneCOne.QueryError`: a nested
member path that would need a join, a custom comparer, a `Like` pattern, a nested quantifier, a
`Skip` on ACE, a `Skip` without an `OrderBy`, or a connection string naming neither supported
dialect. Refusal is the point. A silent fall back to a client-side scan would produce a query that
looks server-side and quietly loads the table.

## Consequences

One new role, one new error number (`ROneCOne.QueryError`), and one new task operation for the
async fill. No new prog-id, `Declare`, or reference: execution rides `DbCommand`, `WithParameter`,
and `DbDataAdapter` unchanged, so the provider's own failure detail and cancellation behavior come
along for free.

The live suite gained thirty-nine assertions driving both dialects from the same six-row shape,
including a customer name of `x' OR '1'='1` that matches one row literally and cannot widen the
result set. Two live findings are recorded here because neither is obvious. A parameter named
`name` in the identifier validator silently lowercased every bare `Name` in the project, since VBA
keeps one global casing per identifier; `rows!Name` then compiled to `rows!name` and the
case-sensitive column lookup stopped matching, breaking a DataTable test unrelated to SQL. A
contract test now refuses any bang-reachable identifier declared in lower case. Separately, a null
comparison takes the plain VBA `Null`, because `ROneCOne.DBNull` is a cell value rather than an
expression operand.

Joins, grouping, and aggregate projections beyond `Count` are deliberately out of scope for this
change. They need a shape the expression tree does not yet carry, and adding them halfway would
blur the line between what the server does and what Excel does.
