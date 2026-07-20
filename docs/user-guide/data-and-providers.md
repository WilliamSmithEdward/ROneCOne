# Data and providers

ROneCOne provides a typed, in-memory data layer for workbook code and a late-bound provider layer
for local OLE DB or ODBC sources.

## Build a typed table

```vba
Dim people As ROneCOne
Dim row As ROneCOne

Set people = ROneCOne.DataTable("People")
people.Column("Id", vbLong).AutoNumber(100&, 10&).AsPrimaryKey
people.Column("Name", vbString).WithDefault "Unknown"
people.Column "Age", vbLong
people.Column "Note", vbString

Set row = people.Row("Ada", 47&, ROneCOne.DBNull).Add
Debug.Print people.Find(100&).Item("Name")
```

Reading the example:

- `Row(...).Add` creates and attaches a positional row and skips omitted auto-number columns.
- `AsPrimaryKey` creates the common one-column key. Set `PrimaryKey` to a typed list of columns
  when a composite key is needed.
- `Find` uses an index rather than scanning every row.
- `ROneCOne.DBNull` is the explicit database-null value.

Rows track `Detached`, `Added`, `Unchanged`, `Modified`, and `Deleted` states. `AcceptChanges`,
`RejectChanges`, and `GetChanges` follow the familiar DataTable workflow.

## Query a view

```vba
Dim view As ROneCOne

Set view = ROneCOne.DataView(people) _
    .WithFilter(people.Rows!Age.AtLeast(40&)) _
    .WithSort("Name")
```

DataView is a typed sequence, so the same LINQ terminals, enumeration, and materialization rules
apply. `DataSet` and `DataRelation` add named tables, parent/child navigation, uniqueness, and
foreign-key validation.

## Load provider data

```vba
Dim adapter As ROneCOne
Dim command As ROneCOne
Dim connection As ROneCOne
Dim table As ROneCOne

Set connection = ROneCOne.DbConnection(connectionString)
connection.Connect
Set command = ROneCOne.DbCommand(sql, connection).WithTimeout(30&)
Set adapter = ROneCOne.DbDataAdapter(command)
Set table = ROneCOne.DataTable("Results")

Debug.Print adapter.FillAsync(table).Await
connection.Disconnect
```

> [!NOTE]
> VBA reserves `Open` and `Close` as language keywords, so direct calls use `Connect` and
> `Disconnect`; `OpenAsync` retains the .NET-aligned async name.

Parameterized commands, readers, transactions, scalar and non-query execution, adapter commands,
source-column binding, and task-returning provider operations are available without adding an ADO
reference.

### Clean up deterministically

For deterministic cleanup around a zero-argument function, use
`ROneCOne.Using(connection).Run(work)`. Adapter batch updates expose `UseTransaction`,
`ContinueUpdateOnError`, and `LastUpdateErrors`.

### Know your provider's limits

Inspect `connection.AsyncMode` and `connection.SupportsNativeAsync` when behavior depends on
provider capabilities; the built-in late-bound ADO path reports cooperative scheduling rather than
native asynchronous I/O.

Provider capabilities still depend on the selected driver. For example, the Excel ISAM supports
worksheet reads, updates, and inserts but rejects row deletion.

## Where next

- [Practical reference](reference.md) completes the learning path with the everyday operator map.
- [Data technical reference](../data.md) defines exact table, relation, view, and provider
  contracts.
- [Guide index](README.md) returns to the full learning path.
