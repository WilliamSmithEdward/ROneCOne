# Data and providers

ROneCOne provides a typed, in-memory data layer for workbook code and a late-bound provider layer
for local OLE DB or ODBC sources.

## Build a typed table

```vba
Dim people As ROneCOne
Dim row As ROneCOne

Set people = ROneCOne.DataTable("People")
people.Column("Id", vbLong).AutoNumber(100&, 10&).AsUnique
people.Column("Name", vbString).WithDefault "Unknown"
people.Column "Age", vbLong

Set row = people.NewRow
row.Item("Name") = "Ada"
row.Item("Age") = 47&
people.AddRow row
```

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

VBA reserves `Open` and `Close` as language keywords, so direct calls use `Connect` and
`Disconnect`; `OpenAsync` retains the .NET-aligned async name. Parameterized commands, readers,
transactions, scalar and non-query execution, adapter commands, source-column binding, and
task-returning provider operations are available without adding an ADO reference.

Provider capabilities still depend on the selected driver. For example, the Excel ISAM supports
worksheet reads, updates, and inserts but rejects row deletion.

See the [data contracts](../data.md) or return to the [guide index](README.md).
