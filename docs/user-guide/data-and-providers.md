# Data and providers

ROneCOne provides a typed, in-memory data layer for workbook code and a late-bound provider layer
for local OLE DB or ODBC sources.

## Move data between a worksheet and a table

Reading and writing worksheet cells one at a time is slow. The Range bridge moves a whole block
in a single call in each direction.

```vba
Dim sales As ROneCOne
Dim top As ROneCOne

Set sales = ROneCOne.DataTableFromRange(Sheet1.Range("A1:C500"))
Set top = ROneCOne.DataView(sales) _
    .WithFilter(sales.Rows!Region.EqualTo("West")) _
    .WithSort("Amount", True)
top.ToRange Sheet2.Range("A1")
```

`DataTableFromRange` reads the block into a new table, using the first row as column names. To
load into a table whose columns you already typed, use `table.LoadFromRange(range)` instead; each
value is checked and widened against your column types. `ROneCOne.ListFromRange(range)` reads a
single row or column into a list. `ToRange` writes a table, a filtered and sorted `DataView`, or
any scalar sequence back to the sheet in one assignment. Every path uses one bulk call, so a few
thousand cells cost a fraction of what a cell loop would.

## Build a typed table

```vba
Dim people As ROneCOne
Dim row As ROneCOne

Set people = ROneCOne.DataTable("People")
people.Column("Id", vbLong).AutoNumber(100, 10).AsPrimaryKey
people.Column("Name", vbString).WithDefault "Unknown"
people.Column "Age", vbLong
people.Column "Note", vbString

Set row = people.Row("Ada", 47, ROneCOne.DBNull).Add
Debug.Print people.Find(100).Item("Name")
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
    .WithFilter(people.Rows!Age.AtLeast(40)) _
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
Set command = ROneCOne.DbCommand(sql, connection).WithTimeout(30)
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

`connection.AsyncMode` reports `"Native"`: `OpenAsync`, `ExecuteReaderAsync`,
`ExecuteScalarAsync`, and `FillAsync` start the operation inside ADO and the returned Task polls
provider state, so the provider works while Excel stays responsive. `ExecuteNonQueryAsync`,
`UpdateAsync`, and `ReadAsync` run their work inside one cooperative task step instead, because
ADO exposes no reliable async completion for affected-row counts or row-by-row updates.

Provider capabilities still depend on the selected driver. For example, the Excel ISAM supports
worksheet reads, updates, and inserts but rejects row deletion.

### Exchange CSV

A table becomes RFC 4180 text in one call, and CSV text becomes a typed table in one more:

```vba
Dim orders As ROneCOne

ROneCOne.File.WriteAllText "C:\data\orders.csv", table.ToCsv
Set orders = ROneCOne.Csv.DeserializeTable( _
    ROneCOne.File.ReadAllText("C:\data\orders.csv"))
```

Quoting, embedded commas and newlines, and doubled quotes follow RFC 4180 exactly. On the way
in, each column infers one deterministic type (integer, double, boolean, ISO date) or stays
text with its original characters; quoted cells always stay text, so `"00042"` keeps its
zeros. Failures raise the typed `ROneCOne.CsvError`. The
[CSV technical reference](../csv.md) defines every rule.

### Pick a provider

Connection strings choose the driver. The two the test suite exercises:

- Workbooks through the ACE ISAM, shipped with Office:
  `Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\path\book.xlsx;Extended
  Properties="Excel 12.0 Xml;HDR=YES";`
- SQL Server through the Microsoft OLE DB driver:
  `Provider=MSOLEDBSQL;Data Source=localhost;Initial Catalog=tempdb;Integrated
  Security=SSPI;` (use `User ID=...;Password=...;` where SQL authentication is enabled).

If `MSOLEDBSQL` is not installed, the connection open raises "Provider cannot be found"; install
the Microsoft OLE DB Driver for SQL Server or use an ODBC connection string through `MSDASQL`.
Connection failures surface the provider's own error text, such as "Login failed for user" for
rejected credentials or "Could not open a connection to SQL Server" for an unreachable host.

## Where next

- [Practical reference](reference.md) completes the learning path with the everyday operator map.
- [Data technical reference](../data.md) defines exact table, relation, view, and provider
  contracts.
- [Guide index](README.md) returns to the full learning path.
