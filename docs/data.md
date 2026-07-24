# DataTable, DataSet, and providers

> [!TIP]
> New to ROneCOne? Start with the
> [Data and providers user guide](user-guide/data-and-providers.md).

The data layer gives workbook code checked tables, rows, relationships, and local database access.
It uses the same one-class runtime as collections and Tasks. The sections below define the exact
technical contract.

## Worksheet Range bridge

- `ROneCOne.DataTableFromRange(range, headers, name)` reads a rectangular block into a new
  Variant-typed DataTable in one `Range.Value` call, taking column names from the header row when
  present.
- `DataTable.LoadFromRange(range, headers)` appends a block into an existing typed table,
  admitting each value through the column types with numeric widening. It requires the range and
  table column counts to match.
- `ROneCOne.ListFromRange(range)` reads a single row or column into a `List<Variant>`.
- `ToRange(target, writeHeaders)` writes a DataTable grid, a DataView's visible grid, or a scalar
  sequence's column vector to a range in one assignment; object-valued cells are refused.

Range parameters are late-bound `Object` values, so the runtime keeps no compile-time Excel
reference. Values move through `Range.Value`, preserving dates, currency, and Booleans across a
round trip.

## In-memory data

- `DataTable`, `DataColumn`, `Column`, `Row`, `NewRow`, `LoadRow`, `SelectRows`, `CloneTable`,
  `Copy`, `Merge`, and `GetChanges`
- typed columns, defaults, auto-increment values, nullability, uniqueness, and primary keys
- `DataRow` indexing, state tracking, delete, accept/reject, and parent/child navigation
- `DataView.WithFilter`, `WithSort`, enumeration, LINQ, and `ToTable`
- `DataSet`, named table/relation collections, constraint enforcement, and relation validation

Adding a constrained relation validates both existing and future rows. Failed relation admission
does not mutate the dataset.

`Column(...).AsPrimaryKey` is the concise single-column form. `PrimaryKey` accepts a typed column
sequence for composite keys. `Find(key...)` uses a maintained hash index for both forms, and
single- or multi-column `DataRelation` navigation uses cached parent and child indexes that are
invalidated by table version changes. `Row(values...).Add` creates and attaches a positional row;
auto-number columns may be omitted. `ROneCOne.DBNull` expresses a deliberate database null without
confusing it with an omitted optional argument.

## Providers

The provider surface late-binds `ADODB.Connection` and `ADODB.Command`, so no workbook reference
is required. It includes parameter inference and explicit ADO types, size, precision, scale,
direction, source-column/version binding, command timeout/type, readers, transactions, adapter
fill/update commands, deterministic disposal, and task-returning calls.

`DbDataAdapter.Update` selects Insert, Update, or Delete commands from each row state, binds current
or original values, executes the command, and accepts successful changes. Actual SQL and mutation
support remain properties of the chosen OLE DB or ODBC provider.

`ROneCOne.Using(resource).Run(work)` provides deterministic cleanup and retains both work and
disposal failures when both occur. Adapter updates can opt into `UseTransaction` and
`ContinueUpdateOnError`; `LastUpdateErrors` retains row-level failures from a continued batch.
Transactions commit only after the whole admitted batch succeeds and roll back on failure.

The late-bound provider layer reports its async mode honestly: `AsyncMode` is `"Native"`.
`OpenAsync`, `ExecuteReaderAsync`, `ExecuteScalarAsync`, and `FillAsync` start the operation
inside ADO with the async option and the returned Task polls provider state cooperatively, so
the provider works while Excel's thread stays free between polls. `ExecuteNonQueryAsync`,
`UpdateAsync`, and `ReadAsync` complete their work inside one cooperative task step because ADO
reports no usable async completion signal for records-affected counts or row-by-row updates.
VBA itself never runs on more than one thread.

[Back to the documentation index](README.md)
