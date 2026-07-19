# DataTable, DataSet, and providers

The data layer uses the same tagged one-class kernel as collections and tasks. Tables own typed
columns and rows; datasets own tables and relations; adapters bridge provider records into those
objects.

## In-memory data

- `DataTable`, `DataColumn`, `Column`, `NewRow`, `LoadRow`, `SelectRows`, `CloneTable`, `Copy`,
  `Merge`, and `GetChanges`
- typed columns, defaults, auto-increment values, nullability, uniqueness, and primary keys
- `DataRow` indexing, state tracking, delete, accept/reject, and parent/child navigation
- `DataView.WithFilter`, `WithSort`, enumeration, LINQ, and `ToTable`
- `DataSet`, named table/relation collections, constraint enforcement, and relation validation

Adding a constrained relation validates both existing and future rows. Failed relation admission
does not mutate the dataset.

## Providers

The provider surface late-binds `ADODB.Connection` and `ADODB.Command`, so no workbook reference
is required. It includes parameter inference and explicit ADO types, size, precision, scale,
direction, source-column/version binding, command timeout/type, readers, transactions, adapter
fill/update commands, deterministic disposal, and task-returning calls.

`DbDataAdapter.Update` selects Insert, Update, or Delete commands from each row state, binds current
or original values, executes the command, and accepts successful changes. Actual SQL and mutation
support remain properties of the chosen OLE DB or ODBC provider.

[Data user guide](user-guide/data-and-providers.md)
