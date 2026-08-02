# Excel Tables (ListObjects)

Hand over the Table.

```vba
Dim sales As ROneCOne

Set sales = ROneCOne.Table(Sheet1.ListObjects("Sales"))
sales.Rows.Count
```

Not `.Range`, not `.DataBodyRange`. The `ListObject` itself. Every worksheet
entry point takes one now: `DataTableFromRange`, `ListFromRange`,
`LoadFromRange`, and `ToRange`. A single table column works too, so
`ListFromRange(salesTable.ListColumns("Amount"))` reads that column into a list.

`ROneCOne.Table` does one thing the others do not: the table it hands back
remembers which Excel Table it came from, so it can re-read itself and write
itself back.

> Before 1.9.0 these bridges took only a `Range`, and passing a `ListObject`
> raised run-time error 438. If you are reading older code that calls
> `DataTableFromRange(lo.Range)`, that still works; it is just no longer
> necessary.

## Reading

```vba
Dim salesTable As ListObject

Set salesTable = Sheet1.ListObjects("Sales")

' Attached: can Refresh and WriteBack
Set sales = ROneCOne.Table(salesTable)

' Detached: an ordinary snapshot
Set sales = ROneCOne.DataTableFromRange(salesTable)

' Body only, so the columns are named Column1, Column2, ...
Set body = ROneCOne.DataTableFromRange(salesTable, False)

' One column into a typed list
Set amounts = ROneCOne.ListFromRange(salesTable.ListColumns("Amount"))

' Or into columns you declared yourself
typed.LoadFromRange salesTable
```

A totals row is not data. It stays out of `Rows` and out of the column reads,
so turning totals on does not change what you get back.

## Querying

`Rows` is an ordinary sequence, so the whole operator set applies.

```vba
Dim byRegion As ROneCOne

sales.Rows.Count(sales.Rows!Region.EqualTo("West"))
sales.Rows.Where(sales.Rows!Amount.AtLeast(100)).OrderBy("Rep").JoinText(", ", "Rep")
sales.Rows.OrderByDescending("Amount").Take(2).Sum("Amount")
sales.Rows.OrderBy("Region").ThenByDescending("Amount")
sales.Rows.SelectItems("Region", vbString).Distinct.Order
sales.Rows.Sum("Amount")
sales.Rows.Average("Amount")
sales.Rows.FirstOrDefault(sales.Rows!Rep.EqualTo("Cy"))
sales.Rows.AnyItem(sales.Rows!Amount.AtLeast(150))

Set byRegion = sales.Rows.GroupBy("Region")
byRegion.First.Key & " " & byRegion.First.Sum("Amount")
```

`sales.Rows!Region` is bang syntax, a readable way to name a column when
building a condition. It means the same as `sales.Rows.Condition("Region")`.

A `DataView` is the sheet-oriented alternative: a live filtered and sorted
window that knows how to write itself back.

```vba
Set topWest = ROneCOne.DataView(sales) _
    .WithFilter(sales.Rows!Region.EqualTo("West")) _
    .WithSort("Amount", True)
```

## Mapping to your own objects

Give ROneCOne a factory that makes one empty instance, and it fills the
properties whose names match your column headers. Matching is
case-insensitive.

```vba
' In a standard module, Public so ROneCOne.Func can resolve it by name
Public Function NewSalesRow() As SalesRow
    Set NewSalesRow = New SalesRow
End Function
```

```vba
Dim factory As ROneCOne
Dim objects As ROneCOne
Dim rebuilt As ROneCOne

Set factory = ROneCOne.Func("Module1.NewSalesRow").Takes().Returns(vbObject)
Set objects = sales.ToObjects(factory)

objects.Item(0).Rep                                   ' "Ada", a real SalesRow
objects.Where(objects.Condition("Region").EqualTo("West")).Count

' And back: name the properties to read, each becomes a column
Set rebuilt = ROneCOne.DataTableFromObjects( _
    objects, Array("Region", "Rep", "Amount"), "Rebuilt")
```

The trip is lossless. `rebuilt.ToJson = sales.ToJson` is `True`.

## JSON and CSV

Tables, single rows, and views all serialize, in both directions.

```vba
sales.ToJson                          ' the whole table as a JSON array
sales.ToJson(True)                    ' the same, two-space indented
sales.Rows.Item(0).ToJson             ' {"Region":"West","Rep":"Ada","Amount":120}
topWest.ToJson                        ' only the filtered, sorted rows
sales.ToCsv                           ' RFC 4180 text, header row first

Set sales = ROneCOne.Json.DeserializeTable(jsonText, "Sales")
Set sales = ROneCOne.Csv.DeserializeTable(csvText, "Sales")
```

## Writing back

`WriteBack` puts rows into the Excel Table the value came from and resizes the
table to fit. It returns the number of rows written.

```vba
' Keep only the West rows, biggest first, and make the sheet match
sales.WriteBack topWest

' Or write the table's own rows after changing them
sales.LoadRow Array("South", "Fay", 99)
sales.WriteBack

' Re-read the sheet in place, in case something else changed it
sales.Refresh
```

`ToRange` accepts a Table as its target and does exactly the same thing, so a
view can drive the write without going through the attached table:

```vba
topWest.ToRange salesTable
```

Four behaviors are worth knowing, because Excel makes them awkward and
ROneCOne handles them for you:

- **Shrinking clears up after itself.** Excel leaves the vacated cells
  populated underneath a table that got smaller. `WriteBack` clears them.
- **Writing no rows empties the table.** Excel refuses to resize a table to
  its header alone, so the body is deleted instead. The table and its headers
  survive.
- **A totals row survives.** Excel places it inconsistently across a grow and
  a shrink, so totals are switched off around the resize and restored after.
- **Name, style, and headers are preserved** across every resize.

The column count has to match. Writing a five-column table into a
three-column Excel Table raises rather than silently truncating.

`Refresh` and `WriteBack` need a table that came from `ROneCOne.Table`, since
that is what remembers the origin. On any other table they raise and tell you
so.

## See it running

The Excel Tables demo workbook builds a real table and works through all of
the above, live, with the result next to what was expected. Download it from
the [release page](https://github.com/WilliamSmithEdward/ROneCOne/releases/latest).

## Related

- [Data and providers](data-and-providers.md) for typed tables, constraints, and relations
- [Collections and LINQ](collections-and-linq.md) for the full operator set
- [JSON and typed objects](json-and-objects.md) for binding rules and partial reads

[Back to the user guide](README.md) | [Documentation index](../README.md)
