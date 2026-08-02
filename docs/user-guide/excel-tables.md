# Excel Tables (ListObjects)

An Excel Table is a `ListObject`, and a `ListObject` is not a `Range`. Every ROneCOne bridge takes
a `Range`, so you never hand over the table itself. You hand over one of its ranges.

That single fact is the whole obstacle. Pass the table directly and you get run-time error 438,
"Object doesn't support this property or method", which reads like ROneCOne rejecting your table
when it is really VBA rejecting the type.

```vba
Dim sales As ROneCOne

' Wrong: raises 438
Set sales = ROneCOne.DataTableFromRange(Sheet1.ListObjects("Sales"))

' Right: hand over a range belonging to the table
Set sales = ROneCOne.DataTableFromRange(Sheet1.ListObjects("Sales").Range)
```

## The three ranges you will use

| Range | What it covers | Use it when |
|---|---|---|
| `listObject.Range` | Header row plus all data rows | You want the headers to become column names |
| `listObject.DataBodyRange` | Data rows only | You will supply names yourself, or there are none |
| `listObject.ListColumns("Amount").DataBodyRange` | One column of data | You want a single typed list |

```vba
Dim amounts As ROneCOne
Dim bodyOnly As ROneCOne
Dim salesTable As ListObject

Set salesTable = Sheet1.ListObjects("Sales")
Set sales = ROneCOne.DataTableFromRange(salesTable.Range)
Set bodyOnly = ROneCOne.DataTableFromRange(salesTable.DataBodyRange, False)
Set amounts = ROneCOne.ListFromRange(salesTable.ListColumns("Amount").DataBodyRange)
```

With `headers:=False` there are no names to read, so the columns are called `Column1`, `Column2`,
and so on.

## After that it is an ordinary table

Once the data is in a `DataTable`, nothing about it is table-specific any more. `Rows` is a
sequence, so every operator that works on a list works here.

```vba
Dim byRegion As ROneCOne

' Filter, sort, and project
sales.Rows.Where(sales.Rows!Amount.AtLeast(100)) _
    .OrderBy("Rep").JoinText(", ", "Rep")

' Sort by two keys, then keep the top few
sales.Rows.OrderBy("Region").ThenByDescending("Amount").Take(3)

' Aggregate the whole table, or only the matching rows
sales.Rows.Sum("Amount")
sales.Rows.Count(sales.Rows!Region.EqualTo("West"))

' Group, then total each group
Set byRegion = sales.Rows.GroupBy("Region")
byRegion.First.Key & " " & byRegion.First.Sum("Amount")
```

`sales.Rows!Region` is bang syntax, a readable way to name a column when building a condition. It
means the same thing as `sales.Rows.Condition("Region")`.

A `DataView` is the sheet-oriented alternative. It stays attached to the table and knows how to
write itself back to cells.

```vba
Dim topWest As ROneCOne

Set topWest = ROneCOne.DataView(sales) _
    .WithFilter(sales.Rows!Region.EqualTo("West")) _
    .WithSort("Amount", True)
topWest.ToRange Sheet2.Range("A1")
```

## Map rows onto your own class

Give ROneCOne a factory that makes one empty instance, and it fills the properties whose names
match your column headers. Matching is case-insensitive.

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

' And back the other way: name the properties to read
Set rebuilt = ROneCOne.DataTableFromObjects( _
    objects, Array("Region", "Rep", "Amount"), "Rebuilt")
```

## JSON and CSV

Tables, single rows, and views all serialize.

```vba
sales.ToJson                          ' the whole table as a JSON array
sales.ToJson(True)                    ' the same, two-space indented
sales.Rows.Item(0).ToJson             ' {"Region":"West","Rep":"Ada","Amount":120}
topWest.ToJson                        ' only the filtered, sorted rows
sales.ToCsv                           ' RFC 4180 text, header row first

Set sales = ROneCOne.Json.DeserializeTable(jsonText, "Sales")
Set sales = ROneCOne.Csv.DeserializeTable(csvText, "Sales")
```

## Writing back and growing the table

`ToRange` writes headers plus rows in one bulk assignment, so a few thousand cells cost a fraction
of a cell loop.

Growing the table is Excel's job. Add the row through the `ListObject`, fill the new cells, then
read the table again. A `DataTable` is a snapshot taken when you read it, not a live link, so an
instance you already hold will not notice the new row.

```vba
salesTable.ListRows.Add
With salesTable.DataBodyRange
    .Cells(salesTable.ListRows.Count, 1).Value = "South"
    .Cells(salesTable.ListRows.Count, 2).Value = "Fay"
    .Cells(salesTable.ListRows.Count, 3).Value = 99
End With
Set sales = ROneCOne.DataTableFromRange(salesTable.Range)
```

## See it running

The Excel Tables demo workbook builds a real table, then works through every step above with the
live result next to the expected one. Download it from the
[release page](https://github.com/WilliamSmithEdward/ROneCOne/releases/latest).

## Related

- [Data and providers](data-and-providers.md) for typed tables, constraints, and relations
- [Collections and LINQ](collections-and-linq.md) for the full operator set
- [JSON and typed objects](json-and-objects.md) for binding rules and partial reads

[Back to the user guide](README.md) | [Documentation index](../README.md)
