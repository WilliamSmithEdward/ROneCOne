# JSON and typed objects

Parse JSON the way C# uses System.Text.Json: deserialize into navigable values, bind onto your
own classes, or land an array of objects straight into a typed DataTable. Serialize any
collection, table, or row back to JSON text. No references, no add-ins, no third-party module.

## Parse and navigate

```vba
Dim tree As ROneCOne

Set tree = ROneCOne.Json.Deserialize( _
    "{""name"":""Ada"",""age"":36,""tags"":[""x"",""y""]}")
Debug.Print tree.Item("name")            ' Ada
Debug.Print tree.Item("age")             ' 36
Debug.Print tree.Item("tags").Item(1)    ' y
Debug.Print tree.Count
```

A JSON object becomes an ordered dictionary with instant key lookup; a JSON array becomes a
list; numbers, text, booleans, and nulls arrive as ordinary values. Everything you know about
querying collections applies to the tree immediately. Malformed JSON raises a typed error
(`ROneCOne.JsonError`) whose message carries the offending position.

## Serialize anything back

```vba
Debug.Print tree.ToJson                          ' compact
Debug.Print ROneCOne.Json.Serialize(tree, True)  ' indented
Debug.Print ROneCOne.ListOf(vbLong, 1, 2, 3).ToJson   ' [1,2,3]
```

Member order is preserved, numbers always use a dot regardless of locale, dates travel as ISO
8601 text, and `Null` becomes `null`. Dictionaries, lists, queries, tables, rows, and views all
serialize; so do plain VBA Collections and one-dimensional arrays.

## JSON into a DataTable

```vba
Dim table As ROneCOne

Set table = ROneCOne.Json.DeserializeTable( _
    "[{""id"":1,""name"":""Ada""},{""id"":2,""name"":""Bo""}]")
Debug.Print table.Rows.Count             ' 2, with typed Id and Name columns
```

Column types are inferred from the values, nested objects become dotted columns
(`"owner.city"`), nested arrays are kept as JSON text cells, and missing members become
database nulls. An envelope is one argument away:
`Json.DeserializeTable(body, "Items", "$.data.items")`. From there, `table.ToRange` puts the
web on a worksheet.

## Bind to your own classes

VBA cannot invent instances of your classes, so you hand the runtime a tiny factory once, and
it does the rest by property name:

```vba
Public Function NewCustomer() As Customer
    Set NewCustomer = New Customer
End Function

Dim factory As ROneCOne
Dim customers As ROneCOne

Set factory = ROneCOne.Func("Module1.NewCustomer").Takes().Returns(vbObject)
Set customers = ROneCOne.Json.DeserializeObjects(body, factory)
Debug.Print customers.Item(0).CustomerName
```

`DeserializeInto(json, instance)` fills one existing object; scalar members bind by name, and
null, object, and array members are left alone. The same factory idea maps tables to objects
and back:

```vba
Set people = table.ToObjects(factory)                       ' rows to instances
Set table = ROneCOne.DataTableFromObjects(people, _
    Array("CustomerName", "Age", "City"))                    ' instances to rows
```

The property list stands in for the reflection VBA does not have; everything else is inferred.

[Back to the guide index](README.md) | [Exact JSON semantics](../json.md)
