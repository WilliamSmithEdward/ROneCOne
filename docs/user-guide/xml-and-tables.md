# XML documents and tables

Read the XML your systems still speak. `ROneCOne.Xml` wraps the MSXML6 parser Windows already
ships, with XPath queries, typed table extraction, and the same secure defaults .NET uses:
document type definitions rejected, external references never resolved.

## Parse and navigate

```vba
Dim doc As ROneCOne
Dim book As ROneCOne

Set doc = ROneCOne.Xml.Parse( _
    "<catalog><book id=""1""><title>Dune</title></book></catalog>")
Debug.Print doc.Name                         ' catalog
For Each book In doc.Elements("book")
    Debug.Print book.GetAttribute("id"), book.SelectSingleNode("title").Value
Next book
```

`Elements` and `SelectNodes` return ordinary typed lists, so the whole LINQ surface applies.
`SelectSingleNode` returns `Nothing` when nothing matches; test with `Is Nothing`.

## XPath does the searching

```vba
Debug.Print doc.SelectNodes("//book[@id='1']").Count
Debug.Print doc.SelectSingleNode("//book[@id='1']/title").Value
```

Documents with a default namespace need a prefix mapping once, at parse time:

```vba
Set doc = ROneCOne.Xml.Parse(feedText, "xmlns:a='http://www.w3.org/2005/Atom'")
Debug.Print doc.SelectNodes("//a:entry").Count
```

## Straight to a DataTable

One call turns repeated elements into a typed table, with the same deterministic column
inference CSV uses; attributes and simple child elements become columns, and gaps load as
database nulls:

```vba
Dim orders As ROneCOne

Set orders = ROneCOne.Xml.DeserializeTable(exportText, "Orders", "//order")
Debug.Print orders.Rows.Count, orders.Columns.Count
' From there: orders.ToRange puts the whole thing on a worksheet.
```

And back out again. `ToXml` writes the `DataTable.WriteXml` shape and round-trips:

```vba
Debug.Print orders.ToXml()
```

## Files and failure

`ROneCOne.Xml.Load(path)` reads a file honoring its declared encoding. Anything malformed
raises the typed `ROneCOne.XmlError` with the parser's line, position, and reason; a document
carrying a `DOCTYPE` is rejected by design.

```vba
On Error Resume Next
Set doc = ROneCOne.Xml.Load(configPath)
If Err.Number = ROneCOne.XmlError Then Debug.Print Err.Description
On Error GoTo 0
```

## Where next

- [XML technical reference](../xml.md) defines the parser posture, table bridges, and every
  member exactly.
- [Data and providers](data-and-providers.md) covers the DataTables these bridges fill.
- [Guide index](README.md) returns to the full learning path.
