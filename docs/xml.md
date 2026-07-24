# XML

Exact semantics for the XML surface. For the workflow-first introduction, read
[XML documents and tables](user-guide/xml-and-tables.md).

## Surface

`ROneCOne.Xml.Parse(text, [selectionNamespaces])` parses a document and returns its document
element as a node value; `Load(path, [selectionNamespaces])` does the same from a file,
honoring the file's declared encoding. A node exposes:

| Member | Behavior |
|---|---|
| `Name` | The element name |
| `Value` | Concatenated descendant text, entities decoded |
| `GetAttribute(name)` | The attribute's text; raises `XmlError` when absent |
| `HasAttribute(name)` | True when the attribute is present |
| `Elements([name])` | Child elements as a queryable typed list, optionally filtered by name |
| `SelectNodes(xpath)` | Every XPath match as a queryable typed list |
| `SelectSingleNode(xpath)` | The first match, or `Nothing` when nothing matches |
| `OuterXml` | The node's serialized markup |

Failures raise error number `ROneCOne.XmlError` from source `ROneCOne.XmlException`. A parse
failure carries the parser's line, position, and reason; an invalid XPath query re-raises
through the same channel.

## Parser posture

The engine is MSXML6 (`MSXML2.DOMDocument.6.0`), which ships with Windows. The runtime keeps
its secure defaults: document type definitions stay prohibited (a document with a `DOCTYPE`
fails to parse), external references are never resolved, and parsing is synchronous and
non-validating. Whitespace-only text between elements is not preserved.

XPath is the selection language. Elements in a default namespace are invisible to plain XPath;
map a prefix through the optional `selectionNamespaces` argument
(`"xmlns:p='urn:example'"`) and query with that prefix (`//p:item`). `GetAttribute` and
`Elements` compare names exactly, including any prefix.

## XML into a DataTable

`ROneCOne.Xml.DeserializeTable(text, [tableName], [rowsPath])` builds a typed DataTable. The
row path is an XPath selecting one element per row; the default `/*/*` takes every child of the
document element. Columns are the union, in first-seen order, of each row's attributes and
simple child elements (elements whose children are text only); nested elements never become
columns, and a simple child element wins over a same-named attribute. A missing or empty cell
loads as a database null.

Column types come from the same deterministic inference CSV uses: whole numbers widen through
`Long`, `LongLong`, and `Double`, booleans are `true`/`false`, dates are validated ISO text,
and a column that mixes shapes stays text with its original characters. A row path that matches
nothing, or rows with no usable columns, raise `XmlError`. Namespaced documents go through
`Parse` plus `SelectNodes` instead; `DeserializeTable` targets plain documents.

## A DataTable into XML

`table.ToXml([rootName], [rowName])` writes compact element-per-column XML in the
`DataTable.WriteXml` shape: `rootName` defaults to `NewDataSet`, `rowName` defaults to the
table's name. Text cells escape `&`, `<`, and `>`; numbers write invariantly; dates write the
`yyyy-MM-ddTHH:mm:ss` stamp; booleans write `true`/`false`; a database null omits its element
entirely, so `ToXml` output round-trips through `DeserializeTable` with nulls intact. Root,
row, and column names must be legal element names (a letter or underscore, then letters,
digits, `-`, `_`, or `.`); anything else is refused.

[Back to the documentation index](README.md)
