# ADR 0022: XML over MSXML6 with DataTable bridges

Status: accepted, 2026-07-24

## Context

XML remains the wire format of enterprise systems: SOAP services, configuration files, exports
from line-of-business tools, and Office's own document formats. The runtime already exchanges
JSON and CSV with the data layer; XML was the last major text format without a surface. MSXML6
(`MSXML2.DOMDocument.6.0`) ships with every supported Windows, and a probe confirmed everything
the surface needs: XPath is the default selection language, `parseError` carries code, line,
position, and reason, namespaced queries work through the `SelectionNamespaces` property, and
the security defaults are right (DTDs prohibited with code C00CE584, external references
unresolved). The probe also pinned the behaviors the design must respect: `async` and
`validateOnParse` default to True and must be switched off, elements in a default namespace are
invisible to plain XPath until a prefix is mapped, and a node's `text` concatenates descendant
text without separators.

## Decision

`ROneCOne.Xml.Parse(text, [selectionNamespaces])` and `Load(path, [selectionNamespaces])` wrap
a secured MSXML6 document (synchronous, non-validating, externals unresolved, DTD prohibition
untouched) and return the document element as a node value. Nodes expose `Name`, `Value`
(concatenated text), `GetAttribute`/`HasAttribute` (the `XmlElement` names; `Attribute` cannot
be a VBA member name because it collides with the class file format's attribute lines),
`Elements([name])`, `SelectNodes(xpath)` returning a queryable typed list, `SelectSingleNode`
returning `Nothing` on a miss like the DOM it wraps, and `OuterXml`. Parse failures raise the
typed `ROneCOne.XmlError` from source `ROneCOne.XmlException` carrying the parser's line,
position, and reason; invalid XPath re-raises through the same channel.

The data bridges mirror JSON and CSV. `ROneCOne.Xml.DeserializeTable(text, [tableName],
[rowsPath])` selects one element per row (default `/*/*`, every child of the document element)
and builds a typed DataTable whose columns are the union, in first-seen order, of row
attributes and simple child elements; nested elements never become columns, an element beats a
same-named attribute, and absent or empty cells load as database nulls. Column typing reuses
the CSV inference verbatim, so XML, CSV, and JSON agree on every value shape.
`table.ToXml([rootName], [rowName])` writes compact element-per-column rows in the
`DataTable.WriteXml` shape, omitting null cells, escaping text, and validating element names,
so `ToXml` output round-trips through `DeserializeTable`.

## Consequences

`XML_PROG_ID` joins the source-contract whitelist, which also pins the secure posture: the
runtime must keep `resolveExternals` False and may never re-enable DTDs. Namespaced documents
require an explicit prefix mapping through `selectionNamespaces`; `DeserializeTable` targets
plain documents and namespaced tables go through `Parse` plus `SelectNodes`. `Parse` and
`DeserializeTable` join the existing role-dispatched members shared with the DateTime, TimeSpan,
JSON, and CSV surfaces. The live suite asserts navigation, predicates, entity decoding,
namespace mapping, DTD rejection, malformed-markup and bad-XPath failures, file load with
declared encodings, typed table extraction with null gaps, and the serialize round trip.
