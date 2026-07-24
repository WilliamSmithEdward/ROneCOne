# ROneCOne user guide

This guide starts with the result you want and keeps the machinery out of the way. You do not
need to understand delegates, expression trees, or COM to begin. Start with a demo, import one
class module, and copy the smallest example that solves your problem.

## Pick a path

| Guide | What you will accomplish | Typical first read |
|---|---|---:|
| [Getting started](getting-started.md) | Run a demo and add ROneCOne to a workbook | 5 minutes |
| [Collections and LINQ](collections-and-linq.md) | Query and summarize data | 10 minutes |
| [Delegates and expressions](delegates-and-expressions.md) | Build reusable behavior | 10 minutes |
| [Events and exceptions](events-and-exceptions.md) | Coordinate changes and failures | 10 minutes |
| [Tasks and async](tasks-and-async.md) | Coordinate work, cancellation, and progress | 10 minutes |
| [Data and providers](data-and-providers.md) | Model, query, and load tabular data | 15 minutes |
| [HTTP and web data](http-and-web.md) | Call web APIs with awaitable requests | 10 minutes |
| [JSON and typed objects](json-and-objects.md) | Parse, serialize, and bind JSON | 10 minutes |
| [Files and folders](files-and-folders.md) | Read and write files with real encodings | 10 minutes |
| [Processes and commands](processes-and-commands.md) | Await command lines with captured output | 5 minutes |
| [Text, hashing, and encoding](text-and-hashing.md) | Regex, formatting, digests, base64, and hex | 10 minutes |
| [Dates and times](dates-and-times.md) | Parse ISO 8601 and epochs, convert zones, add durations | 10 minutes |
| [XML documents and tables](xml-and-tables.md) | Query XML with XPath and land it in DataTables | 10 minutes |
| [Practical reference](reference.md) | Find names, defaults, and limits | As needed |

## Fastest route to the first result

1. Download the latest demo workbooks from the
   [release page](https://github.com/WilliamSmithEdward/ROneCOne/releases/latest).
2. Open the Collections demo and run its main macro.
3. Follow [Getting started](getting-started.md) to import `ROneCOne.cls` into a copy of your own
   workbook.
4. Choose one recipe from [Collections and LINQ](collections-and-linq.md).

The examples lead with the concise form you should normally use. The focused
[technical documents](../README.md) remain available when you need exact semantics or the
canonical expansion behind that syntax.

## What you need

- Windows x64
- Microsoft 365 Excel
- A macro-enabled `.xlsm`, `.xlsb`, or `.xlam` file
- Permission to run VBA in that file

ROneCOne itself does not require an installer, an add-in, an external library, or trusted
programmatic access to the VBIDE. The network is used only by the HTTP client, and only for the
URLs you request.

[Back to the project overview](../../README.md) | [Documentation index](../README.md)
