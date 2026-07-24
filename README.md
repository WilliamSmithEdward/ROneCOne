# ROneCOne

**Give VBA the love it deserves.**

[![Latest release](https://img.shields.io/github/v/release/WilliamSmithEdward/ROneCOne)](https://github.com/WilliamSmithEdward/ROneCOne/releases/latest)
[![MIT license](https://img.shields.io/github/license/WilliamSmithEdward/ROneCOne)](LICENSE)
[![Microsoft 365 Excel on Windows x64](https://img.shields.io/badge/Excel-Microsoft_365_Windows_x64-217346)](docs/user-guide/getting-started.md)
[![Single class module runtime](https://img.shields.io/badge/runtime-one_class_module-0078D4)](src/ROneCOne.cls)

ROneCOne is the standard library Excel VBA never had, shaped after C# and .NET. Import one
class module, `ROneCOne.cls`, and ordinary workbook code can query typed collections with
LINQ, await tasks and web requests, handle failure with structured Try and Catch, model data
in tables, and exchange JSON, CSV, and XML.

Everything you already built keeps working. There is no installer, no add-in, no external
reference, and no separate runtime to manage: the entire library is one file that travels inside
any workbook that imports it. Your macros call it like any other VBA class.

## What it looks like

```vba
Dim scores As ROneCOne
Dim strongScores As ROneCOne

Set scores = ROneCOne.ListOf(vbLong, 90, 72, 88, 95)

Set strongScores = scores _
    .Where(scores.Element.AtLeast(85)) _
    .OrderDescending _
    .ToList

MsgBox "Scores at or above 85: " & strongScores.JoinText(", ")
```

That is a typed list, filtered, sorted, and displayed without a loop, counter, or temporary
array. The same fluent style carries through object queries such as `.Where("Age").AtLeast(40)`,
Tasks, typed events, and structured Try, Catch, and Finally flows.

## What becomes possible

| You want to... | ROneCOne gives you... |
|---|---|
| Work with data cleanly | Typed collections and LINQ-style filtering, shaping, sorting, and aggregation |
| Reuse behavior | Delegates and expression-based functions that can be passed, combined, and invoked |
| Coordinate async work | Await-style Tasks with delays, timeouts, cancellation, progress, and continuations |
| Model business data | DataTable, DataSet, DataView, relations, change tracking, and providers |
| Call web APIs | An HttpClient with awaitable verbs, overlapped downloads, and typed failures |
| Exchange JSON | Serialize and deserialize trees, tables, and your own classes, System.Text.Json style |
| Read and write files | File, Directory, and Path surfaces with real UTF-8, System.IO style |
| Open and make zips | Read any zip and write archives in pure VBA |
| Watch and log | Await folder changes and log runs to a file, service style |
| Exchange CSV | RFC 4180 round trips between DataTables and CSV text with typed columns |
| Exchange XML | XPath queries over secured MSXML6 and typed DataTable bridges both ways |
| Run command lines | Awaitable processes with exit codes and captured output, never a frozen Excel |
| Match and hash text | Regular expressions, SHA and HMAC digests, and base64 or hex encoding |
| Work with timestamps | DateTimeOffset-style instants: ISO 8601, Unix epochs, zone conversion, durations |
| Format and build text | String.Format-style invariant formatting, StringBuilder, GUIDs, crypto randomness |
| Build responsive workbook logic | Typed events with predictable subscription and delivery |
| Handle failure deliberately | Structured Try, Catch, and Finally flows |
| Deploy without a platform project | One importable class module with no outside dependency |

The design follows familiar C# and .NET names and behavior wherever VBA permits it, and
IntelliSense descriptions keep the surface discoverable inside the editor.

## What it needs

- Windows x64 Microsoft 365 Excel
- A macro-enabled workbook or add-in: `.xlsm`, `.xlsb`, or `.xlam`
- One imported file: [`src/ROneCOne.cls`](src/ROneCOne.cls)

Everything runs locally inside your one Excel process. ROneCOne requires no installer, no
external library, and no runtime VBIDE trust, and it sends no telemetry. The network is touched
only by the HTTP client, and only for URLs you request. It never launches a second Excel.

## Start in three steps

1. **[Run a demo](https://github.com/WilliamSmithEdward/ROneCOne/releases/latest)** - each
   release ships self-contained demo workbooks with visible, worksheet-by-worksheet results.
2. **[Follow the getting-started guide](docs/user-guide/getting-started.md)** - import one class
   file and run your first query in about five minutes.
3. **[Pick a recipe](docs/user-guide/README.md)** - short guides cover collections, delegates,
   events, exceptions, tasks, and data access.

## Documentation

| Surface | Where |
|---|---|
| Guided learning path | [User guide](docs/user-guide/README.md) |
| Everyday operator lookup | [Practical reference](docs/user-guide/reference.md) |
| Exact contracts and semantics | [Technical documentation index](docs/README.md) |
| Design and runtime boundaries | [Architecture](docs/architecture.md) |
| Release history | [Changelog](CHANGELOG.md) |

## License

ROneCOne is available under the [MIT License](LICENSE).
