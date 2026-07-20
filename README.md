# ROneCOne

**Give VBA the love it deserves.**

[![Latest release](https://img.shields.io/github/v/release/WilliamSmithEdward/ROneCOne)](https://github.com/WilliamSmithEdward/ROneCOne/releases/latest)
[![MIT license](https://img.shields.io/github/license/WilliamSmithEdward/ROneCOne)](LICENSE)
[![Microsoft 365 Excel on Windows x64](https://img.shields.io/badge/Excel-Microsoft_365_Windows_x64-217346)](docs/user-guide/getting-started.md)
[![Single class module runtime](https://img.shields.io/badge/runtime-one_class_module-0078D4)](src/ROneCOne.cls)

ROneCOne is a modern programming layer for Excel VBA. Import one class module, `ROneCOne.cls`,
and ordinary workbook code gains typed collections, LINQ-style queries, parallel tasks, typed
events, structured error handling, and an in-memory data layer with local database access - the
everyday toolkit of C# and .NET, written as plain, compilable VBA.

Everything you already built keeps working. There is no installer, no add-in, no external
reference, and no separate runtime to manage: the entire library is one file that travels inside
any workbook that imports it. Your macros call it like any other VBA class, and your workbook
stays a workbook.

## What it looks like

```vba
Dim scores As ROneCOne
Dim strongScores As ROneCOne

Set scores = ROneCOne.ListOf(vbLong, CLng(90), CLng(72), CLng(88), CLng(95))

Set strongScores = scores _
    .Where(scores.Element.AtLeast(CLng(85))) _
    .OrderDescending _
    .ToList

MsgBox "Scores at or above 85: " & strongScores.JoinText(", ")
```

That is a typed list, filtered, sorted, and displayed without a loop, counter, or temporary
array. The same fluent style carries through object queries such as `.Where("Age").AtLeast(40)`,
parallel Tasks, typed events, and structured Try, Catch, and Finally flows.

## What becomes possible

| You want to... | ROneCOne gives you... |
|---|---|
| Work with data cleanly | Typed collections and LINQ-style filtering, shaping, sorting, and aggregation |
| Reuse behavior | Delegates and expression-based functions that can be passed, combined, and invoked |
| Run calculations in parallel | Native Tasks, await-style flow, cancellation, progress, and continuations |
| Model business data | DataTable, DataSet, DataView, relations, change tracking, and providers |
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
external library, no network access, and no runtime VBIDE trust, and it sends no telemetry.
Safe calculations can use Windows worker threads; Excel, VBA procedures, workbook objects, and
COM stay on Excel's thread, and parallel work never opens another Excel.

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
