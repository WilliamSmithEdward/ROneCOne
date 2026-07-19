# ROneCOne

## Give VBA the love it deserves

Excel carries serious work. ROneCOne gives its built-in language a modern programming layer
without asking you to replace the workbooks, workflows, or skills you already trust.

One class module brings generic collections, LINQ, tasks, typed data, reusable behavior, events,
and structured failure handling into ordinary VBA. There is no add-in to deploy, no service to
connect, and no second runtime to install.

**[See it work in the latest demo release](https://github.com/WilliamSmithEdward/ROneCOne/releases/latest)**

## What becomes possible

| You want to... | ROneCOne gives you... |
|---|---|
| Work with data cleanly | Typed collections and LINQ-style filtering, shaping, sorting, and aggregation |
| Reuse behavior | Delegates and expression-based functions that can be passed, combined, and invoked |
| Coordinate work | Tasks, await-style flow, cancellation, progress, and continuations |
| Model business data | DataTable, DataSet, DataView, relations, change tracking, and providers |
| Build responsive workbook logic | Typed events with predictable subscription and delivery |
| Handle failure deliberately | Structured Try, Catch, and Finally flows |
| Deploy without a platform project | One importable class module with no outside dependency |

ROneCOne runs locally inside one Windows x64 Microsoft 365 Excel process. It works in
macro-enabled workbooks and add-ins, requires no runtime VBIDE access, generates no source code,
and sends no telemetry.

## A modern runtime for serious workbooks

ROneCOne includes mutable, concurrent-style, and immutable generic collection families; a broad
LINQ surface; delegates and expression lambdas; tasks and await-style coordination; cancellation
and progress; typed events; structured exceptions; and an in-memory data layer with OLE DB and
ODBC access. IntelliSense descriptions keep the surface discoverable in the editor.

The design follows familiar C# and .NET names and behavior wherever VBA permits it. The concise
form leads every demo; the guide also shows the underlying contract when you need to debug or
extend it.

## Choose your starting point

1. **[Run the demos](https://github.com/WilliamSmithEdward/ROneCOne/releases/latest)** to see the
   capabilities working in Excel.
2. **[Follow the user guide](docs/user-guide/)** for a short, indexed path from import to useful
   workbook code.
3. **[Download the runtime](src/ROneCOne.cls)** when you are ready to add it to a project.

Technical details remain available in the
**[architecture](docs/architecture.md)** and focused reference documents. Release history is in
the **[changelog](CHANGELOG.md)**.

## License

ROneCOne is available under the [MIT License](LICENSE).
