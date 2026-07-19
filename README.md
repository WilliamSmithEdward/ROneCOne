# ROneCOne

## Give VBA the love it deserves

Excel carries serious work. ROneCOne gives its built-in language a modern programming layer
without asking you to replace the workbooks, workflows, or skills you already trust.

One class module brings modern collections, expressive data queries, reusable behavior, typed
events, and structured error handling into ordinary VBA. There is no add-in to deploy, no service
to connect, and no second runtime to install.

**[See it work in the latest demo release](https://github.com/WilliamSmithEdward/ROneCOne/releases/latest)**

## What becomes possible

| You want to... | ROneCOne gives you... |
|---|---|
| Work with data cleanly | Typed collections and LINQ-style filtering, shaping, sorting, and aggregation |
| Reuse behavior | Delegates and expression-based functions that can be passed, combined, and invoked |
| Build responsive workbook logic | Typed events with predictable subscription and delivery |
| Handle failure deliberately | Structured Try, Catch, and Finally flows |
| Deploy without a platform project | One importable class module with no outside dependency |

ROneCOne runs locally inside one Windows x64 Microsoft 365 Excel process. It works in
macro-enabled workbooks and add-ins, requires no runtime VBIDE access, generates no source code,
and sends no telemetry.

## Available now, designed to grow

Today's runtime includes generic collections, LINQ-style queries, expressive predicates,
delegates, expression lambdas, typed events, and structured exceptions. IntelliSense descriptions
make the surface discoverable inside the editor.

The direction is a fuller modern language experience inside VBA: broader generic collections,
task-based workflows, async and await-style coordination, cancellation, progress, disposables,
and safe concurrency within one Excel process.

Each capability must preserve the same promise: one file, local execution, predictable behavior,
and as little ceremony as VBA can support.

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
