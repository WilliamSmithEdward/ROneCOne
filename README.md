# ROneCOne

ROneCOne is an experimental, one-file VBA runtime that makes ordinary Excel VBA feel more like
modern C#. Version 0.1.0 ships a complete delegate and expression-lambda foundation with no
runtime install, add-in, network access, external library, or trusted VBIDE access.

## Try it

Open [`demo/ROneCOne_Demo.xlsm`](demo/ROneCOne_Demo.xlsm), press `Alt+F8`, and run
`RunROneCOneDemo`. The workbook executes six examples and a 10,000-call benchmark in one Excel
process.

To use the runtime in another workbook, import [`src/ROneCOne.cls`](src/ROneCOne.cls) through the
VBE's **File > Import File** command. That one class is the entire deployed runtime.

```vba
Dim x As ROneCOne
Dim square As ROneCOne

Set x = ROneCOne.Parameter(vbLong)
Set square = ROneCOne.Lambda(x.Multiply(x), x)

Debug.Print square(CLng(9))      ' 81: default-member delegate call
Debug.Print square.Run(CLng(9))  ' 81: explicit form
```

## Version 0.1.0

- immutable expression trees and string-free anonymous lambdas
- runtime-typed parameters and deterministic contract errors
- unary and binary arithmetic, comparisons, concatenation, and Boolean negation
- short-circuit `AndAlso` and `OrElse` semantics
- object method delegates through `FromMethod`
- scalar and object return values
- delegate composition through `PipeTo`
- callable default-member syntax such as `square(9)`
- IntelliSense descriptions embedded in the exported class

VBA classes cannot declare a public member named `Invoke` because it collides with inherited COM
`IDispatch.Invoke`. ROneCOne uses `Run` as the explicit name and marks it as the default member,
which enables the more C#-like `square(9)` call form.

## Runtime contract

- Windows x64 Microsoft 365 Excel
- `.xlsm`, `.xlsb`, and `.xlam` deployment targets
- one imported runtime file and one Excel application process
- no runtime code generation or VBIDE trust
- no elevation, telemetry, or implicit network traffic
- existing VBA remains usable unchanged
- future local logging is opt-in and never transmits data

## Development

Use Python 3.10 or later. Development dependencies are isolated from the shipped VBA runtime.

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.venv\Scripts\python.exe -m unittest discover -s tests\python -v
.venv\Scripts\pyvbaanalysis.exe src\ROneCOne.cls tests\vba\DelegateFixture.cls `
    tests\vba\TestDelegates.bas --no-inline-suppression --format text
.venv\Scripts\python.exe tools\build_test_workbook.py
powershell -ExecutionPolicy Bypass -File tools\run_excel_tests.ps1
```

The Excel harness observes Win32 and UI Automation surfaces owned by its exact task process. It
captures and dismisses modal dialogs, records selected VBE code for compiler faults, fails fast on
break mode, and enforces a hard deadline so a hidden Excel instance cannot hang indefinitely. See
[`docs/development.md`](docs/development.md).

The real `%USERPROFILE%\.ronecone.env` is private and optional. The committed
`.ronecone.env.example` reserves the local-only logging settings for the future logging slice;
version 0.1.0 does not emit runtime logs.

## Roadmap

Features ship one complete vertical slice at a time: structured exceptions, runtime-generic
collections and query operators, tasks and async/await, events, disposables, and native-safe
parallel operations. See [`docs/architecture.md`](docs/architecture.md).

## License

MIT. See [`LICENSE`](LICENSE).
