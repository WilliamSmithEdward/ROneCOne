# Security policy

## Reporting a vulnerability

Report a vulnerability privately through GitHub: open the repository's
[Security tab](https://github.com/WilliamSmithEdward/ROneCOne/security) and choose **Report a
vulnerability**. Please do not open a public issue for one. Say which release you used, include
the code or workbook that shows the problem, and describe what you expected to happen.

The report stays private while it is looked into. A fix ships in a new release, and the advisory
is published with it.

## Supported versions

Fixes ship in the next release. Only the latest release receives them; earlier releases are not
patched separately.

## What the runtime can reach

`ROneCOne.cls` runs only when your code calls it. It has no `Workbook_Open`, `Auto_Open`, or event
handler that Excel starts on its own, and it reaches the machine only through the members your
code uses:

| Reach | Only through |
|---|---|
| Network | `HttpClient` over WinHTTP for the URLs you request, and `Xml.Load` if you pass it a URL |
| Processes | `Process.RunAsync`, which runs your command line through `cmd.exe` and captures its output in temporary files that it deletes, and `Process.StartSession`, which keeps `cmd.exe` running and talks to it over pipes |
| Files | `File`, `Directory`, `Path`, `Csv`, `ZipFile`, `Logger`, `FileWatcher`, `Xml.Load`, and `HttpClient.DownloadFileAsync`, on paths you pass |
| Databases | `DbConnection` and the other ADO members, with the connection string you supply |
| Native code | `Declare PtrSafe` calls into `kernel32`, `oleaut32`, `ole32`, and `bcrypt` |

It sends no telemetry, never reads or writes the registry, never uses the VBIDE at run time, and
never starts a second Excel. The demo workbooks go further only because their demos do: the HTTP
demo downloads from `https://pokeapi.co`, the Process demo runs commands through `cmd.exe`, the
Zip demo runs PowerShell's `Compress-Archive` and `Expand-Archive`, and several demos write
scratch files beside the workbook and delete them.

## Verifying a release

Every release carries `vX.Y.Z-sha256.txt`, the SHA-256 hash of `ROneCOne.cls` and each demo
workbook, and from 1.10.2 on, `vX.Y.Z-security-report.md`. The report lists what olevba and
MacroRaptor find in every one of those files, their hashes, and whether the results match the
reviewed baseline at the release tag.

## How the code is scanned

The [Security workflow](.github/workflows/security.yml) runs `tools/security_scan.py` on every
push and pull request. It scans `ROneCOne.cls` and every demo workbook with olevba and
MacroRaptor (mraptor), both from oletools, and compares each file's olevba findings and mraptor
flags with `tools/security_baseline.json`. The run fails on any change in either direction, on
any code that runs on its own, such as a `Workbook_Open` handler, and on any P-code its source
does not explain. A new finding is accepted only by reviewing it and updating the baseline in the
same change. The baseline never records code that runs on its own or P-code its source does not
explain.

olevba and mraptor report some results on this code by design:

- **Suspicious keywords** such as `Shell`, `Lib`, `CreateObject`, `Kill`, and `SaveToFile` name
  the capabilities in the table above.
- **IOCs** are the libraries and programs those capabilities use: `bcrypt.dll`, `oleaut32.dll`,
  and `cmd.exe`.
- **Hex and Base64 strings** are ordinary words and numbers in the code, such as `SHA1` and
  `2147483647`, that happen to decode.
- **VBA stomping** is flagged on the demo workbooks because pcodedmp prints eight names with a
  trailing type character, such as `payload$`, that the source never spells. Every name without
  its suffix is in the source. The scan checks every other name and string in the P-code against
  the source, and a missing one fails the run.
- **mraptor's W and X flags** mark code that writes files or memory, such as `SaveToFile` and
  `Kill`, and code that runs something outside VBA, such as `CreateObject`, `Shell`, and the
  `Declare` statements. Every file is `-WX`. mraptor calls a file suspicious only when its A
  flag, for code that runs on its own, joins W or X, and no file sets A, so every verdict is
  Macro OK.

The [CI workflow](.github/workflows/ci.yml) also runs pyVBAanalysis over the sources and every
demo workbook.
