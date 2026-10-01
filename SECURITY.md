# Security policy

## Reporting a vulnerability

Report a vulnerability privately, not in a public issue or pull request:
[open a private report](https://github.com/WilliamSmithEdward/ROneCOne/security/advisories/new).
Only the maintainer sees it. Include the release you used, the code or
workbook that shows the problem, and what you expected to happen, with
credentials and private data removed.

A confirmed vulnerability is fixed in a release on the GitHub releases
page, and the advisory is published with it,
crediting you unless you ask otherwise.

## Supported versions

Only the latest release on the GitHub releases page receives security
fixes. Older releases are not maintained separately; update when a fix
ships.

## Scope

The runtime is `src/ROneCOne.cls`, one VBA class that runs inside Excel
with the rights of the person who opens the workbook. The demo workbooks
in `demo/` embed it. A vulnerability here is the runtime reaching
something your code did not ask it to, or doing more with what you pass
it than its documentation says.

### What the runtime can reach

`ROneCOne.cls` runs only when your code calls it. It has no
`Workbook_Open`, `Auto_Open`, or event handler that Excel starts on its
own, and it reaches the machine only through the members your code uses:

| Reach | Only through |
|---|---|
| Network | `HttpClient` over WinHTTP for the URLs you request, and `Xml.Load` if you pass it a URL |
| Processes | `Process.RunAsync`, which runs your command line through `cmd.exe` and captures its output in temporary files that it deletes, and `Process.StartSession`, which keeps `cmd.exe` running and talks to it over pipes |
| Files | `File`, `Directory`, `Path`, `Csv`, `ZipFile`, `Logger`, `FileWatcher`, `Xml.Load`, and `HttpClient.DownloadFileAsync`, on paths you pass |
| Databases | `DbConnection` and the other ADO members, with the connection string you supply |
| Native code | `Declare PtrSafe` calls into `kernel32`, `oleaut32`, `ole32`, and `bcrypt` |

It sends no telemetry, never reads or writes the registry, never uses the
VBIDE at run time, and never starts a second Excel. The demo workbooks go
further only because their demos do: the HTTP demo downloads from
`https://pokeapi.co`, the Process demo runs commands through `cmd.exe`,
the Zip demo runs PowerShell's `Compress-Archive` and `Expand-Archive`,
and several demos write scratch files beside the workbook and delete them.

## How the code is checked

Three workflows check every pull request and every push to `main`, and
their gates decide whether a change can merge: **CI passed**,
**Security passed** and **Malware scan passed**. A gate passes only when
every job before it did, and any unexpected finding fails it, whatever its
severity. Security and Malware scan also run daily at 19:17 UTC.

- **Code:** olevba and mraptor (MacroRaptor), both from oletools, scan
  `ROneCOne.cls` and every demo workbook. `tools/security_scan.py` compares
  each file's olevba findings and mraptor flags with the reviewed baseline
  and fails on any change in either direction, on any code that runs on its
  own, such as a `Workbook_Open` handler, and on any P-code its source does
  not explain. olevba's own VBA stomping flag is replaced by that last
  check, which sets aside the trailing type character pcodedmp prints on
  some names (`payload$`) and fails on any other P-code name or string the
  source lacks. The results are in the job's log, not in code scanning. CI
  also runs pyVBAanalysis over the sources and every demo workbook.
- **Workflows:** zizmor audits the GitHub Actions workflows; a finding fails
  Security.
- **Dependencies:** there is nothing to audit. The runtime has no
  dependencies, and the Python tools the workflows run come from hash-locked
  files (see Pinning and updates).
- **Malware:** ClamAV, with signatures freshclam fetches and verifies on
  every run, and YARA-X, with the YARA Forge rules pinned to a release and
  its SHA-256, scan the shipped `src/ROneCOne.cls` and every
  `demo/*.xlsm`. YARA-X scans each workbook as a file and also scans its
  decompressed ZIP members, including `xl/vbaProject.bin`; ClamAV handles
  its own archive inspection. A new detection fails Malware scan, and so
  does a scanner, signature update, rule download, compilation, or workbook
  read error. A match from a public collection can be heuristic: it
  warrants review, not an automatic malware verdict.
- **OpenSSF Scorecard** rates the repository's security practices on every
  change to `main` and weekly, and the README badge shows the result.
  Some of its checks do not fit this project. A single maintainer cannot
  have a second person approve every change. The class module and
  workbooks are built locally rather than by CI, so a release carries
  `vX.Y.Z-sha256.txt` and the security report rather than a build
  provenance signature. Fuzzing does not apply: ROneCOne is VBA, which runs
  only inside Office.

## Accepted findings

A finding is fixed, or accepted with a written reason in
`tools/security_baseline.json` (olevba and mraptor) or
`tools/malware_exceptions.json` (ClamAV and YARA-X).

The baseline records each shipped file's olevba findings and mraptor flags
by file name, not by the file's content. A new finding is accepted only
by reviewing it and updating the baseline in the same change, and a
recorded finding that no longer appears fails the scan too. The baseline
never records code that runs on its own or P-code its source does not
explain.

A malware exception names the `scanner`, `path`, `detection`, the shipped
file's `sha256`, and a specific `reason`. For a workbook member, `path`
has the form `demo/Name.xlsm!xl/vbaProject.bin`, and its hash is the whole
workbook's hash. An exception applies only to that detection in those
exact bytes, so a changed file needs another review, and an exception that
no longer matches fails the scan. A failed signature download or scanner
error is never bypassed with an exception.

zizmor keeps its exceptions in `.github/zizmor.yml` or inline beside the
line they excuse, each with its reason. There are none: the repository has
no `.github/zizmor.yml` and no inline exceptions.

The baseline holds these results, by design:

- **Suspicious keywords** such as `Shell`, `Lib`, `CreateObject`, `Kill`,
  and `SaveToFile` name the capabilities in the table under Scope.
- **IOCs** are the libraries and programs those capabilities use:
  `bcrypt.dll`, `oleaut32.dll`, and `cmd.exe`.
- **Hex and Base64 strings** are ordinary words and numbers in the code,
  such as `SHA1` and `2147483647`, that happen to decode.
- **VBA stomping** is flagged on the demo workbooks because pcodedmp
  prints eight names with a trailing type character, such as `payload$`,
  that the source never spells. Every name without its suffix is in the
  source.
- **mraptor's W and X flags** mark code that writes files or memory, such
  as `SaveToFile` and `Kill`, and code that runs something outside VBA,
  such as `CreateObject`, `Shell`, and the `Declare` statements. Every file
  is `-WX`. mraptor calls a file suspicious only when its A flag, for code
  that runs on its own, joins W or X, and no file sets A, so every verdict
  is Macro OK.

The malware exception list is empty. The 1.10.2 demo workbooks matched
YARA Forge rule `ARKBIRD_SOLG_TA505_Maldoc_21Nov_2` on four ordinary
Office/VBA library reference strings; its sample-specific paths and long
payload strings did not match. Repackaging the 1.10.3 workbooks removed
the match. ClamAV found no infections in the repackaged files when scanned
with engine 1.5.4 on 2026-09-29.

A rule that matches VBA's `ReDim` or `ReDim Preserve` is matching expected
code: `SessionAppendBytes` doubles a byte buffer's capacity before
appending data, and the CSV parser grows its field and quote-state arrays
when a row has more than eight fields. Such a match still needs its full
condition and the surrounding code reviewed before any exception is added.

## Pinning and updates

Everything the workflows run is pinned: actions to full commit SHAs,
runners to named OS releases, the ClamAV image to a digest, Python tools
(oletools, YARA-X, zizmor, pyVBAanalysis) to hash-locked lock files in
`.github/requirements/`, the local development tools to the hash-locked
`requirements-dev.txt`, and the YARA Forge rules to a release and its
SHA-256 in `.github/security/yara.json`. ClamAV's signatures change too
often to pin, so freshclam fetches and verifies them on every run.

Dependabot proposes updates to GitHub Actions, the Python lock files
(`.github/requirements/` and `requirements-dev.txt`) and the ClamAV image
once a version is a week old (the owner's own packages, such as
pyVBAanalysis, at once), and at once for a security advisory. The Update
YARA rules workflow proposes new YARA pins each week. A minor or patch
update, and the YARA pull request, merges itself once CI, Security and
Malware scan pass; a third-party major version waits for review.

## Releases

The class module and demo workbooks are built locally, and the GitHub
release carries `ROneCOne.cls`, every demo workbook, and
`vX.Y.Z-sha256.txt`, the SHA-256 of each of those files.

Publishing the release starts `.github/workflows/release-security.yml`. It
downloads the release's `.cls` and `.xlsm` files, scans them with olevba
and mraptor against the baseline at the release tag, and attaches
`vX.Y.Z-security-report.md`: what olevba and mraptor find in every file,
their hashes, and whether the results match the reviewed baseline.
Releases from 1.10.2 on carry it. Started by hand with a release's tag,
the workflow is a dry run and attaches nothing.

### Verifying a release

Compare a download's SHA-256 with `vX.Y.Z-sha256.txt` or the security
report:

```powershell
Get-FileHash .\ROneCOne.cls -Algorithm SHA256
```

## Repository settings

<!-- repo-standards:begin security-settings. Copied from WilliamSmithEdward/repo-standards, templates/security/settings-block.md. Change it there; the weekly rescan fails a copy that differs. -->
- `main` accepts changes only through a pull request that passes
  **CI passed**, **Security passed** and **Malware scan passed**. The
  ruleset has no bypass, for the owner either, and refuses force-pushes and
  deleting the branch.
- A `v*` release tag cannot be moved or deleted once pushed, except by a
  repository admin.
- A workflow that uses an action not pinned to a full commit SHA fails to
  run. Workflow tokens are read-only unless a job is granted more for
  itself.
- Secret scanning with push protection, Dependabot alerts and security
  updates, and private vulnerability reporting are on.
<!-- repo-standards:end -->
