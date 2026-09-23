# Changelog

All notable changes to ROneCOne are documented here. The format is based on
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and release numbering follows
[Semantic Versioning](https://semver.org/spec/v2.0.0.html). Downloadable assets and SHA-256
checksums for each version are on the
[releases page](https://github.com/WilliamSmithEdward/ROneCOne/releases).

## Unreleased

### Changed

- `ROneCOne.cls` no longer declares variables that nothing reads, found by the dead-code checks
  in pyVBAanalysis 2.2.0. Two linked-list node fields unused since 1.0.0, four unused locals, and
  an error source the HTTP poll saved but never raised are gone, and eight calls that parked a
  result in a throwaway variable are now plain statements. Behavior is unchanged. The Collections,
  Data, HTTP, and Tasks demos drop the same throwaway variables, and the Data and Tasks worksheets
  show the plain calls.

- The Collections, Data, and Excel Table demos no longer declare locals such as
  `Dim Age As Variant` for the names they reach through bang syntax. pyVBAanalysis versions
  before 2.2.0 reported `!Age` as an undeclared variable without them.

## 1.10.0 - 2026-09-22

### Fixed

- The demo workbooks fit their text in Excel. The renders used to review them draw text in
  narrower fonts than Excel does, so they hid most of what was wrong. Measured with Excel's own
  AutoFit, the 1.9.1 workbooks split or clipped 74 cells, soft-wrapped 50 code lines
  mid-expression, and cut off text in rows too short to hold it, among them every capability
  demo's title band. Code lines longer than their column now break with a VBA line
  continuation, each example row takes its height from its longest cell, every table row has an
  explicit height, and value columns widen only where a timestamp, an escaped string, or a
  joined list needs the room. Five snippets that already carried a statement onto a second line
  lacked the ` _` continuation VBA requires, and now end the line with it. The two SHA-256
  digests on the Text sheet still wrap mid-string, inside rows sized to hold them.

- Expected Booleans no longer show as checkboxes beside the plain TRUE the macro writes. The
  workbook builders stored a JavaScript Boolean as an Excel checkbox cell, and now write an
  expected Boolean as `=TRUE` or `=FALSE`.

- The capability demos' Examples titles read naturally, such as "Live HTTP examples" and "Live
  file and CSV examples" in place of "Live http + async examples" and "Live files + csv
  examples". Their Status columns show PASS in green and CHECK in red, as the Delegates and
  Collections demos already did, and NOT RUN now shows in amber in all sixteen demos.

- The HTTP benchmark labels its overlapped and sequential timings and shows both to six
  decimals, and its note and the demo module's comments say what overlapping can save instead
  of promising a saving the measurement does not always show. The Collections benchmark
  subtitle names both scales it measures.

### Changed

- The Excel Table demo serializes its sample row with `ToJson(True)`, so the JSON shows indented
  and wraps at its own line breaks.

## 1.9.1 - 2026-09-22

### Changed

- Every shipped module now opens with the release it belongs to and the full MIT license, directly
  below `Option Explicit`. That covers `ROneCOne.cls` and each module packaged into the demo
  workbooks. A demo module copied out of its workbook used to carry no license and no version,
  and the runtime's license sat on line 208, below the Windows declarations.
  `tools/stamp_release_headers.py` writes the header from the newest dated heading in this
  changelog and from `LICENSE`, and a source contract fails while any module disagrees with
  either.

### Fixed

- Importing `ROneCOne.cls` no longer rewrites the host's own modules, closing
  [issue #6](https://github.com/WilliamSmithEdward/ROneCOne/issues/6). VBA keeps one spelling
  per identifier across a project, and a declaration anywhere sets it. The runtime declared more
  than a thousand parameters and locals in lowercase whose names Excel, Office, or VBA spell
  otherwise, so `.Value` became `.value` and `.Text` became `.text` in every module of the
  project, and each export carried the change as diff noise. Exported through the VBE beside a
  host module that declares none of those names, 24 names in the host module came back recased,
  `Count`, `Item`, `Rows`, and `Value` among them, and the runtime recased its own members too:
  `Capacity`, `Columns`, and `Result` came back in lowercase. The same rule made ROneCOne and
  ModernJsonInVBA recase each other when they shared a project.

  Public parameters keep their .NET names and take the spelling the type libraries use, as
  Excel's own parameters do: `Add(Value)`, `Item(Index)`, `WhereNot(Predicate)`,
  `Contains(Value, Comparer)`. VBA matches named arguments regardless of case, so no call site
  changes. Parameters of Private and Friend members, locals, UDT fields, and Declare parameters
  were renamed to names no reference defines, such as `value` to `itemValue`, `index` to `idx`,
  and `result` to `outcome`. The JSON parser's private types, adapted from ModernJsonInVBA and
  still named after its `JsonReader` and Public `JsonTextBuilder`, are now `JsonScanState` and
  `JsonWriteBuffer`.

  Two members keep their .NET spelling where Office capitalizes the whole word, `Guid` and
  `Xml`, so a host that writes `GUID` or `.XML` still sees those two recased.
  `tests/python/test_casing.py` holds the runtime to one spelling per name and to the spelling
  the default references use, and `tools/run_casing_roundtrip.ps1` exports a host project
  through the VBE and requires every token back as written.

  Replacing an older class restores the names this runtime declares, but not a name it never
  declares: VBA keeps that spelling after the class that set it is removed. Measured on a host
  saved beside 1.9.0 and then upgraded, 64 names such as `.cells`, `.names`, and `.width` kept
  the old spelling. The upgrade section of
  [Getting started](docs/user-guide/getting-started.md#upgrading-from-an-earlier-release)
  has a procedure to paste and delete that restores all of them.

- The demo modules no longer recase Excel's names or ROneCOne's own inside their workbooks.
  `Dim json As String` in the HTTP demo turned `ROneCOne.Json` into `ROneCOne.json`, and the
  `description` parameter in every demo's failure handler recased `Err.Description`. Locals and
  parameters across all eighteen demo modules now take names no reference defines, and the
  worksheet snippets quote the new names. Snippets that used a colliding placeholder, such as
  `path`, `text`, or `task`, use one that does not. The casing tests and the VBE round trip
  cover the demos beside the runtime, and every demo workbook was rebuilt and rerun.

- The VBA examples in the documentation no longer teach names that recase a reader's project.
  Copied into a module, `Dim task As ROneCOne` recased ROneCOne's `Task`, `Dim row` recased
  Excel's `.Row`, and a variable named after the procedure it wraps, such as `skipBadRow` for
  `SkipBadRow`, recased that procedure. The examples across nineteen pages now use names no
  library or ROneCOne member defines, matching the demos where an example mirrors one, and prose
  that names one of those variables follows it. `test_casing.py` checks every example, one
  spelling per page included.

## 1.9.0 - 2026-08-02

### Added

- Excel Tables are a first-class input. `DataTableFromRange`, `ListFromRange`, `LoadFromRange`,
  and `ToRange` all accept a `ListObject` or a single `ListColumn` directly. Previously they took
  only a `Range`, so passing a Table raised run-time error 438 and the workaround was to pass
  `listObject.Range` yourself. The runtime had never heard of the type: searching the source for
  `ListObject` returned nothing, and neither did the documentation. With headers wanted the slice
  stops after the last body row, so a totals row is never read as data, and an empty Table yields
  its columns with no rows rather than failing.

- `ROneCOne.Table(listObject)` reads a Table that stays attached to where it came from.
  `Refresh` re-reads it in place, and `WriteBack` writes rows into the Table and resizes it to
  fit. `ToRange` given a Table as its target performs the same write, so a `DataView` can drive
  it. Excel makes this awkward in four specific ways, all measured live before anything was
  written: it refuses to resize a Table to its header alone, it leaves the vacated cells
  populated when a Table shrinks, it places a totals row inside the requested range on a grow and
  outside it on a shrink, and it preserves name, style, and headers throughout. `WriteBack`
  handles the first three and relies on the fourth. See
  [ADR 0030](docs/decisions/0030-excel-table-surface.md).

- A dedicated Excel Table demo workbook and an
  [Excel Tables guide](docs/user-guide/excel-tables.md). Forty-four examples take the Table
  through reading, the full LINQ set over its rows, mapping out to instances of your own class
  and back with the round trip proved, JSON and CSV in both directions, and writing back into the
  Table while it grows and shrinks on the sheet.

### Fixed

- Every JSON serialization left a stray error behind, closing
  [issue #5](https://github.com/WilliamSmithEdward/ROneCOne/issues/5). This is the `IsArray`
  hazard from #4 in a second intrinsic: `VarType` also evaluates its argument in a value context,
  so a Variant holding a runtime value was dereferenced through its default member, `Run`, which
  rejected the zero arguments that dereference supplies. `VarType` then reported `vbObject`
  precisely because that call failed, so serialization returned the right text while leaving `Err`
  dirty for any caller running under `On Error Resume Next`. The JSON writer opens with a `VarType`
  test on the value it was handed, so lists, dictionaries, tables, rows, and views were all
  affected; `ToCsv` was not. All 52 type tests now go through a guarded `VarTypeOf` helper that
  checks `IsObject` first, and a source contract asserts `VarType` appears exactly once in the
  runtime, inside that helper. The guard also removes an accidental dependency: the runtime reached
  its object branches only because the default member happened to fail, which would not hold for a
  foreign object whose default member succeeds. One behavior changes as a result, in favour of the
  documented contract: `Json.Serialize(someRange)` used to serialize the cell's value and now
  raises, as the docs have always said it would.

- `ToJson(True)` ignored indentation for tables, rows, and views, contradicting the documented
  promise that indented output is available for every serializable value. Only dictionaries and
  sequences carried line structure; the data roles fell through to the compact writer and returned
  text identical to `ToJson()`. The indented writer now covers them.

### Changed

- The source contract that pins `Friend` member access now rejects any `.Item(...).Member` chain
  reaching a `Friend` member, not only reads through a declared `Object` or `Variant` local. Both
  `Collection.Item` and this class's own `Item` return a `Variant`, so such a chain binds late and
  raises 438 at run time with no late-bound declaration in sight. That is exactly how the defect
  from [issue #3](https://github.com/WilliamSmithEdward/ROneCOne/issues/3) reappeared inside the
  new write-back path during development, and the analyzer rather than the contract caught it.

## 1.8.1 - 2026-08-01

### Fixed

- `IsArray` left a stray error behind when handed a Variant holding a runtime value, closing
  [issue #4](https://github.com/WilliamSmithEdward/ROneCOne/issues/4). `IsArray` evaluates its
  argument in a value context, so VBA dereferenced the object through its default member, `Run`,
  which rejected the zero arguments that dereference supplies. The raise was swallowed wherever a
  caller had `On Error Resume Next` active, so calls returned correct answers while leaving `Err`
  dirty for the caller to trip over later. Every array test now goes through a guarded
  `IsArrayValue` helper that checks `IsObject` first, and a source contract asserts `IsArray`
  appears exactly once in the runtime, inside that helper. Two further call sites were exposed to
  the same hazard, in `DeserializeOnly` and in primary-key argument building.

- `OneOf` given a single ROneCOne sequence raised run-time error 438 instead of building the
  membership condition, closing
  [issue #3](https://github.com/WilliamSmithEdward/ROneCOne/issues/3). `IsSequenceContainer` read
  the Friend `InternalRole` through an `Object` local, and VBA keeps Friend members off the
  IDispatch interface, so the read compiled clean and failed only when executed, even though the
  caller is this very class. Binding through a typed `ROneCOne` local fixes it. Every previous
  `OneOf` test passed an array or two scalars, so nothing reached the branch;
  `customers.Where("City").OneOf(cities)` now works and is covered live. A source contract sweeps
  the whole runtime for Friend members read through `Object` or `Variant` locals so the class of
  defect cannot return.

## 1.8.0 - 2026-07-28

### Added

- LINQ that runs on the server: `connection.Queryable(tableName)` compiles the runtime's own
  expression trees into parameterized SQL, so a filter happens in the database instead of after
  loading every row into Excel. `Where`, `OrderBy`, `OrderByDescending`, `ThenBy`,
  `ThenByDescending`, `SelectColumns`, `Take`, and `Skip` compose immutably; `Count`, `AnyItem`,
  `FirstOrDefault`, `ToDataTable`, `ToDataTableAsync`, and `CountAsync` execute; `ToSqlString` and
  `SqlParameterValues` render exactly what will be sent. Every captured constant leaves as a `?`
  marker, so a value can never be read as SQL. A `Null` becomes `IS NULL`, string helpers become
  dialect-escaped `LIKE`, and `IsIn` becomes an `IN` list. SQL Server and ACE differ in row
  limits, offsets, wildcard escaping, and the uppercase function, all settled by live probe.
  Anything untranslatable refuses with the typed `ROneCOne.QueryError` rather than falling back to
  a client-side scan. See [ADR 0027](docs/decisions/0027-linq-to-sql-over-expression-trees.md).
- Interactive command sessions: `ROneCOne.Process.StartSession(command, [workingDirectory],
  [encodingName])` keeps one `cmd.exe` alive so a workbook can hold a conversation with it.
  `WriteAsync` and `WriteLineAsync` queue input, `CloseInput` signals end of file,
  `ReadLineAsync` and `ReadErrorLineAsync` await the next line on each stream separately,
  `ReadAvailable` and `ReadErrorAvailable` take what is buffered without waiting, `ReadToEndAsync`
  collects the rest, and `WaitForExitAsync` resolves to the exit code, alongside `HasExited`,
  `ExitCode`, `ProcessId`, and `KillProcess`. Nothing blocks Excel: reads size themselves from
  `PeekNamedPipe`, and writes ride an overlapped named pipe so an oversized payload pends instead
  of freezing the host. A line read takes an optional timeout, because a prompt arrives with no
  trailing newline. See [ADR 0028](docs/decisions/0028-interactive-process-sessions.md).

- JSON partial reads, closing
  [issue #1](https://github.com/WilliamSmithEdward/ROneCOne/issues/1): a large response is often
  almost entirely members you will never read, and the parser used to build all of them.
  `Json.DeserializeOnly(text, paths)` walks the document once and materializes only the requested
  paths, preserving structure so navigation matches `Deserialize` exactly, and
  `Json.DeserializeAt(text, path)` returns the single value one path addresses. `DeserializeTable`
  and `DeserializeObjects` now apply `arrayPath` during the scan instead of parsing the whole
  document and navigating the result. A missing path is omitted by `DeserializeOnly` and raises
  from `DeserializeAt`. See [ADR 0029](docs/decisions/0029-json-partial-reads.md).

### Fixed

- `DeserializeInto` and `DeserializeObjects` no longer raise on a JSON member the target class
  does not model. `CallByName` had no error trap, so error 438 made binding unusable against a
  real API response; unknown members are now ignored, matching System.Text.Json.

### Host boundaries

- VBA reserves `Kill` for its file-deletion statement, so a session terminates through
  `KillProcess`, joining `Connect`, `Disconnect`, `YieldOnce`, `IsIn`, and `SingleItem`.
- VBA keeps one global casing per identifier, so declaring a lowercase `name` anywhere rewrites
  every bare `Name` in the project and breaks bang syntax against a case-sensitive column lookup.
  A source contract now refuses that declaration.

## 1.7.0 - 2026-07-24

### Added

- Fourteenth demo workbook. `ROneCOne_Zip_Demo.xlsm` builds a small tree, zips it, reads an
  entry without extracting, proves interop both ways against PowerShell (opening a
  Compress-Archive deflate archive and handing its own output to Expand-Archive), shows the
  directory-traversal guard refusing a hostile entry, and gates a benchmark that inflates a
  1,000-line archived file. All offline.
- Zip archives over a pure-VBA engine, in the spirit of System.IO.Compression:
  `ROneCOne.ZipFile.OpenRead` lists and reads entries (`Entries`, `GetEntry`, and per-entry
  `FullName`, `Name`, `Length`, `CompressedLength`, `ReadAllText`, `ReadAllBytes`,
  `ExtractToFile`), inflating stored and deflated data with CRC-32 verification;
  `CreateFromDirectory` writes a store-only archive and `ExtractToDirectory` unpacks with a
  directory-traversal guard and no overwrite. Zip64 and encrypted archives are refused with the
  typed `ROneCOne.ZipError`. The engine was built against generated fixtures covering every
  deflate block flavor and verified live against PowerShell Compress-Archive and Expand-Archive
  both directions. See [ADR 0023](docs/decisions/0023-zip-over-a-pure-vba-engine.md).
- A file logger in the spirit of Microsoft.Extensions.Logging: `ROneCOne.Logger(path,
  [minimumLevel])` writes UTC millisecond-stamped, level-coded, composite-formatted lines with
  `LogTrace` through `LogCritical`, `IsEnabled`, and `MinimumLevel`; sub-minimum calls never
  touch the disk and each written line is appended and flushed to survive a crash. See
  [ADR 0024](docs/decisions/0024-file-logger.md).
- Awaitable folder watching: `ROneCOne.FileWatcher(folder, [filter]).WaitForChangeAsync` returns
  a Task that resolves on the next Created, Changed, or Deleted file, snapshotting at call time
  and rescanning on a throttle while awaited; cancellation and timeouts compose like every task.
  See [ADR 0025](docs/decisions/0025-awaitable-file-watching.md).
- URL and HTML escaping, and process input: `ROneCOne.Uri.EscapeDataString` /
  `UnescapeDataString` (RFC 3986 over UTF-8) and `ROneCOne.WebUtility.HtmlEncode` / `HtmlDecode`
  (the five entities plus numeric references, through surrogate pairs); `Process.RunAsync` gains
  an optional `standardInput` that is written and closed at start so filters like `sort` see
  end-of-file. See [ADR 0026](docs/decisions/0026-escaping-helpers-and-process-input.md).

## 1.6.0 - 2026-07-24

### Added

- Twelfth and thirteenth demo workbooks, plus a wider text demo.
  `ROneCOne_DateTime_Demo.xlsm` parses ISO 8601 and epoch timestamps offline, re-views
  instants across offsets, clamps calendar arithmetic, subtracts instants into durations,
  shows the typed rejection of impossible dates, and gates a benchmark that round-trips
  1,000 timestamps. `ROneCOne_Xml_Demo.xlsm` parses and queries XML offline with XPath and
  namespaces, refuses a DOCTYPE with the typed error, lands a catalog in a typed DataTable,
  writes it back with `ToXml`, and gates a benchmark extracting a 1,000-row table.
  `ROneCOne_Text_Demo.xlsm` adds invariant `Strings.Format` rows, a `StringBuilder` chain,
  and GUID plus random-byte minting.
- XML in the spirit of System.Xml: `ROneCOne.Xml.Parse(text, [selectionNamespaces])` and
  `Load(path, [selectionNamespaces])` wrap a secured MSXML6 document (DTDs prohibited,
  external references unresolved) and return nodes with `Name`, `Value`, `GetAttribute`,
  `HasAttribute`, `Elements`, `SelectNodes`, `SelectSingleNode`, and `OuterXml`; failures
  raise the typed `ROneCOne.XmlError` with the parser's line, position, and reason.
  `Xml.DeserializeTable(text, [tableName], [rowsPath])` turns row elements into a typed
  DataTable through the same deterministic inference CSV uses, and `table.ToXml` writes
  round-trippable element-per-column rows in the DataTable.WriteXml shape, omitting nulls.
  See [ADR 0022](docs/decisions/0022-xml-over-msxml6.md).
- Invariant text formatting and identity helpers: `ROneCOne.Strings.Format` implements the
  String.Format grammar (`{index,alignment:format}`, brace escapes, `G N F D X P` specifiers,
  date tokens) with invariant output on every machine; `ROneCOne.StringBuilder()` builds long
  text in linear time with `Append`, `AppendLine`, `AppendFormat`, `Length`, `Clear`, and
  `ToString`; `ROneCOne.Guid.NewGuid` mints canonical version 4 GUIDs; and
  `ROneCOne.RandomNumberGenerator` draws crypto-grade `GetBytes` and rejection-sampled
  `GetInt32` from Windows CNG. Malformed templates raise the typed `ROneCOne.FormatError`.
  See [ADR 0021](docs/decisions/0021-invariant-formatting-and-identity.md).
- Dates and times in the spirit of DateTimeOffset: `ROneCOne.DateTime` parses full ISO 8601
  with offsets (`Parse`/`TryParse`), reads the clock (`UtcNow`, `Now`, `Today`), converts Unix
  epochs and VBA dates both ways, does calendar arithmetic with month-end clamping, re-views
  instants (`ToUniversalTime`, `ToLocalTime`, `ToOffset`), compares by instant, and formats
  round-trip ISO text or a `yyyy MM dd HH mm ss fff` token pattern. `ROneCOne.TimeSpan`
  carries signed durations with totals, .NET-signed components, algebra, and a d.hh:mm:ss
  text round trip. Windows performs every zone conversion per instant through kernel32; the
  runtime hard-codes no offsets and no daylight saving rules. Text failures raise the typed
  `ROneCOne.FormatError`. See
  [ADR 0020](docs/decisions/0020-datetime-over-kernel32-time.md).
- Regular expressions in the spirit of System.Text.RegularExpressions:
  `ROneCOne.Regex(pattern, [ignoreCase], [multiLine])` over the in-box script engine, with
  `IsMatch`, `Match` (a typed match carrying `Success`, `Value`, `FirstIndex`, `Length`, and a
  `Groups` list), `Matches` (a queryable list), `Replace` with `$1` group substitution, and
  `Split` that ignores zero-length matches. Invalid patterns raise the typed
  `ROneCOne.RegexError` at creation. See
  [ADR 0018](docs/decisions/0018-regex-over-vbscript-regexp.md).
- Hashing and encoding in the spirit of System.Security.Cryptography and System.Convert:
  `ROneCOne.Hash` computes `Sha256`, `Sha512`, `Sha1`, `Md5`, and `HmacSha256` through Windows
  CNG (no reference, no registration), over text (as UTF-8) or bytes; `ROneCOne.Convert` adds
  `ToBase64String`, `FromBase64String`, `ToHexString`, and `FromHexString`. Digests match the
  FIPS 180 and RFC 4231 vectors. See
  [ADR 0019](docs/decisions/0019-hashing-over-windows-cng.md).
- `HttpClient.DownloadFileAsync(url, filePath)` saves a response body straight to disk,
  composing the byte-array transfer with the file layer and faulting on non-success like the
  other body-consuming verbs.
- Ninth and tenth demo workbooks. `ROneCOne_Files_Demo.xlsm` walks the file and CSV surface
  offline under one self-cleaning folder: UTF-8 and byte-order-mark round trips, lines as
  collections, path helpers, recursive wildcard enumeration, a typed CSV file round trip with
  null preservation, and a gated benchmark round-tripping 1,000 rows through a CSV file.
  `ROneCOne_Process_Demo.xlsm` runs cmd.exe built-ins only: awaited echo with captured
  output, exit-code passthrough, separated standard error, per-command working directories,
  WhenAll overlap, honest failure shapes, and a gated benchmark overlapping three commands.
- Awaitable shell commands: `ROneCOne.Process.RunAsync(command, [workingDirectory],
  [cancellationToken])` starts the command immediately through the Windows Script Host and
  returns a Task whose result carries `ExitCode`, `StandardOutput`, and `StandardError`.
  Output streams redirect to scratch files that the runtime reads and deletes, so no output
  volume can deadlock Excel; several commands overlap under `WhenAll`; cancellation
  terminates the process. A command that fails reports through its exit code like
  System.Diagnostics.Process; only a start failure raises the typed `ROneCOne.ProcessError`.
  See [ADR 0017](docs/decisions/0017-awaitable-processes-over-wscript-exec.md).
- CSV exchange over the data layer: `table.ToCsv` and `ROneCOne.Csv.Serialize` write a
  DataTable or DataView to RFC 4180 text with minimal quoting, invariant numbers, ISO 8601
  dates, and a round-trippable distinction between database nulls (empty field) and empty
  strings (quoted pair); `ROneCOne.Csv.DeserializeTable` parses CSV into a typed DataTable
  with strict quote discipline, tolerant line endings, and deterministic column inference
  (integers widening to `LongLong` and `Double`, booleans, validated ISO dates; quoted cells
  and mixed columns stay text with their original characters). Failures raise the typed
  `ROneCOne.CsvError`. The number reader behind JSON and CSV now converts fraction and
  exponent text with the locale-independent `Val` instead of `CDbl`. See
  [ADR 0016](docs/decisions/0016-csv-exchange-over-the-data-layer.md).
- A file system layer in the spirit of System.IO: `ROneCOne.File` reads and writes text
  (UTF-8 by default with no byte-order mark, byte-order marks honored on read, `utf-16` and
  `windows-1252` on request), lines, and bytes, and copies, moves, and deletes with
  System.IO's exact missing/existing semantics; `ROneCOne.Directory` creates parents in one
  call, deletes empty-or-recursive, and enumerates files and folders sorted with `*`/`?`
  patterns and optional recursion; `ROneCOne.Path` joins and dissects path text. Failures
  raise the typed `ROneCOne.IOError` from source `ROneCOne.IOException`. See
  [ADR 0015](docs/decisions/0015-file-system-layer-over-adodb-stream.md).
- Native provider async: `OpenAsync`, `ExecuteReaderAsync`, `ExecuteScalarAsync`, and
  `FillAsync` now start their operation inside ADO with the async option at call time and the
  returned Task polls provider state cooperatively, so a slow query no longer freezes Excel
  while awaited. Failures surface the provider's own error detail, cancellation cancels the
  in-flight ADO operation, and a live SQL Server timing proof gates the overlap. `AsyncMode`
  reports `"Native"`; `ExecuteNonQueryAsync`, `UpdateAsync`, and `ReadAsync` keep documented
  single-step execution because ADO exposes no reliable async completion for them. The live
  suite now also exercises SQL Server on localhost through `MSOLEDBSQL` with integrated
  security: parameterized inserts, transactions, async failure shapes, and cancellation. See
  [ADR 0014](docs/decisions/0014-native-provider-async-over-adodb.md).

### Removed

- `SupportsNativeAsync` is gone: with the execute and open verbs genuinely native it stopped
  describing anything a single Boolean could state truthfully. `AsyncMode` and the provider
  documentation carry the exact per-verb behavior instead.

## 1.5.0 - 2026-07-23

### Added

- JSON serialization in the spirit of System.Text.Json: `ROneCOne.Json.Serialize` (compact or
  indented) and `Deserialize` into runtime-native values, where a JSON object becomes an
  ordered String-to-Variant dictionary and an array becomes a Variant list;
  `DeserializeTable` lands an array of objects in a typed DataTable with dotted columns for
  nested objects and an optional `$.data.items`-style path; `DeserializeInto`,
  `DeserializeObjects`, `ToObjects`, and `DataTableFromObjects` bind JSON and DataTables to
  your own classes through a factory delegate and explicit property names; and a `ToJson`
  convenience serializes collections, tables, rows, and views. Parsing is RFC 8259 strict
  with position-carrying typed errors (`ROneCOne.JsonError`), and the reader and writer
  mechanics are adapted, with permission, from
  [ModernJsonInVBA](https://github.com/WilliamSmithEdward/ModernJsonInVBA). See
  [ADR 0013](docs/decisions/0013-json-in-the-spirit-of-system-text-json.md).
- An eighth demo workbook, `ROneCOne_Json_Demo.xlsm`, walks the whole JSON surface offline:
  twelve validated examples cover parsing and navigation, compact and indented round-trips,
  DataTable landing with dotted columns and envelope paths, class binding in both directions,
  and RFC strictness, with a gated benchmark round-tripping a 1,000-row table.

### Fixed

- Primary-key and unique-column enforcement no longer rebuilds or scans per operation. A new
  row slots into the constraint indexes incrementally, a field write validates only its own
  column with token-index probes instead of row scans, and only key-affecting edits or deletes
  mark the indexes for one lazy rebuild. A live scenario of 2,000 loads, 2,000 edits, and
  2,000 finds on a table with a primary key and a unique column previously exceeded the
  harness's 30-second deadline; it now runs inside a new 1.5-second release gate. See
  [ADR 0012](docs/decisions/0012-incremental-constraint-indexes.md).
- List and generic-collection materialization builds one snapshot Collection instead of two,
  and an `HttpResponse` releases its byte-array copy once `GetByteArrayAsync` hands the result
  to its task.

## 1.4.0 - 2026-07-23

### Fixed

- Keyed mutation no longer rebuilds the whole hash index per operation. Replacing the value of
  an existing key now updates its single slot, removal deletes the slot in place with a
  backward-shift cluster repair and slides only the affected canonical indexes (tail removals
  skip the slide), and bulk removal loops defer one rebuild to the next probe. A live scenario
  of 10,000 in-place updates plus 2,000 keyed removals on a 10,000-entry dictionary previously
  exceeded the harness's 30-second deadline; it now completes in 0.6 to 1.3 seconds against a
  new 1.5-second release gate. The tally pattern `d(k) = d(k) + 1` is O(1) per call. See
  [ADR 0009](docs/decisions/0009-in-place-hash-index-maintenance.md).
- Indexed list writes no longer rebuild the For Each mirror per assignment: lists now share the
  version-checked lazy enumeration the other collections already used, and every eager mirror
  write is gone. Ten thousand `list(i) = value` assignments dropped from exceeding the
  30-second harness deadline to 0.1 seconds. See
  [ADR 0010](docs/decisions/0010-version-cached-positional-access.md).
- `Rows`, `Columns`, `Tables`, `Relations`, and `PrimaryKey` cache their snapshot against the
  owner's structural version instead of rebuilding it on every access, so `For i:
  table.Rows.Item(i)` is linear. Field edits advance a separate data version, which keeps
  read-write row loops linear while views still refresh. Repeated access now returns the same
  snapshot object until the owner structurally changes.
- Positional reads on materialized generic collections (sets, queues, dictionary entries) index
  their backing arrays directly instead of materializing per access. Deferred queries and live
  views intentionally keep per-read materialization: a query always sees the latest data.
- Re-filtering or re-sorting a DataView that had already been enumerated served stale results,
  because `WithFilter` and `WithSort` never advanced the view's version. Both now do, and the
  view's staleness key includes its own version.
- `AddColumn` on a table that already had rows left every existing row one cell short of the
  schema, so the next constraint sweep or field access failed with "Subscript out of range".
  A late column now backfills every existing row with the column's default cell.
- The keyed update path reuses the probe's found slot instead of probing again, and
  custom-comparer membership tests scan the element array directly instead of allocating a
  snapshot Collection per call.

### Changed

- Demo modules and documentation examples drop the `CLng(...)`, `CDbl(...)`, and `&`-suffix
  ceremony on numeric literals that lossless widening made unnecessary; plain literals now
  appear at every typed boundary, matching how the library is meant to be used.

### Added

- An awaitable HTTP client shaped like System.Net.Http: `ROneCOne.HttpClient()` with
  `BaseAddress`, `Timeout`, and `DefaultRequestHeader`; `GetAsync`, `GetStringAsync`,
  `GetByteArrayAsync`, `PostAsync`, `PutAsync`, `PatchAsync`, `DeleteAsync`, and `SendAsync`
  (which transmits any method, standard or custom), each returning a cooperative Task; and an
  `HttpResponse` with `StatusCode`, `ReasonPhrase`,
  `IsSuccessStatusCode`, `EnsureSuccessStatusCode`, `Content`, `Header`, and `AllHeaders`.
  Requests ride `WinHttp.WinHttpRequest.5.1` in process (no references or installs), overlap
  in flight under `WhenAll`, honor cancellation tokens by aborting the transport, and fault
  with a typed `HttpRequestException` (`ROneCOne.HttpRequestError`) on non-success text and
  byte reads. See
  [ADR 0011](docs/decisions/0011-awaitable-http-client-over-winhttp.md).
- A seventh demo workbook, `ROneCOne_Http_Demo.xlsm`, downloading live data from
  https://pokeapi.co with a sequential-versus-overlapped benchmark, plus an
  [HTTP user guide](docs/user-guide/http-and-web.md) and an
  [HTTP reference](docs/http.md) documenting the failure model and the `Application.Run`
  error boundary.
- A keyed-mutation scenario in the live benchmark harness with its own release gate, and a live
  consistency contract covering in-place updates, removals, re-adds, enumeration after mutation,
  the concurrent dictionary surface, and hash-set element replacement.
- Live benchmark scenarios with release gates for indexed list writes, positional row loops
  through `Rows`, and positional set reads, plus consistency contracts for lazy list
  enumeration, snapshot caching, view refresh after refilter and field edits, and late-column
  backfill.

## 1.3.0 - 2026-07-20

### Changed

- Element and key storage now uses doubling dynamic arrays instead of a VBA `Collection`, so a
  positional read on a materialized list or generic collection is O(1), an append is amortized
  O(1), and a value replacement is O(1). A `for i: list(i)` loop over ten thousand elements drops
  from quadratic (it previously exceeded the test deadline) to linear. The public API is
  unchanged. See [ADR 0005](docs/decisions/0005-array-backed-collection-storage.md).
- Rebuilt the Tasks demo and guide around cooperative coordination: scheduling, `WhenAll`,
  continuations, cancellation, progress, timeouts, and yielding.
- The task benchmark now measures 1,000 cooperative `Task.Run(...).Await` lifecycles.
- Rewrote the README and the documentation surface for a general audience: a code-led README,
  an indexed `docs/` directory, and standardized user guides and references, with no
  parallelism claims anywhere.
- Rewrote the inline documentation in every demo module in a tutorial voice, with a
  plain-language header explaining each concept and how to run its macro, so a non-technical
  reader can follow every example.

### Added

- `ToDisplayString`, which renders a deferred expression, lambda, or delegate as readable
  pseudocode (for example `(x.Age >= 40)`) so a query can be inspected while debugging.
- A worksheet Range bridge built on single bulk `Range.Value` calls:
  `ROneCOne.DataTableFromRange`, `DataTable.LoadFromRange`, `ROneCOne.ListFromRange`, and a
  role-dispatched `ToRange` that writes a DataTable grid, a DataView's visible grid, or a scalar
  sequence's column vector. Range parameters are late-bound so the runtime keeps no compile-time
  Excel reference. See [ADR 0004](docs/decisions/0004-worksheet-range-bridge.md).
- Lossless numeric widening at every type-admission point (list and collection elements,
  dictionary and keyed-collection keys, lambda parameters, delegate arguments and results,
  DataColumn values, primary-key `Find`, progress values, and completion results). Plain
  numeric literals no longer need `CLng`/`CDbl`-style wrappers; a value is admitted when it
  promotes to the declared type without loss and is stored as that type. Narrowing,
  cross-family, and Boolean/Date/String conversions still raise `TypeMismatchError` atomically.
  See [ADR 0003](docs/decisions/0003-lossless-numeric-widening.md).
- Architecture decision records under `docs/decisions/`: the native-slice removal, the
  `Task.Run` rename, lossless numeric widening, the worksheet Range bridge, array-backed
  storage, the measured deferral of query-plan and DISPID optimizations, the declined
  background completion pump, and the reaffirmed single-file runtime (ADR 0001 through 0008).

### Removed

- The native `Task.Run` execution slice: the expression verifier, bytecode compiler, embedded
  x64 worker kernel, thread-pool and virtual-memory declares, and the `TaskRun`,
  `ExecutionMode`, `WorkerThreadId`, and `CurrentThreadId` members. Direct measurement showed
  the parallel path delivered no practical gain, and removing the embedded machine code
  simplifies audit and trust. Signature-bound native delegates through `DispCallFunc` are
  unchanged. See [ADR 0001](docs/decisions/0001-remove-native-task-run.md).
- `RunOnExcel`: with one execution model there is nothing left for the name to distinguish.
  `Task.Run(work, optional token)` is the single scheduling call and executes cooperatively on
  Excel's thread. See
  [ADR 0002](docs/decisions/0002-task-run-names-the-cooperative-scheduler.md).
- The kernel builder `tools/build_native_task_kernel.py` and its source contracts.
- The unused `mExpressionText` field, the write-only collection builder flag, the bracketed
  `[Empty]` alias (`EmptyOf` remains), the duplicate public `TaskFromResult`
  (`Task.FromResult` remains), the expired `DemoCustomerQuery` packager shim, and the reserved
  `.ronecone.env.example` schema the runtime never read.

## 1.2.0 - 2026-07-19

### Added

- Hot `Task.Run` execution for verified numeric and Boolean expression lambdas on the Windows
  thread pool, with inferred result types, cancellation checks, native fault translation, and
  ordered `WhenAll` composition.
- `Task.RunOnExcel` for explicit cooperative execution of VBA procedures, workbook operations,
  COM, continuations, and other Excel-owned work.
- `ExecutionMode`, `WorkerThreadId`, and `CurrentThreadId` diagnostics that make task placement
  observable without relying on timing guesses.
- A reproducible x64 worker-kernel builder and a source contract that requires the embedded kernel
  to match its verified generated bytes exactly.

### Changed

- Split task scheduling into an honest native-compute path and an Excel-thread path; unsupported
  `Task.Run` work now fails before submission instead of silently running sequentially.
- Reworked the Tasks demo and guide around parallel forecast calculations and safe workbook work.

### Safety boundary

- Worker threads cannot invoke VBA, Excel, COM, objects, strings, or arbitrary addresses. The
  embedded interpreter executes only verified private bytecode over task-owned scalar data.
- Worker code is changed from read/write to read/execute before submission. Bytecode and context
  memory remain non-executable, and task teardown waits for callbacks before releasing memory.

## 1.1.0 - 2026-07-19

### Added

- Open-address hash indexes with direct value slots for default-equality dictionaries, hash sets,
  keyed and ordered collections, concurrent and immutable variants, and lookups.
- `Capacity`, `EnsureCapacity`, and `TrimExcess` contracts with live reuse-after-`Clear` coverage.
- `Task.YieldOnce`, `WaitAsync`, direct `WhenAll(task...)` and `WhenAny(task...)` inputs,
  nonblocking `Task.Exception`, and AggregateException inspection and handling.
- Disposable cancellation registrations, deterministic `Using(resource).Run(work)`, typed timeout
  and cancellation errors, and monotonic task deadlines.
- `DataColumn.AsPrimaryKey`, `DataTable.Row(...).Add`, explicit `ROneCOne.DBNull`, indexed
  single/composite primary-key `Find`, and indexed multi-column relation navigation.
- Provider capability reporting plus transactional and continue-on-error DataAdapter options.
- Live 10,000-item hash build/read and 100,000-lookup performance gates.

### Changed

- Replaced linear default dictionary/set lookup with average O(1) hashing and linear sorted
  map/set lookup with binary search. Custom equality comparers retain a truthful linear path.
- Hardened the cooperative scheduler against deadline wraparound, same-task reentrant waits,
  dependency cycles, callback failures, and leaked registrations.
- Made `WhenAll` retain every child fault while preserving ordinary VBA error identity on Await.
- Made primary-key and relation indexes rebuild atomically when table versions change.
- Made provider task APIs report cooperative execution instead of implying native asynchronous I/O.
- Expanded the live host contract to 414 assertions and refreshed the Tasks, Data, and Collections
  living demos around the shortest proven syntax.

### Host boundaries

- VBA does not resolve a class member named `Yield`, so the legal C#-aligned form is `YieldOnce`.
- Tasks and provider calls remain cooperative on Excel's owning thread; no second Excel application
  or unsafe worker-thread COM access is used.

## 1.0.0 - 2026-07-19

### Added

- The full generic collection family: dictionary, hash set, queue, stack, linked list, sorted and
  ordered maps/sets, priority queue, observable/read-only/keyed collections, concurrent-style and
  blocking collections, and immutable collections with builders.
- A substantially broader LINQ surface covering conversion, grouping, joining, flattening,
  partitioning, set-by-key, indexing, counting, and aggregation operators.
- Cooperative Task workflows with await-style coordination, delay, continuations, WhenAll,
  WhenAny, cancellation, callbacks, progress, and typed completion sources.
- DataTable, DataColumn, DataRow, DataView, DataSet, DataRelation, constraints, change tracking,
  copying, merging, selection, and relation navigation.
- Late-bound OLE DB and ODBC connections, commands, typed parameters, readers, transactions,
  adapters, source-column binding, update commands, deterministic disposal, and Task-returning
  provider operations.
- Dedicated Tasks and Data/Providers living demo workbooks and popup-aware VBE source-selection
  diagnostics.

### Changed

- Optimized composite ordering with cached scalar keys; the 10,000-element live Excel scenario
  remains below its 1.0-second release gate without weakening the threshold.
- Allowed universal delegates to target explicit methods on ROneCOne runtime values.
- Made relation admission validate existing parent and child rows before mutating a DataSet.
- Unified task waiting under one cooperative scheduler so `WhenAny` observes every candidate and
  completion sources, `WhenAll`, delays, and continuations share terminal-state semantics.
- Expanded the live suites to 58 delegate, 187 collection, 71 advanced-collection, and 68
  task/data/provider assertions.

### Host boundaries

- VBA reserves `Open` and `Close` as procedure names, so direct provider calls use `Connect` and
  `Disconnect`; `OpenAsync` keeps the .NET-aligned async name.
- Tasks coordinate cooperatively on Excel's owning thread. They never launch another Excel
  application and do not claim unsafe CPU-parallel VBA execution.
- Provider mutation support follows the selected driver; the Excel ISAM supports reads, updates,
  and inserts but rejects linked-sheet row deletion.

## 0.9.0 - 2026-07-19

### Added

- The complete .NET-aligned ordering family: `Order`, `OrderDescending`, `OrderBy`,
  `OrderByDescending`, `ThenBy`, and `ThenByDescending`.
- Independent comparers for every ordering level, including direct ordering of user-defined class
  instances through an explicit comparer.
- Live contracts for stable composite keys, reset and continuation rules, deferred execution,
  selector-once evaluation, Null, Boolean, ordinal strings, and rejected mixed Variant coercion.
- A 10,000-element composite-ordering benchmark with a 1.0-second isolated-suite gate and a
  2.5-second living-workbook gate sized from fresh Excel-process measurements.

### Changed

- Replaced quadratic insertion sorting with a stable O(n log n) bottom-up merge sort that caches
  every selected key exactly once per element and enumeration.
- Removed the earlier identity-order aliases in favor of the .NET names `Order` and
  `OrderDescending`; the current API, demos, tests, and documentation move together.
- Made `ThenBy` legal only on an immediately active ordered query and made a new primary ordering
  operation replace the prior ordering chain.
- Expanded the live collection suite from 122 to 146 assertions and upgraded the existing
  Collections workbook with composite ordering and a measured ordering release gate.

## 0.8.0 - 2026-07-19

### Added

- Composable collection-membership expressions through `IsIn`, `NotIn`, `ContainsMember`, and
  expression arguments to `List<T>.Contains`; `OneOf` now accepts arrays and collections.
- Null-safe `?.` member paths with deterministic Null propagation through the remaining chain.
- `Both`, `Either`, `Negated`, and `WhereNot` predicate composition sugar plus case-insensitive
  equality, prefix, suffix, containment, and pattern helpers.
- Predicate-aware `Count`, `FirstOrDefault`, `LastOrDefault`, `SingleItem`, `SingleOrDefault`, and
  `None` terminals.
- Nested collection quantifiers through `AnyMatch`, `AllMatch`, `NoneMatch`, `WhereAny`,
  `WhereAll`, and `WhereNone`.
- `EqualityComparer` and `Comparer` factories with comparer-aware containment, distinctness,
  sorting, extreme-value operators, and `SequenceEqual`.
- Reusable `Always`, `Never`, `Match`, and `NotMatch` predicate factories.
- A repeated object-member benchmark alongside the scalar query benchmark.

### Changed

- Cached normalized member names directly on expression nodes, removing one captured-value node
  and one recursive evaluation per member access.
- Expanded the live collection suite from 91 to 122 assertions and the Collections workbook from
  nineteen to twenty-five formula-verified examples.
- Replaced explicit nullable-object guard chains in the primary demo with `?.` paths.
- Recorded the Excel-host grammar boundary that forbids `In` and `Single` as procedure names;
  the legal public forms are `IsIn` and `SingleItem`.

## 0.7.0 - 2026-07-19

### Added

- Deferred contextual filters through `Where("Member").AtLeast(...)` and the complete comparison,
  set, string, Boolean, null, and pattern condition vocabulary.
- Reusable `Condition(memberPath)` expressions backed by one stable typed element parameter per
  sequence, including composed multi-member predicates and dotted object paths.
- Native VBA bang-member expressions such as `customers!Age` and `With customers: !Age` through
  the existing default-member dispatch surface.
- Member-name selectors for projection, ordering, aggregation, and text joining.
- `DistinctBy`, `MinBy`, and `MaxBy` key-based LINQ operators.
- `Predicate` and `WhereMethod` inference of `Func<T, Boolean>` from a sequence's generic type.
- Eleven live user-class LINQ scenarios in the Collections workbook.

### Changed

- Made `Contains` serve both typed sequence containment and C#-style string expression containment.
- Updated the primary demo surface to remove explicit element variables wherever member context is
  sufficient, while retaining and documenting the canonical expression/delegate forms.
- Expanded the live collection suite from 66 to 91 assertions.
- Confirmed through the popup-aware VBE harness that both direct and `With`-scoped bang syntax
  compile and execute in Microsoft 365 x64 Excel.

## 0.6.0 - 2026-07-19

### Added

- `Execute` as a statement-form Action invocation surface, including inline native reference
  wrappers without dummy result variables.
- Typed mutable events through `EventOf`, fluent `Subscribe`, last-match `Unsubscribe`, snapshot
  `Emit`, and `HandlerCount`.
- Structured immutable `Try`, catch-all and error-number `Catch`, `Finally`, captured exception
  metadata, deterministic rethrow, and cleanup precedence.
- Explicit populated `ListOf(T, items...)`, inferred `ListFrom(first, rest...)`, and live-tested
  empty `ListLike(example)` construction.
- Atomic `AddRange` inputs from typed ROneCOne sequences, VBA arrays, and `Collection` values.
- `ForEach(Action)`, predicate-optional `Exists`, and `JoinText` terminals.
- Separate Events and Exceptions living workbooks with source, packaging, execution, rendering,
  and same-process benchmark contracts.

### Changed

- Admitted every universal unary `Func` role into LINQ normalization, including workbook
  procedures, native functions, compositions, and multicast delegates.
- Reworked delegate and collection demos to use one-expression signatures, inline `ByRef`, Action
  `Execute`, collection initializers, inferred user-class lists, `ForEach`, and `JoinText`.
- Expanded live totals to 58 delegate/event/exception assertions and 66 collection assertions.
- Recorded and eliminated VBA's host-only invalid-`ParamArray` forwarding failure; the popup-aware
  harness closed the compiler modal and terminated only its task-owned Excel process.
- Made every bounded Excel launcher tolerate duplicate `Path`/`PATH` host environments and removed
  machine-specific Node paths from workbook rendering.

## 0.5.0 - 2026-07-18

### Added

- Universal `Func` and `Action` factories for expression trees, object methods, callable objects,
  and workbook procedures.
- Immutable `Takes` / `Returns` signatures, C#-style `DynamicInvoke`, target/method/signature
  metadata, and exact user-class type descriptors.
- Immutable multicast `Combine`, `Remove`, `GetInvocationList`, deterministic invocation order,
  and last-result semantics.
- Fail-closed Windows x64 `Native` / `NativeAction` dispatch through `DispCallFunc`.
- True native `ByRef` with `RefOf` and typed wrappers for VBA numeric, Boolean, Date, and Currency
  variables.
- Live-host contracts for callable objects, workbook procedures, object identity returns,
  multicast removal, native function pointers, `ByRef` mutation, and rejected incomplete calls.
- A complete universal-delegate guide and an eleven-scenario living delegate workbook.

### Changed

- Replaced the narrow `FromMethod` adapter with the single `Func` / `Action` construction model.
- Routed LINQ method delegates through universal typed factories.
- Made non-native `ByRef` fail before dispatch because `CallByName` and `Application.Run` cannot
  preserve the original variable identity.
- Made invalid `Action`-to-`Func` and `Action.Returns(...)` conversions fail during construction.
- Expanded the Excel harness to report exact delegate assertion failures as well as modal VBE
  compiler selections.

Demos, tests, and documentation use the preferred universal model without compatibility shims.

## 0.4.0 - 2026-07-18

### Added

- `Element` as the clear, LINQ-aligned name for the typed sequence expression parameter.

### Changed

- Updated the collections demo, workbook examples, quick start, and deep documentation to lead
  with `Element` rather than the less familiar `It` terminology.
- Made `Element` the canonical implementation used by identity sorting.
- Removed the less familiar `It` sequence-parameter name in favor of `Element`.

The API intentionally prefers the clearer public surface over compatibility aliases.

## 0.3.0 - 2026-07-18

### Added

- `Var` and `VarLike` aliases plus `.AsFunc` for inferred, typed delegate parameters.
- Automatic left-to-right parameter inference when `Lambda` is called without an explicit
  parameter list.
- Sequence-typed `It` parameters and scalar object-member expressions through `value("Member")`
  or the explicit `Member` API.
- `Map`, `Exists`, identity ordering, `AtLeast`, and `AtMost` syntax sugar.
- Deterministic `MemberAccessError` handling for invalid object-member expressions.
- Live Excel coverage for concise primitive queries, user-class queries, member failures, inferred
  unary/binary delegates, and `VarLike` object delegates.

### Changed

- Reworked both living demos to lead with the shortest clear syntax while keeping canonical forms
  in the deeper documentation.
- Removed the demo-only `DemoCustomerQuery` adapter; user-defined-class filtering, projection,
  ordering, quantification, and aggregation now use direct typed member expressions.
- Expanded the live totals to 19 delegate and 52 collection assertions without changing either
  performance gate.
- Recorded maximal safe syntax sugar as a foundational repository product direction.

Canonical `Parameter`, `ParameterLike`, explicit `Lambda`, `SelectItems`, `AnyItem`, and selector
APIs remain available for explicit construction and debugging.

## 0.2.1 - 2026-07-18

### Added

- A dedicated **User Class LINQ** worksheet with six live, formula-verified `DemoCustomer`
  scenarios: strict typing, deferred filtering, projection, object ordering, quantifiers, and an
  aggregate projection.
- `Age` and `City` fields on the demo customer model plus a demo-only `DemoCustomerQuery` adapter
  containing named predicates and selectors.
- Source-contract tests for demo organization, readability, workbook content, and VBA packaging.

### Changed

- Reorganized both demo VBA modules into small, purpose-focused procedures with comments that
  explain the important design decisions.
- Expanded the collections runner from one example sheet and eight checks to two sheets and
  fourteen checks.
- Extended final workbook rendering and module round-trip verification for the new demo surface.
- Expanded `.gitignore` for local secrets, development environments, generated workbooks, Office
  recovery files, task logs, and editor metadata.

The shipped one-file `ROneCOne.cls` runtime and its public API are unchanged from 0.2.0.

## 0.2.0 - 2026-07-18

### Added

- Strict runtime-generic `List<T>` values for primitive and exact user-defined class types.
- Zero-based default/explicit indexing, typed mutation, atomic `AddRange`, and nested `For Each`.
- Deferred LINQ pipelines with filtering, projection, ordering, slicing, distinct, append,
  prepend, and reverse operators.
- Immediate quantifier, element, numeric aggregate, count, list, and array terminals.
- Typed `Range` and `Repeat` sequence factories.
- Live Excel tests for source mutation, user-class identity/projection, enumeration refresh, and
  deterministic type failures.
- A 10,000-element collection benchmark with a 0.75-second release gate.
- Separate living workbooks for delegates and collections/LINQ, each independently built,
  executed, rendered, and popup-monitored.
- The complete MIT License embedded directly in the shipped `ROneCOne.cls`.

### Changed

- Generic collections moved ahead of structured exceptions in the feature order.
- Demo/test workers now report exact collection assertion failures and support capability-specific
  workbooks.

## 0.1.0 - 2026-07-18

### Added

- One-file `ROneCOne.cls` tagged-object runtime.
- Typed expression parameters and immutable expression-tree lambdas.
- Default-member delegate calls and explicit `Run` calls.
- Arithmetic, comparison, concatenation, short-circuit Boolean, and negation expressions.
- Method delegates, scalar/object returns, and delegate composition.
- Source-contract tests, whole-project static analysis, live Excel tests, and benchmarks.
- Popup-adaptive Excel/VBE introspection with selected-code diagnostics and hard timeouts.
- Formula-backed, macro-enabled living demo workbook.
