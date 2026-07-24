# ADR 0015: A System.IO-style file layer over ADODB.Stream and Scripting

Status: accepted, 2026-07-23

## Context

VBA's native file statements read and write in the system ANSI codepage, so UTF-8 (the
encoding of essentially every modern file exchange) silently corrupts anything past ASCII.
The usual fixes are a FileSystemObject reference plus hand-rolled `ADODB.Stream` incantations
scattered through user code. The runtime already replicates `System.Net.Http` and
`System.Text.Json`; files are the remaining leg of the data-in, data-out story.

A live probe established the stream mechanics: `ADODB.Stream` writes UTF-8 with a 3-byte
mark that a `Position = 3` binary copy removes cleanly, reads BOM-less UTF-8 correctly when
told the charset, strips a leading mark when one exists, and writes UTF-16 with its `FF FE`
mark. Missing files raise 3002 from the stream.

## Decision

Three factory surfaces mirror `System.IO`: `ROneCOne.File` (`ReadAllText`, `WriteAllText`,
`AppendAllText`, `ReadAllLines`, `WriteAllLines`, `ReadAllBytes`, `WriteAllBytes`, `Exists`,
`Delete`, `Copy`, `Move`), `ROneCOne.Directory` (`Exists`, `CreateDirectory`, `Delete`,
`GetFiles`, `GetDirectories`), and `ROneCOne.Path` (`Combine`, the name and extension
helpers, `GetFullPath`, `GetTempPath`). Encodings are a closed set (`utf-8` default,
`utf-16`, `windows-1252`); a byte-order mark always wins on read; UTF-8 writes carry no mark;
appends never embed one. Failures raise one typed number, `ROneCOne.IOError`, from source
`ROneCOne.IOException`, and the missing/existing/overwrite semantics copy `System.IO`
deliberately, including silent `File.Delete` of a missing file and refusal to delete a
non-empty folder without `recursive`.

VBA has no method overloading, so the four member names the runtime already used elsewhere
(`Exists`, `Delete`, `Copy`, `Combine`) dispatch on the instance's role, the same way `State`
and `Disconnect` already serve multiple roles. `Combine` now returns `Variant` so the path
role can return text while delegate roles keep returning multicast objects; data-row `Delete`
and table `Copy` reject stray arguments instead of ignoring them. Enumeration results are
sorted with case-insensitive text comparison so no test or workflow ever depends on file
system order, and `GetFiles` patterns translate to VBA `Like` with `[` and `#` escaped.

## Consequences

Two prog-ids join the source-contract whitelist: `Scripting.FileSystemObject` and the
const-based `ADODB.Stream`. The live suite adds a file system contract under a scratch root
beside the test workbook: UTF-8 round trips with no mark, UTF-16 marks honored on read and
write, windows-1252 single-byte output, line splitting and terminal-newline rules, byte round
trips, copy/move/delete semantics including the typed failures, sorted and recursive
enumeration, and the path helper table. Everything runs offline and cleans up after itself.
