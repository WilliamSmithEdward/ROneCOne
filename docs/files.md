# Files

Exact semantics for the System.IO-style surface. For the workflow-first introduction, read
[Files and folders](user-guide/files-and-folders.md).

## Surfaces

`ROneCOne.File`, `ROneCOne.Directory`, and `ROneCOne.Path` return role instances mirroring the
static classes of `System.IO`. Text and byte transput run through `ADODB.Stream`; folder
operations run through the Scripting runtime. Both are in-process prog-ids: no references,
installs, or add-ins.

## Text encoding

`ReadAllText`, `WriteAllText`, `AppendAllText`, `ReadAllLines`, and `WriteAllLines` accept an
optional encoding of `utf-8` (default), `utf-16`, or `windows-1252`; anything else raises
`InvalidArgumentError`. On read, a byte-order mark overrides the requested encoding: `FF FE`
decodes as UTF-16 little endian, `FE FF` as UTF-16 big endian, and `EF BB BF` as UTF-8. On
write, UTF-8 output carries no byte-order mark (matching .NET's default), UTF-16 keeps its
mark at the start of a file, and appended text never embeds one mid-file.

`ReadAllLines` accepts CRLF, LF, and CR line endings and, like .NET, ignores one final line
terminator. `WriteAllLines` accepts a ROneCOne sequence, VBA array, or `Collection` and
terminates every line with CRLF, including the last.

## Failure contract

File system failures raise error number `ROneCOne.IOError` from source
`ROneCOne.IOException`, carrying the path and the underlying provider description. The
semantics follow `System.IO`:

| Operation | Behavior |
|---|---|
| `File.ReadAllText` / `ReadAllLines` / `ReadAllBytes` on a missing file | Raises |
| `File.Delete` on a missing file | Silent |
| `File.Copy` onto an existing file | Raises unless `overwrite:=True` |
| `File.Move` onto an existing file | Raises |
| `Directory.CreateDirectory` | Creates missing parents; existing folder is silent |
| `Directory.Delete` on a missing folder | Raises |
| `Directory.Delete` on a non-empty folder | Raises unless `recursive:=True` |
| `Directory.GetFiles` / `GetDirectories` on a missing folder | Raises |

## Enumeration

`GetFiles(path, [searchPattern], [allDirectories])` and `GetDirectories` match the pattern
against entry names case-insensitively with `*` and `?` wildcards, return full paths, and sort
results with case-insensitive text comparison so output order never depends on the file
system. Passing `allDirectories:=True` walks every subfolder.

## Path helpers

`Path.Combine` joins any number of parts with `\`, skips empty parts, and restarts at a rooted
part, like `System.IO.Path.Combine`. `GetFileName`, `GetDirectoryName`, `GetExtension`
(including its dot), `GetFileNameWithoutExtension`, and `ChangeExtension` are pure text
operations accepting both separators; `ChangeExtension` with empty text removes the extension.
`GetFullPath` resolves against the process working directory and `GetTempPath` ends with a
separator. Path helpers never touch the disk; only `GetFullPath` and `GetTempPath` consult the
environment.

## Watching a folder for changes

`ROneCOne.FileWatcher(folder, [filter])` mirrors `FileSystemWatcher` on the task scheduler.
`WaitForChangeAsync([cancellationToken])` returns a Task that resolves on the next change to a
value with `ChangeType` (`Created`, `Changed`, or `Deleted`) and `Name`. The baseline snapshot
is taken when `WaitForChangeAsync` is called, so a change that lands before the await is still
seen; the filter uses the same `*` and `?` matching as `Directory.GetFiles`. A missing folder
raises `IOError`.

This is a poll, not an event subscription: while the Task is awaited the folder is rescanned at
most four times a second, and a file both created and deleted between two scans is not observed.
Cancellation and `WaitAsync` timeouts compose like every other task.

## Reading and writing zip archives

The zip surface pairs naturally with the file layer; see [zip archives](compression.md).
`ROneCOne.ZipFile.OpenRead(path)` lists and reads entries without extracting,
`ExtractToDirectory` unpacks with a directory-traversal guard, and `CreateFromDirectory` writes
a stored archive. Reading inflates deflated entries produced by any standard tool.

[Back to the documentation index](README.md)
