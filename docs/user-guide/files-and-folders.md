# Files and folders

Read and write files the way C# does, without `Open ... For Input` statements, ANSI-only
codepages, or a FileSystemObject reference. `ROneCOne.File`, `ROneCOne.Directory`, and
`ROneCOne.Path` mirror `System.IO`, and every text operation speaks UTF-8 by default.

## Read and write text

```vba
Dim report As String

ROneCOne.File.WriteAllText "C:\data\report.txt", "Total: 42"
report = ROneCOne.File.ReadAllText("C:\data\report.txt")
```

UTF-8 is the default and the output carries no byte-order mark, so other tools read it
cleanly. Reading auto-detects a byte-order mark, so UTF-16 files decode correctly even when
you did not ask. Pass `"utf-16"` or `"windows-1252"` as the optional last argument when a
legacy consumer demands it.

Lines work as collections:

```vba
Dim lines As ROneCOne

ROneCOne.File.WriteAllLines "C:\data\names.txt", Array("Ada", "Grace")
Set lines = ROneCOne.File.ReadAllLines("C:\data\names.txt")
Debug.Print lines.Item(0)                     ' Ada
Debug.Print lines.Count                       ' 2
```

`ReadAllLines` returns an ordinary `ListOf(vbString)`, so the whole LINQ surface applies.
`ReadAllBytes` and `WriteAllBytes` round-trip binary content as `Byte` arrays.

## Manage files and folders

```vba
ROneCOne.Directory.CreateDirectory "C:\data\archive\2026"
ROneCOne.File.Copy "C:\data\report.txt", "C:\data\archive\2026\report.txt"
ROneCOne.File.Move "C:\data\old.txt", "C:\data\archive\old.txt"
ROneCOne.File.Delete "C:\data\temp.txt"
ROneCOne.Directory.Delete "C:\data\scratch", True
```

`CreateDirectory` creates missing parents in one call. `Copy` refuses to overwrite unless you
pass `True`; `Move` refuses always; `File.Delete` of a missing file is silent, exactly like
`System.IO.File.Delete`. Failures raise error number `ROneCOne.IOError` with the path in the
message, so `Catch`-style handlers can key on one number.

Enumerate with patterns, sorted and optionally recursive:

```vba
Dim workbooks As ROneCOne

Set workbooks = ROneCOne.Directory.GetFiles("C:\data", "*.xlsx", True)
```

## Build paths without string surgery

```vba
Dim target As String

target = ROneCOne.Path.Combine(ROneCOne.Path.GetTempPath(), "ronecone", "out.json")
Debug.Print ROneCOne.Path.GetFileNameWithoutExtension(target)   ' out
Debug.Print ROneCOne.Path.ChangeExtension(target, "csv")        ' ...\out.csv
```

`Combine` inserts separators only where needed and restarts at a rooted part, so joining user
input never produces `C:\data\\file` or silently escapes a folder.

## Where next

- [Files technical reference](../files.md) defines encodings, failure semantics, and
  enumeration order exactly.
- [Data and providers](data-and-providers.md) shows the tables your files often feed.
- [Guide index](README.md) returns to the full learning path.
