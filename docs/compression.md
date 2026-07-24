# Zip archives

Exact semantics for the zip surface. For the workflow-first introduction, read
[Files and folders](user-guide/files-and-folders.md).

## Surface

`ROneCOne.ZipFile` reads and writes archives in pure VBA, with no reference, add-in, or Shell
automation.

| Member | Behavior |
|---|---|
| `ZipFile.CreateFromDirectory(sourceDir, archivePath)` | Zips a folder tree into a new stored archive |
| `ZipFile.ExtractToDirectory(archivePath, targetDir)` | Extracts every entry, refusing existing files |
| `ZipFile.OpenRead(archivePath)` | Opens an archive for reading, returning an archive value |
| `archive.Entries` | The entries as a queryable typed list |
| `archive.GetEntry(fullName)` | One entry by exact name, or `Nothing` |
| `entry.FullName` / `entry.Name` | The path inside the archive, and its last segment |
| `entry.Length` / `entry.CompressedLength` | The uncompressed and stored byte counts |
| `entry.ReadAllText([encoding])` / `entry.ReadAllBytes` | The entry's content, without extracting to disk |
| `entry.ExtractToFile(path, [overwrite])` | Writes one entry to a file |

Failures raise error number `ROneCOne.ZipError` from source `ROneCOne.ZipException`.

## Reading

`OpenRead` parses the central directory once. Each entry read inflates the stored or deflated
data and verifies it against the entry's recorded CRC-32; a mismatch raises `ZipError`.
`ReadAllText` sniffs a byte-order mark exactly like `File.ReadAllText`, so a UTF-8 or UTF-16
entry decodes correctly, and the optional encoding names the fallback when no mark is present.
`ExtractToFile` refuses to overwrite unless `overwrite` is `True`.

The engine is RFC 1951 inflate (stored, fixed-Huffman, and dynamic-Huffman blocks). It reads
archives produced by PowerShell `Compress-Archive`, .NET, the zip command line, and the common
desktop tools.

## Writing

`CreateFromDirectory` walks the folder tree and writes every file as a stored (uncompressed)
entry, with forward-slash names, a UTF-8 name flag when a name needs it, and DOS timestamps.
Empty folders become directory entries. Stored entries are larger than compressed ones but are
read by every unzipper, including `Expand-Archive`. The destination must not already exist.

## Security and limits

`ExtractToDirectory` routes every entry name through a traversal guard: an absolute path, a
drive letter, a backslash, or a `..` segment is refused with `ZipError` before anything is
written, so a zip-slip archive cannot escape the target folder. Entries are still listable
through `OpenRead`; only extraction to disk is guarded.

Refused with `ZipError`: Zip64 archives (4 GB or larger, or 65,535 or more entries), encrypted
entries, multi-disk archives, and compression methods other than store and deflate. Writing is
store-only; the surface does not deflate.

[Back to the documentation index](README.md)
