# ADR 0023: Zip archives over a pure-VBA engine

Status: accepted, 2026-07-24

## Context

Zip is the last file operation the runtime could not perform. It already downloads, hashes, and
stores files, but the archives those files usually travel in stayed opaque, and Office's own
formats (`.xlsx`, `.docx`) are zips. VBA offers nothing built in. The one in-box interop route,
`Shell.Application.NameSpace(...).CopyHere`, was rejected on principle: it has no completion
signal (the copy runs on a background thread the caller cannot join without a sleep-and-hope
loop) and no error channel (a failed extraction is silent), which is exactly the kind of
nondeterminism the Task scheduler and the file layer were built to avoid. Depending on
`.NET`'s `System.IO.Compression` through COM repeats the fragility ADR 0019 already rejected
for crypto: it is not reliably registered.

The remaining route is arithmetic, and the zip and DEFLATE formats are small enough to
implement directly: a central directory of fixed-layout records, CRC-32 from a 256-entry
table, and RFC 1951 inflate (stored, fixed-Huffman, and dynamic-Huffman blocks with a 32 KB
back-reference window). A scratchpad engine was built and iterated against python-generated
fixtures covering every block flavor, long back-references, multi-entry archives with
subdirectories and empty files, a UTF-8 flagged name, and hostile inputs, until every content
byte and CRC matched in Excel.

## Decision

`ROneCOne.ZipFile` reads and writes archives in pure VBA. `OpenRead` parses the central
directory and returns an archive whose `Entries` is a queryable list and whose `GetEntry` finds
one by exact name (`Nothing` on a miss); each entry exposes `FullName`, `Name`, `Length`,
`CompressedLength`, `ReadAllText`, `ReadAllBytes`, and `ExtractToFile`. Reading inflates stored
and deflated entries and verifies every entry against its stored CRC-32. `CreateFromDirectory`
walks a folder tree and writes a store-only archive (uncompressed entries, which every unzipper
reads; the larger size is documented, not hidden) with UTF-8 names and DOS timestamps.
`ExtractToDirectory` refuses to overwrite existing files and routes every entry name through a
traversal guard that rejects absolute paths, drive letters, backslashes, and `..` segments, so
a zip-slip archive cannot escape the target folder. Zip64 archives (4 GB or larger) and
encrypted entries are refused with the typed `ROneCOne.ZipError` from source
`ROneCOne.ZipException`; multi-disk archives and unknown compression methods are refused the
same way.

## Consequences

No prog-id and no `Declare` are added; the engine is arithmetic over a byte array read through
the existing file layer, sharing `IoDecodeText` for entry names and `IoWriteAllBytesCore` for
output. A private `InflateState` type and the CRC and Huffman tables join the class. The tables
are built lazily on first use and, because every zip entry is its own instance, the extract
path ensures them itself rather than assuming a directory parse ran first; this was a live-only
defect (a stored round trip never inflates, so it hid until a deflated PowerShell entry was
read through an entry instance) and is the reason the engine is validated live, not just
offline. The live suite adds a zip contract: a create-open-extract round trip, per-entry
reads, overwrite and traversal refusals, the Zip64 and encryption rejections, and interop in
both directions against PowerShell `Compress-Archive` (whose deflate output the engine
inflates) and `Expand-Archive` (which reads the engine's output).
