# ADR 0026: Escaping helpers and process standard input

Status: accepted, 2026-07-24

## Context

Two everyday gaps sat next to the HTTP and process surfaces. Building a query string by hand is
a common source of bugs because VBA has no percent-encoder, and emitting text into HTML needs
entity encoding the runtime did not offer. Separately, `Process.RunAsync` could capture a
command's output but could not feed it input, so the large family of filter tools that read
standard input (`sort`, `findstr`, `certutil`) was out of reach. A probe confirmed the process
mechanics: writing to the `Exec` object's `StdIn` and then closing it lets a filter see
end-of-file and complete, and the close is what matters most, because a filter with an open,
empty input waits forever.

## Decision

`ROneCOne.Uri.EscapeDataString` percent-encodes over UTF-8 bytes with the RFC 3986 unreserved
set (`A-Z a-z 0-9 - . _ ~`) preserved, and `UnescapeDataString` decodes valid `%XX` escapes
through UTF-8 while leaving malformed escapes untouched, matching `Uri.UnescapeDataString`.
`ROneCOne.WebUtility.HtmlEncode` encodes the five markup-sensitive characters (`&`, `<`, `>`,
`"`, `'`) plus every non-ASCII character as a numeric reference, decoding through surrogate
pairs; `HtmlDecode` reverses the five named entities and numeric decimal and hex references and
passes unknown entities through unchanged, matching `WebUtility`'s lenient posture.
`Process.RunAsync` gains an optional `standardInput` string, written to the process at start and
then always closed, so a filter sees end-of-file whether or not input was supplied. A private
`IoDecodeText` is added as the mirror of the existing `IoEncodeText`, turning raw bytes into
text through one `ADODB.Stream`; the escaping helpers and, later, the zip surface share it.

## Consequences

No new prog-id, `Declare`, or error number: the escaping surfaces are pure VBA over the existing
stream encoders, and the process change is one `StdIn` write and close on the already-open `Exec`
object. `Uri` and `WebUtility` join the role-dispatched factory members. The live suite extends
the process contract with a `sort` and a `findstr` fed through standard input and a
closed-empty-input case that must not hang, and adds an escaping contract: percent round trips
including UTF-8 and preserved unreserved characters, malformed escapes left alone, HTML encoding
of markup and non-ASCII and a surrogate pair, lenient decoding of named and numeric entities,
and the role guards.
