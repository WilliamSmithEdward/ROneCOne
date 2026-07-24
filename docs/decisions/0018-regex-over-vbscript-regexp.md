# ADR 0018: Regular expressions over VBScript.RegExp

Status: accepted, 2026-07-23

## Context

VBA has no regular expression support in the language. The usual answer is a reference to
"Microsoft VBScript Regular Expressions 5.5", which the one-file, no-reference invariant
forbids. The engine behind that reference, `VBScript.RegExp`, is an in-box prog-id creatable
late-bound with no registration, so it fits the same whitelist pattern as WinHTTP, ADO, the
ADODB stream, and WScript.Shell.

A live probe established the engine's semantics: `Execute` returns a matches collection with
`FirstIndex`, `Length`, `Value`, and `SubMatches`; `Global` toggles first-only versus all
matches; an unmatched optional group reports `Empty`; `a*` against `bab` yields four
zero-length-inclusive matches; `IgnoreCase` and `MultiLine` behave; and an invalid pattern
raises 5017/5020 only when the pattern is first used, not when assigned.

## Decision

`ROneCOne.Regex(pattern, [ignoreCase], [multiLine])` wraps the engine with the
`System.Text.RegularExpressions` verb set: `IsMatch`, `Match`, `Matches`, `Replace`, `Split`,
and a `Pattern` accessor. `Match` and `Matches` return a dedicated match role
(`ROLE_REGEX_MATCH`) exposing `Success`, `Value`, `FirstIndex`, `Length`, and `Groups`, where
`Groups` is a `String` list with the whole match at index 0 and captures after, so an
unmatched group reads as empty text rather than `Empty`. `Matches` returns a `ListOf` match
values so the LINQ surface composes over results. `Split` skips zero-length matches, matching
.NET's behavior and avoiding a piece between every character for a pattern like `a*`.

Invalid patterns are surfaced eagerly: `ConfigureRegex` runs a throwaway `Test` at creation so
a bad pattern raises the typed `ROneCOne.RegexError` (source `ROneCOne.RegexException`) where
the mistake is, not at some later call site. The dialect limits (indexed groups only, no
lookbehind, `$1` substitution only) are documented rather than papered over, because they are
the engine's, not the wrapper's.

The public `Replace` member introduced a VBA name-resolution trap: a class member named like
an intrinsic shadows every unqualified call to that intrinsic within the class, and VBA strips
the `$` type suffix during resolution, so bare `Replace$` calls elsewhere rebound to the
two-argument member and broke at runtime. Every internal intrinsic call is now `VBA.`
qualified. The static analyzer does not catch this; only the live host does.

## Consequences

`VBScript.RegExp` joins the const-based prog-id whitelist. The live suite adds a regex
contract: group extraction, `Matches`/`Replace`/`Split`, case and multiline flags, the
failed-match value shape, and the eager bad-pattern error. All of it runs offline.
