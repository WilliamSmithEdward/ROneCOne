# Regular expressions

Exact semantics for the regular expression surface. For the workflow-first introduction, read
[Text, hashing, and encoding](user-guide/text-and-hashing.md).

## Surface

`ROneCOne.Regex(pattern, [ignoreCase], [multiLine])` compiles a pattern and returns a regex
value with `Pattern`, `IsMatch`, `Match`, `Matches`, `Replace`, and `Split`. `Match` returns a
match value with `Success`, `Value`, `FirstIndex` (zero-based), `Length`, and `Groups` (a
`String` list: element 0 is the whole match, then each capture group in order). An invalid
pattern raises error number `ROneCOne.RegexError` from source `ROneCOne.RegexException` at
creation, not at first use.

## Engine and dialect

The engine is the in-box script runtime's expression object (`VBScript.RegExp`), so the
dialect is that engine's, not .NET's:

- Indexed capture groups only; no named groups, lookbehind, atomic groups, or possessive
  quantifiers. Lookahead (`(?=...)`, `(?!...)`) and non-capturing groups (`(?:...)`) work.
- `Replace` expands `$1` through `$9` and `$&`; there is no `${name}` form.
- An unmatched optional group reports as empty text in `Groups`.
- Case folding is opt-in through `ignoreCase`; `multiLine` makes `^` and `$` match at line
  breaks. Both default to off.

## Method semantics

| Method | Behavior |
|---|---|
| `IsMatch(input)` | True if the pattern matches anywhere |
| `Match(input)` | The first match, or a value whose `Success` is False |
| `Matches(input)` | Every match as a `ListOf` regex-match values, ready for LINQ |
| `Replace(input, replacement)` | Replaces every match, expanding group references |
| `Split(input)` | Splits at each non-empty match; zero-length matches never split |

`Split` deliberately ignores zero-length matches so a pattern such as `a*` never manufactures
empty pieces between characters. Consecutive delimiters still yield the empty fields between
them, matching `System.Text.RegularExpressions.Regex.Split`.

[Back to the documentation index](README.md)
