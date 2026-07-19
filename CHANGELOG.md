# Changelog

## 0.4.0 - 2026-07-18

### Added

- `Element` as the clear, LINQ-aligned name for the typed sequence expression parameter.

### Changed

- Updated the collections demo, workbook examples, quick start, and deep documentation to lead
  with `Element` rather than the less familiar `It` terminology.
- Made `Element` the canonical implementation used by identity sorting.
- Removed the less familiar `It` sequence-parameter name under the current pre-stability API policy.

The pre-stability API intentionally prefers the clearer public surface over compatibility aliases.

## 0.3.0 - 2026-07-18

### Added

- `Var` and `VarLike` aliases plus `.AsFunc` for inferred, typed delegate parameters.
- Automatic left-to-right parameter inference when `Lambda` is called without an explicit
  parameter list.
- Sequence-typed `It` parameters and scalar object-member expressions through `value("Member")`
  or the explicit `Member` API.
- `Map`, `Exists`, `Sorted`, `SortedDescending`, `AtLeast`, and `AtMost` syntax sugar.
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

All existing `Parameter`, `ParameterLike`, explicit `Lambda`, `FromMethod`, `SelectItems`,
`AnyItem`, and selector APIs remain backward compatible.

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
