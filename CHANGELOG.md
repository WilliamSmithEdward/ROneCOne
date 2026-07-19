# Changelog

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
