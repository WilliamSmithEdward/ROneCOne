# Changelog

## 0.9.0 - 2026-07-19

### Added

- The complete .NET-aligned ordering family: `Order`, `OrderDescending`, `OrderBy`,
  `OrderByDescending`, `ThenBy`, and `ThenByDescending`.
- Independent comparers for every ordering level, including direct ordering of user-defined class
  instances through an explicit comparer.
- Live contracts for stable composite keys, reset and continuation rules, deferred execution,
  selector-once evaluation, Null, Boolean, ordinal strings, and rejected mixed Variant coercion.
- A 10,000-element composite-ordering benchmark with a 1.0-second isolated-suite gate and a
  2.5-second living-workbook gate sized from fresh Excel-process measurements.

### Changed

- Replaced quadratic insertion sorting with a stable O(n log n) bottom-up merge sort that caches
  every selected key exactly once per element and enumeration.
- Removed the earlier identity-order aliases in favor of the .NET names `Order` and
  `OrderDescending`; the current API, demos, tests, and documentation move together.
- Made `ThenBy` legal only on an immediately active ordered query and made a new primary ordering
  operation replace the prior ordering chain.
- Expanded the live collection suite from 122 to 146 assertions and upgraded the existing
  Collections workbook with composite ordering and a measured ordering release gate.

## 0.8.0 - 2026-07-19

### Added

- Composable collection-membership expressions through `IsIn`, `NotIn`, `ContainsMember`, and
  expression arguments to `List<T>.Contains`; `OneOf` now accepts arrays and collections.
- Null-safe `?.` member paths with deterministic Null propagation through the remaining chain.
- `Both`, `Either`, `Negated`, and `WhereNot` predicate composition sugar plus case-insensitive
  equality, prefix, suffix, containment, and pattern helpers.
- Predicate-aware `Count`, `FirstOrDefault`, `LastOrDefault`, `SingleItem`, `SingleOrDefault`, and
  `None` terminals.
- Nested collection quantifiers through `AnyMatch`, `AllMatch`, `NoneMatch`, `WhereAny`,
  `WhereAll`, and `WhereNone`.
- `EqualityComparer` and `Comparer` factories with comparer-aware containment, distinctness,
  sorting, extreme-value operators, and `SequenceEqual`.
- Reusable `Always`, `Never`, `Match`, and `NotMatch` predicate factories.
- A repeated object-member benchmark alongside the scalar query benchmark.

### Changed

- Cached normalized member names directly on expression nodes, removing one captured-value node
  and one recursive evaluation per member access.
- Expanded the live collection suite from 91 to 122 assertions and the Collections workbook from
  nineteen to twenty-five formula-verified examples.
- Replaced explicit nullable-object guard chains in the primary demo with `?.` paths.
- Recorded the Excel-host grammar boundary that forbids `In` and `Single` as procedure names;
  the legal public forms are `IsIn` and `SingleItem`.

## 0.7.0 - 2026-07-19

### Added

- Deferred contextual filters through `Where("Member").AtLeast(...)` and the complete comparison,
  set, string, Boolean, null, and pattern condition vocabulary.
- Reusable `Condition(memberPath)` expressions backed by one stable typed element parameter per
  sequence, including composed multi-member predicates and dotted object paths.
- Native VBA bang-member expressions such as `customers!Age` and `With customers: !Age` through
  the existing default-member dispatch surface.
- Member-name selectors for projection, ordering, aggregation, and text joining.
- `DistinctBy`, `MinBy`, and `MaxBy` key-based LINQ operators.
- `Predicate` and `WhereMethod` inference of `Func<T, Boolean>` from a sequence's generic type.
- Eleven live user-class LINQ scenarios in the Collections workbook.

### Changed

- Made `Contains` serve both typed sequence containment and C#-style string expression containment.
- Updated the primary demo surface to remove explicit element variables wherever member context is
  sufficient, while retaining and documenting the canonical expression/delegate forms.
- Expanded the live collection suite from 66 to 91 assertions.
- Confirmed through the popup-aware VBE harness that both direct and `With`-scoped bang syntax
  compile and execute in Microsoft 365 x64 Excel.

## 0.6.0 - 2026-07-19

### Added

- `Execute` as a statement-form Action invocation surface, including inline native reference
  wrappers without dummy result variables.
- Typed mutable events through `EventOf`, fluent `Subscribe`, last-match `Unsubscribe`, snapshot
  `Emit`, and `HandlerCount`.
- Structured immutable `Try`, catch-all and error-number `Catch`, `Finally`, captured exception
  metadata, deterministic rethrow, and cleanup precedence.
- Explicit populated `ListOf(T, items...)`, inferred `ListFrom(first, rest...)`, and live-tested
  empty `ListLike(example)` construction.
- Atomic `AddRange` inputs from typed ROneCOne sequences, VBA arrays, and `Collection` values.
- `ForEach(Action)`, predicate-optional `Exists`, and `JoinText` terminals.
- Separate Events and Exceptions living workbooks with source, packaging, execution, rendering,
  and same-process benchmark contracts.

### Changed

- Admitted every universal unary `Func` role into LINQ normalization, including workbook
  procedures, native functions, compositions, and multicast delegates.
- Reworked delegate and collection demos to use one-expression signatures, inline `ByRef`, Action
  `Execute`, collection initializers, inferred user-class lists, `ForEach`, and `JoinText`.
- Expanded live totals to 58 delegate/event/exception assertions and 66 collection assertions.
- Recorded and eliminated VBA's host-only invalid-`ParamArray` forwarding failure; the popup-aware
  harness closed the compiler modal and terminated only its task-owned Excel process.
- Made every bounded Excel launcher tolerate duplicate `Path`/`PATH` host environments and removed
  machine-specific Node paths from workbook rendering.

## 0.5.0 - 2026-07-18

### Added

- Universal `Func` and `Action` factories for expression trees, object methods, callable objects,
  and workbook procedures.
- Immutable `Takes` / `Returns` signatures, C#-style `DynamicInvoke`, target/method/signature
  metadata, and exact user-class type descriptors.
- Immutable multicast `Combine`, `Remove`, `GetInvocationList`, deterministic invocation order,
  and last-result semantics.
- Fail-closed Windows x64 `Native` / `NativeAction` dispatch through `DispCallFunc`.
- True native `ByRef` with `RefOf` and typed wrappers for VBA numeric, Boolean, Date, and Currency
  variables.
- Live-host contracts for callable objects, workbook procedures, object identity returns,
  multicast removal, native function pointers, `ByRef` mutation, and rejected incomplete calls.
- A complete universal-delegate guide and an eleven-scenario living delegate workbook.

### Changed

- Replaced the narrow `FromMethod` adapter with the single `Func` / `Action` construction model.
- Routed LINQ method delegates through universal typed factories.
- Made non-native `ByRef` fail before dispatch because `CallByName` and `Application.Run` cannot
  preserve the original variable identity.
- Made invalid `Action`-to-`Func` and `Action.Returns(...)` conversions fail during construction.
- Expanded the Excel harness to report exact delegate assertion failures as well as modal VBE
  compiler selections.

Demos, tests, and documentation use the preferred universal model without compatibility shims.

## 0.4.0 - 2026-07-18

### Added

- `Element` as the clear, LINQ-aligned name for the typed sequence expression parameter.

### Changed

- Updated the collections demo, workbook examples, quick start, and deep documentation to lead
  with `Element` rather than the less familiar `It` terminology.
- Made `Element` the canonical implementation used by identity sorting.
- Removed the less familiar `It` sequence-parameter name in favor of `Element`.

The API intentionally prefers the clearer public surface over compatibility aliases.

## 0.3.0 - 2026-07-18

### Added

- `Var` and `VarLike` aliases plus `.AsFunc` for inferred, typed delegate parameters.
- Automatic left-to-right parameter inference when `Lambda` is called without an explicit
  parameter list.
- Sequence-typed `It` parameters and scalar object-member expressions through `value("Member")`
  or the explicit `Member` API.
- `Map`, `Exists`, identity ordering, `AtLeast`, and `AtMost` syntax sugar.
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

Canonical `Parameter`, `ParameterLike`, explicit `Lambda`, `SelectItems`, `AnyItem`, and selector
APIs remain available for explicit construction and debugging.

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
