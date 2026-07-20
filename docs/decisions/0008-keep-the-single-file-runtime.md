# ADR 0008: Keep the single-file runtime

Status: accepted, 2026-07-20

## Context

After the 1.3.0 feature batch we evaluated relaxing the one-file invariant to ship a small
companion `.bas` module alongside `ROneCOne.cls`. VBA reserves three powers for standard
modules that class modules never get: `AddressOf`, name resolution for `Application.OnTime`
and `Application.Run`, and visibility to worksheet formulas and Alt+F8. A companion module
would therefore unlock, in descending order of value:

1. A cell-formula surface: spilling UDFs that run VBA-defined delegates and query the
   in-memory data layer from worksheet cells, reaching users who never write macros.
2. A background completion pump hosting timer callbacks, reversing ADR 0007 and enabling
   fire-and-forget `OnCompleted` continuations with no `Await`.
3. `Auto_Open` and `Auto_Close` lifecycle hooks for safe timer cleanup and Function Wizard
   registration via `Application.MacroOptions`.
4. Runtime-owned `AddressOf` for native same-thread callbacks.
5. Alt+F8-visible entry points, public constants, and terse global constructors.

None of this changes the threading verdict: timer callbacks run on Excel's thread, so tasks
would remain cooperative with zero parallelism gain.

## Decision

Keep the runtime a single predeclared class. The import-one-file identity is the product
contract, and the project rejects degraded modes, so a companion module would have to be a
hard requirement with fail-fast version pairing rather than an optional extra. That makes
the change a two-file product, not a one-file product with a bonus. The strongest bundle in
its favor, the formula surface plus the pump, serves an audience the project is not
currently targeting; items 3 through 5 never justify the break on their own.

## Consequences

ADR 0007 stands: no background pump, and task chains terminate in a bounded `Await`. No
worksheet UDF surface ships; workbook authors who want one can write thin one-line wrappers
in their own standard modules, which VBA has always allowed. Revisit only if formula users
become a target audience, and treat any future companion module as a product-contract
change with mandatory pairing, not an add-on.
