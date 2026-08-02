# 0030. A first-class Excel Table surface

Date: 2026-08-02

## Status

Accepted.

## Context

The worksheet bridge took a `Range` and nothing else. An Excel Table is a
`ListObject`, which is not a `Range` and has no `Value`, so handing one to
`DataTableFromRange` raised run-time error 438 from the first line of
`ReadRangeGrid`. The runtime had never heard of the type: a search for
`ListObject` over the source returned zero hits, and the documentation never
mentioned it either.

The failure reads as the runtime rejecting the table rather than VBA rejecting
the type, and 438 gives no hint about what to pass instead. Tables are how most
people keep tabular data in a workbook, so this was the first wall a new user
hit. The advice, when it existed at all, was to pass `listObject.Range`
yourself, which works but makes the table a second-class input and leaves the
user to know that a totals row is included in `.Range` and must be trimmed.

Writing back was worse. `ToRange` writes a grid at a target cell, so pointing
it at a table's range overwrites cells without changing the table's extent. The
table then disagrees with its own contents, and any rows the new data did not
cover stay behind inside it.

## Decision

Recognize the type at the boundary and give it a surface.

Every worksheet entry point resolves its argument through one helper.
A `ListObject` or a `ListColumn` becomes a `Range`; anything else passes
through untouched, so the cost is one `TypeName` test. With headers wanted, the
slice is `source.Range.Resize(1 + owner.ListRows.Count)`, which stops after the
last body row because `.Range` spans the totals row and a total is not data. An
empty table still yields its header row, which reads back as columns and no
rows rather than as a failure.

`ROneCOne.Table(listObject)` returns a DataTable that keeps a reference to the
Table it came from. Everything a DataTable can already do keeps working, and
two members are added: `Refresh` re-reads in place, and `WriteBack` writes rows
into the Table and resizes it to fit. `ToRange` given a Table as its target
performs the same write, so a DataView can drive it directly.

`Refresh` is a `Sub` that mutates in place rather than a `Function` returning a
fresh table, because a bare `Refresh` statement on a Function would compile and
silently do nothing.

## Consequences

Excel's resize behavior is awkward in four specific ways, all measured live
before any of this was written rather than inferred from documentation:

- Resizing to a header row alone raises 1004, because a table must keep at
  least one data row. Writing zero rows therefore goes through
  `DataBodyRange.Delete`, which leaves the table and its headers in place and
  reports `ListRows.Count` as 0 with `DataBodyRange` as `Nothing`.
- Shrinking leaves the vacated cells populated below the table, so the old tail
  is cleared explicitly after the resize.
- A totals row lands inside the requested range when a table grows and outside
  it when one shrinks. Rather than encode that asymmetry, totals are switched
  off around the resize and restored afterwards.
- Name, style, and header text all survive a resize, so none of them need
  saving and restoring.

The column count must match, and a mismatch raises rather than truncating.

One behavior changed for existing callers: `ToRange` given a `ListObject` used
to raise 438 and now resizes and fills the table. Nothing could have depended
on the old behavior, since it never succeeded.

The write-back path shares `InternalLiveDataRows` with the JSON writer, which
is `Friend` rather than `Private` because both ask another instance for its
rows. That is the shape that reintroduced the defect from
[issue #3](https://github.com/WilliamSmithEdward/ROneCOne/issues/3) during
implementation: `Collection.Item` returns a `Variant`, so reading a `Friend`
member straight off it binds late and raises 438 at run time with no `Object`
local anywhere in the code. The source contract that pins `Friend` reads was
extended to reject any `.Item(...).Member` chain, not only late-bound
declarations.

## Alternatives considered

**Leave it, and document `pass .Range`.** This is what the previous release
did, by writing a guide around the workaround. It teaches the user to carry a
detail the runtime is better placed to handle, and it does not solve write-back
at all.

**A separate role and class-like surface for tables.** A new role would have
meant reimplementing or forwarding every DataTable member for the attached
case. Keeping the DataTable role and adding one field gets the entire existing
surface for free, and the two new members are the only things that need the
origin.

**Detect the type inside `ReadRangeGrid`.** That is one function rather than
four call sites, but `ReadRangeGrid` does not know whether headers were asked
for, and the header question decides between `.Range` and `.DataBodyRange`.
Resolving at each entry point keeps the decision where the answer is known.
