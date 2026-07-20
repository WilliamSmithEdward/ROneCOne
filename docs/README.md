# ROneCOne documentation

Everything here describes one product: a single importable class module,
[`src/ROneCOne.cls`](../src/ROneCOne.cls), that gives Excel VBA generic collections, LINQ,
delegates, tasks, typed events, structured exceptions, and an in-memory data layer.

Choose the path that matches what you are trying to do.

## Learn and build

Start here if you write workbook code and want results quickly. The
[user guide](user-guide/README.md) is short, indexed, and code-led:

| Guide | What you will accomplish |
|---|---|
| [Getting started](user-guide/getting-started.md) | Run a demo and add ROneCOne to a workbook |
| [Collections and LINQ](user-guide/collections-and-linq.md) | Query and summarize data |
| [Delegates and expressions](user-guide/delegates-and-expressions.md) | Build reusable behavior |
| [Events and exceptions](user-guide/events-and-exceptions.md) | Coordinate changes and failures |
| [Tasks and async](user-guide/tasks-and-async.md) | Coordinate work, cancellation, and progress |
| [Data and providers](user-guide/data-and-providers.md) | Model, query, and load tabular data |
| [Practical reference](user-guide/reference.md) | Find names, defaults, and limits fast |

## Exact contracts

Read these when you need precise semantics, edge cases, or the canonical form behind the concise
syntax. Each technical document links back to its user guide.

| Reference | Covers |
|---|---|
| [Collections and LINQ](collections.md) | Typing rules, deferred execution, comparers, ordering, canonical expansions |
| [Delegates](delegates.md) | Call targets, signatures, multicast, native invocation, true ByRef |
| [Events](events.md) | Subscription, delivery order, snapshot and failure semantics |
| [Exceptions](exceptions.md) | Catch ordering, rethrow, cleanup precedence, captured metadata |
| [Tasks](tasks.md) | Execution modes, coordination, cancellation, memory and thread safety |
| [Data and providers](data.md) | Tables, relations, views, provider capabilities, deterministic cleanup |

## Understand the design

- [Architecture](architecture.md) explains the one-file runtime invariant, each capability slice,
  and the state and isolation rules.
- [Development and Excel safety](development.md) documents the four validation gates, the
  popup-adaptive Excel harness, benchmarks, and how the demo workbooks are rebuilt.
- [Decision records](decisions/) capture significant architecture decisions with their context
  and consequences.

## Release history

- [Changelog](../CHANGELOG.md) records every released change.
- [Release notes](releases/) archive the published notes and SHA-256 checksums for each version.

[Back to the project overview](../README.md)
