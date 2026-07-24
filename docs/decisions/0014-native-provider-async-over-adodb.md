# ADR 0014: Native provider async over polled ADODB state

Status: accepted, 2026-07-23

## Context

The provider layer's task-returning verbs (`OpenAsync`, `ExecuteReaderAsync`,
`ExecuteScalarAsync`, `ExecuteNonQueryAsync`, `FillAsync`, `UpdateAsync`, `ReadAsync`) executed
their ADO calls synchronously inside one cooperative task step. That was documented honestly
(`SupportsNativeAsync = False`, `AsyncMode = "Cooperative"`), but it meant a slow query froze
Excel for its full duration even when awaited. ADO itself offers `adAsyncConnect` and
`adAsyncExecute` (both value 16) with a pollable `State` bitmask, the same shape as the WinHTTP
transport behind the HTTP client (ADR 0011).

A live probe against a local SQL Server default instance (`MSOLEDBSQL`, integrated security)
and against the ACE ISAM established the mechanics:

- `Open(..., adAsyncConnect)` returns with `State = 2` (connecting) and settles to open.
- `Command.Execute(, , adAsyncExecute)` returns immediately with the recordset in
  `State = 4` (executing); a 400 ms `WAITFOR` produced roughly forty thousand free polling
  iterations on the VBA thread before completion, and the value read correctly afterward.
- Touching fields before completion raises cleanly (3265), so results must stay gated behind
  the state poll.
- ACE accepts both async options as well, with correct results, so one code path serves both
  providers with no capability branching.
- Asynchronous non-query execution is not trustworthy: the `RecordsAffected` argument came
  back -1 immediately and the connection state never exposed a pollable executing phase, so
  there is no reliable way to both overlap and report an affected-row count.
- `Command.Cancel` on an in-flight async execution closes the recordset and leaves the
  connection usable.

## Decision

`OpenAsync` starts the connection with `adAsyncConnect` at call time; `ExecuteReaderAsync`,
`ExecuteScalarAsync`, and `FillAsync` build the ADO command and start it with `adAsyncExecute`
at call time. Each returns a hot task of a new `TASK_DB_ASYNC` kind whose advance step polls the
ADO `State` bitmask and harvests the result (the reader itself, the first field, or a filled
DataTable) only after the executing and fetching bits clear. A closed state before completion is
a failure; the provider's `Errors` collection supplies the raised number, source, and
description. Cancellation calls `Cancel` on the pending command or connection, mirroring the
HTTP abort path.

`ExecuteNonQueryAsync`, `UpdateAsync`, and `ReadAsync` keep the cooperative single-step
execution, because affected-row fidelity and row-by-row update semantics have no usable native
async signal. `AsyncMode` now reports `"Native"` and its documentation names those three
exceptions; `SupportsNativeAsync` is removed rather than left behind as a half-true Boolean.

## Consequences

The live suite gains a SQL Server contract against the local default instance over
`MSOLEDBSQL` with `Integrated Security=SSPI`, so no credential appears in the repository:
parameterized multi-row inserts, typed reader and fill results, commit and rollback,
cancellation mid-`WAITFOR`, async failure surfacing for bad statements and bad logins, and a
timing proof that a 300 ms server-side wait overlaps 250 ms of busy local VBA work (total under
a 450 ms ceiling, where serialized execution needs at least 550 ms). The suite therefore
requires a reachable localhost SQL Server in addition to Excel, ACE, and pokeapi.co network
access; development.md records the prerequisite. The ACE-based provider tests and the Data
demo run unchanged on the same native path, and the demo's capability rows now show
`AsyncMode = "Native"` and the live connection state.
