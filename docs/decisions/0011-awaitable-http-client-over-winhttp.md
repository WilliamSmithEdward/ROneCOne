# ADR 0011: An awaitable HTTP client over polled WinHTTP

Status: accepted, 2026-07-23

## Context

Calling a web API from VBA normally means hand-rolling `WinHttp.WinHttpRequest.5.1` calls or
importing a third-party module, and doing it without freezing Excel means doing neither well.
ROneCOne already has a cooperative Task surface with `Await`, `WhenAll`, continuations,
cancellation, and timeouts, so the design goal was a System.Net.Http-shaped client whose
requests are ordinary awaitable tasks.

A live probe of the transport (through the same late-bound COM surface VBA uses) established
the mechanics: with `Open(..., async:=True)`, `WaitForResponse(0)` is a clean non-blocking
poll, returning False in flight and True on completion, and raising only for hard transport
failures; `Abort` resets the object; and several sent requests genuinely overlap in flight,
with three requests completing together in roughly the time of the slowest.

## Decision

Wrap WinHTTP behind `ROneCOne.HttpClient()` with the .NET verb set (`GetAsync`,
`GetStringAsync`, `GetByteArrayAsync`, `PostAsync`, `PutAsync`, `DeleteAsync`, `SendAsync`),
client state (`BaseAddress`, `Timeout`, `DefaultRequestHeader`), and an `HttpResponse` role
(`StatusCode`, `ReasonPhrase`, `IsSuccessStatusCode`, `EnsureSuccessStatusCode`, `Content`,
`Header`, `AllHeaders`). Every verb sends immediately and returns a hot task of a new
`TASK_HTTP_SEND` kind whose advance step polls `WaitForResponse(0)`; completion captures the
status, headers, and body and releases the transport. `GetStringAsync` and `GetByteArrayAsync`
fault with a typed `HttpRequestException` (`ROneCOne.HttpRequestError`) on non-success, like
their .NET namesakes; `GetAsync` always hands back the response. Cancellation aborts the
in-flight transport. Event-driven completion (`WithEvents`) was rejected because it requires a
compile-time reference, violating the no-references invariant; polling from the cooperative
scheduler needs nothing.

The honest concurrency framing is part of the decision: transfers overlap inside WinHTTP's own
worker threads while VBA remains single-threaded and collects results cooperatively. The demo
and docs present the overlap as saved network wait, never as parallel VBA.

## The Application.Run boundary

Building the demo exposed a platform boundary worth recording: a VBA error cannot propagate
upward across `Application.Run`. When a delegate bound to a standard-module procedure by name
raises, or a runtime raise (such as an awaited fault) unwinds through one, Excel surfaces its
runtime-error dialog in that context instead of reaching the caller's handler or a `Catch`.
Expression-bodied and object-method delegates propagate normally, and errors trapped inside the
named procedure itself work normally. The documented pattern for web failures is therefore
await-site handling, which also mirrors where C# places `try { await ... } catch`. The demo,
guide, and reference all teach that pattern.

## Consequences

Twenty live assertions cover the surface against https://pokeapi.co (authorized for testing and
demos): success and failure statuses, text and byte bodies, headers, `WhenAll` overlap,
cancellation, argument validation, and the typed fault path. A seventh demo workbook shows the
client end to end, including a benchmark that times the same three downloads sequentially and
overlapped. Those runs need internet access; every other gate runs offline. The prog-id
whitelist in the source contract now pins the one const-based `CreateObject` to
`WinHttp.WinHttpRequest.5.1`, and the runtime still initiates no network traffic on its own.
