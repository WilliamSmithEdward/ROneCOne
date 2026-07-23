# HTTP client

Exact semantics for the awaitable HTTP surface. For the workflow-first introduction, read
[HTTP and web data](user-guide/http-and-web.md).

## Transport and lifecycle

The client wraps `WinHttp.WinHttpRequest.5.1`, created late-bound in process, so no reference,
add-in, or install is required. Every verb opens the request asynchronously, applies the
client's timeouts and default headers, sends immediately, and returns a hot task. The transfer
proceeds on WinHTTP's own worker threads; the VBA side completes the task cooperatively by
polling the transport during `Await` (or any scheduler pass), so several requests overlap in
flight while VBA itself stays single-threaded. ROneCOne makes no VBA parallelism claim.

Cancellation aborts the underlying transport and the task reports `IsCanceled`. Hard transport
failures (name resolution, connection loss, TLS) fault the task with the WinHTTP message under
error number `ROneCOne.HttpRequestError` and source `ROneCOne.HttpRequestException`.

## Client surface

| Member | Semantics |
|---|---|
| `ROneCOne.HttpClient()` | Creates a client with a 30,000 ms timeout and no default headers |
| `BaseAddress` | Prefixed to relative URLs; a relative URL without it raises an invalid-argument error. Absolute URLs (containing `://`) ignore it |
| `Timeout` | Milliseconds applied to resolve, connect, send, and receive; must be positive |
| `DefaultRequestHeader(name, value)` | Adds a header sent with every request; returns the client for chaining |
| `GetAsync(url, [token])` | Task of `HttpResponse`; never faults on HTTP status |
| `GetStringAsync(url, [token])` | Task of `String`; faults with `HttpRequestException` on any non-2xx status |
| `GetByteArrayAsync(url, [token])` | Task of a byte array; faults on non-2xx like `GetStringAsync` |
| `PostAsync(url, body, [contentType], [token])` | Task of `HttpResponse`; sets `Content-Type` only when given |
| `PutAsync(url, body, [contentType], [token])` | Task of `HttpResponse` |
| `PatchAsync(url, body, [contentType], [token])` | Task of `HttpResponse` |
| `DeleteAsync(url, [token])` | Task of `HttpResponse` |
| `SendAsync(method, url, [body], [contentType], [token])` | Any method, standard or custom, uppercased; body sent only when supplied. WinHTTP imposes no verb blocklist, so OPTIONS, HEAD, and site-specific verbs all transmit |

## Response surface

| Member | Semantics |
|---|---|
| `StatusCode` | The numeric HTTP status |
| `ReasonPhrase` | The status text |
| `IsSuccessStatusCode` | True for 200 through 299 |
| `EnsureSuccessStatusCode` | Returns the response, or raises `HttpRequestException` |
| `Content` | The body as text |
| `Header(name)` | One response header, case-insensitive, or an empty string |
| `AllHeaders` | The raw header block |

Headers and body are captured when the transfer completes; the response holds no live transport
afterwards.

## Failure semantics and the Application.Run boundary

`Await` on a faulted request raises the inner error (number `ROneCOne.HttpRequestError`), and
`task.Exception` returns the .NET-aligned `AggregateException` wrapper whose `InnerExceptions`
holds it. Handle web failures at the await site, the same place C# puts `try { await ... }
catch`.

An error cannot cross an `Application.Run` boundary upward. When a delegate bound to a
standard-module procedure by name raises (including a runtime raise unwinding through it, such
as an awaited fault), Excel surfaces its runtime-error dialog inside that context instead of
propagating to a caller's handler or a `Catch`. This is VBA platform behavior, not a runtime
choice: expression-bodied and object-method delegates propagate normally, and raises trapped
inside the named procedure itself work normally. Recover from web failures at the await site
rather than letting them unwind out of a by-name procedure.

## Privacy

The client contacts only the URLs you request. The runtime never initiates network traffic on
its own, and nothing about your workbook is transmitted beyond the requests you write.

[Back to the documentation index](README.md)
