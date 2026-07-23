# HTTP and web data

Call a web API from VBA the way C# calls one from System.Net.Http: create a client, start a
request, await the result. No references, no add-ins, no JSON library to install. The transfer
runs inside Windows while your workbook stays responsive, and you collect the answer with
`Await`.

## Download one resource

```vba
Dim client As ROneCOne
Dim response As ROneCOne

Set client = ROneCOne.HttpClient()
client.BaseAddress = "https://pokeapi.co/api/v2/"
client.Timeout = 20000

Set response = client.GetAsync("pokemon/pikachu").Await
Debug.Print response.StatusCode      ' 200
Debug.Print response.ReasonPhrase    ' OK
Debug.Print response.Header("Content-Type")
Debug.Print Left$(response.Content, 80)
```

`GetAsync` starts the request and returns a task immediately; `Await` hands back the response
when it arrives. A relative URL rides on `BaseAddress`, so switching servers is one assignment.

## Just the text

```vba
Dim json As String

json = client.GetStringAsync("pokemon/ditto").Await
```

`GetStringAsync` resolves straight to the body and refuses to return text from a failed
request: a non-2xx status raises the typed `HttpRequestException` instead.

## Overlap several downloads

```vba
Dim replies As ROneCOne

Set replies = ROneCOne.Task.WhenAll( _
    client.GetAsync("pokemon/bulbasaur"), _
    client.GetAsync("pokemon/charmander"), _
    client.GetAsync("pokemon/squirtle")).Await
Debug.Print replies.Item(0).StatusCode
```

Each `GetAsync` starts its transfer at once, so all three are in flight together and the total
wait is roughly the slowest single download. The overlap happens inside WinHTTP; your VBA still
runs on one thread.

## Handle failure like try / await / catch

```vba
Dim task As ROneCOne

Set task = client.GetStringAsync("pokemon/missingno")
On Error Resume Next
task.Await
If Err.Number = ROneCOne.HttpRequestError Then
    Debug.Print "fell back to cached data"
End If
On Error GoTo 0
```

Handle the error at the await site, exactly where a C# `try { await ... } catch` would sit. The
task also keeps the full story in `IsFaulted` and `Exception`, mirroring `Task.Exception` in
.NET. One boundary to know about: an error that unwinds out of a procedure the runtime invokes
by name (through `Application.Run`) cannot reach a `Catch`; Excel shows its own error dialog
instead. Keep await-and-recover logic at the call site, as above.

## Send, post, and cancel

```vba
Dim source As ROneCOne
Dim task As ROneCOne

Set task = client.PostAsync("collector/notes", "{""note"":1}", "application/json")

Set source = ROneCOne.CancellationTokenSource
Set task = client.GetAsync("pokemon/eevee", source.Token)
source.Cancel          ' aborts the transfer; the task reports IsCanceled
```

`PostAsync`, `PutAsync`, `PatchAsync`, and `DeleteAsync` cover the named verbs;
`SendAsync("REPORT", url)` sends any method a server understands, standard or custom. Every
verb accepts the same cancellation tokens as the rest of the task surface, and
`GetByteArrayAsync` downloads binary content.

The client only ever contacts URLs you pass it. The runtime itself never transmits anything.

[Back to the guide index](README.md) | [Exact HTTP semantics](../http.md)
