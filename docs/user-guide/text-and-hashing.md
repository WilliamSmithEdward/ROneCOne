# Text, hashing, and encoding

Match and reshape text with real regular expressions, compute the same digests every other
platform produces, and move between bytes and base64 or hex, all without a reference or a
helper module. `ROneCOne.Regex`, `ROneCOne.Hash`, and `ROneCOne.Convert` mirror
`System.Text.RegularExpressions`, `System.Security.Cryptography`, and `System.Convert`.

## Match and extract with regex

```vba
Dim email As ROneCOne
Dim hit As ROneCOne

Set email = ROneCOne.Regex("(\w+)@(\w+)\.(\w+)")
If email.IsMatch(cell) Then
    Set hit = email.Match(cell)
    Debug.Print hit.Value            ' ada@example.com
    Debug.Print hit.Groups.Item(1)   ' ada
End If
```

`Matches` returns an ordinary typed list, so the whole LINQ surface applies:

```vba
Dim domains As ROneCOne

Set domains = ROneCOne.Regex("@(\w+)") _
    .Matches(allText)
Debug.Print domains.Count
```

`Replace` expands `$1`-style group references, and `Split` breaks on a pattern without
inventing empty pieces at zero-length matches:

```vba
Debug.Print ROneCOne.Regex("\s*,\s*").Split("a, b ,c").Count   ' 3
```

## Hash and sign

```vba
Dim digest As String

digest = ROneCOne.Convert.ToHexString(ROneCOne.Hash.Sha256("hello"))
```

Text hashes as its UTF-8 bytes, so the result matches what Python, C#, or a shell `sha256sum`
would print. Pass a `Byte` array to hash binary content, such as a file read with
`ROneCOne.File.ReadAllBytes`. `HmacSha256(key, message)` signs a request payload:

```vba
Dim signature As String

signature = ROneCOne.Convert.ToBase64String( _
    ROneCOne.Hash.HmacSha256(secretKey, requestBody))
client.DefaultRequestHeader "X-Signature", signature
```

## Encode bytes

`Convert.ToBase64String` / `FromBase64String` and `ToHexString` / `FromHexString` round-trip
`Byte` arrays to text and back. These pair naturally with the hash and file surfaces, and with
`HttpClient.DownloadFileAsync`, which saves a response body straight to disk:

```vba
ROneCOne.HttpClient().DownloadFileAsync( _
    "https://example.com/data.bin", "C:\data\data.bin").Await
```

## Where next

- [Regex technical reference](../regex.md) defines the dialect and every method exactly.
- [Hashing and encoding reference](../hashing.md) lists digests, vectors, and encoding rules.
- [HTTP and web data](http-and-web.md) covers the client that `DownloadFileAsync` extends.
- [Guide index](README.md) returns to the full learning path.
