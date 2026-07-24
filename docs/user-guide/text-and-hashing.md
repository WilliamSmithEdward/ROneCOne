# Text, hashing, and encoding

Match and reshape text with real regular expressions, format values the same way on every
machine, build long text in linear time, compute the same digests every other platform
produces, and move between bytes and base64 or hex, all without a reference or a helper
module. `ROneCOne.Regex`, `ROneCOne.Strings`, `ROneCOne.StringBuilder`, `ROneCOne.Hash`, and
`ROneCOne.Convert` mirror their .NET namesakes.

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

## Format values predictably

`Strings.Format` speaks the `String.Format` grammar and always writes invariant text: a period
decimal separator and comma grouping, whatever the machine locale says. That makes output safe
to parse back, diff, or ship:

```vba
Debug.Print ROneCOne.Strings.Format( _
    "{0} ran {1:N1} km in {2}", "Ada", 12.375, _
    ROneCOne.TimeSpan.FromMinutes(63))        ' Ada ran 12.4 km in 01:03:00
Debug.Print ROneCOne.Strings.Format("{0:D4} | {1:X}", 42, 255)   ' 0042 | FF
Debug.Print ROneCOne.Strings.Format("{0:yyyy-MM-dd}", Date)
```

Alignment pads inside a column (`{0,10}` right, `{0,-10}` left), `{{` and `}}` write literal
braces, and anything malformed raises the typed `ROneCOne.FormatError`.

## Build long text without the slowdown

Concatenating strings in a loop is quadratic in VBA. `StringBuilder` is the linear fix, with
the fluent surface you know:

```vba
Dim report As ROneCOne

Set report = ROneCOne.StringBuilder()
For Each row In rows
    report.AppendFormat "{0,-12}{1,8:N2}", row.Item("name"), row.Item("total")
    report.AppendLine
Next row
Debug.Print report.ToString
```

## Identify and randomize

`Guid.NewGuid` mints correlation ids; `RandomNumberGenerator` draws crypto-grade bytes and
uniform integers (there is no seed; this is the cryptographic source):

```vba
Debug.Print ROneCOne.Guid.NewGuid                          ' 8-4-4-4-12, version 4
Debug.Print ROneCOne.RandomNumberGenerator.GetInt32(1, 7)  ' a fair die
Debug.Print ROneCOne.Convert.ToHexString( _
    ROneCOne.RandomNumberGenerator.GetBytes(16))           ' a 128-bit token
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
- [Formatting, building, and identity reference](../strings.md) defines the format grammar,
  builder, GUID, and random semantics exactly.
- [Hashing and encoding reference](../hashing.md) lists digests, vectors, and encoding rules.
- [HTTP and web data](http-and-web.md) covers the client that `DownloadFileAsync` extends.
- [Guide index](README.md) returns to the full learning path.
