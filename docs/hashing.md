# Hashing and encoding

Exact semantics for the hashing and byte-encoding surfaces. For the workflow-first
introduction, read [Text, hashing, and encoding](user-guide/text-and-hashing.md).

## Hashing

`ROneCOne.Hash` computes digests through Windows CNG (`bcrypt.dll`), which ships on every
supported Windows and needs no registration or reference:

| Method | Digest |
|---|---|
| `Sha256(value)` | 32-byte SHA-256 |
| `Sha512(value)` | 64-byte SHA-512 |
| `Sha1(value)` | 20-byte SHA-1 |
| `Md5(value)` | 16-byte MD5 |
| `HmacSha256(key, value)` | 32-byte HMAC-SHA256 |

Every method accepts either a `String`, hashed as its UTF-8 bytes (matching what other
platforms produce for the same text), or a `Byte` array, hashed verbatim. Each returns a
`Byte` array; pair it with `Convert.ToHexString` or `ToBase64String` for text. Digests match
the published vectors (FIPS 180 for the SHA family and MD5, RFC 4231 for HMAC). SHA-1 and MD5
are provided for interoperability with legacy systems, not for security decisions.

## Encoding

`ROneCOne.Convert` moves between `Byte` arrays and text in pure VBA:

- `ToBase64String(bytes)` / `FromBase64String(text)` implement standard base64 with `=`
  padding. Decoding tolerates whitespace, like `System.Convert.FromBase64String`, but rejects
  any other non-alphabet character, misplaced padding, or a length that is not a multiple of
  four.
- `ToHexString(bytes)` produces uppercase hexadecimal; `FromHexString(text)` accepts either
  case and raises on an odd length or a non-hex digit.

Both raise `InvalidArgumentError` on malformed input.

[Back to the documentation index](README.md)
