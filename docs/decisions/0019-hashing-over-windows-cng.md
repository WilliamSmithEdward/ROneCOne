# ADR 0019: Hashing over Windows CNG

Status: accepted, 2026-07-23

## Context

Hashing rounds out the transport and exchange story: checksums for the file layer (the release
ceremony itself publishes SHA-256s) and HMAC request signing composed with the HTTP client. VBA
offers nothing built in. The obvious COM route, the .NET Framework crypto classes
(`System.Security.Cryptography.SHA256Managed` and kin), turned out not to be creatable: a live
probe got automation error 0x80131700 for every one of them, because they are not registered
as COM-visible on this machine, and depending on the .NET Framework being COM-registered is
exactly the fragile assumption the runtime avoids.

The probe's second candidate succeeded cleanly: the Windows CNG primitives in `bcrypt.dll`,
called through `Declare`. `BCryptOpenAlgorithmProvider` / `CreateHash` / `HashData` /
`FinishHash` produced the exact FIPS 180 vector for SHA-256 of "abc", the correct SHA-512,
SHA-1, and MD5 digests, and the RFC 4231 HMAC-SHA256 vector using the `BCRYPT_ALG_HANDLE_HMAC`
flag with the key passed to `CreateHash`. A 1 MB hash was instant. `bcrypt.dll` ships on every
supported Windows and needs no registration.

## Decision

`ROneCOne.Hash` computes `Sha256`, `Sha512`, `Sha1`, `Md5`, and `HmacSha256` through CNG
`Declare`s, joining the existing `DispCallFunc`/`CopyMemory`/`Sleep` declarations the class
already carries. Each accepts a `String`, hashed as its UTF-8 bytes so results match every
other platform, or a `Byte` array hashed verbatim, and returns a `Byte` array. A CNG status
failure raises through the standard contract-error path. `ROneCOne.Convert` provides the
text pairing in pure VBA: `ToBase64String` / `FromBase64String` with standard padding and
whitespace tolerance, and `ToHexString` / `FromHexString`. Reusing CNG rather than the .NET
COM classes keeps the surface working on a machine with no COM-registered framework, which is
the common case.

`DownloadFileAsync` lands in the same change: it composes the existing `GetByteArrayAsync`
task with `WriteAllBytes`, adding a `File` result mode to the HTTP task that writes the body to
a path once the transfer completes and faults on non-success like the other body-consuming
verbs.

## Consequences

No new prog-id is added; the CNG functions are `Declare`s against a system DLL. The live suite
adds a hashing contract asserting the published SHA and HMAC vectors, byte-versus-text
equivalence, and the base64 and hex round trips with their malformed-input rejections, plus a
`DownloadFileAsync` assertion against the authorized pokeapi host. The base64 decoder carried a
32-bit accumulator overflow on long inputs, caught and fixed during implementation; an empty
string passed as VBA's null BSTR to `ADODB.Stream.WriteText` raised error 3001, so the file
layer's text encoder now short-circuits empty content.
