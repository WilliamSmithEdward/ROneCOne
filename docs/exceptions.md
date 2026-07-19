# Structured exceptions

New to ROneCOne? Start with the
[Events and exceptions user guide](user-guide/events-and-exceptions.md).

ROneCOne provides structured `Try`, `Catch`, and `Finally` flow over universal Actions. The
operation is an immutable builder inside `ROneCOne.cls`; executing it does not generate source,
access the VBIDE, or change Excel security settings.

## Catch and finally

```vba
Dim attempt As ROneCOne

Set attempt = ROneCOne.Try(work) _
    .Catch(errorHandler) _
    .Finally(cleanup)

attempt.Execute
```

`work` and `cleanup` are zero-argument Actions. A Catch Action may take no argument or one captured
exception. `ROneCOne.Exception` is the readable type prototype for that signature:

```vba
Set errorHandler = ROneCOne.Action("Demo.HandleError") _
    .Takes(ROneCOne.Exception)

Public Sub HandleError(ByVal errorInfo As Variant)
    Debug.Print errorInfo.ErrorNumber
    Debug.Print errorInfo.Message
End Sub
```

Captured exceptions expose `ErrorNumber`, `ErrorSource`, `Message`, `HelpFile`, and `HelpContext`.
They contain only the local VBA `Err` state and are never logged or transmitted by the runtime.

## Filtered catches

The two-argument form matches an exact VBA error number:

```vba
Set attempt = ROneCOne.Try(work) _
    .Catch(vbObjectError + 100, expectedHandler) _
    .Catch(fallbackHandler) _
    .Finally(cleanup)
```

Catches are tested in construction order. The one-argument form is catch-all. A filtered miss
preserves the pending exception, runs `Finally`, and rethrows it with its captured metadata.

## Execution semantics

- success skips all catches and runs `Finally` once;
- the first matching catch handles the pending exception;
- an error raised by a catch replaces the original pending exception;
- `Finally` runs after success, a handled error, an unhandled error, or a failed catch;
- an error raised by `Finally` takes precedence and propagates;
- the operation can be executed repeatedly and nested inside another operation.

The captured error is the error surfaced by the delegate boundary. Expression Actions preserve
ordinary VBA runtime errors directly. Object and workbook-procedure delegates retain their
documented late-bound dispatch behavior.

The independently executable
[`ROneCOne_Exceptions_Demo.xlsm`](../demo/ROneCOne_Exceptions_Demo.xlsm) demonstrates catch-all and
filtered catches, rethrow behavior, successful finalization, and a same-process benchmark.
